param(
  [Parameter(Mandatory=$true)][string]$InputReleaseRoot,
  [Parameter(Mandatory=$true)][string]$OutputRoot,
  [string]$ProductVersion=''
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest

function Assert-NewOutput([string]$output,[string[]]$inputs){
 $full=[IO.Path]::GetFullPath($output).TrimEnd('\')
 if(Test-Path -LiteralPath $full){throw 'OUTPUT_EXISTS_REFUSE_REPLACEMENT'}
 foreach($input in $inputs){
  $source=[IO.Path]::GetFullPath($input).TrimEnd('\')
  if($full.Equals($source,[StringComparison]::OrdinalIgnoreCase)-or$source.StartsWith($full+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'OUTPUT_OVERLAPS_INPUT'}
 }
 $cursor=$full
 while($cursor){
  if(Test-Path -LiteralPath $cursor){$item=Get-Item -LiteralPath $cursor -Force;if(($item.Attributes-band[IO.FileAttributes]::ReparsePoint)-ne0){throw 'OUTPUT_REPARSE_PATH_REJECTED'}}
  $next=Split-Path -Parent $cursor;if($next-eq$cursor){break};$cursor=$next
 }
}
Assert-NewOutput $OutputRoot @($InputReleaseRoot)
$buildResultPath=Join-Path $InputReleaseRoot 'build-result.json'
if(Test-Path -LiteralPath $buildResultPath -PathType Leaf){
  $buildResult=Get-Content -LiteralPath $buildResultPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $builtVersion=[string]$buildResult.ProductVersion
  if([string]::IsNullOrWhiteSpace($builtVersion)){throw 'BUILD_RESULT_PRODUCT_VERSION_MISSING'}
  if([string]::IsNullOrWhiteSpace($ProductVersion)){$ProductVersion=$builtVersion}
  elseif($ProductVersion-ne$builtVersion){throw ('MANAGED_SETUP_VERSION_MISMATCH expected='+$builtVersion+' actual='+$ProductVersion)}
}
elseif([string]::IsNullOrWhiteSpace($ProductVersion)){
  throw 'PRODUCT_VERSION_REQUIRED_WHEN_BUILD_RESULT_MISSING'
}


function Sha([string]$p){(Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToUpperInvariant()}
function Assert-File([string]$p,[string]$label){if(!(Test-Path -LiteralPath $p -PathType Leaf)){throw ($label+'_MISSING: '+$p)}}

$releaseZip=Join-Path $InputReleaseRoot ('WordAgent-'+$ProductVersion+'-win-x64.zip')
$installScript=Join-Path $InputReleaseRoot 'setup-payload\Install-WordAgent.ps1'
Assert-File $releaseZip 'RELEASE_ZIP'
Assert-File $installScript 'INSTALL_SCRIPT'

$candidates=@(
  (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
  (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
)
$csc=@($candidates|Where-Object{Test-Path -LiteralPath $_ -PathType Leaf}|Select-Object -First 1)
if($csc.Count-ne1){throw 'CSC_TOOLCHAIN_UNAVAILABLE__NO_TOOL_INSTALL_ALLOWED'}
$cscPath=[string]$csc[0]

Assert-NewOutput $OutputRoot @($InputReleaseRoot)
New-Item -ItemType Directory -Path $OutputRoot|Out-Null

$setupExe=Join-Path $OutputRoot ('WordAgent-Setup-'+$ProductVersion+'-win-x64.exe')
$outZip=Join-Path $OutputRoot ('WordAgent-'+$ProductVersion+'-win-x64.zip')
Copy-Item -LiteralPath $releaseZip -Destination $outZip -Force

$source=Join-Path $OutputRoot 'WordAgent.ManagedBootstrapper.cs'
$src=@'
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Win32;

[assembly: AssemblyTitle("WordAgent Setup")]
[assembly: AssemblyProduct("WordAgent")]
[assembly: AssemblyCompany("WordAgent")]
[assembly: AssemblyVersion("2.0.0.0")]
[assembly: AssemblyFileVersion("2.0.0.0")]

internal static class WordAgentManagedSetup
{
    private const string ScriptResource = "WordAgent.InstallScript";
    private const string ZipResource = "WordAgent.ReleaseZip";

    private static string Quote(string value)
    {
        if (value == null) return "\"\"";
        var result = new System.Text.StringBuilder("\"");
        int slashes = 0;
        foreach (char c in value)
        {
            if (c == '\\') { slashes++; continue; }
            if (c == '"') result.Append('\\', slashes * 2 + 1);
            else result.Append('\\', slashes);
            result.Append(c);
            slashes = 0;
        }
        result.Append('\\', slashes * 2);
        return result.Append('"').ToString();
    }

    private static void ExtractResource(string name, string destination)
    {
        Assembly asm = Assembly.GetExecutingAssembly();
        using (Stream src = asm.GetManifestResourceStream(name))
        {
            if (src == null) throw new InvalidOperationException("RESOURCE_MISSING_" + name);
            using (FileStream dst = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                src.CopyTo(dst);
            }
        }
    }

    private static int RunChild(string fileName, string arguments, string workingDirectory)
    {
        ProcessStartInfo psi = new ProcessStartInfo();
        psi.FileName = fileName;
        psi.Arguments = arguments;
        psi.WorkingDirectory = workingDirectory;
        psi.UseShellExecute = false;
        psi.CreateNoWindow = true;
        psi.RedirectStandardOutput = true;
        psi.RedirectStandardError = true;
        using (Process p = new Process())
        {
            p.StartInfo = psi;
            if (!p.Start()) return 253;
            var stdoutTask = p.StandardOutput.ReadToEndAsync();
            var stderrTask = p.StandardError.ReadToEndAsync();
            if (!p.WaitForExit(900000))
            {
                try { p.Kill(); } catch { }
                return 254;
            }
            stdoutTask.GetAwaiter().GetResult();
            stderrTask.GetAwaiter().GetResult();
            return p.ExitCode;
        }
    }


    private static void CleanupLegacyUninstallRegistration()
    {
        const string legacyPath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\PersianWordAgent";
        using (RegistryKey currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default))
        {
            using (RegistryKey existing = currentUser.OpenSubKey(legacyPath, false))
            {
                if (existing == null) return;
            }
            currentUser.DeleteSubKeyTree(legacyPath, false);
        }
    }

    private static int ParseInt(string value)
    {
        int n;
        if (!Int32.TryParse(value, out n)) throw new ArgumentException("INVALID_INTEGER");
        return n;
    }

    public static int Main(string[] args)
    {
        try
        {
            string installRoot = null;
            bool noRegister = false;
            bool simulateMissingRelease = false;

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (String.Equals(a, "--diagnostic-child-exit", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    int wanted = ParseInt(args[++i]);
                    return RunChild(Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe",
                        "/d /c exit " + wanted.ToString(), Environment.CurrentDirectory);
                }
                else if (String.Equals(a, "--install-root", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    installRoot = args[++i];
                }
                else if (String.Equals(a, "--no-register", StringComparison.OrdinalIgnoreCase))
                {
                    noRegister = true;
                }
                else if (String.Equals(a, "--diagnostic-simulate-missing-release", StringComparison.OrdinalIgnoreCase))
                {
                    simulateMissingRelease = true;
                }
                else if (String.Equals(a, "/Q", StringComparison.OrdinalIgnoreCase) ||
                         String.Equals(a, "/S", StringComparison.OrdinalIgnoreCase) ||
                         String.Equals(a, "--silent", StringComparison.OrdinalIgnoreCase))
                {
                    // Accepted compatibility switches; installer is silent by design.
                }
                else if (a.StartsWith("/InstallRoot=", StringComparison.OrdinalIgnoreCase))
                {
                    installRoot = a.Substring("/InstallRoot=".Length);
                }
                else { throw new ArgumentException("UNRECOGNIZED_OR_INCOMPLETE_ARGUMENT"); }
            }

            string tempRoot = Path.Combine(Path.GetTempPath(), "WordAgentSetup-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
            try
            {
                string script = Path.Combine(tempRoot, "Install-WordAgent.ps1");
                string zip = Path.Combine(tempRoot, "WordAgent-Release.zip");
                ExtractResource(ScriptResource, script);
                if (!simulateMissingRelease) ExtractResource(ZipResource, zip);

                string ps = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.System),
                    "WindowsPowerShell\\v1.0\\powershell.exe");
                if (!File.Exists(ps)) ps = "powershell.exe";

                string psArgs =
                    "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " + Quote(script) +
                    " -Silent";
                if (!String.IsNullOrWhiteSpace(installRoot))
                    psArgs += " -InstallRoot " + Quote(installRoot);
                if (noRegister)
                    psArgs += " -NoRegister";

                int childExit = RunChild(ps, psArgs, tempRoot);
                if (childExit != 0) return childExit;
                if (!noRegister)
                {
                    // Legacy registration is preserved unless separately identity-qualified.
                }
                return 0;
            }
            finally
            {
                try { Directory.Delete(tempRoot, true); } catch { }
            }
        }
        catch
        {
            return 252;
        }
    }
}
'@
Set-Content -LiteralPath $source -Value $src -Encoding UTF8

$args=@(
  '/nologo',
  '/target:winexe',
  '/platform:anycpu',
  '/optimize+',
  ('/out:'+$setupExe),
  ('/resource:'+$installScript+',WordAgent.InstallScript'),
  ('/resource:'+$releaseZip+',WordAgent.ReleaseZip'),
  $source
)
& $cscPath @args
if($LASTEXITCODE-ne0){throw ('CSC_MANAGED_SETUP_COMPILE_EXIT_'+$LASTEXITCODE)}
Assert-File $setupExe 'MANAGED_SETUP'

$sourceSha=Sha $source
# Generated bootstrapper source is retained as build evidence.

$result=[ordered]@{
  Schema='WORDAGENT_MANAGED_SETUP_BUILD_RESULT_V1'
  ProductVersion=$ProductVersion
  InstallerContainer='MANAGED_EMBEDDED_RESOURCE_BOOTSTRAPPER'
  Compiler=$cscPath
  CompilerVersion=[string](Get-Item -LiteralPath $cscPath).VersionInfo.FileVersion
  CompilerSHA256=Sha $cscPath
  InputReleaseZip=$releaseZip
  InputReleaseZipSHA256=Sha $releaseZip
  InputInstallScript=$installScript
  InputInstallScriptSHA256=Sha $installScript
  SetupExe=$setupExe
  SetupExeSHA256=Sha $setupExe
  SetupExeSize=(Get-Item -LiteralPath $setupExe).Length
  OutputReleaseZip=$outZip
  OutputReleaseZipSHA256=Sha $outZip
  EmbeddedSourceSHA256=$sourceSha
  SigningStatus=[string](Get-AuthenticodeSignature -FilePath $setupExe).Status
}
$result|ConvertTo-Json -Depth 20|Set-Content -LiteralPath (Join-Path $OutputRoot 'managed-build-result.json') -Encoding UTF8
$result
