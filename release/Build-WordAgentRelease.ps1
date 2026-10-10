param(
 [Parameter(Mandatory=$true)][string]$ProjectRoot,
 [Parameter(Mandatory=$true)][string]$ToolingRoot,
 [Parameter(Mandatory=$true)][string]$OutputRoot,
 [string]$ProductVersion='',
 [int]$Sequence=0,
 [string]$EngineVersion='',
 [switch]$LegacyIExpress
)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Sha([string]$p){(Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToUpperInvariant()}
function Run([string]$exe,[string[]]$argv,[string]$label,[int]$timeout=900){$i=[Diagnostics.ProcessStartInfo]::new();$i.FileName=$exe;$i.WorkingDirectory=$ProjectRoot;$i.UseShellExecute=$false;$i.CreateNoWindow=$true;$i.RedirectStandardOutput=$true;$i.RedirectStandardError=$true;foreach($a in $argv){[void]$i.ArgumentList.Add($a)};$p=[Diagnostics.Process]::new();$p.StartInfo=$i;if(!$p.Start()){throw ($label+' start failed')};$o=$p.StandardOutput.ReadToEndAsync();$e=$p.StandardError.ReadToEndAsync();if(!$p.WaitForExit($timeout*1000)){try{$p.Kill($true)}catch{};throw ($label+' timeout')};$so=$o.GetAwaiter().GetResult();$se=$e.GetAwaiter().GetResult();if($p.ExitCode-ne0){throw ($label+' exit='+$p.ExitCode+' stderr='+$se)};[pscustomobject]@{Stdout=$so;Stderr=$se}}
$ProjectRoot=[IO.Path]::GetFullPath($ProjectRoot);$ToolingRoot=[IO.Path]::GetFullPath($ToolingRoot);$OutputRoot=[IO.Path]::GetFullPath($OutputRoot)
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
Assert-NewOutput $OutputRoot @($ProjectRoot,$ToolingRoot)
if($ProjectRoot-ne$ToolingRoot){throw 'NONCANONICAL_TOOLING_ROOT'};
$versionProps=Join-Path $ProjectRoot 'Directory.Build.props'
if(!(Test-Path -LiteralPath $versionProps -PathType Leaf)){throw 'VERSION_AUTHORITY_MISSING'}
[xml]$versionXml=Get-Content -LiteralPath $versionProps -Raw -Encoding UTF8
$versionNode=$versionXml.SelectSingleNode('/Project/PropertyGroup/WordAgentVersion')
if($null-eq$versionNode -or [string]::IsNullOrWhiteSpace($versionNode.InnerText)){throw 'WORDAGENT_VERSION_MISSING'}
$canonicalVersion=$versionNode.InnerText.Trim()
if($canonicalVersion-notmatch'^[0-9]+\.[0-9]+\.[0-9]+(-[A-Za-z0-9.-]+)?$'){throw 'INVALID_CANONICAL_VERSION'}
$sequenceNode=$versionXml.SelectSingleNode('/Project/PropertyGroup/WordAgentReleaseSequence')
if($null-eq$sequenceNode){throw 'RELEASE_SEQUENCE_AUTHORITY_MISSING'}
$canonicalSequence=[int]$sequenceNode.InnerText
if($canonicalSequence-le0){throw 'RELEASE_SEQUENCE_INVALID'}
if($Sequence-eq0){$Sequence=$canonicalSequence}
if($Sequence-ne$canonicalSequence){throw 'RELEASE_SEQUENCE_NOT_CANONICAL'}
if([string]::IsNullOrWhiteSpace($ProductVersion)){$ProductVersion=$canonicalVersion}
if([string]::IsNullOrWhiteSpace($EngineVersion)){$EngineVersion=$ProductVersion}
if($ProductVersion-ne$canonicalVersion){throw ('PRODUCT_VERSION_NOT_CANONICAL expected='+$canonicalVersion+' actual='+$ProductVersion)}
if($EngineVersion-ne$canonicalVersion){throw ('ENGINE_VERSION_NOT_CANONICAL expected='+$canonicalVersion+' actual='+$EngineVersion)}

Assert-NewOutput $OutputRoot @($ProjectRoot,$ToolingRoot);New-Item -ItemType Directory -Path $OutputRoot|Out-Null
$dotnet=(Get-Command dotnet.exe -ErrorAction Stop).Source
$desktopProj=Join-Path $ProjectRoot 'src\WordAgent.Desktop\WordAgent.Desktop.csproj';$cliProj=Join-Path $ProjectRoot 'src\WordAgent.Cli\WordAgent.Cli.csproj'
$productRoot=Join-Path $OutputRoot ('WordAgent-'+$ProductVersion);$desktopOut=Join-Path $productRoot 'desktop';$engineOut=Join-Path $desktopOut 'dist\win-x64';New-Item -ItemType Directory -Force -Path $engineOut|Out-Null
Run $dotnet @('restore',$desktopProj,'--locked-mode','-r','win-x64') 'desktop restore' 600|Out-Null;Run $dotnet @('publish',$desktopProj,'-c','Release','-r','win-x64','--self-contained','true','--no-restore','-o',$desktopOut,'-warnaserror') 'desktop publish' 900|Out-Null
Run $dotnet @('restore',$cliProj,'--locked-mode','-r','win-x64') 'cli restore' 600|Out-Null;Run $dotnet @('publish',$cliProj,'-c','Release','-r','win-x64','--self-contained','true','--no-restore','-o',$engineOut,'-warnaserror') 'cli publish' 900|Out-Null
$engine=Join-Path $engineOut 'WordAgent.exe';if(!(Test-Path $engine)){throw 'published Engine missing'};$v=Run $engine @('version') 'engine version' 30;if(($v.Stdout.Trim()-split"`r?`n")[0].Trim()-ne$EngineVersion){throw 'published Engine version mismatch'}
$upDir=Join-Path $productRoot 'updater';New-Item -ItemType Directory -Force -Path $upDir|Out-Null;Copy-Item (Join-Path $ToolingRoot 'updater\WordAgentUpdater.ps1') (Join-Path $upDir 'WordAgentUpdater.ps1');Copy-Item (Join-Path $ToolingRoot 'setup\Install-WordAgent.ps1.template') (Join-Path $upDir 'Install-WordAgent.ps1.template');Copy-Item (Join-Path $ToolingRoot 'updater\channel.json') (Join-Path $upDir 'channel.json');Copy-Item (Join-Path $ToolingRoot 'setup\Uninstall-WordAgent.ps1') (Join-Path $productRoot 'Uninstall-WordAgent.ps1')
$files=@();foreach($f in @(Get-ChildItem $productRoot -Recurse -File|Sort-Object FullName)){$files += [ordered]@{Path=$f.FullName.Substring($productRoot.Length+1).Replace('\','/');Size=$f.Length;SHA256=Sha $f.FullName}}
$manifest=[ordered]@{Product='WordAgent';ProductVersion=$ProductVersion;Sequence=$Sequence;Channel='rc';EngineVersion=$EngineVersion;Runtime='win-x64';BuiltAt=(Get-Date).ToString('o');Files=$files};$manifest|ConvertTo-Json -Depth 20|Set-Content -LiteralPath (Join-Path $productRoot 'release-manifest.json') -Encoding UTF8
$zip=Join-Path $OutputRoot ('WordAgent-'+$ProductVersion+'-win-x64.zip');[IO.Compression.ZipFile]::CreateFromDirectory($productRoot,$zip,[IO.Compression.CompressionLevel]::Optimal,$true);$zipSha=Sha $zip
$setupPayload=Join-Path $OutputRoot 'setup-payload';New-Item -ItemType Directory -Force -Path $setupPayload|Out-Null;Copy-Item $zip (Join-Path $setupPayload 'WordAgent-Release.zip');$tpl=Get-Content (Join-Path $ToolingRoot 'setup\Install-WordAgent.ps1.template') -Raw -Encoding UTF8;$tpl=$tpl.Replace('__RELEASE_ZIP_SHA256__',$zipSha);Set-Content -LiteralPath (Join-Path $setupPayload 'Install-WordAgent.ps1') -Value $tpl -Encoding UTF8
$wrapperText=@"
@echo off
setlocal
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Install-WordAgent.ps1" -Silent %*
set "rc=%ERRORLEVEL%"
exit /b %rc%
"@
Set-Content -LiteralPath (Join-Path $setupPayload 'Install-WordAgent.cmd') -Value $wrapperText -Encoding ASCII
$setupExe=$null;if($LegacyIExpress){$iexpress=(Get-Command iexpress.exe -ErrorAction Stop).Source
$setupExe=Join-Path $OutputRoot ('WordAgent-Setup-'+$ProductVersion+'-win-x64.exe');$sed=Join-Path $OutputRoot 'WordAgent-Setup.sed';$src=$setupPayload.TrimEnd('\')+'\';$sedText=@"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=$setupExe
FriendlyName=WordAgent Setup $ProductVersion
AppLaunched=cmd.exe /d /c Install-WordAgent.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=cmd.exe /d /c Install-WordAgent.cmd
UserQuietInstCmd=cmd.exe /d /c Install-WordAgent.cmd
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$src
[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
[Strings]
FILE0=WordAgent-Release.zip
FILE1=Install-WordAgent.ps1
FILE2=Install-WordAgent.cmd
"@
Set-Content -LiteralPath $sed -Value $sedText -Encoding ASCII
Run $iexpress @('/N',$sed) 'IExpress setup build' 300|Out-Null;if(!(Test-Path $setupExe)){throw 'Setup EXE missing after IExpress'}
}
$result=[ordered]@{InstallerContainer='MANAGED_SETUP_SEPARATE_STEP';SourceRoot=$ProjectRoot;ToolingRoot=$ToolingRoot;BuildPropsSHA256=Sha $versionProps;ProductVersion=$ProductVersion;Sequence=$Sequence;EngineVersion=$EngineVersion;ReleaseZip=$zip;ReleaseZipSHA256=$zipSha;SetupExe=$setupExe;SetupExeSHA256=$(if($setupExe){Sha $setupExe}else{$null});ProductRoot=$productRoot;ManifestSHA256=Sha (Join-Path $productRoot 'release-manifest.json')};$result|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $OutputRoot 'build-result.json') -Encoding UTF8

