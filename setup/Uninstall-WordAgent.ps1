param(
 [string]$InstallRoot=$PSScriptRoot,
 [switch]$NoRegister,
 [switch]$Force
)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$InstallRoot=[IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
foreach($protected in @([IO.Path]::GetPathRoot($InstallRoot).TrimEnd('\'),$env:USERPROFILE,$env:LOCALAPPDATA,$env:APPDATA,$env:ProgramFiles,${env:ProgramFiles(x86)},$env:WINDIR)){
 if($protected-and$InstallRoot.Equals(([string]$protected).TrimEnd('\'),[StringComparison]::OrdinalIgnoreCase)){throw 'PROTECTED_UNINSTALL_ROOT'}
}
$cursor=$InstallRoot
while($cursor){if(Test-Path -LiteralPath $cursor){$item=Get-Item -LiteralPath $cursor -Force;if(($item.Attributes-band[IO.FileAttributes]::ReparsePoint)-ne0){throw 'UNINSTALL_REPARSE_PATH_REJECTED'}};$next=Split-Path -Parent $cursor;if($next-eq$cursor){break};$cursor=$next}
$marker=Join-Path $InstallRoot '.wordagent-install-marker.json'
if(!(Test-Path -LiteralPath $marker -PathType Leaf)){throw 'WORDAGENT_INSTALL_MARKER_MISSING'}
$identity=Get-Content -LiteralPath $marker -Raw -Encoding UTF8|ConvertFrom-Json
if([string]$identity.Product-ne'WordAgent'){throw 'UNINSTALL_PRODUCT_MISMATCH'}
if(!$Force){$d=@(Get-Process -Name 'WordAgent.Desktop','WordAgent' -ErrorAction SilentlyContinue);if($d.Count-ne0){throw 'WORDAGENT_PROCESS_ACTIVE'}}
$backup=$InstallRoot+'.uninstalled-'+[guid]::NewGuid().ToString('N').Substring(0,8)
if(Test-Path -LiteralPath $backup){throw 'UNINSTALL_BACKUP_COLLISION'}
$links=@();$registry='HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\WordAgent';$removeRegistry=$false
if(!$NoRegister){
 $ws=New-Object -ComObject WScript.Shell
 $paths=@((Join-Path ([Environment]::GetFolderPath('DesktopDirectory')) 'WordAgent.lnk'),(Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\WordAgent.lnk'),(Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\WordAgent\WordAgent.lnk'))
 foreach($path in $paths){if(Test-Path -LiteralPath $path){$target=$ws.CreateShortcut($path).TargetPath;if($target-and[IO.Path]::GetFullPath($target).StartsWith($InstallRoot+'\',[StringComparison]::OrdinalIgnoreCase)){$links+=@{Path=$path;Bytes=[IO.File]::ReadAllBytes($path)}}}}
 if(Test-Path -LiteralPath $registry){$key=Get-Item -LiteralPath $registry;$location=[string]$key.GetValue('InstallLocation');$removeRegistry=$location-and[IO.Path]::GetFullPath($location).TrimEnd('\').Equals($InstallRoot,[StringComparison]::OrdinalIgnoreCase)}
}
Move-Item -LiteralPath $InstallRoot -Destination $backup
try{
 if(!$NoRegister){
  foreach($link in $links){Remove-Item -LiteralPath $link.Path -Force}
  if($removeRegistry){Remove-Item -LiteralPath $registry -Recurse -Force}
 }
}catch{
 $errorRecord=$_
 foreach($link in $links){[IO.File]::WriteAllBytes($link.Path,$link.Bytes)}
 if(!(Test-Path -LiteralPath $InstallRoot)){Move-Item -LiteralPath $backup -Destination $InstallRoot}
 throw $errorRecord
}
Write-Output 'UNINSTALL_STATUS=PASS_ACTIVE_PATH_REMOVED_BACKUP_RETAINED'
Write-Output ('ROLLBACK_BACKUP='+$backup)
exit 0
