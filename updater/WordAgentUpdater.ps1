# WordAgent verified updater (R20261010F private candidate).
# Local mode accepts a separately pinned SHA256; GitHub mode pins account/repo ID,
# requires immutable release metadata + GitHub asset digest, never executes remote scripts.
param(
  [string]$InstallRoot=(Join-Path $env:LOCALAPPDATA 'Programs\WordAgent'),
  [string]$ReleaseZip='',
  [string]$ExpectedSha256='',
  [string]$GitHubReleaseTag='',
  [switch]$CheckOnly,
  [switch]$NoRegister
)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$Repo='GOD13emad/PersianWordAgent'
$RepoId=1337917140L
$OwnerId=94546844L
function Get-Sha([string]$path){(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant()}
function Assert-ExpectedSha([string]$s){
 if($s-notmatch'^[a-fA-F0-9]{64}$'){throw 'UPDATER_EXPECTED_SHA256_REQUIRED'}
}
function Require-ArchiveManifest([string]$archivePath){
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $zip=[IO.Compression.ZipFile]::OpenRead($archivePath)
 try{
  if($zip.Entries.Count-gt10000){throw 'UPDATER_ZIP_ENTRY_LIMIT'}
  $directories=@($zip.Entries|Where-Object {$_.FullName -notmatch '/$' -and $_.FullName -match '^[^/]+/release-manifest\.json$'})
  if($directories.Count-ne1){throw 'UPDATER_SINGLE_RELEASE_MANIFEST_REQUIRED'}
  $reader=[IO.StreamReader]::new($directories[0].Open())
  try{$json=$reader.ReadToEnd()|ConvertFrom-Json}finally{$reader.Dispose()}
  if($json.Product-ne'WordAgent' -or $json.ProductVersion-notmatch '^2\.0\.0-rc[0-9]+$' -or [int]$json.Sequence-le0){
   throw 'UPDATER_RELEASE_IDENTITY_INVALID'
  }
  if(@($json.Files).Count-lt100){throw 'UPDATER_RELEASE_PAYLOAD_TOO_SMALL'}
  return $json
 }finally{$zip.Dispose()}
}
$canonical=[IO.Path]::GetFullPath($InstallRoot).TrimEnd('\')
$default=[IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'Programs\WordAgent')).TrimEnd('\')
if(!(Test-Path -LiteralPath $canonical -PathType Container)){throw 'UPDATER_TARGET_NOT_INSTALLED'}
$marker=Join-Path $canonical '.wordagent-install-marker.json'
if(!(Test-Path $marker -PathType Leaf)){throw 'UPDATER_MARKER_MISSING'}
$old=Get-Content $marker -Raw -Encoding utf8|ConvertFrom-Json
if($old.Product-ne'WordAgent' -or [int]$old.Sequence-le0){throw 'UPDATER_INSTALLED_IDENTITY_INVALID'}
if($NoRegister -and $canonical.Equals($default,[StringComparison]::OrdinalIgnoreCase)){
 throw 'UPDATER_REGISTERED_INSTALL_REQUIRES_REGISTRATION'
}
$localMode=![string]::IsNullOrWhiteSpace($ReleaseZip)
$remoteMode=![string]::IsNullOrWhiteSpace($GitHubReleaseTag)
if($localMode-eq$remoteMode){throw 'UPDATER_CHOOSE_EXACTLY_ONE_TRUST_MODE'}
if($remoteMode -and $ExpectedSha256){throw 'UPDATER_REMOTE_SHA_MUST_BE_GITHUB_ASSET_DIGEST'}
if($localMode){Assert-ExpectedSha $ExpectedSha256}
$downloadUrl=''
$assetName=''
if($remoteMode){
 if($GitHubReleaseTag-notmatch'^v2\.0\.0-rc[0-9]+$'){throw 'UPDATER_GITHUB_TAG_INVALID'}
 $headers=@{'Accept'='application/vnd.github+json';'User-Agent'='WordAgent-Updater/2.0';'X-GitHub-Api-Version'='2026-03-10'}
 $repository=Invoke-RestMethod -Uri ("https://api.github.com/repos/"+$Repo) -Headers $headers -TimeoutSec 20
 if([long]$repository.id-ne$RepoId -or [long]$repository.owner.id-ne$OwnerId -or
  [string]$repository.full_name -cne $Repo){throw 'UPDATER_GITHUB_REPO_IDENTITY_MISMATCH'}
 $url="https://api.github.com/repos/$Repo/releases/tags/$GitHubReleaseTag"
 $release=Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 20
 if($release.tag_name-cne$GitHubReleaseTag -or $release.draft -or $release.immutable -ne $true){
  throw 'UPDATER_GITHUB_RELEASE_NOT_IMMUTABLE_OR_UNPUBLISHED'
 }
 $assetName='WordAgent-'+$GitHubReleaseTag.Substring(1)+'-win-x64.zip'
 $assets=@($release.assets|Where-Object { $_.name -ceq $assetName -and $_.state -eq 'uploaded'})
 if($assets.Count-ne1){throw 'UPDATER_GITHUB_ASSET_NOT_UNIQUE'}
 $asset=$assets[0]
 if([string]$asset.digest-notmatch'^sha256:([0-9a-fA-F]{64})$'){
  throw 'UPDATER_GITHUB_ASSET_DIGEST_MISSING'
 }
 $ExpectedSha256=$Matches[1].ToUpperInvariant()
 $downloadUrl=[string]$asset.browser_download_url
 $canonicalUrl='https://github.com/'+$Repo+'/releases/download/'+$GitHubReleaseTag+'/'+$assetName
 if($downloadUrl-cne$canonicalUrl){throw 'UPDATER_GITHUB_DOWNLOAD_URL_NOT_PINNED'}
 if([long]$asset.size-lt1000000 -or [long]$asset.size-gt1073741824){
  throw 'UPDATER_GITHUB_ASSET_SIZE_INVALID'
 }
}
$template=Join-Path $PSScriptRoot 'Install-WordAgent.ps1.template'
if(!(Test-Path $template -PathType Leaf)){throw 'UPDATER_TRUSTED_LOCAL_INSTALLER_TEMPLATE_MISSING'}
$body=Get-Content -LiteralPath $template -Raw -Encoding UTF8
if(([regex]::Matches($body,'__RELEASE_ZIP_SHA256__')).Count-ne1){
 throw 'UPDATER_TRUSTED_TEMPLATE_ANCHOR_INVALID'
}
$work=Join-Path ([IO.Path]::GetTempPath()) ('wordagent-updater-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -ErrorAction Stop|Out-Null
try{
 $candidate=Join-Path $work 'WordAgent-Release.zip'
 if($remoteMode){
  Invoke-WebRequest -Uri $downloadUrl -OutFile $candidate -MaximumRedirection 5 -TimeoutSec 600 -ErrorAction Stop
 }else{
  $source=[IO.Path]::GetFullPath($ReleaseZip)
  if(!(Test-Path $source -PathType Leaf)){throw 'UPDATER_LOCAL_ARCHIVE_NOT_FOUND'}
  if((Get-Item $source).Length-gt1073741824){throw 'UPDATER_LOCAL_ARCHIVE_TOO_LARGE'}
  Copy-Item -LiteralPath $source -Destination $candidate -ErrorAction Stop
 }
 if((Get-Sha $candidate)-ne$ExpectedSha256.ToUpperInvariant()){
  throw 'UPDATER_ASSET_SHA256_MISMATCH'
 }
 $next=Require-ArchiveManifest $candidate
 if($remoteMode -and ('v'+$next.ProductVersion)-cne $GitHubReleaseTag){
  throw 'UPDATER_REMOTE_MANIFEST_VERSION_MISMATCH'
 }
 if([int]$next.Sequence-le[int]$old.Sequence){throw 'UPDATER_DOWNGRADE_OR_SAME_SEQUENCE_REFUSED'}
 Write-Output ("UPDATER_PREFLIGHT_PASS OLD="+$old.Version+" NEXT="+$next.ProductVersion+" SEQ="+$next.Sequence+" SHA256="+$ExpectedSha256)
 if($CheckOnly){Write-Output 'UPDATER_CHECK_ONLY_PASS NO_INSTALL';return}
 if(@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count-ne0 -or
    @(Get-Process WordAgent.Desktop -ErrorAction SilentlyContinue).Count-ne0){
  throw 'UPDATER_CLOSE_WORD_AND_WORDAGENT_BEFORE_INSTALL'
 }
 $generated=Join-Path $work 'Install-WordAgent.ps1'
 [IO.File]::WriteAllText($generated,$body.Replace('__RELEASE_ZIP_SHA256__',$ExpectedSha256),[Text.UTF8Encoding]::new($false))
 $argv=@('-NoProfile','-NonInteractive','-File',$generated,'-InstallRoot',$canonical,'-Silent')
 if($NoRegister){$argv+='-NoRegister'}
 & pwsh.exe @argv
 if($LASTEXITCODE-ne0){throw ('UPDATER_TRUSTED_INSTALLER_FAILED_EXIT_'+$LASTEXITCODE)}
 $post=Get-Content $marker -Raw -Encoding UTF8|ConvertFrom-Json
 if([int]$post.Sequence-ne[int]$next.Sequence -or $post.Version-ne$next.ProductVersion){
  throw 'UPDATER_INSTALLED_MARKER_MISMATCH'
 }
 Write-Output ("UPDATER_INSTALL_PASS VERSION="+$post.Version+" SEQUENCE="+$post.Sequence)
}finally{
 if(Test-Path $work){Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue}
}
