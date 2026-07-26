param([Parameter(Mandatory=$true)][string]$ConfigPath)

$ErrorActionPreference = 'Stop'
$Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$UpdateRoot = Join-Path $env:APPDATA 'Signarama\Illustrator Helper\updates'
$LogDirectory = Join-Path $UpdateRoot 'logs'
New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null

function Write-UpdateLog([string]$EventName, [hashtable]$Details = @{}) {
  $Record = @{timestamp=(Get-Date).ToUniversalTime().ToString('o'); event=$EventName}
  foreach($Key in $Details.Keys) {$Record[$Key] = $Details[$Key]}
  $File = Join-Path $LogDirectory ((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd') + '.jsonl')
  Add-Content -LiteralPath $File -Value ($Record | ConvertTo-Json -Compress) -Encoding UTF8
}

function Assert-SafeArchive([string]$PackagePath) {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $Archive = [IO.Compression.ZipFile]::OpenRead($PackagePath)
  try {
    if($Archive.Entries.Count -eq 0) {throw 'Update archive is empty'}
    foreach($Entry in $Archive.Entries) {
      $Name = $Entry.FullName.Replace('\', '/')
      if($Name.StartsWith('/') -or $Name -match '^[A-Za-z]:' -or ($Name.Split('/') -contains '..')) {
        throw "Unsafe archive path: $Name"
      }
    }
  } finally {$Archive.Dispose()}
}

function Assert-StagedPlugin([string]$Directory) {
  $ManifestPath = Join-Path $Directory 'CSXS\manifest.xml'
  if(!(Test-Path -LiteralPath $ManifestPath) -or !(Test-Path -LiteralPath (Join-Path $Directory 'index.html'))) {throw 'Required plugin files are missing'}
  $Xml = Get-Content -LiteralPath $ManifestPath -Raw
  $Id = [regex]::Match($Xml, 'ExtensionBundleId="([^"]+)"')
  $Version = [regex]::Match($Xml, 'ExtensionBundleVersion="([^"]+)"')
  if(!$Id.Success -or $Id.Groups[1].Value -ne $Config.pluginId) {throw 'Plugin ID does not match'}
  if(!$Version.Success -or $Version.Groups[1].Value -ne $Config.targetVersion) {throw 'Plugin version does not match'}
}

$Staging = Join-Path $UpdateRoot ("staging\{0}-{1}" -f $Config.targetVersion, [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
$Backup = Join-Path $UpdateRoot ("backups\{0}" -f $Config.installedVersion)
$Replaced = $false

try {
  Write-UpdateLog 'update-started' @{installedVersion=$Config.installedVersion; targetVersion=$Config.targetVersion; installer='powershell'}
  Write-Host ''
  Write-Host 'Waiting for user to close Illustrator. Do not close this PowerShell window.' -ForegroundColor Yellow
  Write-Host ''
  $Deadline = (Get-Date).AddMinutes(30)
  while(Get-Process -Name Illustrator -ErrorAction SilentlyContinue) {
    if((Get-Date) -gt $Deadline) {throw 'Illustrator remains open after 30 minutes'}
    Start-Sleep -Seconds 2
  }
  Write-Host 'Illustrator is closed. Installing update...' -ForegroundColor Cyan
  Assert-SafeArchive $Config.packagePath
  New-Item -ItemType Directory -Path $Staging -Force | Out-Null
  Expand-Archive -LiteralPath $Config.packagePath -DestinationPath $Staging -Force
  Assert-StagedPlugin $Staging
  New-Item -ItemType Directory -Path (Split-Path $Backup -Parent) -Force | Out-Null
  if(Test-Path -LiteralPath $Backup) {Remove-Item -LiteralPath $Backup -Recurse -Force}
  Move-Item -LiteralPath $Config.installPath -Destination $Backup
  $Replaced = $true
  try {
    Move-Item -LiteralPath $Staging -Destination $Config.installPath
    Assert-StagedPlugin $Config.installPath
  } catch {
    if(Test-Path -LiteralPath $Config.installPath) {Remove-Item -LiteralPath $Config.installPath -Recurse -Force}
    if(Test-Path -LiteralPath $Backup) {Move-Item -LiteralPath $Backup -Destination $Config.installPath}
    Write-UpdateLog 'rollback-complete' @{targetVersion=$Config.targetVersion; rollbackResult='success'; message=$_.Exception.Message}
    throw
  }
  Remove-Item -LiteralPath $ConfigPath -Force
  Write-UpdateLog 'installation-complete' @{installedVersion=$Config.installedVersion; targetVersion=$Config.targetVersion; backupPath=$Backup; installationResult='success'}
  Write-Host ''
  Write-Host 'Update installed. You can now reopen Illustrator.' -ForegroundColor Green
  Start-Sleep -Seconds 8
} catch {
  Write-UpdateLog 'installation-failed' @{installedVersion=$Config.installedVersion; targetVersion=$Config.targetVersion; installationResult='failed'; errorCode='INSTALL_FAILED'; message=$_.Exception.Message}
  Write-Host ''
  Write-Host ("Update failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
  Write-Host 'See the Updates developer log after reopening Illustrator.' -ForegroundColor Yellow
  Start-Sleep -Seconds 15
  exit 1
}
