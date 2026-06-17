$ErrorActionPreference = "Stop"

$sourceApp = Join-Path $PSScriptRoot "app"
$installDir = Join-Path $env:LOCALAPPDATA "TataInventoryApp"

if (!(Test-Path $sourceApp)) {
    throw "Installer payload not found: $sourceApp"
}

New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Copy-Item -Path (Join-Path $sourceApp "*") -Destination $installDir -Recurse -Force

$shortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Tata Inventory App.lnk"
$targetPath = Join-Path $installDir "Run_Tata_Inventory_App.bat"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetPath
$shortcut.WorkingDirectory = $installDir
$shortcut.Description = "Tata Inventory App"
$shortcut.Save()

Write-Host "Tata Inventory App installed to:"
Write-Host "  $installDir"
Write-Host ""
Write-Host "A desktop shortcut was created:"
Write-Host "  $shortcutPath"
Write-Host ""
Write-Host "Run the app from the desktop shortcut, or run:"
Write-Host "  $targetPath"
