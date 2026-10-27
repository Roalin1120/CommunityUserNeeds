$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $MyInvocation.MyCommand.Path
$targetBat = Join-Path $project 'launch_with_browser.bat'
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop 'WhatsApp Export Admin.lnk'

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetBat
$shortcut.WorkingDirectory = $project
$shortcut.Description = 'Start local Flask app and open http://127.0.0.1:5000'
$shortcut.Save()

Write-Host "Created: $shortcutPath"
