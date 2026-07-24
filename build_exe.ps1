Set-Location $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$exePath = Join-Path $PSScriptRoot "dist\Document Scanner.exe"
$iconPath = Join-Path $PSScriptRoot "scanner_icon.ico"

function Set-Shortcut {
	param(
		[Parameter(Mandatory = $true)][string]$ShortcutPath,
		[Parameter(Mandatory = $true)][string]$TargetPath,
		[Parameter(Mandatory = $true)][string]$IconPath,
		[string]$WorkingDirectory
	)

	$shell = New-Object -ComObject WScript.Shell
	$shortcut = $shell.CreateShortcut($ShortcutPath)
	$shortcut.TargetPath = $TargetPath
	$shortcut.WorkingDirectory = $WorkingDirectory
	$shortcut.IconLocation = "$IconPath,0"
	$shortcut.Save()
}

& $python -m pip install -r requirements.txt
& $python -c "from scanner_app import create_default_icon; create_default_icon('scanner_icon.ico', overwrite=True)"
& $python -m PyInstaller --onefile --windowed --icon=scanner_icon.ico --name 'Document Scanner' scanner_app.py

if (Test-Path $exePath) {
	# Copy EXE and icon to local drive so Windows allows pinning to taskbar
	$localDir = Join-Path $env:LOCALAPPDATA "Document Scanner"
	$localExe = Join-Path $localDir "Document Scanner.exe"
	$localIcon = Join-Path $localDir "scanner_icon.ico"
	$shortcutTarget = $exePath
	$shortcutIcon = $iconPath
	$shortcutWorkingDir = (Split-Path -Parent $exePath)
	$usingLocalCopy = $false
	if (-not (Test-Path $localDir)) {
		New-Item -ItemType Directory -Path $localDir | Out-Null
	}
	try {
		Copy-Item -Path $exePath -Destination $localExe -Force -ErrorAction Stop
		Copy-Item -Path $iconPath -Destination $localIcon -Force -ErrorAction Stop
		$shortcutTarget = $localExe
		$shortcutIcon = $localIcon
		$shortcutWorkingDir = $localDir
		$usingLocalCopy = $true
		Write-Host "Copied EXE to: $localExe" -ForegroundColor Cyan
	}
	catch {
		Write-Host "Local EXE copy failed (likely running). Falling back to dist EXE for shortcuts." -ForegroundColor Yellow
	}

	$userDesktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Document Scanner.lnk"
	Set-Shortcut -ShortcutPath $userDesktopShortcut -TargetPath $shortcutTarget -IconPath $shortcutIcon -WorkingDirectory $shortcutWorkingDir
	Write-Host "Updated shortcut: $userDesktopShortcut" -ForegroundColor Yellow

	$publicDesktopShortcut = "C:\Users\Public\Desktop\Document Scanner.lnk"
	if (Test-Path $publicDesktopShortcut) {
		Set-Shortcut -ShortcutPath $publicDesktopShortcut -TargetPath $shortcutTarget -IconPath $shortcutIcon -WorkingDirectory $shortcutWorkingDir
		Write-Host "Updated shortcut: $publicDesktopShortcut" -ForegroundColor Yellow
	}

	if ($usingLocalCopy) {
		Write-Host "Shortcuts target local AppData build." -ForegroundColor Green
	}
	else {
		Write-Host "Shortcuts target dist build to ensure latest version is used." -ForegroundColor Green
	}

	# Flush Windows icon cache so the desktop icon redraws immediately
	$cacheDir = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
	Get-ChildItem "$cacheDir\iconcache*.db" -ErrorAction SilentlyContinue | Remove-Item -Force
	$legacyCache = "$env:LOCALAPPDATA\IconCache.db"
	if (Test-Path $legacyCache) { Remove-Item $legacyCache -Force }
	try {
		$sig = '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);'
		Add-Type -MemberDefinition $sig -Namespace WinAPI -Name Shell32 -ErrorAction SilentlyContinue
		[WinAPI.Shell32]::SHChangeNotify(0x8000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
		Write-Host "Icon cache flushed." -ForegroundColor Cyan
	}
	catch { }
}

Write-Host "Build complete. EXE and desktop shortcut are up to date." -ForegroundColor Green
