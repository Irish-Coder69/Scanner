# Document Scanner

Windows desktop app for scanning documents to a chosen folder.

## Features

- Auto-detects connected WIA scanners
- Scanner details panel: model, connection type, device ID, and status
- Destination folder picker with Browse button
- Remembers selected destination folder and file type between launches
- Requires a file name before scanning
- Supports PDF, PNG, and JPG output
- Multi-page PDF support
- Busy-scanner handling with retry prompt
- Prevents duplicate file names by auto-appending counters (example: file_1.pdf)
- Startup cleanup option removes files older than 1 year in the selected folder

## Requirements

- Windows with WIA-compatible scanner
- Python 3.13+

## Run From Source

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\scanner_app.py
```

## Build EXE

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

The build script now also refreshes the desktop shortcut:

- Rebuilds dist\Document Scanner.exe
- Updates the user desktop shortcut target to the newest EXE
- Updates the shortcut icon to scanner_icon.ico
- Updates the public desktop shortcut too, if it already exists

## Installer (optional)

If you want an installer package, build the EXE first and then compile installer.iss with Inno Setup.

## Versioning And GitHub Releases

- The app version is stored in `version.json`.
- Bump `version.json` before every GitHub release.
- Tag releases as `v<version>` (example: `v1.0.1`).

### Release workflow

A GitHub Actions workflow is included at `.github/workflows/release.yml`.

On tag push (`v*`), it will:

1. Validate that tag name matches `version.json`
2. Build `Document Scanner.exe`
3. Sync `installer.iss` `AppVersion` from `version.json`
4. Build installer with Inno Setup
5. Publish release asset: `installer.exe`

### Optional code signing (recommended)

To reduce "unrecognized download" warnings, configure code signing in GitHub Secrets:

1. `CODE_SIGN_CERT_BASE64`: Base64-encoded `.pfx` certificate bytes
2. `CODE_SIGN_CERT_PASSWORD`: Password for that `.pfx`

When these secrets are present, the workflow signs both:

1. `dist/Document Scanner.exe`
2. `installer_output/installer.exe`

### Example release commands

```powershell
git add .
git commit -m "Release v1.0.1"
git tag v1.0.1
git push origin main --tags
```

## Troubleshooting

### "running scripts is disabled on this system"

If you see this PowerShell error when running `build_exe.ps1`:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

The `-ExecutionPolicy Bypass` flag temporarily allows the script to run without changing your system settings.

### "The scanner is busy"

If you get a busy error during a scan:

1. Close any other scanning applications (Epson Scan, Windows Scan, etc.)
2. Wait a few seconds
3. Click Retry in the dialog

The app will retry up to 3 times before giving up.

### Scanner not detected

**Check:**

- Scanner is plugged in and powered on
- Scanner drivers are installed (WIA-compatible)
- Click the "Refresh" button in the app to re-scan
- Check Device Manager to confirm the scanner appears

### Missing dependencies

If the app crashes on startup with import errors, reinstall requirements:

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade -r requirements.txt
```

### File saves to wrong location

The app remembers the last destination folder. If scanning to an unexpected location:

1. Click "Browse" and select the correct folder
2. The new folder will be saved for future scans

### No desktop shortcut after build

Run the build script again with:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

The script will automatically create or update your desktop shortcut.

### Pinning to the taskbar

Use this stable method:

1. Launch **Document Scanner** from the desktop shortcut
2. While it is running, right-click the app icon in the taskbar
3. Select **"Pin to taskbar"**

Note: Depending on Windows version/policy, the desktop shortcut context menu may not show a direct "Pin to taskbar" option.
