import json
import os
import queue
import re
import tempfile
import threading
import time
import urllib.error
import urllib.request
import urllib.parse
import webbrowser
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Literal, TypedDict, cast

try:
    import pythoncom
    import win32com.client
    from PIL import Image, ImageDraw, ImageTk
except ImportError:
    pythoncom = None
    win32com = None
    Image = None
    ImageDraw = None
    ImageTk = None

SCANNER_DEVICE_TYPE = 1
WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
WIA_FORMAT_JPG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
WIA_CURRENT_INTENT_PROPERTY = 6146
WIA_HORIZONTAL_DPI_PROPERTY = 6147
WIA_VERTICAL_DPI_PROPERTY = 6148
WIA_INTENT_COLOR = 1
WIA_INTENT_GRAYSCALE = 2
WIA_INTENT_BLACKWHITE = 4
WIA_DOCUMENT_HANDLING_STATUS_PROPERTY = 3087
WIA_DOCUMENT_HANDLING_SELECT_PROPERTY = 3088
WIA_DPS_FEEDER = 1
WIA_DPS_FLATBED = 2
DEFAULT_UPDATE_MANIFEST_URL = "https://api.github.com/repos/Irish-Coder69/Scanner/releases/latest"
PROJECT_RELEASES_URL = "https://github.com/Irish-Coder69/Scanner/releases/latest"
PROJECT_COPYRIGHT = "Copyright (c) 2026 Judson M. Fitzpatrick - Irish_Coder's_Programing"


def load_version_info() -> dict[str, str]:
    default_info = {
        "version": "1.0.0",
        "update_manifest_url": "",
    }
    candidate_paths: list[Path] = []

    # Installed onefile EXE: keep version.json next to the executable.
    if getattr(sys, "frozen", False):
        try:
            candidate_paths.append(Path(sys.executable).with_name("version.json"))
        except Exception:
            pass

    # Source/dev mode fallback.
    candidate_paths.append(Path(__file__).with_name("version.json"))

    # PyInstaller extraction fallback when bundled with --add-data.
    meipass = getattr(sys, "_MEIPASS", "")
    if isinstance(meipass, str) and meipass:
        candidate_paths.append(Path(meipass) / "version.json")

    try:
        for version_file in candidate_paths:
            if not version_file.exists():
                continue

            with open(version_file, "r", encoding="utf-8") as handle:
                raw = json.load(handle)
            if not isinstance(raw, dict):
                continue

            version = str(raw.get("version", "")).strip()
            update_manifest_url = str(raw.get("update_manifest_url", "")).strip()
            if version:
                default_info["version"] = version
            if update_manifest_url:
                default_info["update_manifest_url"] = update_manifest_url
            break
    except Exception:
        pass
    return default_info


_VERSION_INFO = load_version_info()
APP_VERSION = _VERSION_INFO["version"]
UPDATE_MANIFEST_URL = os.environ.get("DOC_SCANNER_UPDATE_URL", _VERSION_INFO["update_manifest_url"]).strip() or DEFAULT_UPDATE_MANIFEST_URL


class ScannerRecord(TypedDict):
    name: str
    connection: str
    device_id: str
    status: str
    adf_supported: bool
    adf_ready: bool


class SettingsData(TypedDict, total=False):
    folder: str
    format: str
    dpi: int
    pages: int
    adf_delay_seconds: int


ScanResult = tuple[Literal["ok", "cancelled", "error"], str | Exception | None]


def sanitize_filename(name: str) -> str:
    invalid = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in invalid else ch for ch in name).strip()
    return cleaned.rstrip(".")


def get_unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 1
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def make_wia_safe_temp_path(suffix: str) -> str:
    fd, temp_name = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    os.unlink(temp_name)
    return temp_name


def detect_connection_type(info: Any) -> str:
    values: list[str] = []

    try:
        values.append(str(info.DeviceID))
    except Exception:
        pass

    try:
        for prop in info.Properties:
            try:
                prop_name = str(prop.Name).strip().lower()
                if prop_name in {"name", "port", "server", "description", "pnp id", "device id"}:
                    values.append(f"{prop_name}:{prop.Value}")
            except Exception:
                continue
    except Exception:
        pass

    blob = " | ".join(values).lower()

    if any(x in blob for x in ["usb", "usbscan", "vid_", "pid_"]):
        return "USB"
    if any(x in blob for x in ["wsd", "tcp", "ip_", "\\\\", "network", "ethernet", "wifi", "wi-fi"]):
        return "Network"
    if "bluetooth" in blob:
        return "Bluetooth"

    return "Unknown"


def get_scanner_status(info: Any) -> str:
    try:
        device_id = str(getattr(info, "DeviceID", "")).strip()
        if device_id:
            return "Ready"
    except Exception:
        pass
    return "Unknown"


def is_busy_error(exc: Exception) -> bool:
    error_text = str(exc).lower()
    return "device is busy" in error_text or "wia device is busy" in error_text or "-2145320954" in error_text


def is_adf_empty_error(exc: Exception) -> bool:
    error_text = str(exc).lower()
    return (
        "no documents left in the document feeder" in error_text
        or "-2145320957" in error_text
    )


def _extract_semver(version_text: str) -> tuple[int, int, int]:
    """
    Extract a comparable semantic version triplet from free-form text.
    Examples handled: "1.2.3", "v1.2.3", "release-1.2.3-beta".
    """
    text = version_text.strip()
    if not text:
        return (0, 0, 0)

    match = re.search(r"(\d+)(?:\.(\d+))?(?:\.(\d+))?", text)
    if not match:
        return (0, 0, 0)

    major = int(match.group(1) or 0)
    minor = int(match.group(2) or 0)
    patch = int(match.group(3) or 0)
    return (major, minor, patch)


def _choose_latest_version_text(manifest: dict[str, Any]) -> str:
    """
    Read version text from supported endpoints.
    Priority: explicit custom manifest version, then GitHub tag, then release name.
    """
    candidates = [
        str(manifest.get("version", "")).strip(),
        str(manifest.get("tag_name", "")).strip(),
        str(manifest.get("name", "")).strip(),
    ]

    for raw in candidates:
        if not raw:
            continue
        cleaned = raw[1:] if raw.lower().startswith("v") else raw
        if _extract_semver(cleaned) != (0, 0, 0):
            return cleaned

    return ""


def _choose_download_url(manifest: dict[str, Any]) -> str:
    """
    Pick the best download target from supported manifest formats.
    """
    custom_url = str(manifest.get("download_url", "")).strip()
    if custom_url:
        return custom_url

    assets = manifest.get("assets")
    if isinstance(assets, list):
        for asset in assets:
            if not isinstance(asset, dict):
                continue
            name = str(asset.get("name", "")).strip().lower()
            browser_url = str(asset.get("browser_download_url", "")).strip()
            if not browser_url:
                continue
            if name == "installer.exe":
                return browser_url

        # Fallback to the first downloadable asset URL.
        for asset in assets:
            if isinstance(asset, dict):
                browser_url = str(asset.get("browser_download_url", "")).strip()
                if browser_url:
                    return browser_url

    return str(manifest.get("html_url", "")).strip()


def _read_prop_value(props: Any, property_id: int) -> int | None:
    try:
        for prop in props:
            try:
                if int(getattr(prop, "PropertyID", -1)) == property_id:
                    return int(prop.Value)
            except Exception:
                continue
    except Exception:
        return None
    return None


def detect_adf_capability(device: Any) -> tuple[bool, bool]:
    """
    Return (adf_supported, adf_ready) based on WIA document handling properties.
    """
    select_value = _read_prop_value(getattr(device, "Properties", []), WIA_DOCUMENT_HANDLING_SELECT_PROPERTY)
    status_value = _read_prop_value(getattr(device, "Properties", []), WIA_DOCUMENT_HANDLING_STATUS_PROPERTY)

    adf_supported = False
    adf_ready = False

    if select_value is not None:
        adf_supported = bool(select_value & WIA_DPS_FEEDER)

    if status_value is not None:
        adf_ready = bool(status_value & WIA_DPS_FEEDER)
        adf_supported = adf_supported or bool(status_value & WIA_DPS_FEEDER)

    return adf_supported, adf_ready


def get_connected_scanners() -> list[ScannerRecord]:
    if win32com is None or pythoncom is None:
        return []

    try:
        pythoncom_module = cast(Any, pythoncom)
        win32_client = cast(Any, win32com.client)
        pythoncom_module.CoInitialize()
        manager = win32_client.Dispatch("WIA.DeviceManager")
        scanners: list[ScannerRecord] = []

        for info in manager.DeviceInfos:
            try:
                if int(info.Type) != SCANNER_DEVICE_TYPE:
                    continue

                scanner_name = "Unknown scanner"
                for prop in info.Properties:
                    try:
                        if str(prop.Name).lower() == "name":
                            scanner_name = str(prop.Value)
                            break
                    except Exception:
                        continue

                device_id = str(getattr(info, "DeviceID", ""))
                adf_supported = False
                adf_ready = False
                try:
                    device = info.Connect()
                    adf_supported, adf_ready = detect_adf_capability(device)
                except Exception:
                    pass
                scanners.append(
                    {
                        "name": scanner_name,
                        "connection": detect_connection_type(info),
                        "device_id": device_id,
                        "status": get_scanner_status(info),
                        "adf_supported": adf_supported,
                        "adf_ready": adf_ready,
                    }
                )
            except Exception:
                continue

        return scanners
    except Exception:
        return []


# ---------- Settings persistence ----------

_APPDATA = os.environ.get("APPDATA", "")
CONFIG_DIR = Path(_APPDATA) / "DocumentScanner" if _APPDATA else Path.home() / "AppData" / "Roaming" / "DocumentScanner"
CONFIG_FILE = CONFIG_DIR / "settings.json"


def load_settings() -> SettingsData:
    """Return persisted settings, or an empty dict if none found."""
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
                raw = json.load(fh)
                if not isinstance(raw, dict):
                    return {}
                raw_map = cast(dict[str, object], raw)

                settings: SettingsData = {}
                folder = raw_map.get("folder")
                if isinstance(folder, str):
                    settings["folder"] = folder

                save_format = raw_map.get("format")
                if isinstance(save_format, str):
                    settings["format"] = save_format

                dpi = raw_map.get("dpi")
                if isinstance(dpi, int):
                    settings["dpi"] = dpi

                pages = raw_map.get("pages")
                if isinstance(pages, int):
                    settings["pages"] = pages

                adf_delay_seconds = raw_map.get("adf_delay_seconds")
                if isinstance(adf_delay_seconds, int):
                    settings["adf_delay_seconds"] = adf_delay_seconds

                return settings
    except Exception:
        pass
    return {}


def save_settings(data: SettingsData) -> None:
    """Write settings to the config file, silently ignoring errors."""
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
    except Exception:
        pass


# ------------------------------------------


def create_default_icon(path: str = "scanner_icon.ico", overwrite: bool = True) -> None:
    if Image is None or ImageDraw is None:
        return

    # Packaged builds already carry an embedded icon and may run from folders
    # where writing scanner_icon.ico is blocked.
    if getattr(sys, "frozen", False):
        return

    icon_path = Path(path)
    if icon_path.exists() and not overwrite:
        return

    img = Image.new("RGBA", (256, 256), (20, 52, 96, 255))
    draw = ImageDraw.Draw(img)

    draw.rounded_rectangle((30, 54, 225, 205), radius=20, fill=(242, 244, 247, 255), outline=(29, 41, 57, 255), width=5)
    draw.rectangle((58, 26, 196, 78), fill=(173, 184, 196, 255), outline=(55, 66, 79, 255), width=4)
    draw.rectangle((64, 108, 190, 174), fill=(255, 255, 255, 255), outline=(167, 176, 186, 255), width=3)
    draw.rectangle((146, 90, 195, 108), fill=(255, 183, 77, 255), outline=(198, 122, 24, 255), width=2)
    draw.line((78, 138, 176, 138), fill=(46, 110, 194, 255), width=6)
    draw.line((78, 154, 162, 154), fill=(46, 110, 194, 255), width=6)
    draw.ellipse((86, 88, 116, 118), fill=(46, 110, 194, 255))
    draw.ellipse((122, 88, 152, 118), fill=(46, 110, 194, 255))

    try:
        img.save(icon_path, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
    except PermissionError:
        # Startup should continue even if icon file is locked or read-only.
        return
    except OSError:
        # Some environments deny writes to the working folder; icon generation is optional.
        return


class ScannerApp(tk.Tk):
    _MAX_ADF_EMPTY_RETRIES = 5

    def __init__(self):
        super().__init__()
        self.title("Document Scanner")
        self.resizable(True, True)

        self.scanner_var = tk.StringVar(value="Checking for scanner...")

        _cfg = load_settings()
        _saved_folder = _cfg.get("folder", "")
        _default_folder = (
            _saved_folder
            if _saved_folder and Path(_saved_folder).exists()
            else str(Path.home() / "Documents")
        )
        self.folder_var = tk.StringVar(value=_default_folder)
        self.filename_var = tk.StringVar()
        self.format_var = tk.StringVar(value=_cfg.get("format", "PDF"))
        self.status_var = tk.StringVar(value="Ready")
        self.dialog_status_var = tk.StringVar(value="Click Scan Document to open the Windows scan dialog.")
        self.dialog_step_var = tk.StringVar(value="Step 1 of 4: Ready to scan")
        self.dialog_next_var = tk.StringVar(value="Next action: Enter a file name and click Scan Document.")
        self.dialog_progress_var = tk.DoubleVar(value=0.0)
        saved_dpi = int(_cfg.get("dpi", 300))
        if saved_dpi < 100:
            saved_dpi = 100
        if saved_dpi > 1200:
            saved_dpi = 1200
        self.dpi_var = tk.StringVar(value=str(saved_dpi))
        saved_pages = int(_cfg.get("pages", 1))
        if saved_pages < 1:
            saved_pages = 1
        if saved_pages > 100:
            saved_pages = 100
        self.pages_var = tk.StringVar(value=str(saved_pages))
        saved_adf_delay = int(_cfg.get("adf_delay_seconds", 4))
        if saved_adf_delay < 1:
            saved_adf_delay = 1
        if saved_adf_delay > 10:
            saved_adf_delay = 10
        self.adf_delay_var = tk.StringVar(value=str(saved_adf_delay))

        self.model_var = tk.StringVar(value="N/A")
        self.connection_var = tk.StringVar(value="N/A")
        self.device_id_var = tk.StringVar(value="N/A")
        self.ready_var = tk.StringVar(value="N/A")
        self.adf_var = tk.StringVar(value="N/A")
        self._source_options = ["Auto", "Flatbed"]
        self.scan_source_var = tk.StringVar(value="Auto")
        self.scan_quality_var = tk.StringVar(value="Color")
        self.preview_status_var = tk.StringVar(value="No preview yet. Click Scan Preview.")
        self.is_scanning = False
        self.preview_photo: Any | None = None
        self.preview_label: ttk.Label | None = None
        self.progress_pct_var = tk.StringVar(value="")
        self._progress_timer_id: str | None = None
        self._progress_start: float = 0.0
        self._progress_est: float = 10.0
        self._scan_cancel_requested = False
        self._active_batch_source: str | None = None

        # Auto-save whenever folder or format changes
        self.folder_var.trace_add("write", self._on_setting_change)
        self.format_var.trace_add("write", self._on_setting_change)
        self.dpi_var.trace_add("write", self._on_setting_change)
        self.pages_var.trace_add("write", self._on_setting_change)
        self.adf_delay_var.trace_add("write", self._on_setting_change)

        # Save on normal window close
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.setup_style()
        self.setup_icon()
        self.setup_menu()
        self.build_ui()
        self.detect_scanner()
        self.bind("<Return>", lambda event: self.scan_document())
        self.after(0, lambda: self.state("zoomed"))
        self.after(200, self.cleanup_old_files)
        self.after(300, lambda: self.check_for_updates(silent=True))

    def setup_menu(self) -> None:
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Refresh Scanner", command=self.detect_scanner)
        file_menu.add_separator()
        file_menu.add_command(label="Close Program", command=self._on_close)

        scan_settings_menu = tk.Menu(menubar, tearoff=0)
        scan_settings_menu.add_command(label="Image Quality...", command=self.open_quality_settings)
        scan_settings_menu.add_command(label="Resolution (DPI)...", command=self.open_dpi_settings)
        scan_settings_menu.add_command(label="Pages to Scan...", command=self.open_pages_settings)
        scan_settings_menu.add_command(label="Source...", command=self.open_source_settings)
        scan_settings_menu.add_command(label="ADF Page Delay (seconds)...", command=self.open_adf_delay_settings)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Check for Updates (GitHub)", command=lambda: self.check_for_updates(silent=False))
        help_menu.add_command(label="Open GitHub Releases", command=lambda: webbrowser.open(PROJECT_RELEASES_URL))

        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="About Document Scanner", command=self.show_about)

        menubar.add_cascade(label="File", menu=file_menu)
        menubar.add_cascade(label="Scan Settings", menu=scan_settings_menu)
        menubar.add_cascade(label="Help", menu=help_menu)
        menubar.add_cascade(label="About", menu=about_menu)
        self.config(menu=menubar)

    def show_about(self) -> None:
        messagebox.showinfo(
            "About Document Scanner",
            "\n".join([
                f"Document Scanner v{APP_VERSION}",
                "",
                "A Windows desktop application for scanning documents to PDF, PNG, and JPG.",
                "Built with WIA scanner support, multi-page ADF scanning, and update checks.",
                "",
                "Created by:",
                "Judson M. Fitzpatrick",
                "Irish_Coder's_Programing",
                "",
                PROJECT_COPYRIGHT,
                "",
                f"Updates are checked from GitHub:\n{DEFAULT_UPDATE_MANIFEST_URL}",
            ]),
        )

    def _open_setting_picker(self, title: str, prompt: str, variable: tk.StringVar, values: list[str]) -> None:
        if self.is_scanning:
            self.status_var.set("Finish the current scan before changing settings.")
            return

        if not values:
            return

        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=prompt, wraplength=320, justify="left").grid(row=0, column=0, sticky="w")

        initial_value = variable.get()
        if initial_value not in values:
            initial_value = values[0]
        selected = tk.StringVar(value=initial_value)

        combo = ttk.Combobox(frame, textvariable=selected, values=values, state="readonly", width=26)
        combo.grid(row=1, column=0, sticky="ew", pady=(8, 10))
        combo.focus_set()

        actions = ttk.Frame(frame)
        actions.grid(row=2, column=0, sticky="e")

        def _apply() -> None:
            variable.set(selected.get())
            self.status_var.set(f"{title} updated: {selected.get()}")
            dialog.destroy()

        ttk.Button(actions, text="Apply", command=_apply).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="Cancel", command=dialog.destroy).grid(row=0, column=1)

        frame.columnconfigure(0, weight=1)
        dialog.bind("<Return>", lambda _event: _apply())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        dialog.wait_window(dialog)

    def open_quality_settings(self) -> None:
        self._open_setting_picker(
            "Image Quality",
            "Choose the capture quality mode for scanning.",
            self.scan_quality_var,
            ["Color", "Grayscale", "BlackWhite", "Custom"],
        )

    def open_dpi_settings(self) -> None:
        self._open_setting_picker(
            "Resolution (DPI)",
            "Choose scan resolution. Higher DPI increases detail and file size.",
            self.dpi_var,
            [str(dpi) for dpi in range(100, 1201, 100)],
        )

    def open_pages_settings(self) -> None:
        self._open_setting_picker(
            "Pages to Scan",
            "Choose how many pages to capture in this scan run.",
            self.pages_var,
            [str(page) for page in range(1, 101)],
        )

    def open_source_settings(self) -> None:
        self._open_setting_picker(
            "Scan Source",
            "Choose source mode. ADF option appears only when feeder is supported.",
            self.scan_source_var,
            self._source_options,
        )

    def open_adf_delay_settings(self) -> None:
        self._open_setting_picker(
            "ADF Page Delay",
            "Choose delay between pages. Increase this if page 2+ is blank or unstable.",
            self.adf_delay_var,
            [str(second) for second in range(1, 11)],
        )

    def check_for_updates(self, silent: bool) -> None:
        if not UPDATE_MANIFEST_URL:
            if not silent:
                messagebox.showinfo(
                    "Check for Updates",
                    "Update source is not configured.\n\n"
                    "Set DOC_SCANNER_UPDATE_URL to a JSON endpoint returning at least: {\"version\": \"x.y.z\"}.",
                )
            return

        self.status_var.set("Checking for updates...")

        def _worker() -> None:
            try:
                parsed_url = urllib.parse.urlparse(UPDATE_MANIFEST_URL)
                query_items = urllib.parse.parse_qsl(parsed_url.query, keep_blank_values=True)
                query_items.append(("_ts", str(int(time.time()))))
                fresh_url = urllib.parse.urlunparse(
                    parsed_url._replace(query=urllib.parse.urlencode(query_items))
                )

                request = urllib.request.Request(
                    fresh_url,
                    headers={
                        "User-Agent": "DocumentScanner",
                        "Accept": "application/json",
                        "Cache-Control": "no-cache",
                        "Pragma": "no-cache",
                    },
                )
                with urllib.request.urlopen(request, timeout=8) as response:
                    payload = response.read().decode("utf-8")
                manifest = json.loads(payload)

                # Supports both a custom manifest ({"version": "x.y.z"}) and
                # GitHub Releases API response ({"tag_name": "vX.Y.Z"}).
                latest_version = _choose_latest_version_text(manifest)

                download_url = _choose_download_url(manifest)

                notes = str(manifest.get("notes", "")).strip()
                if not notes:
                    notes = str(manifest.get("body", "")).strip()

                if not latest_version:
                    raise RuntimeError("Update manifest did not include a version.")

                has_update = _extract_semver(latest_version) > _extract_semver(APP_VERSION)
                self.after(0, lambda: self._handle_update_result(silent, has_update, latest_version, download_url, notes, None))
            except Exception as exc:
                self.after(0, lambda: self._handle_update_result(silent, False, "", "", "", exc))

        threading.Thread(target=_worker, daemon=True).start()

    def _handle_update_result(
        self,
        silent: bool,
        has_update: bool,
        latest_version: str,
        download_url: str,
        notes: str,
        error: Exception | None,
    ) -> None:
        if error is not None:
            self.status_var.set("Unable to check for updates.")
            if not silent:
                messagebox.showwarning("Check for Updates", f"Unable to check for updates.\n\n{error}")
            return

        if has_update:
            self.status_var.set(f"Update available: {latest_version}")
            message = [
                f"A new version is available: {latest_version}",
                f"Current version: {APP_VERSION}",
            ]
            if notes:
                message.append(f"\nRelease notes:\n{notes}")
            if download_url:
                message.append("\nThe installer can be downloaded directly in-app.")
                should_download = messagebox.askyesno(
                    "Update Available",
                    "\n".join(message) + "\n\nDownload installer now?",
                )
                if should_download:
                    self._download_update_installer(download_url, latest_version)
            else:
                messagebox.showinfo("Update Available", "\n".join(message))
        else:
            self.status_var.set("Application is up to date.")
            if not silent:
                messagebox.showinfo(
                    "Check for Updates",
                    f"You are up to date (version {APP_VERSION}).\nLatest available: {latest_version or APP_VERSION}",
                )

    def _download_update_installer(self, download_url: str, latest_version: str) -> None:
        downloads_dir = Path.home() / "Downloads"
        downloads_dir.mkdir(parents=True, exist_ok=True)

        parsed = urllib.parse.urlparse(download_url)
        suggested_name = Path(parsed.path).name.strip()
        if not suggested_name:
            suggested_name = f"DocumentScannerSetup_v{latest_version}.exe"
        if not suggested_name.lower().endswith(".exe"):
            suggested_name = f"{suggested_name}.exe"

        target_path = get_unique_path(downloads_dir / suggested_name)
        self.status_var.set(f"Downloading update: {target_path.name}...")
        self.dialog_status_var.set(f"Downloading installer to {target_path}")

        def _worker() -> None:
            try:
                request = urllib.request.Request(
                    download_url,
                    headers={
                        "User-Agent": "DocumentScanner",
                        "Accept": "application/octet-stream,*/*",
                    },
                )
                with urllib.request.urlopen(request, timeout=120) as response, open(target_path, "wb") as handle:
                    while True:
                        chunk = response.read(1024 * 256)
                        if not chunk:
                            break
                        handle.write(chunk)
                self.after(0, lambda: self._finish_update_download(target_path, None, download_url))
            except Exception as exc:
                self.after(0, lambda: self._finish_update_download(target_path, exc, download_url))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish_update_download(self, target_path: Path, error: Exception | None, download_url: str) -> None:
        if error is not None:
            self.status_var.set("Update download failed.")
            messagebox.showwarning(
                "Download Failed",
                "Unable to download the installer directly.\n\n"
                f"Error: {error}\n\n"
                "No browser page was opened. Please try again from Check for Updates.",
            )
            return

        self.status_var.set(f"Update downloaded: {target_path.name}")
        self.dialog_status_var.set(f"Installer downloaded: {target_path}")
        messagebox.showinfo(
            "Download Complete",
            f"Installer downloaded to:\n{target_path}",
        )

    def setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Dialog.TLabelframe", padding=12)
        style.configure("Dialog.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("DialogHeader.TLabel", font=("Segoe UI", 9), foreground="#4f4f4f")
        style.configure("DialogStep.TLabel", font=("Segoe UI", 10, "bold"), foreground="#004aad")
        style.configure("DialogSection.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("DialogMuted.TLabel", font=("Segoe UI", 9), foreground="#5e5e5e")

    def setup_icon(self) -> None:
        try:
            ico_path = Path(__file__).with_name("scanner_icon.ico")
            if ico_path.exists():
                cast(Any, self).iconbitmap(str(ico_path))
        except Exception:
            pass

    def build_ui(self):
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)

        left_col = ttk.Frame(outer)
        right_col = ttk.Frame(outer)
        preview_col = ttk.Frame(outer)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right_col.grid(row=0, column=1, sticky="nsew", padx=8)
        preview_col.grid(row=0, column=2, sticky="nsew", padx=(8, 0))

        ttk.Label(left_col, text="Document Scanner", font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )

        ttk.Label(
            left_col,
            text="Scan documents directly to a selected folder.",
            foreground="#555555",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 14))

        ttk.Label(left_col, text="Detected scanner:").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        ttk.Label(left_col, textvariable=self.scanner_var, relief="solid", width=56).grid(
            row=2, column=1, sticky="ew", padx=6, pady=6
        )
        self.refresh_button = ttk.Button(left_col, text="Refresh", command=self.detect_scanner)
        self.refresh_button.grid(row=2, column=2, padx=6, pady=6)

        details = ttk.LabelFrame(left_col, text="Scanner Details", padding=10)
        details.grid(row=3, column=0, columnspan=3, sticky="ew", padx=6, pady=8)

        ttk.Label(details, text="Model:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(details, textvariable=self.model_var).grid(row=0, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(details, text="Connection:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(details, textvariable=self.connection_var).grid(row=1, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(details, text="Device ID:").grid(row=2, column=0, sticky="nw", padx=6, pady=4)
        ttk.Label(details, textvariable=self.device_id_var, wraplength=560).grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(details, text="Status:").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(details, textvariable=self.ready_var).grid(row=3, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(details, text="ADF:").grid(row=4, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(details, textvariable=self.adf_var).grid(row=4, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(left_col, text="Destination folder:").grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 0))
        folder_row = ttk.Frame(left_col)
        folder_row.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 4))
        ttk.Entry(folder_row, textvariable=self.folder_var, width=96).grid(row=0, column=0, sticky="ew")
        self.browse_button = ttk.Button(folder_row, text="Browse", command=self.browse_folder)
        self.browse_button.grid(row=1, column=0, sticky="w", pady=(6, 0))
        folder_row.columnconfigure(0, weight=1)

        ttk.Label(left_col, text="File name:").grid(row=6, column=0, sticky="w", padx=6, pady=6)
        self.filename_entry = ttk.Entry(left_col, textvariable=self.filename_var, width=58)
        self.filename_entry.grid(row=6, column=1, sticky="ew", padx=6, pady=6)

        ttk.Label(left_col, text="File type:").grid(row=7, column=0, sticky="w", padx=6, pady=6)
        ttk.Combobox(
            left_col,
            textvariable=self.format_var,
            values=["PDF", "PNG", "JPG"],
            state="readonly",
            width=12,
        ).grid(row=7, column=1, sticky="w", padx=6, pady=6)

        settings_summary = ttk.LabelFrame(left_col, text="Scan Settings", padding=8)
        settings_summary.grid(row=8, column=0, columnspan=3, sticky="ew", padx=6, pady=(6, 4))
        ttk.Label(
            settings_summary,
            text="Use the Scan Settings menu to adjust options in separate dialogs.",
            foreground="#555555",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=4, pady=(0, 8))
        ttk.Label(settings_summary, text="Quality:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, textvariable=self.scan_quality_var).grid(row=1, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, text="DPI:").grid(row=2, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, textvariable=self.dpi_var).grid(row=2, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, text="Pages:").grid(row=3, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, textvariable=self.pages_var).grid(row=3, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, text="Source:").grid(row=4, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, textvariable=self.scan_source_var).grid(row=4, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, text="ADF delay (s):").grid(row=5, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(settings_summary, textvariable=self.adf_delay_var).grid(row=5, column=1, sticky="w", padx=4, pady=2)

        actions = ttk.Frame(left_col)
        actions.grid(row=9, column=1, sticky="w", padx=(2, 6), pady=(8, 10))

        self.scan_button = ttk.Button(actions, text="Scan Document", command=self.scan_document)
        self.scan_button.grid(row=0, column=0, padx=(0, 8), pady=(0, 2))
        self.preview_button = ttk.Button(actions, text="Scan Preview", command=self.scan_preview)
        self.preview_button.grid(row=0, column=1, padx=(0, 8), pady=(0, 2))
        self.clear_button = ttk.Button(actions, text="Clear Name", command=self.clear_name)
        self.clear_button.grid(row=0, column=2, padx=(0, 8), pady=(0, 2))
        self.close_button = ttk.Button(actions, text="Close", command=self._on_close)
        self.close_button.grid(row=0, column=3, pady=(0, 2))

        ttk.Separator(left_col, orient="horizontal").grid(row=10, column=0, columnspan=3, sticky="ew", pady=10)

        ttk.Label(left_col, text="Status:").grid(row=11, column=0, sticky="nw", padx=6, pady=6)
        ttk.Label(left_col, textvariable=self.status_var, foreground="#004aad").grid(
            row=11, column=1, columnspan=2, sticky="w", padx=6, pady=6
        )

        ttk.Label(left_col, text="Scan Progress:").grid(row=12, column=0, sticky="w", padx=6, pady=(0, 6))
        progress_frame = ttk.Frame(left_col)
        progress_frame.grid(row=12, column=1, columnspan=2, sticky="ew", padx=6, pady=(0, 6))
        self.scan_progressbar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            length=200,
            maximum=100,
        )
        self.scan_progressbar.pack(side="left", fill="x", expand=True)
        ttk.Label(progress_frame, textvariable=self.progress_pct_var, width=5, anchor="e").pack(side="left", padx=(4, 0))

        ttk.Label(
            left_col,
            text="The selected folder stays in place until you change it manually.",
            foreground="#666666",
        ).grid(row=13, column=0, columnspan=3, sticky="w", padx=6, pady=(8, 0))

        left_col.columnconfigure(1, weight=1)

        # Right-side panel mirrors the scan dialog workflow that appears during scan.
        dialog_panel = ttk.LabelFrame(right_col, text="Scan Dialog", style="Dialog.TLabelframe")
        dialog_panel.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            dialog_panel,
            text="Windows opens a scan dialog during capture. This panel keeps that workflow visible in-app.",
            style="DialogHeader.TLabel",
            wraplength=320,
            justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        ttk.Label(dialog_panel, textvariable=self.dialog_step_var, style="DialogStep.TLabel").grid(
            row=1, column=0, sticky="w", pady=(0, 6)
        )
        ttk.Progressbar(
            dialog_panel,
            orient="horizontal",
            mode="determinate",
            variable=self.dialog_progress_var,
            maximum=100,
        ).grid(row=2, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(dialog_panel, text="Dialog status:", style="DialogSection.TLabel").grid(
            row=3, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(
            dialog_panel,
            textvariable=self.dialog_status_var,
            relief="solid",
            padding=8,
            wraplength=320,
            justify="left",
        ).grid(row=4, column=0, sticky="ew", pady=(0, 8))

        ttk.Label(dialog_panel, textvariable=self.dialog_next_var, style="DialogMuted.TLabel", wraplength=320, justify="left").grid(
            row=5, column=0, sticky="w", pady=(0, 10)
        )

        dialog_actions = ttk.Frame(dialog_panel)
        dialog_actions.grid(row=6, column=0, sticky="ew", pady=(0, 10))
        self.dialog_scan_button = ttk.Button(dialog_actions, text="Scan", command=self.scan_document)
        self.dialog_scan_button.grid(row=0, column=0, padx=(0, 8))
        self.dialog_cancel_button = ttk.Button(dialog_actions, text="Cancel", command=self.cancel_dialog_action)
        self.dialog_cancel_button.grid(row=0, column=1)
        self.dialog_cancel_button.state(["disabled"])

        ttk.Separator(dialog_panel, orient="horizontal").grid(row=7, column=0, sticky="ew", pady=(0, 10))

        preview = ttk.LabelFrame(dialog_panel, text="Dialog Preview", padding=8)
        preview.grid(row=8, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(preview, text="Source:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(preview, textvariable=self.model_var, width=28).grid(row=0, column=1, sticky="w", pady=2)

        ttk.Label(preview, text="Color mode:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(preview, textvariable=self.scan_quality_var).grid(row=1, column=1, sticky="w", pady=2)

        ttk.Label(preview, text="Output type:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(preview, textvariable=self.format_var).grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(preview, text="Resolution:").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(preview, textvariable=self.dpi_var).grid(row=3, column=1, sticky="w", pady=2)

        preview.columnconfigure(1, weight=1)

        ttk.Label(dialog_panel, text="Selected scanner:", style="DialogSection.TLabel").grid(
            row=9, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(dialog_panel, textvariable=self.scanner_var, wraplength=320, justify="left").grid(
            row=10, column=0, sticky="w", pady=(0, 8)
        )

        ttk.Label(dialog_panel, text="Selected output:", style="DialogSection.TLabel").grid(
            row=11, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(dialog_panel, textvariable=self.format_var).grid(row=12, column=0, sticky="w", pady=(0, 8))

        ttk.Label(dialog_panel, text="Scan quality:", style="DialogSection.TLabel").grid(
            row=13, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(dialog_panel, textvariable=self.scan_quality_var).grid(row=14, column=0, sticky="w")

        dialog_panel.columnconfigure(0, weight=1)

        preview_panel = ttk.LabelFrame(preview_col, text="Preview", style="Dialog.TLabelframe")
        preview_panel.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            preview_panel,
            text="Use Scan Preview to capture a single page and review framing before saving.",
            style="DialogHeader.TLabel",
            wraplength=320,
            justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.preview_label = ttk.Label(
            preview_panel,
            text="Preview image will appear here",
            relief="solid",
            anchor="center",
            width=44,
            padding=8,
        )
        self.preview_label.grid(row=1, column=0, sticky="nsew")

        ttk.Label(preview_panel, textvariable=self.preview_status_var, style="DialogMuted.TLabel", wraplength=320, justify="left").grid(
            row=2, column=0, sticky="w", pady=(10, 0)
        )

        preview_panel.columnconfigure(0, weight=1)
        preview_panel.rowconfigure(1, weight=1)

        right_col.columnconfigure(0, weight=1)
        preview_col.columnconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1, uniform="maincols")
        outer.columnconfigure(1, weight=1, uniform="maincols")
        outer.columnconfigure(2, weight=1, uniform="maincols")
        outer.rowconfigure(0, weight=1)

    def update_dialog_stage(self, stage: str):
        stages = {
            "ready": (
                "Step 1 of 4: Ready to scan",
                "Scan dialog idle. Click Scan Document to start.",
                "Next action: Enter a file name and click Scan Document.",
                0,
            ),
            "opening": (
                "Step 2 of 4: Opening scanner dialog",
                "Opening scanner dialog...",
                "Next action: Wait for the Windows scan dialog to appear.",
                35,
            ),
            "scanning": (
                "Step 3 of 4: Scanning in progress",
                "Scanner dialog opened. Complete the scan in the Windows dialog window.",
                "Next action: Confirm scan settings and complete capture.",
                70,
            ),
            "saved": (
                "Step 4 of 4: Scan complete",
                self.dialog_status_var.get(),
                "Next action: Enter a new file name for the next scan.",
                100,
            ),
            "error": (
                "Step 4 of 4: Scan failed",
                self.dialog_status_var.get(),
                "Next action: Fix the issue and click Scan Document again.",
                100,
            ),
        }

        step, status, next_action, progress = stages.get(stage, stages["ready"])
        self.dialog_step_var.set(step)
        self.dialog_status_var.set(status)
        self.dialog_next_var.set(next_action)
        self.dialog_progress_var.set(progress)

    def set_scanning_state(self, scanning: bool, message: str | None = None):
        self.is_scanning = scanning
        if scanning:
            self._scan_cancel_requested = False
        else:
            self._active_batch_source = None

        for button in [self.scan_button, self.preview_button, self.clear_button, self.refresh_button, self.browse_button, self.close_button]:
            if scanning:
                button.state(["disabled"])
            else:
                button.state(["!disabled"])

        if scanning:
            self.dialog_scan_button.state(["disabled"])
            self.dialog_cancel_button.state(["!disabled"])
        else:
            self.dialog_scan_button.state(["!disabled"])
            self.dialog_cancel_button.state(["disabled"])

        if message is not None:
            self.status_var.set(message)
            self.dialog_status_var.set(message)

        if scanning:
            self.update_dialog_stage("scanning")
            self.scan_progressbar.configure(mode="determinate")
            self.scan_progressbar["value"] = 0
            self.progress_pct_var.set("0%")
        else:
            self._reset_progress()
            if message is None:
                self.update_dialog_stage("ready")

        self.update()

    # ------------------------------------------------------------------
    # Progress helpers
    # ------------------------------------------------------------------

    def _estimate_scan_duration(self) -> float:
        """Return rough estimated scan seconds based on selected DPI."""
        dpi = self.get_selected_dpi()
        return max(3.0, min(90.0, 2.0 + dpi / 20.0))

    def _start_progress_ticks(self, estimated_seconds: float) -> None:
        self._progress_start = time.monotonic()
        self._progress_est = max(estimated_seconds, 1.0)
        self._tick_progress()

    def _tick_progress(self) -> None:
        elapsed = time.monotonic() - self._progress_start
        ratio = elapsed / self._progress_est
        # Linear ramp to 85 % over the estimated time, then slowly crawl toward 94 %
        if ratio <= 1.0:
            pct = ratio * 85.0
        else:
            extra = ratio - 1.0
            pct = 85.0 + min(extra * 5.0, 9.0)
        pct = min(pct, 94.0)
        self.scan_progressbar["value"] = pct
        self.progress_pct_var.set(f"{int(pct)}%")
        self._progress_timer_id = self.after(200, self._tick_progress)

    def _complete_progress(self) -> None:
        """Snap the bar to 100 % and cancel the tick timer."""
        if self._progress_timer_id:
            self.after_cancel(self._progress_timer_id)
            self._progress_timer_id = None
        self.scan_progressbar["value"] = 100
        self.progress_pct_var.set("100%")
        self.update_idletasks()

    def _reset_progress(self) -> None:
        """Cancel the tick timer and clear the bar."""
        if self._progress_timer_id:
            self.after_cancel(self._progress_timer_id)
            self._progress_timer_id = None
        self.scan_progressbar["value"] = 0
        self.progress_pct_var.set("")

    # ------------------------------------------------------------------
    # Background WIA acquisition thread
    # ------------------------------------------------------------------

    def _apply_scan_source_to_device(self, device: Any, source: str) -> None:
        desired_flag = WIA_DPS_FLATBED
        if source in {"ADF", "ADF_LOCKED"}:
            desired_flag = WIA_DPS_FEEDER
        elif source == "Auto":
            adf_supported, adf_ready = detect_adf_capability(device)
            # Auto should only force feeder when pages are actually available.
            if adf_supported and adf_ready:
                desired_flag = WIA_DPS_FEEDER

        try:
            for prop in getattr(device, "Properties", []):
                try:
                    if int(getattr(prop, "PropertyID", -1)) == WIA_DOCUMENT_HANDLING_SELECT_PROPERTY:
                        current_value = int(prop.Value)
                        if desired_flag == WIA_DPS_FEEDER:
                            # Force feeder mode by setting FEEDER and clearing FLATBED.
                            prop.Value = (current_value | WIA_DPS_FEEDER) & ~WIA_DPS_FLATBED
                        else:
                            prop.Value = (current_value | WIA_DPS_FLATBED) & ~WIA_DPS_FEEDER
                        break
                except Exception:
                    continue
        except Exception:
            pass

    def _select_wia_item(self, device: Any, source: str) -> Any | None:
        try:
            item_count = int(device.Items.Count)
        except Exception:
            return None

        if item_count < 1:
            return None

        preferred_indices = list(range(1, item_count + 1))
        if source in {"ADF", "ADF_LOCKED"}:
            preferred_indices = [2, 1] + [index for index in preferred_indices if index not in {1, 2}]
        elif source == "Auto":
            adf_supported, adf_ready = detect_adf_capability(device)
            if adf_supported and adf_ready:
                preferred_indices = [2, 1] + [index for index in preferred_indices if index not in {1, 2}]
            else:
                preferred_indices = [1, 2] + [index for index in preferred_indices if index not in {1, 2}]

        for index in preferred_indices:
            try:
                return device.Items[index]
            except Exception:
                continue

        return None

    def _should_use_feeder(self, device: Any, source: str) -> bool:
        if source in {"ADF", "ADF_LOCKED"}:
            return True
        if source == "Auto":
            adf_supported, adf_ready = detect_adf_capability(device)
            return adf_supported and adf_ready
        return False

    def _resolve_batch_source(self, total_pages: int) -> str:
        selected_source = self.scan_source_var.get()
        if selected_source == "ADF":
            return "ADF_LOCKED"
        if selected_source == "Auto" and total_pages > 1 and "ADF" in self._source_options:
            return "ADF_LOCKED"
        return selected_source

    def _get_page_transition_delay_ms(self) -> int:
        source = self.scan_source_var.get()
        adf_delay_ms = self.get_selected_adf_delay_seconds() * 1000
        if source == "ADF":
            return adf_delay_ms
        if source == "Auto":
            return max(1000, int(adf_delay_ms * 0.7))
        return 250

    def _wia_acquire_thread(
        self,
        result_q: queue.Queue[ScanResult],
        requested_format: str,
        dpi: int,
        quality: str,
        source: str,
    ) -> None:
        """
        Run entirely in a background thread.
        Acquires one image via WIA and puts the result into *result_q* as
        a 2-tuple:  ("ok", temp_path_str) | ("cancelled", None) | ("error", exc)
        """
        quality_map = {
            "Color": WIA_INTENT_COLOR,
            "Grayscale": WIA_INTENT_GRAYSCALE,
            "BlackWhite": WIA_INTENT_BLACKWHITE,
            "Black & White": WIA_INTENT_BLACKWHITE,
        }
        temp_name: str | None = None
        result: ScanResult = ("error", RuntimeError("Scan failed."))
        try:
            if self._scan_cancel_requested:
                result = ("cancelled", None)
                return

            if pythoncom is None or win32com is None:
                result = ("error", RuntimeError("Scanning components are not installed."))
                return

            pythoncom_module = cast(Any, pythoncom)
            win32_client = cast(Any, win32com.client)
            pythoncom_module.CoInitialize()
            attempts = 6
            for attempt in range(1, attempts + 1):
                if self._scan_cancel_requested:
                    result = ("cancelled", None)
                    break

                image = None
                device: Any | None = None
                try:
                    manager = win32_client.Dispatch("WIA.DeviceManager")
                    for info in manager.DeviceInfos:
                        if int(info.Type) != SCANNER_DEVICE_TYPE:
                            continue
                        device = info.Connect()
                        self._apply_scan_source_to_device(device, source)
                        item = self._select_wia_item(device, source)
                        if item is None:
                            continue
                        try:
                            for prop in item.Properties:
                                pid = int(prop.PropertyID)
                                if pid in {WIA_HORIZONTAL_DPI_PROPERTY, WIA_VERTICAL_DPI_PROPERTY}:
                                    prop.Value = dpi
                                elif pid == WIA_CURRENT_INTENT_PROPERTY and quality in quality_map:
                                    prop.Value = quality_map[quality]
                        except Exception:
                            pass
                        image = item.Transfer(requested_format)
                        break

                    if image is None:
                        raise RuntimeError("No scanner item found.")
                except Exception as direct_exc:
                    if self._scan_cancel_requested:
                        result = ("cancelled", None)
                        break

                    is_busy = is_busy_error(direct_exc)
                    feeder_empty = is_adf_empty_error(direct_exc)
                    feeder_requested = source == "ADF"
                    feeder_mode = feeder_requested
                    if device is not None:
                        feeder_mode = self._should_use_feeder(device, source)

                    # ADF drivers can throw transient transfer faults mid-feed.
                    # Retry feeder scans even when the driver does not classify the error as "busy".
                    should_retry = attempt < attempts and (is_busy or (feeder_mode and not feeder_empty))

                    if should_retry:
                        wait_seconds = 1.0 * attempt if is_busy else 1.5 * attempt
                        time.sleep(wait_seconds)
                        continue

                    if feeder_mode:
                        raise

                    dialog = win32_client.Dispatch("WIA.CommonDialog")
                    image = dialog.ShowAcquireImage(
                        SCANNER_DEVICE_TYPE, 0, 0, requested_format, False, True, False,
                    )

                if self._scan_cancel_requested:
                    result = ("cancelled", None)
                    break

                if image is None:
                    result = ("cancelled", None)
                    break

                temp_name = make_wia_safe_temp_path(".png")
                image.SaveFile(temp_name)

                if self._scan_cancel_requested:
                    if os.path.exists(temp_name):
                        try:
                            os.remove(temp_name)
                        except Exception:
                            pass
                    result = ("cancelled", None)
                    break

                result = ("ok", temp_name)
                temp_name = None
                break

            if result[0] == "error" and isinstance(result[1], Exception) and str(result[1]) == "Scan failed.":
                result = ("error", RuntimeError("Unable to communicate with the scanner."))
        except Exception as exc:
            if self._scan_cancel_requested:
                result = ("cancelled", None)
            else:
                result = ("error", exc)
            if temp_name and os.path.exists(temp_name):
                try:
                    os.remove(temp_name)
                except Exception:
                    pass
        finally:
            try:
                if pythoncom is not None:
                    cast(Any, pythoncom).CoUninitialize()
            except Exception:
                pass
            result_q.put(result)

    # ------------------------------------------------------------------
    # Threaded image scan helpers
    # ------------------------------------------------------------------

    def _start_threaded_image_scan(self, folder: "Path", filename: str, save_type: str) -> None:
        self._start_threaded_image_batch_scan(folder, filename, save_type, 1)

    def _start_threaded_image_batch_scan(self, folder: "Path", filename: str, save_type: str, total_pages: int) -> None:
        saved_paths: list[Path] = []
        self._start_threaded_image_page_scan(folder, filename, save_type, 1, total_pages, saved_paths)

    def _start_threaded_image_page_scan(
        self,
        folder: "Path",
        filename: str,
        save_type: str,
        page_num: int,
        total_pages: int,
        saved_paths: list[Path],
        empty_retry_count: int = 0,
    ) -> None:
        page_num = max(page_num, len(saved_paths) + 1)
        extension = "jpg" if save_type == "JPG" else save_type.lower()
        page_suffix = f"_p{page_num}" if total_pages > 1 else ""
        final_path = get_unique_path(folder / f"{filename}{page_suffix}.{extension}")
        requested_format = WIA_FORMAT_JPG if save_type == "JPG" else WIA_FORMAT_PNG
        dpi_snap = self.get_selected_dpi()
        quality_snap = self.scan_quality_var.get()
        source_snap = self._active_batch_source or self.scan_source_var.get()
        result_q: queue.Queue[ScanResult] = queue.Queue()
        self.status_var.set(f"Scanning page {page_num} of {total_pages}...")
        self.dialog_status_var.set(f"Scanning page {page_num} of {total_pages}...")
        self._start_progress_ticks(self._estimate_scan_duration())
        t = threading.Thread(
            target=self._wia_acquire_thread,
            args=(result_q, requested_format, dpi_snap, quality_snap, source_snap),
            daemon=True,
        )
        t.start()
        self.after(
            100,
            lambda: self._poll_image_page_save(
                result_q,
                final_path,
                save_type,
                filename,
                page_num,
                total_pages,
                saved_paths,
                folder,
                empty_retry_count,
            ),
        )

    def _poll_image_page_save(
        self,
        result_q: queue.Queue[ScanResult],
        final_path: "Path",
        save_type: str,
        filename: str,
        page_num: int,
        total_pages: int,
        saved_paths: list[Path],
        folder: Path,
        empty_retry_count: int,
    ) -> None:
        try:
            status, value = result_q.get_nowait()
        except queue.Empty:
            self.after(
                100,
                lambda: self._poll_image_page_save(
                    result_q,
                    final_path,
                    save_type,
                    filename,
                    page_num,
                    total_pages,
                    saved_paths,
                    folder,
                    empty_retry_count,
                ),
            )
            return

        if self._scan_cancel_requested or status == "cancelled":
            if status == "ok" and isinstance(value, str) and os.path.exists(value):
                try:
                    os.remove(value)
                except Exception:
                    pass
            self.status_var.set("Scan cancelled.")
            self.dialog_status_var.set("Scan cancelled from dialog.")
            self.dialog_next_var.set("Next action: Enter a file name and click Scan Document.")
            self.update_dialog_stage("ready")
            if self.ready_var.get() in {"Scanning...", "Busy"}:
                self.ready_var.set("Ready")
            self._scan_cancel_requested = False
            self.set_scanning_state(False)
            return

        self._complete_progress()
        if status == "error":
            exc = value if isinstance(value, Exception) else RuntimeError("Scan failed.")
            active_source = self._active_batch_source or self.scan_source_var.get()
            if is_adf_empty_error(exc) and active_source in {"ADF", "ADF_LOCKED", "Auto"}:
                scanned_pages = len(saved_paths)
                remaining_pages = total_pages - scanned_pages
                if remaining_pages > 0 and empty_retry_count < self._MAX_ADF_EMPTY_RETRIES:
                    retry_num = empty_retry_count + 1
                    self.status_var.set(
                        f"Waiting for page {page_num} to feed from ADF... (retry {retry_num}/{self._MAX_ADF_EMPTY_RETRIES})"
                    )
                    self.dialog_status_var.set(
                        f"ADF reported empty while waiting for page {page_num}. Retrying automatically ({retry_num}/{self._MAX_ADF_EMPTY_RETRIES})..."
                    )
                    delay_ms = max(2500, self._get_page_transition_delay_ms() + (retry_num * 750))
                    self.after(
                        delay_ms,
                        lambda: self._start_threaded_image_page_scan(
                            folder,
                            filename,
                            save_type,
                            page_num,
                            total_pages,
                            saved_paths,
                            retry_num,
                        ),
                    )
                    return

                if remaining_pages > 0:
                    continue_scan = messagebox.askyesno(
                        "Scan More Pages?",
                        "The feeder is empty.\n\n"
                        "Click Yes to load more pages and continue scanning.\n"
                        "Click No to save the pages already scanned.",
                    )
                    if continue_scan:
                        self.status_var.set(f"Waiting for more pages. Continuing with page {page_num}...")
                        self.dialog_status_var.set(f"Continuing scan for page {page_num} after refill.")
                        self.after(
                            max(1500, self._get_page_transition_delay_ms()),
                            lambda: self._start_threaded_image_page_scan(
                                folder,
                                filename,
                                save_type,
                                page_num,
                                total_pages,
                                saved_paths,
                                0,
                            ),
                        )
                        return

                if saved_paths:
                    self.filename_var.set("")
                    self.status_var.set(f"ADF emptied after {scanned_pages} page(s). Saved scanned pages.")
                    self.dialog_status_var.set(f"ADF emptied. Saved {scanned_pages} page(s) successfully.")
                    self.update_dialog_stage("saved")
                    preview_list = "\n".join(str(path.name) for path in saved_paths[:10])
                    more = f"\n...and {scanned_pages - 10} more" if scanned_pages > 10 else ""
                    messagebox.showinfo(
                        "ADF Empty",
                        f"No more pages were found in the feeder.\n\nSaved {scanned_pages} page(s) to:\n{folder}\n\n{preview_list}{more}",
                    )
                    if self.ready_var.get() in {"Scanning...", "Busy"}:
                        self.ready_var.set("Ready")
                    self.set_scanning_state(False)
                    return

                self.status_var.set("ADF is empty.")
                self.dialog_status_var.set("No pages were found in the feeder.")
                self.update_dialog_stage("ready")
                messagebox.showwarning(
                    "ADF Empty",
                    "The feeder is empty.\n\nLoad paper in the ADF and scan again.",
                )
                if self.ready_var.get() in {"Scanning...", "Busy"}:
                    self.ready_var.set("Ready")
                self.set_scanning_state(False)
                return

            self.status_var.set("Scan failed.")
            self.dialog_status_var.set("Scan failed. See error dialog for details.")
            self.update_dialog_stage("error")
            if is_busy_error(exc):
                messagebox.showerror(
                    "Scanner Busy",
                    "The scanner is busy.\n\nClose any other scan app or scanner window, wait a few seconds, and try again.",
                )
            elif self.scan_source_var.get() == "ADF":
                messagebox.showerror(
                    "ADF Scan Error",
                    "Unable to scan from ADF.\n\n"
                    "Please confirm paper is loaded in the feeder and the scanner driver is set to feeder mode, then try again.",
                )
            else:
                messagebox.showerror("Scan Error", f"Unable to complete the scan.\n\n{exc}")
        elif status == "ok":
            if not isinstance(value, str) or Image is None:
                self.status_var.set("Failed to save scan.")
                messagebox.showerror("Save Error", "Scan succeeded but no image data was returned.")
                if self.ready_var.get() in {"Scanning...", "Busy"}:
                    self.ready_var.set("Ready")
                self.set_scanning_state(False)
                return

            temp_path = value
            try:
                with Image.open(temp_path) as img:
                    pil_img = img.convert("RGB").copy()
                if save_type == "JPG":
                    pil_img.save(str(final_path), "JPEG", quality=95)
                else:
                    pil_img.save(str(final_path), "PNG")
                saved_paths.append(final_path)
            except Exception as exc:
                self.status_var.set("Failed to save scan.")
                messagebox.showerror("Save Error", f"Scan succeeded but could not be saved.\n\n{exc}")
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

            if len(saved_paths) == page_num and page_num < total_pages:
                next_page = len(saved_paths) + 1
                delay_ms = self._get_page_transition_delay_ms()
                self.after(
                    delay_ms,
                    lambda: self._start_threaded_image_page_scan(
                        folder,
                        filename,
                        save_type,
                        next_page,
                        total_pages,
                        saved_paths,
                        0,
                    ),
                )
                return

            if len(saved_paths) == total_pages:
                self.filename_var.set("")
                if total_pages == 1:
                    self.status_var.set(f"Saved successfully: {saved_paths[0]}")
                    self.dialog_status_var.set(f"Scan saved successfully: {saved_paths[0].name}")
                    messagebox.showinfo("Success", f"Document saved to:\n{saved_paths[0]}")
                else:
                    self.status_var.set(f"Saved {total_pages} pages to {folder}")
                    self.dialog_status_var.set(f"Saved {total_pages} pages successfully.")
                    preview_list = "\n".join(str(path.name) for path in saved_paths[:10])
                    more = f"\n...and {total_pages - 10} more" if total_pages > 10 else ""
                    messagebox.showinfo(
                        "Success",
                        f"Saved {total_pages} pages to:\n{folder}\n\n{preview_list}{more}",
                    )
                self.update_dialog_stage("saved")
                self.after(50, lambda: self.filename_entry.focus_set())
        if self.ready_var.get() in {"Scanning...", "Busy"}:
            self.ready_var.set("Ready")
        self.set_scanning_state(False)

    def _start_threaded_preview(self) -> None:
        dpi_snap = self.get_selected_dpi()
        quality_snap = self.scan_quality_var.get()
        source_snap = self._active_batch_source or self.scan_source_var.get()
        result_q: queue.Queue[ScanResult] = queue.Queue()
        self._start_progress_ticks(self._estimate_scan_duration())
        t = threading.Thread(
            target=self._wia_acquire_thread,
            args=(result_q, WIA_FORMAT_PNG, dpi_snap, quality_snap, source_snap),
            daemon=True,
        )
        t.start()
        self.after(100, lambda: self._poll_preview_result(result_q))

    def _poll_preview_result(self, result_q: queue.Queue[ScanResult]) -> None:
        try:
            status, value = result_q.get_nowait()
        except queue.Empty:
            self.after(100, lambda: self._poll_preview_result(result_q))
            return

        if self._scan_cancel_requested or status == "cancelled":
            if status == "ok" and isinstance(value, str) and os.path.exists(value):
                try:
                    os.remove(value)
                except Exception:
                    pass
            self.preview_status_var.set("Preview canceled.")
            self.status_var.set("Preview canceled.")
            self.dialog_status_var.set("Preview canceled from dialog.")
            self.dialog_next_var.set("Next action: Enter a file name and click Scan Document.")
            self.update_dialog_stage("ready")
            if self.ready_var.get() in {"Previewing...", "Scanning...", "Busy"}:
                self.ready_var.set("Ready")
            self._scan_cancel_requested = False
            self.set_scanning_state(False)
            return

        self._complete_progress()
        if status == "error":
            exc = value if isinstance(value, Exception) else RuntimeError("Preview failed.")
            self.preview_status_var.set("Preview failed.")
            self.status_var.set("Preview failed.")
            self.dialog_status_var.set("Preview failed. See error dialog for details.")
            self.update_dialog_stage("error")
            if is_busy_error(exc):
                messagebox.showerror(
                    "Scanner Busy",
                    "The scanner is busy.\n\nPlease close other scan windows, wait a few seconds, and try preview again.",
                )
            else:
                messagebox.showerror("Preview Error", f"Unable to capture preview.\n\n{exc}")
        elif status == "ok":
            if not isinstance(value, str) or Image is None or ImageTk is None or self.preview_label is None:
                self.preview_status_var.set("Preview failed.")
                messagebox.showerror("Preview Error", "Preview data was returned but could not be displayed.")
                if self.ready_var.get() in {"Previewing...", "Scanning...", "Busy"}:
                    self.ready_var.set("Ready")
                self.set_scanning_state(False)
                return

            temp_path = value
            try:
                with Image.open(temp_path) as img:
                    preview = img.convert("RGB")
                    preview.thumbnail((360, 460))
                    self.preview_photo = ImageTk.PhotoImage(preview)
                    self.preview_label.configure(image=self.preview_photo, text="")
                    self.preview_status_var.set(f"Preview captured at {self.get_selected_dpi()} DPI.")
                self.status_var.set("Preview captured. Adjust settings or file name, then scan.")
                self.dialog_status_var.set("Preview captured. Ready to scan and save.")
                self.update_dialog_stage("ready")
            except Exception as exc:
                self.preview_status_var.set("Preview failed.")
                messagebox.showerror("Preview Error", f"Preview captured but could not be displayed.\n\n{exc}")
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass
        if self.ready_var.get() in {"Previewing...", "Scanning...", "Busy"}:
            self.ready_var.set("Ready")
        self.set_scanning_state(False)

    def _start_threaded_pdf_scan(self, folder: Path, filename: str, total_pages: int) -> None:
        final_path = get_unique_path(folder / f"{filename}.pdf")
        if Image is None:
            raise RuntimeError("Pillow is required to build PDFs.")

        pages: list[Any] = []
        temp_files: list[str] = []
        self._start_threaded_pdf_page_scan(final_path, pages, temp_files, 1, total_pages, 0)

    def _start_threaded_pdf_page_scan(
        self,
        final_path: Path,
        pages: list[Any],
        temp_files: list[str],
        page_num: int,
        total_pages: int,
        empty_retry_count: int,
    ) -> None:
        page_num = max(page_num, len(pages) + 1)
        dpi_snap = self.get_selected_dpi()
        quality_snap = self.scan_quality_var.get()
        source_snap = self._active_batch_source or self.scan_source_var.get()
        result_q: queue.Queue[ScanResult] = queue.Queue()
        self.status_var.set(f"Scanning page {page_num} of {total_pages}...")
        self.dialog_status_var.set(f"Scanning page {page_num} of {total_pages} for PDF...")
        self._start_progress_ticks(self._estimate_scan_duration())
        worker = threading.Thread(
            target=self._wia_acquire_thread,
            args=(result_q, WIA_FORMAT_PNG, dpi_snap, quality_snap, source_snap),
            daemon=True,
        )
        worker.start()
        self.after(100, lambda: self._poll_pdf_page_result(result_q, final_path, pages, temp_files, page_num, total_pages, empty_retry_count))

    def _poll_pdf_page_result(
        self,
        result_q: queue.Queue[ScanResult],
        final_path: Path,
        pages: list[Any],
        temp_files: list[str],
        page_num: int,
        total_pages: int,
        empty_retry_count: int,
    ) -> None:
        try:
            status, value = result_q.get_nowait()
        except queue.Empty:
            self.after(100, lambda: self._poll_pdf_page_result(result_q, final_path, pages, temp_files, page_num, total_pages, empty_retry_count))
            return

        if self._scan_cancel_requested or status == "cancelled":
            self._cleanup_pdf_temp_files(temp_files)
            pages.clear()
            self.status_var.set("Scan cancelled.")
            self.dialog_status_var.set("PDF scan cancelled from dialog.")
            self.dialog_next_var.set("Next action: Enter a file name and click Scan Document.")
            self.update_dialog_stage("ready")
            if self.ready_var.get() in {"Scanning...", "Busy"}:
                self.ready_var.set("Ready")
            self._scan_cancel_requested = False
            self.set_scanning_state(False)
            return

        self._complete_progress()

        if status == "error":
            exc = value if isinstance(value, Exception) else RuntimeError("PDF scan failed.")
            active_source = self._active_batch_source or self.scan_source_var.get()
            if is_adf_empty_error(exc) and active_source in {"ADF", "ADF_LOCKED", "Auto"}:
                scanned_pages = len(pages)
                remaining_pages = total_pages - scanned_pages
                if remaining_pages > 0 and empty_retry_count < self._MAX_ADF_EMPTY_RETRIES:
                    retry_num = empty_retry_count + 1
                    self.status_var.set(
                        f"Waiting for page {page_num} to feed from ADF... (retry {retry_num}/{self._MAX_ADF_EMPTY_RETRIES})"
                    )
                    self.dialog_status_var.set(
                        f"ADF reported empty while waiting for PDF page {page_num}. Retrying automatically ({retry_num}/{self._MAX_ADF_EMPTY_RETRIES})..."
                    )
                    delay_ms = max(2500, self._get_page_transition_delay_ms() + (retry_num * 750))
                    self.after(
                        delay_ms,
                        lambda: self._start_threaded_pdf_page_scan(
                            final_path,
                            pages,
                            temp_files,
                            page_num,
                            total_pages,
                            retry_num,
                        ),
                    )
                    return

                if remaining_pages > 0:
                    continue_scan = messagebox.askyesno(
                        "Scan More Pages?",
                        "The feeder is empty.\n\n"
                        "Click Yes to load more pages and continue scanning.\n"
                        "Click No to save the pages already scanned.",
                    )
                    if continue_scan:
                        self.status_var.set(f"Waiting for more pages. Continuing with page {page_num}...")
                        self.dialog_status_var.set(f"Continuing PDF scan for page {page_num} after refill.")
                        self.after(
                            max(1500, self._get_page_transition_delay_ms()),
                            lambda: self._start_threaded_pdf_page_scan(
                                final_path,
                                pages,
                                temp_files,
                                page_num,
                                total_pages,
                                0,
                            ),
                        )
                        return

                if pages:
                    self.status_var.set("ADF emptied. Saving scanned PDF pages...")
                    self.dialog_status_var.set("ADF emptied. Finalizing PDF with scanned pages.")
                    self._finish_pdf_scan(final_path, pages, temp_files)
                    return

                self._cleanup_pdf_temp_files(temp_files)
                self.status_var.set("ADF is empty.")
                self.dialog_status_var.set("No pages were found in the feeder.")
                self.update_dialog_stage("ready")
                messagebox.showwarning(
                    "ADF Empty",
                    "The feeder is empty.\n\nLoad paper in the ADF and scan again.",
                )
                if self.ready_var.get() in {"Scanning...", "Busy"}:
                    self.ready_var.set("Ready")
                self.set_scanning_state(False)
                return

            self._cleanup_pdf_temp_files(temp_files)
            self.status_var.set("Scan failed.")
            self.dialog_status_var.set("Scan failed. See error dialog for details.")
            self.update_dialog_stage("error")
            if is_busy_error(exc):
                messagebox.showerror(
                    "Scanner Busy",
                    "The scanner is busy.\n\nPlease close Epson Scan or any other scan window, wait a few seconds, and try again.",
                )
            elif self.scan_source_var.get() == "ADF":
                messagebox.showerror(
                    "ADF Scan Error",
                    "Unable to scan from ADF.\n\n"
                    "Please confirm paper is loaded in the feeder and the scanner driver is set to feeder mode, then try again.",
                )
            else:
                messagebox.showerror("Scan Error", f"Unable to complete the scan.\n\n{exc}")
            if self.ready_var.get() in {"Scanning...", "Busy"}:
                self.ready_var.set("Ready")
            self.set_scanning_state(False)
            return

        if not isinstance(value, str) or Image is None:
            self._cleanup_pdf_temp_files(temp_files)
            self.status_var.set("Scan failed.")
            self.dialog_status_var.set("Scan failed. No image data was returned.")
            self.update_dialog_stage("error")
            if self.ready_var.get() in {"Scanning...", "Busy"}:
                self.ready_var.set("Ready")
            self.set_scanning_state(False)
            return

        temp_path = value
        try:
            with Image.open(temp_path) as img:
                pages.append(img.convert("RGB").copy())
            temp_files.append(temp_path)
        except Exception as exc:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            self._cleanup_pdf_temp_files(temp_files)
            self.status_var.set("Scan failed.")
            self.dialog_status_var.set("Unable to prepare scanned page for PDF.")
            self.update_dialog_stage("error")
            messagebox.showerror("PDF Error", f"Scanned page could not be prepared for the PDF.\n\n{exc}")
            if self.ready_var.get() in {"Scanning...", "Busy"}:
                self.ready_var.set("Ready")
            self.set_scanning_state(False)
            return

        if page_num < total_pages:
            next_page = len(pages) + 1
            delay_ms = self._get_page_transition_delay_ms()
            self.after(
                delay_ms,
                lambda: self._start_threaded_pdf_page_scan(final_path, pages, temp_files, next_page, total_pages, 0),
            )
            return

        self._finish_pdf_scan(final_path, pages, temp_files)

    def _cleanup_pdf_temp_files(self, temp_files: list[str]) -> None:
        for temp_name in temp_files:
            if os.path.exists(temp_name):
                try:
                    os.remove(temp_name)
                except Exception:
                    pass

    def _finish_pdf_scan(self, final_path: Path, pages: list[Any], temp_files: list[str]) -> None:
        if not pages:
            self._cleanup_pdf_temp_files(temp_files)
            self.status_var.set("Scan cancelled.")
            self.dialog_status_var.set("PDF scan cancelled from dialog.")
            self.update_dialog_stage("ready")
            if self.ready_var.get() in {"Scanning...", "Busy"}:
                self.ready_var.set("Ready")
            self.set_scanning_state(False)
            return

        try:
            self.scan_progressbar["value"] = 90
            self.progress_pct_var.set("90%")
            self.status_var.set("Saving PDF...")
            self.dialog_status_var.set(f"Saving PDF: {final_path.name}")
            self.update_idletasks()
            pages[0].save(final_path, save_all=True, append_images=pages[1:])
            self._complete_progress()
            self.filename_var.set("")
            self.status_var.set(f"Saved successfully: {final_path}")
            self.dialog_status_var.set(f"PDF saved successfully: {final_path.name} ({len(pages)} page(s))")
            self.update_dialog_stage("saved")
            messagebox.showinfo("Success", f"PDF saved to:\n{final_path}")
            self.after(50, lambda: self.filename_entry.focus_set())
        finally:
            self._cleanup_pdf_temp_files(temp_files)

        if self.ready_var.get() in {"Scanning...", "Busy"}:
            self.ready_var.set("Ready")
        self.set_scanning_state(False)

    def get_selected_dpi(self) -> int:
        try:
            return min(1200, max(100, int(self.dpi_var.get())))
        except Exception:
            return 300

    def get_selected_pages(self) -> int:
        try:
            return min(100, max(1, int(self.pages_var.get())))
        except Exception:
            return 1

    def get_selected_adf_delay_seconds(self) -> int:
        try:
            return min(10, max(1, int(self.adf_delay_var.get())))
        except Exception:
            return 4

    def apply_scan_settings(self, item: Any) -> None:
        dpi_value = self.get_selected_dpi()
        quality = self.scan_quality_var.get()
        quality_map = {
            "Color": WIA_INTENT_COLOR,
            "Grayscale": WIA_INTENT_GRAYSCALE,
            "BlackWhite": WIA_INTENT_BLACKWHITE,
        }

        try:
            for prop in item.Properties:
                try:
                    prop_id = int(getattr(prop, "PropertyID", -1))
                    if prop_id in {WIA_HORIZONTAL_DPI_PROPERTY, WIA_VERTICAL_DPI_PROPERTY}:
                        prop.Value = dpi_value
                    elif prop_id == WIA_CURRENT_INTENT_PROPERTY and quality in quality_map:
                        prop.Value = quality_map[quality]
                except Exception:
                    continue
        except Exception:
            # Some scanner drivers lock or ignore these properties; scanning can still continue.
            pass

    def acquire_image_direct(self, requested_format: str) -> Any:
        if win32com is None:
            raise RuntimeError("Scanning components are not installed.")

        win32_client = cast(Any, win32com.client)
        manager = win32_client.Dispatch("WIA.DeviceManager")

        for info in manager.DeviceInfos:
            if int(info.Type) != SCANNER_DEVICE_TYPE:
                continue

            device = info.Connect()
            item = self._select_wia_item(device, self.scan_source_var.get())
            if item is None:
                continue

            self.apply_scan_settings(item)
            return item.Transfer(requested_format)

        raise RuntimeError("No scanner item found.")

    def acquire_image(self, requested_format: str) -> Any:
        attempts = 0

        while attempts < 3:
            attempts += 1
            try:
                if pythoncom is None or win32com is None:
                    raise RuntimeError("Scanning components are not installed.")

                pythoncom_module = cast(Any, pythoncom)
                win32_client = cast(Any, win32com.client)
                pythoncom_module.CoInitialize()
                try:
                    image = self.acquire_image_direct(requested_format)
                    self.dialog_status_var.set("Scanning directly with in-app settings.")
                except Exception:
                    dialog = win32_client.Dispatch("WIA.CommonDialog")
                    image = dialog.ShowAcquireImage(
                        SCANNER_DEVICE_TYPE,
                        0,
                        0,
                        requested_format,
                        attempts > 1,
                        True,
                        False,
                    )
                    self.dialog_status_var.set("Scanner driver opened Windows scan dialog for this scan.")
                
                self.ready_var.set("Ready")
                return image
            except Exception as exc:
                if is_busy_error(exc):
                    self.ready_var.set("Busy")
                    retry = messagebox.askretrycancel(
                        "Scanner Busy",
                        "The scanner is busy.\n\nClose any other scan app or scanner window, wait a moment, then click Retry.",
                    )
                    if retry:
                        self.status_var.set("Retrying scanner...")
                        self.update_idletasks()
                        time.sleep(1.5)
                        continue
                    raise RuntimeError(
                        "The scanner is busy. Close any other scanning app or window, wait a few seconds, and try again."
                    ) from exc
                raise
            finally:
                try:
                    if pythoncom is not None:
                        cast(Any, pythoncom).CoUninitialize()
                except Exception:
                    pass

        raise RuntimeError("Unable to communicate with the scanner.")

    def _show_preview_from_image(self, image: Any) -> None:
        if Image is None or ImageTk is None or self.preview_label is None:
            return

        temp_name = make_wia_safe_temp_path(".png")
        try:
            image.SaveFile(temp_name)
            with Image.open(temp_name) as img:
                preview = img.convert("RGB")
                preview.thumbnail((360, 460))
                self.preview_photo = ImageTk.PhotoImage(preview)
                self.preview_label.configure(image=self.preview_photo, text="")
                self.preview_status_var.set(f"Preview captured at {self.get_selected_dpi()} DPI.")
        finally:
            if os.path.exists(temp_name):
                os.remove(temp_name)

    def scan_preview(self):
        if self.is_scanning:
            self.status_var.set("A scan is already in progress.")
            return

        if win32com is None or pythoncom is None or Image is None or ImageTk is None:
            messagebox.showerror(
                "Missing packages",
                "Please install the required packages:\n\npip install pywin32 pillow",
            )
            return

        if self.scanner_var.get() == "No scanner detected":
            messagebox.showwarning("Scanner not found", "No scanner was detected.")
            return

        self.ready_var.set("Previewing...")
        self.update_dialog_stage("opening")
        self.set_scanning_state(True, "Capturing preview scan...")
        self.after(50, self._do_preview)

    def _do_preview(self):
        # Runs in background thread with live progress
        self._start_threaded_preview()

    def detect_scanner(self):
        scanners = get_connected_scanners()
        if scanners:
            first = scanners[0]
            display = [
                f"{s['name']} [{s['connection']}]"
                + (" [ADF]" if s["adf_supported"] else "")
                for s in scanners
            ]
            self.scanner_var.set(", ".join(display))
            self.model_var.set(first["name"])
            self.connection_var.set(first["connection"])
            self.device_id_var.set(first["device_id"] or "N/A")
            self.ready_var.set(first["status"])
            if first["adf_supported"]:
                state = "Ready" if first["adf_ready"] else "Available (no pages loaded)"
                self.adf_var.set(state)
                self._source_options = ["Auto", "Flatbed", "ADF"]
                if self.scan_source_var.get() not in self._source_options:
                    self.scan_source_var.set("ADF")
            else:
                self.adf_var.set("Not available")
                self._source_options = ["Auto", "Flatbed"]
                if self.scan_source_var.get() not in self._source_options:
                    self.scan_source_var.set("Flatbed")
            self.status_var.set("Scanner detected and ready.")
            self.dialog_status_var.set("Scanner ready. Click Scan Document to open the scan dialog.")
            self.update_dialog_stage("ready")
        else:
            self.scanner_var.set("No scanner detected")
            self.model_var.set("N/A")
            self.connection_var.set("N/A")
            self.device_id_var.set("N/A")
            self.ready_var.set("Not detected")
            self.adf_var.set("N/A")
            self._source_options = ["Auto", "Flatbed"]
            self.scan_source_var.set("Auto")
            self.status_var.set("Connect the scanner and click Refresh.")
            self.dialog_status_var.set("No scanner detected. Connect scanner and click Refresh.")
            self.dialog_step_var.set("Step 1 of 4: Scanner required")
            self.dialog_next_var.set("Next action: Connect scanner, then click Refresh.")

    def cleanup_old_files(self) -> None:
        folder = Path(self.folder_var.get())
        if not folder.exists():
            return

        cutoff = time.time() - (365 * 24 * 60 * 60)
        deleted: list[str] = []

        for f in folder.iterdir():
            try:
                if f.is_file() and f.stat().st_mtime < cutoff:
                    f.unlink()
                    deleted.append(f.name)
            except Exception:
                pass

        if deleted:
            self.status_var.set(f"Cleaned up {len(deleted)} file(s) older than 1 year.")
            messagebox.showinfo(
                "Cleanup Complete",
                f"Deleted {len(deleted)} file(s) older than 1 year from:\n{folder}\n\n"
                + "\n".join(deleted[:20])
                + (f"\n...and {len(deleted) - 20} more" if len(deleted) > 20 else ""),
            )

    def _on_setting_change(self, *_: str) -> None:
        save_settings({
            "folder": self.folder_var.get(),
            "format": self.format_var.get(),
            "dpi": self.get_selected_dpi(),
            "pages": self.get_selected_pages(),
            "adf_delay_seconds": self.get_selected_adf_delay_seconds(),
        })

    def _on_close(self) -> None:
        save_settings({
            "folder": self.folder_var.get(),
            "format": self.format_var.get(),
            "dpi": self.get_selected_dpi(),
            "pages": self.get_selected_pages(),
            "adf_delay_seconds": self.get_selected_adf_delay_seconds(),
        })
        self.destroy()

    def browse_folder(self):
        selected = filedialog.askdirectory(initialdir=self.folder_var.get() or str(Path.home()))
        if selected:
            self.folder_var.set(selected)  # trace fires → save_settings() called automatically

    def cancel_dialog_action(self):
        if self.is_scanning:
            self._scan_cancel_requested = True
            self.dialog_status_var.set("Use Cancel in the Windows scan dialog window to stop capture.")
            self.dialog_next_var.set("Cancel requested. Finish or close the scan dialog to stop the current scan.")
            self.status_var.set("Cancel requested. Finish or close the scan dialog to stop the current scan.")
        else:
            self.dialog_status_var.set("No active scan to cancel.")
            self.update_dialog_stage("ready")

    def clear_name(self):
        self.filename_var.set("")
        self.status_var.set("File name cleared. Ready for next scan.")
        self.dialog_status_var.set("Ready for next scan.")
        self.update_dialog_stage("ready")

    def scan_document(self):
        if self.is_scanning:
            self.status_var.set("A scan is already in progress.")
            self.dialog_status_var.set("A scan is already in progress.")
            return

        if win32com is None or pythoncom is None or Image is None:
            messagebox.showerror(
                "Missing packages",
                "Please install the required packages:\n\npip install pywin32 pillow",
            )
            return

        folder_text = self.folder_var.get().strip()
        filename = sanitize_filename(self.filename_var.get())
        save_type = self.format_var.get().upper()
        pages_to_scan = self.get_selected_pages()

        if pages_to_scan > 1 and save_type in {"PNG", "JPG"}:
            save_type = "PDF"
            self.format_var.set("PDF")
            messagebox.showinfo(
                "Multi-Page Output",
                "Multi-page scans are saved as a single PDF file.",
            )
            self.status_var.set("Switched to PDF so all pages are saved in one file.")

        if self.scanner_var.get() == "No scanner detected":
            messagebox.showwarning("Scanner not found", "No scanner was detected.")
            return

        if not folder_text:
            messagebox.showwarning("Missing folder", "Please choose a destination folder.")
            return

        if not filename:
            messagebox.showwarning("Missing file name", "Please enter a file name before scanning.")
            return

        folder = Path(folder_text)
        folder.mkdir(parents=True, exist_ok=True)

        self.ready_var.set("Scanning...")
        self.update_dialog_stage("opening")
        self.set_scanning_state(True, "Opening scanner...")
        self.after(50, lambda: self._do_scan(folder, filename, save_type, pages_to_scan))

    def _do_scan(self, folder: Path, filename: str, save_type: str, pages_to_scan: int):
        self._active_batch_source = self._resolve_batch_source(pages_to_scan)
        if save_type == "PDF":
            self._start_threaded_pdf_scan(folder, filename, pages_to_scan)
        else:
            # Single image scan — runs in background thread with live progress
            self._start_threaded_image_batch_scan(folder, filename, save_type, pages_to_scan)

    def scan_image(self, folder: Path, filename: str, save_type: str):
        extension = "jpg" if save_type == "JPG" else save_type.lower()
        final_path = get_unique_path(folder / f"{filename}.{extension}")

        requested_format = WIA_FORMAT_JPG if save_type == "JPG" else WIA_FORMAT_PNG
        image = self.acquire_image(requested_format)

        if image is None:
            self.status_var.set("Scan cancelled.")
            self.dialog_status_var.set("Scan cancelled from dialog.")
            self.update_dialog_stage("ready")
            return

        if final_path.exists():
            final_path = get_unique_path(final_path)

        image.SaveFile(str(final_path))
        self.filename_var.set("")
        self.status_var.set(f"Saved successfully: {final_path}")
        self.dialog_status_var.set(f"Scan saved successfully: {final_path.name}")
        self.update_dialog_stage("saved")
        messagebox.showinfo("Success", f"Document saved to:\n{final_path}")
        self.after(50, lambda: self.filename_entry.focus_set())

    def scan_pdf(self, folder: Path, filename: str):
        final_path = get_unique_path(folder / f"{filename}.pdf")
        if Image is None:
            raise RuntimeError("Pillow is required to build PDFs.")

        pages: list[Any] = []
        temp_files: list[str] = []
        page_num = 0

        def _set_pdf_progress(pct: float, msg: str) -> None:
            self.scan_progressbar["value"] = pct
            self.progress_pct_var.set(f"{int(pct)}%")
            self.status_var.set(msg)
            self.update_idletasks()

        while True:
            page_num += 1
            _set_pdf_progress(min(10 + (page_num - 1) * 25, 80), f"Scanning page {page_num}...")
            image = self.acquire_image(WIA_FORMAT_PNG)

            if image is None:
                break

            temp_name = make_wia_safe_temp_path(".png")

            image.SaveFile(temp_name)
            temp_files.append(temp_name)

            with Image.open(temp_name) as img:
                pages.append(img.convert("RGB").copy())

            if not messagebox.askyesno("Scan Another Page", "Would you like to scan another page?"):
                break

        for temp_name in temp_files:
            if os.path.exists(temp_name):
                os.remove(temp_name)

        if not pages:
            self.status_var.set("Scan cancelled.")
            self.dialog_status_var.set("PDF scan cancelled from dialog.")
            self.update_dialog_stage("ready")
            return

        _set_pdf_progress(90, "Saving PDF...")
        pages[0].save(final_path, save_all=True, append_images=pages[1:])
        self._complete_progress()
        self.filename_var.set("")
        self.status_var.set(f"Saved successfully: {final_path}")
        self.dialog_status_var.set(f"PDF saved successfully: {final_path.name}")
        self.update_dialog_stage("saved")
        messagebox.showinfo("Success", f"PDF saved to:\n{final_path}")
        self.after(50, lambda: self.filename_entry.focus_set())


if __name__ == "__main__":
    try:
        create_default_icon("scanner_icon.ico", overwrite=False)
    except Exception:
        # Icon generation is optional and must never block app startup.
        pass
    app = ScannerApp()
    app.mainloop()
