"""WPS discovery and Windows COM export. No install, registry edits or broad kills.

COM availability is probed, not inferred from wps.exe. On macOS/Linux or when the
COM interface is unavailable, use the explicit offline wps_exchange.py route.
"""

from __future__ import annotations
import glob
import os
from pathlib import Path
import re

PROGIDS = ("Kwps.Application", "Wps.Application")


def known_paths(system, env, home):
    paths = []
    if system == "Windows":
        roots = [
            env.get("ProgramFiles", r"C:\Program Files"),
            env.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
            env.get("ProgramW6432", r"C:\Program Files"),
            env.get("LOCALAPPDATA", str(Path(home) / "AppData/Local")),
        ]
        for root in dict.fromkeys(roots):
            for parent in (
                "Kingsoft/WPS Office",
                "Kingsoft/office6",
                "WPS Office",
                "Programs/Kingsoft/WPS Office",
            ):
                base = Path(root) / parent
                paths.append(("wps", str(base / "office6/wps.exe"), "wps-common-install"))
                paths.append(("wps", str(base / "wps.exe"), "wps-common-install"))
                for name in glob.glob(str(base / "*/office6/wps.exe")):
                    paths.append(("wps", name, "wps-versioned-install"))
    elif system == "Darwin":
        for root in (Path("/Applications"), Path(home) / "Applications"):
            for app, exe in (
                ("wpsoffice.app", "wpsoffice"),
                ("WPS Office.app", "wpsoffice"),
                ("WPS Office.app", "wps"),
            ):
                paths.append(
                    ("wps", str(root / app / "Contents/MacOS" / exe), "wps-application-bundle")
                )
    else:
        for path in (
            "/usr/bin/wps",
            "/usr/local/bin/wps",
            "/opt/kingsoft/wps-office/office6/wps",
            "/opt/kingsoft/wps-office/office6/wps.orig",
        ):
            paths.append(("wps", path, "wps-common-install"))
    return paths


def registry_paths():
    if os.name != "nt":
        return [], []
    try:
        import winreg as wr
    except ImportError:
        return [], [{"source": "wps-registry", "status": "unavailable"}]
    found = []
    notes = []
    for hive in (wr.HKEY_CURRENT_USER, wr.HKEY_LOCAL_MACHINE):
        for view in (0, wr.KEY_WOW64_64KEY, wr.KEY_WOW64_32KEY):
            try:
                with wr.OpenKey(
                    hive,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe",
                    0,
                    wr.KEY_READ | view,
                ) as key:
                    value, _ = wr.QueryValueEx(key, "")
                found.append(("wps", str(value).strip('"'), "wps-registry-app-path"))
            except OSError:
                pass
            try:
                with wr.OpenKey(
                    hive, r"SOFTWARE\Kingsoft\Office\6.0\Common", 0, wr.KEY_READ | view
                ) as key:
                    value, _ = wr.QueryValueEx(key, "InstallRoot")
                for suffix in ("wps.exe", "office6/wps.exe"):
                    found.append(("wps", str(Path(value) / suffix), "wps-registry-installroot"))
            except OSError:
                pass
    for progid in PROGIDS:
        try:
            with wr.OpenKey(wr.HKEY_CLASSES_ROOT, progid + r"\CLSID") as key:
                wr.QueryValueEx(key, "")
            notes.append(
                {"source": "WPS COM registration", "status": "registered", "progid": progid}
            )
        except OSError:
            notes.append(
                {"source": "WPS COM registration", "status": "not-visible", "progid": progid}
            )
    return found, notes


def ps_literal(value):
    import base64

    data = base64.b64encode(str(value).encode("utf-8")).decode("ascii")
    return "([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('" + data + "')))"


def com_script(progid, source=None, target=None):
    if progid not in PROGIDS:
        raise ValueError("Unsupported WPS COM ProgID")
    # Never GetActiveObject or Stop-Process: only close our opened document.
    # Some WPS builds reuse an existing process, so do not Quit a pre-existing WPS.
    setup = r"""
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new()
$wps=$null; $document=$null; $alerts=$null; $links=$null; $security=$null
$hadProcess=@(Get-Process -Name wps,wpsoffice -ErrorAction SilentlyContinue).Count -gt 0
try {
  $wps=New-Object -ComObject PROGID_VALUE
  if(-not $hadProcess){$wps.Visible=$false}
  $initialCount=$wps.Documents.Count
  ACTION_CODE
} finally {
  if($null -ne $document){
    try {$document.Close([ref]0)} catch {Write-Warning ('Closing own WPS document failed: '+$_.Exception.Message)} finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)}
  }
  if($null -ne $wps){
    try {
      if($null -ne $links){$wps.Options.UpdateLinksAtOpen=$links}
      if($null -ne $security){$wps.AutomationSecurity=$security}
      if($null -ne $alerts){$wps.DisplayAlerts=$alerts}
      if((-not $hadProcess) -and $wps.Documents.Count -eq 0){$wps.Quit([ref]0)}
    } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wps)}
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"""
    if source is None:
        action = "Write-Output ('WPS '+$wps.Version)"
    else:
        action = r"""
  $alerts=$wps.DisplayAlerts; $wps.DisplayAlerts=0
  $security=$wps.AutomationSecurity; $wps.AutomationSecurity=3
  $links=$wps.Options.UpdateLinksAtOpen; $wps.Options.UpdateLinksAtOpen=$false
  $document=$wps.Documents.Open(SOURCE_VALUE,$false,$true,$false)
  $document.Repaginate()
  $document.ExportAsFixedFormat(TARGET_VALUE,17,$false,0,0,1,1,0,$false,$false,1,$true,$true,$false)
  Write-Output ('WPS '+$wps.Version)
"""
        action = action.replace("SOURCE_VALUE", ps_literal(source)).replace(
            "TARGET_VALUE", ps_literal(target)
        )
    return setup.replace("PROGID_VALUE", ps_literal(progid)).replace("ACTION_CODE", action)


def probe(ps_run):
    attempts = []
    for progid in PROGIDS:
        try:
            proc = ps_run(com_script(progid), 40)
            attempts.append(
                {
                    "progid": progid,
                    "returncode": proc.returncode,
                    "stdout": proc.stdout.strip(),
                    "stderr": proc.stderr.strip(),
                }
            )
            if proc.returncode == 0:
                return {"usable": True, "progid": progid, "attempts": attempts}
        except Exception as exc:
            attempts.append({"progid": progid, "error": str(exc)})
    return {
        "usable": False,
        "attempts": attempts,
        "notice": "发现安装不等于COM可用。改用WPS界面导出交换流程，不要求安装另一套Office。",
    }


def export(docx, pdf, candidate, tmp, ps_run):
    from zipfile import ZipFile
    from shutil import copy2

    # Preview a disposable copy. Never pass user source as a writable COM document.
    copy = Path(tmp) / "wps-input.docx"
    copy2(docx, copy)
    with ZipFile(copy) as z:
        if any("vbaproject" in n.lower() for n in z.namelist()):
            raise ValueError("含宏载荷的文档不能走自动WPS预览，请先明确安全处理。")
    preferred = candidate.get("progid") or candidate.get("probe", {}).get("progid")
    progids = ([preferred] if preferred in PROGIDS else []) + [p for p in PROGIDS if p != preferred]
    errors = []
    for progid in progids:
        Path(pdf).unlink(missing_ok=True)
        try:
            proc = ps_run(com_script(progid, copy, pdf), 180)
            if proc.returncode == 0 and Path(pdf).is_file() and Path(pdf).stat().st_size:
                return {"version": proc.stdout.strip(), "progid": progid}
            errors.append(progid + ": " + proc.stdout + "\n" + proc.stderr)
        except Exception as exc:
            errors.append(progid + ": " + str(exc))
    raise RuntimeError(
        "WPS已找到，但此安装的COM启动/安全设置/导出接口不可用。有已授权桌面工具或有效导出时走wps_exchange.py；否则先尝试其他可用软件，确无替代路径时按no-renderer.md静态交付并说明未完成项，不强迫用户协助。不要改安全策略或杀掉用户进程。\n"
        + "\n".join(errors)
    )
