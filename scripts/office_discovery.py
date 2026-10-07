#!/usr/bin/env python3
"""Discover actual Office installations. PATH misses do not prove absence.

Read-only by default: environment -> PATH -> known locations -> registry.
Launch probes are explicit; never install software, bypass policy, or terminate
unrelated Office processes. A remote sandbox cannot establish a user's host state.
"""

from __future__ import annotations
import argparse, base64, glob, json, os, platform, re, shutil, subprocess, sys
from pathlib import Path


def powershell():
    for name in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        found = shutil.which(name)
        if found:
            return found
    if os.name == "nt":
        found = (
            Path(os.environ.get("SystemRoot", r"C:\Windows"))
            / "System32/WindowsPowerShell/v1.0/powershell.exe"
        )
        if found.is_file():
            return str(found)
    return None


def ps_run(script, timeout=30):
    shell = powershell()
    if not shell:
        raise RuntimeError("当前运行环境未找到可调用的PowerShell；不能据此断言未安装Word。")
    script = (
        "$ProgressPreference='SilentlyContinue'\n[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)\n"
        + script
    )
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    return subprocess.run(
        [shell, "-NoLogo", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )


def known_paths(system=None, environ=None, home=None):
    system = system or platform.system()
    env = os.environ if environ is None else environ
    home = Path(home or Path.home())
    result = []
    if system == "Windows":
        roots = list(
            dict.fromkeys(
                [
                    env.get("ProgramFiles", r"C:\Program Files"),
                    env.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
                    env.get("ProgramW6432", r"C:\Program Files"),
                ]
            )
        )
        for root in roots:
            root = Path(root)
            result += [
                ("libreoffice", str(root / "LibreOffice/program" / exe), "common-install")
                for exe in ("soffice.com", "soffice.exe")
            ]
            for version in ("16", "15", "14"):
                for middle in ("root/", ""):
                    result.append(
                        (
                            "word",
                            str(root / f"Microsoft Office/{middle}Office{version}/WINWORD.EXE"),
                            "common-install",
                        )
                    )
        local = Path(env.get("LOCALAPPDATA", str(home / "AppData/Local")))
        for relative in (
            "Programs/LibreOffice/program/soffice.com",
            "Programs/LibreOffice/program/soffice.exe",
            "LibreOffice/program/soffice.exe",
        ):
            result.append(("libreoffice", str(local / relative), "per-user-install"))
    elif system == "Darwin":
        for root in (Path("/Applications"), home / "Applications"):
            result.append(
                (
                    "libreoffice",
                    str(root / "LibreOffice.app/Contents/MacOS/soffice"),
                    "application-bundle",
                )
            )
            result.append(
                (
                    "word-manual",
                    str(root / "Microsoft Word.app/Contents/MacOS/Microsoft Word"),
                    "application-bundle",
                )
            )
        result += [
            ("libreoffice", p, "package-manager")
            for p in ("/opt/homebrew/bin/soffice", "/usr/local/bin/soffice")
        ]
    else:
        result += [
            ("libreoffice", p, "common-install")
            for p in (
                "/usr/bin/soffice",
                "/usr/bin/libreoffice",
                "/usr/local/bin/soffice",
                "/usr/lib/libreoffice/program/soffice",
                "/opt/libreoffice/program/soffice",
                "/snap/bin/libreoffice",
            )
        ]
        result += [
            ("libreoffice", p, "versioned-install")
            for p in glob.glob("/opt/libreoffice*/program/soffice")
        ]
    from wps_support import known_paths as wps_known_paths

    result.extend(wps_known_paths(system, env, home))
    return result


def registry_paths():
    found = []
    checks = []
    if os.name != "nt":
        return found, [
            {"source": "windows-registry", "status": "not-applicable-to-current-runtime"}
        ]
    try:
        import winreg as wr
    except ImportError:
        return found, [{"source": "windows-registry", "status": "unavailable"}]
    views = list(dict.fromkeys([0, wr.KEY_WOW64_64KEY, wr.KEY_WOW64_32KEY]))
    for hive_name, hive in [("HKCU", wr.HKEY_CURRENT_USER), ("HKLM", wr.HKEY_LOCAL_MACHINE)]:
        for view in views:
            for exe, engine in [
                ("WINWORD.EXE", "word"),
                ("soffice.exe", "libreoffice"),
                ("soffice.com", "libreoffice"),
            ]:
                key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths" + chr(92) + exe
                try:
                    with wr.OpenKey(hive, key, 0, wr.KEY_READ | view) as k:
                        value, _ = wr.QueryValueEx(k, "")
                    found.append(
                        (
                            engine,
                            os.path.expandvars(str(value)).strip('"'),
                            "registry:" + hive_name + ":" + str(view),
                        )
                    )
                except FileNotFoundError:
                    pass
                except OSError as exc:
                    checks.append(
                        {"source": hive_name + ":" + key, "status": "unreadable", "error": str(exc)}
                    )
            try:
                with wr.OpenKey(
                    hive,
                    r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                    0,
                    wr.KEY_READ | view,
                ) as k:
                    value, _ = wr.QueryValueEx(k, "InstallationPath")
                for suffix in ("root/Office16/WINWORD.EXE", "Office16/WINWORD.EXE"):
                    found.append(("word", str(Path(value) / suffix), "registry:ClickToRun"))
            except FileNotFoundError:
                pass
            except OSError as exc:
                checks.append({"source": "ClickToRun", "status": "unreadable", "error": str(exc)})
            try:
                with wr.OpenKey(
                    hive, r"SOFTWARE\LibreOffice\UNO\InstallPath", 0, wr.KEY_READ | view
                ) as k:
                    value, _ = wr.QueryValueEx(k, "")
                for suffix in ("soffice.com", "soffice.exe", "program/soffice.exe"):
                    found.append(("libreoffice", str(Path(value) / suffix), "registry:LibreOffice"))
            except FileNotFoundError:
                pass
            except OSError as exc:
                checks.append(
                    {"source": "LibreOffice registry", "status": "unreadable", "error": str(exc)}
                )
    try:
        with wr.OpenKey(wr.HKEY_CLASSES_ROOT, r"Word.Application\CLSID") as k:
            wr.QueryValueEx(k, "")
        checks.append({"source": "Word.Application COM registration", "status": "registered"})
    except OSError:
        checks.append({"source": "Word.Application COM registration", "status": "not-visible"})
    from wps_support import registry_paths as wps_registry_paths

    wps_found, wps_notes = wps_registry_paths()
    found.extend(wps_found)
    checks.extend(wps_notes)
    checks.append({"source": "windows-registry", "status": "searched-HKCU-HKLM-32-64-bit"})
    return found, checks


def discover(extra_dirs=(), probe=False):
    candidates = []
    checks = []
    seen = set()
    system = platform.system()

    def add(engine, value, source, allow_missing=False):
        if not value:
            return
        value = os.path.expandvars(os.path.expanduser(str(value).strip().strip('"')))
        path = Path(value)
        exists = path.is_file()
        key = (engine, str(path.resolve()) if exists else value)
        if key in seen:
            return
        seen.add(key)
        if exists or allow_missing:
            candidates.append(
                {
                    "engine": engine,
                    "path": value,
                    "source": source,
                    "found": exists,
                    "automation": engine == "libreoffice"
                    or (
                        engine in ("word", "wps")
                        and system == "Windows"
                        and powershell() is not None
                    ),
                }
            )
        checks.append(
            {
                "source": source,
                "engine": engine,
                "path": value,
                "status": "found" if exists else "not-visible",
            }
        )

    for var, engine in [
        ("XAUFE_LIBREOFFICE", "libreoffice"),
        ("LIBREOFFICE_PATH", "libreoffice"),
        ("SOFFICE_PATH", "libreoffice"),
        ("XAUFE_WORD_PATH", "word"),
        ("XAUFE_WPS_PATH", "wps"),
        ("WPS_PATH", "wps"),
    ]:
        if os.environ.get(var):
            add(engine, os.environ[var], "environment:" + var, True)
    for name, engine in [
        ("soffice.com", "libreoffice"),
        ("soffice", "libreoffice"),
        ("libreoffice", "libreoffice"),
        ("soffice.exe", "libreoffice"),
        ("WINWORD.EXE", "word"),
        ("wps.exe", "wps"),
        ("wps", "wps"),
    ]:
        value = shutil.which(name)
        if value:
            add(engine, value, "PATH:" + name)
        else:
            checks.append({"source": "PATH:" + name, "status": "not-on-PATH"})
    for engine, path, source in known_paths():
        add(engine, path, source)
    registry, notes = registry_paths()
    checks.extend(notes)
    for engine, path, source in registry:
        add(engine, path, source)
    for directory in extra_dirs:
        root = Path(directory).expanduser()
        if root.is_file():
            add(
                (
                    "word"
                    if "winword" in root.name.lower()
                    else (
                        "wps"
                        if root.name.lower() in ("wps.exe", "wps", "wpsoffice")
                        else "libreoffice"
                    )
                ),
                str(root),
                "user-specified",
            )
        else:
            for pattern, engine in [
                ("program/soffice.*", "libreoffice"),
                ("LibreOffice/program/soffice.*", "libreoffice"),
                ("Office*/WINWORD.EXE", "word"),
                ("root/Office*/WINWORD.EXE", "word"),
                ("Microsoft Office/root/Office*/WINWORD.EXE", "word"),
                ("office6/wps.exe", "wps"),
                ("*/office6/wps.exe", "wps"),
                ("Kingsoft/WPS Office/*/office6/wps.exe", "wps"),
                ("wps.exe", "wps"),
                ("wps", "wps"),
            ]:
                for path in root.glob(pattern):
                    if path.name.lower() in (
                        "soffice.com",
                        "soffice.exe",
                        "soffice",
                        "winword.exe",
                        "wps.exe",
                        "wps",
                    ):
                        add(engine, str(path), "user-install-directory")
    if not any(c["engine"] == "word" and c["found"] for c in candidates) and any(
        n.get("source") == "Word.Application COM registration" and n.get("status") == "registered"
        for n in notes
    ):
        candidates.append(
            {
                "engine": "word",
                "path": None,
                "source": "registered-COM",
                "found": True,
                "automation": bool(powershell()),
            }
        )
    if not any(c["engine"] == "wps" and c["found"] for c in candidates):
        for note in notes:
            if note.get("source") == "WPS COM registration" and note.get("status") == "registered":
                candidates.append(
                    {
                        "engine": "wps",
                        "path": None,
                        "source": "registered-COM",
                        "found": True,
                        "automation": system == "Windows" and bool(powershell()),
                        "progid": note["progid"],
                    }
                )
                break
    wsl = "microsoft" in platform.release().lower() or "WSL_DISTRO_NAME" in os.environ
    if wsl:
        checks.append(
            {
                "source": "host-boundary",
                "status": "WSL",
                "notice": "此Python处于WSL；宿主机Word可能不在当前PATH。请用Windows PowerShell运行本探测脚本或查看office-discovery.md，不能把WSL不可见说成宿主未安装。",
            }
        )
    if probe:
        for c in candidates:
            if not c["found"]:
                continue
            try:
                if c["engine"] == "libreoffice":
                    r = subprocess.run(
                        [c["path"], "--version"], capture_output=True, text=True, timeout=20
                    )
                    c["probe"] = {
                        "usable": r.returncode == 0,
                        "stdout": r.stdout.strip(),
                        "stderr": r.stderr.strip(),
                    }
                elif c["engine"] == "wps" and c["automation"]:
                    from wps_support import probe as probe_wps

                    c["probe"] = probe_wps(ps_run)
                elif c["engine"] == "word" and c["automation"]:
                    from word_com import run as run_word

                    r = run_word(ps_run, timeout=40)
                    c["probe"] = {
                        "usable": r.returncode == 0,
                        "stdout": r.stdout.strip(),
                        "stderr": r.stderr.strip(),
                    }
            except (OSError, subprocess.SubprocessError) as exc:
                c["probe"] = {"usable": False, "error": str(exc)}
    return {
        "runtime": {"system": system, "python": sys.executable, "wsl": wsl},
        "candidates": candidates,
        "checks": checks,
        "notice": "结论仅覆盖当前工具可访问环境。找到安装但调用失败=权限/启动问题，不等于未安装；云端沙盒不可替用户电脑作不存在结论。",
    }


def usable_candidates(report=None):
    report = report or discover()
    items = [c for c in report["candidates"] if c["found"] and c["automation"]]
    preference = os.environ.get("XAUFE_RENDER_ENGINE", "auto").lower()
    if preference not in ("auto", "word", "libreoffice", "wps"):
        raise ValueError("XAUFE_RENDER_ENGINE须为auto/word/libreoffice/wps。")
    if preference != "auto":
        items = [x for x in items if x["engine"] == preference]
    preferred = "word" if report["runtime"]["system"] == "Windows" else "libreoffice"
    return sorted(
        items, key=lambda x: (not x["source"].startswith("environment:"), x["engine"] != preferred)
    )


def main():
    p = argparse.ArgumentParser(description="先定位真实安装，不用一次which/where失败推断不存在。")
    p.add_argument("--search-dir", action="append", default=[])
    p.add_argument("--probe", action="store_true")
    p.add_argument("--report")
    a = p.parse_args()
    report = discover(a.search_dir, a.probe)
    if a.report:
        Path(a.report).parent.mkdir(parents=True, exist_ok=True)
        Path(a.report).write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
