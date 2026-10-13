"""Run the Excel probe with a wall-clock limit and owned-process cleanup."""

import argparse
import base64
from datetime import datetime, timedelta, timezone
import ctypes
from ctypes import wintypes
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import time


def launch_owned_excel(workbook):
    """Launch one hidden instance with its read-only fixture before acquiring a COM object."""
    import winreg

    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID") as key:
        class_id = winreg.QueryValueEx(key, None)[0]
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{class_id}\LocalServer32") as key:
        registration = winreg.QueryValueEx(key, None)[0]
    match = re.match(r'^\s*"([^"]+)"|^\s*(.*?\.exe)(?:\s|$)', registration, re.IGNORECASE)
    executable = Path(match.group(1) or match.group(2)) if match else None
    if executable is None or not executable.is_file() or executable.name.casefold() != "excel.exe":
        raise RuntimeError("Registered Excel executable could not be verified.")
    startup = subprocess.STARTUPINFO()
    startup.dwFlags = subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = 0
    process = subprocess.Popen([str(executable), "/x", "/automation", "/r", str(workbook)],
                               startupinfo=startup, creationflags=subprocess.CREATE_NO_WINDOW)
    return process


def describe_owned_excel(process):
    """Read the creation identity for a retained process handle."""
    creation, exit_time, kernel, user = (wintypes.FILETIME() for _ in range(4))
    get_times = ctypes.windll.kernel32.GetProcessTimes
    get_times.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    get_times.restype = wintypes.BOOL
    if not get_times(wintypes.HANDLE(int(process._handle)), ctypes.byref(creation),
                     ctypes.byref(exit_time), ctypes.byref(kernel), ctypes.byref(user)):
        raise ctypes.WinError()
    ticks = (creation.dwHighDateTime << 32) | creation.dwLowDateTime
    started = datetime(1601, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=ticks // 10)
    return {"ProcessId": process.pid, "StartTimeUtc": started.isoformat(),
            "Name": "EXCEL", "OwnershipPublishedBeforeCom": True}


def inspect_owner(shell, owner):
    """Inspect one recorded process without touching other Excel instances."""
    process_id = int(owner["ProcessId"])
    script = (
        "$ErrorActionPreference='Stop'; "
        f"$p=Get-Process -Id {process_id} -ErrorAction SilentlyContinue; "
        "if($null -ne $p){[ordered]@{Name=$p.ProcessName;"
        "StartTimeUtc=$p.StartTime.ToUniversalTime().ToString('o')}|ConvertTo-Json -Compress}; exit 0"
    )
    result = subprocess.run([shell, "-NoProfile", "-NonInteractive", "-Command", script],
                            capture_output=True, timeout=15)
    if result.returncode:
        raise RuntimeError("Cannot inspect the recorded Excel process.")
    text = result.stdout.decode("utf-8-sig", errors="strict").strip()
    if not text:
        return False
    current = json.loads(text)
    if current["Name"].casefold() != "excel" or owner["Name"].casefold() != "excel":
        raise RuntimeError("Recorded process is not Excel.")
    actual_start = datetime.fromisoformat(current["StartTimeUtc"])
    expected_start = datetime.fromisoformat(owner["StartTimeUtc"])
    if abs((actual_start - expected_start).total_seconds()) > 0.001:
        raise RuntimeError("Recorded process identity changed; cleanup refused.")
    return True


def main():
    """Bound the child probe and preserve its terminal evidence."""
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook")
    parser.add_argument("result")
    parser.add_argument("--timeout-seconds", type=int, default=180)
    parser.add_argument("--basic", action="store_true",
                        help="Run the basic workbook lifecycle probe; advanced remains the default.")
    parser.add_argument("--probe-script", help="Explicit diagnostic script; defaults to lifecycle verification.")
    parser.add_argument("--fault-after-owner-seconds", type=int, default=0,
                        help="Test-only delay after publishing dedicated Excel ownership.")
    parser.add_argument("--owned-host", action="store_true",
                        help="Compatibility option; the default lifecycle already owns a dedicated Excel process.")
    parser.add_argument("--fault-before-attach-seconds", type=int, default=0,
                        help="Test-only delay before connecting to the parent-owned instance.")
    args = parser.parse_args()
    if args.timeout_seconds < 1 or args.timeout_seconds > 600:
        parser.error("timeout-seconds must be within 1..600")
    if args.fault_after_owner_seconds < 0 or args.fault_after_owner_seconds > 600:
        parser.error("fault-after-owner-seconds must be within 0..600")
    if args.probe_script and args.fault_after_owner_seconds:
        parser.error("fault injection requires the default lifecycle script")
    if args.basic and args.probe_script:
        parser.error("basic mode cannot be combined with an explicit diagnostic probe script")
    if not 0 <= args.fault_before_attach_seconds <= 600:
        parser.error("fault-before-attach-seconds must be within 0..600")
    if args.probe_script and args.owned_host:
        parser.error("owned-host is unavailable with an explicit diagnostic probe script")
    use_owned_host = not bool(args.probe_script)
    if args.fault_before_attach_seconds and not use_owned_host:
        parser.error("fault-before-attach-seconds requires the default lifecycle")
    shell = shutil.which("pwsh")
    if shell is None:
        raise RuntimeError("PowerShell 7 is required for explicit UTF-8 script execution.")
    workbook = Path(args.workbook).resolve()
    result_path = Path(args.result).resolve()
    owner_path = result_path.with_suffix(".owner.json")
    harness_path = result_path.with_suffix(".harness.json")
    log_path = result_path.with_suffix(".log")
    for path in [result_path, owner_path, harness_path, log_path]:
        if path.exists() or path == workbook:
            raise ValueError("Use new result paths; existing evidence is never overwritten.")
    result_path.parent.mkdir(parents=True, exist_ok=True)
    script_name = "verify-excel.ps1" if args.basic else "verify-advanced-excel.ps1"
    script = (Path(args.probe_script).resolve() if args.probe_script
              else Path(__file__).with_name(script_name).resolve())
    command = [shell, "-NoProfile", "-NonInteractive", "-File", str(script),
               "-WorkbookPath", str(workbook), "-ResultPath", str(result_path),
               "-OwnerPath", str(owner_path)]
    if args.fault_after_owner_seconds:
        command.extend(["-FaultAfterOwnerSeconds", str(args.fault_after_owner_seconds)])
    timer = time.perf_counter()
    report = {"command": command, "timeoutSeconds": args.timeout_seconds,
              "timedOut": False, "forcedOwnedExcelCleanup": False,
              "ownedExcelHandleObtained": False, "ownedExcelExited": None,
              "startupError": None, "cleanupError": None, "terminalError": None, "childExitCode": None,
              "childStarted": False}
    owned_process = None
    child = None
    try:
        if use_owned_host:
            owned_process = launch_owned_excel(workbook)
            report["ownedExcelHandleObtained"] = True
            owner = describe_owned_excel(owned_process)
            owner_path.write_text(json.dumps(owner, ensure_ascii=False, indent=2) + "\n",
                                  encoding="utf-8", newline="\n")
            command.extend(["-OwnedProcessId", str(owner["ProcessId"]),
                            "-OwnedStartTimeUtc", owner["StartTimeUtc"]])
            if args.fault_before_attach_seconds:
                command.extend(["-FaultBeforeAttachSeconds", str(args.fault_before_attach_seconds)])
            report["parentOwnedExcel"] = owner
    except Exception as error:
        report["startupError"] = str(error)

    try:
        if report["startupError"] is None:
            # Configure encoding before invoking the script so even parser errors use UTF-8.
            def quote_powershell(value):
                return "'" + value.replace("'", "''") + "'"

            # Parameter names must remain PowerShell tokens; their values remain quoted literals.
            invocation = " & " + quote_powershell(command[4])
            for index in range(5, len(command), 2):
                invocation += " " + command[index] + " " + quote_powershell(command[index + 1])
            prelude = ("[Console]::InputEncoding=[Text.UTF8Encoding]::new($false);"
                       "[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false);"
                       "$OutputEncoding=[Text.UTF8Encoding]::new($false);")
            encoded = base64.b64encode((prelude + invocation).encode("utf-16-le")).decode("ascii")
            report["probeArguments"] = command[4:]
            command = [shell, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]
            report["command"] = command
            with log_path.open("w", encoding="utf-8", newline="\n") as log:
                child = subprocess.Popen(command, stdout=log, stderr=subprocess.STDOUT,
                                         creationflags=subprocess.CREATE_NO_WINDOW)
                report["childStarted"] = True
                try:
                    report["childExitCode"] = child.wait(timeout=args.timeout_seconds)
                except subprocess.TimeoutExpired:
                    report["timedOut"] = True
                    child.kill()
                    report["childExitCode"] = child.wait(timeout=15)
    except Exception as error:
        report["childError"] = str(error)
        if child is not None and child.poll() is None:
            try:
                child.kill()
                child.wait(timeout=15)
            except Exception as cleanup_error:
                report["cleanupError"] = str(cleanup_error)
    finally:
        if child is not None and child.poll() is None:
            try:
                child.kill()
                child.wait(timeout=15)
            except Exception as error:
                report["cleanupError"] = str(error)
        if owned_process is not None:
            try:
                # The retained process handle cannot accidentally target a reused PID.
                if owned_process.poll() is None:
                    owned_process.kill()
                    report["forcedOwnedExcelCleanup"] = True
                owned_process.wait(timeout=15)
                report["ownedExcelExited"] = True
            except Exception as error:
                report["cleanupError"] = str(error)
        elif owner_path.exists():
            try:
                owner = json.loads(owner_path.read_text(encoding="utf-8"))
                alive = inspect_owner(shell, owner)
                if alive:
                    process_id = int(owner["ProcessId"])
                    cleanup = subprocess.run([shell, "-NoProfile", "-NonInteractive", "-Command",
                                              f"Stop-Process -Id {process_id} -Force -ErrorAction Stop"],
                                             capture_output=True, timeout=15)
                    if cleanup.returncode:
                        raise RuntimeError("Owned Excel termination failed.")
                    report["forcedOwnedExcelCleanup"] = True
                report["ownedExcelExited"] = not inspect_owner(shell, owner)
            except Exception as error:
                report["cleanupError"] = str(error)
    report["elapsedSeconds"] = time.perf_counter() - timer
    report["probeStatus"] = "Missing"
    try:
        if result_path.exists():
            terminal = json.loads(result_path.read_text(encoding="utf-8"))
            if not isinstance(terminal, dict):
                raise ValueError("The child probe result must be a JSON object.")
            report["probeStatus"] = terminal.get("Status", "Missing")
    except Exception as error:
        # 子进程产物损坏也保留已经完成的清理证据，不能丢失外围终态。
        report["terminalError"] = str(error)
    passed = (report["startupError"] is None and report.get("childError") is None
              and report["terminalError"] is None
              and not report["timedOut"] and report["childExitCode"] == 0
              and report["probeStatus"] == "Passed" and report["ownedExcelExited"] is True
              and report["cleanupError"] is None)
    report["status"] = "Passed" if passed else "Failed"
    harness_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                            encoding="utf-8", newline="\n")
    print(json.dumps({key: value for key, value in report.items() if key != "command"},
                     ensure_ascii=False))
    if not passed and log_path.exists():
        print(log_path.read_text(encoding="utf-8"))
    return 0 if passed else 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
