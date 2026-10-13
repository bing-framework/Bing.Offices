"""Verify timeout cleanup while another task-owned Excel instance remains open."""

import argparse
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys


def main():
    """Exercise real before/after-connection failures without opening user workbooks."""
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook")
    parser.add_argument("output")
    args = parser.parse_args()
    workbook = Path(args.workbook).resolve()
    output = Path(args.output).resolve()
    if output.exists():
        raise ValueError("Use a new directory; existing evidence is never overwritten.")
    output.mkdir(parents=True)
    runner = Path(__file__).with_name("run-advanced-excel.py")
    spec = importlib.util.spec_from_file_location("owned_excel_runner", runner)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    sentinel_path = output / "existing-instance.xlsx"
    report = {"status": "Running", "existingInstanceWasTaskOwnedSentinel": True,
              "sentinelHandleObtained": False, "sentinelExited": None,
              "sentinelCleanedUp": None, "existingInstance": None,
              "cases": [], "startupError": None, "error": None, "cleanupError": None}
    sentinel = None
    startup_complete = False
    try:
        shutil.copyfile(workbook, sentinel_path)
        sentinel = module.launch_owned_excel(sentinel_path)
        report["sentinelHandleObtained"] = True
        sentinel_owner = module.describe_owned_excel(sentinel)
        report["existingInstance"] = sentinel_owner
        startup_complete = True
        for name, fault_option, limit in [
            ("before-attach", "--fault-before-attach-seconds", "12"),
            ("after-attach", "--fault-after-owner-seconds", "20"),
        ]:
            if sentinel.poll() is not None:
                raise RuntimeError("Existing-instance sentinel exited before the probe.")
            result = output / (name + ".json")
            command = [sys.executable, str(runner), str(workbook), str(result),
                       fault_option, "60", "--timeout-seconds", limit]
            completed = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                       timeout=60)
            (output / (name + ".invocation.log")).write_text(
                completed.stdout.decode("utf-8", errors="strict"), encoding="utf-8", newline="\n")
            harness_path = result.with_suffix(".harness.json")
            owner_path = result.with_suffix(".owner.json")
            harness = (json.loads(harness_path.read_text(encoding="utf-8"))
                       if harness_path.exists() else None)
            owner = (json.loads(owner_path.read_text(encoding="utf-8"))
                     if owner_path.exists() else None)
            existing_alive = sentinel.poll() is None
            connected = owner is not None and owner.get("AcquiredApplication") is True
            passed = (completed.returncode == 1 and harness is not None
                      and harness.get("startupError") is None and owner is not None
                      and harness["timedOut"]
                      and harness["forcedOwnedExcelCleanup"] and harness["ownedExcelExited"]
                      and harness["cleanupError"] is None and existing_alive
                      and owner["ProcessId"] != sentinel_owner["ProcessId"]
                      and (name != "after-attach" or connected))
            report["cases"].append({"name": name, "faultContractPassed": passed,
                                    "existingInstanceStillRunning": existing_alive,
                                    "applicationConnected": connected, "harness": harness,
                                    "owner": owner})
            if not passed:
                raise RuntimeError("Owned-host lifecycle case did not satisfy its fault contract: " + name)
        report["status"] = "Passed"
    except Exception as error:
        report["status"] = "Failed"
        key = "error" if startup_complete else "startupError"
        report[key] = str(error)
    finally:
        if sentinel is not None:
            try:
                # This retained handle belongs to the sentinel created above, never to a user instance.
                if sentinel.poll() is None:
                    sentinel.kill()
                sentinel.wait(timeout=15)
                report["sentinelExited"] = sentinel.poll() is not None
                report["sentinelCleanedUp"] = report["sentinelExited"]
            except Exception as error:
                report["cleanupError"] = str(error)
        report["status"] = ("Passed" if report["status"] == "Passed"
                            and report["sentinelExited"] is True
                            and report["cleanupError"] is None else "Failed")
        (output / "lifecycle.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                                               encoding="utf-8", newline="\n")
    print(json.dumps({"status": report["status"], "error": report["error"],
                      "startupError": report["startupError"],
                      "cleanupError": report["cleanupError"],
                      "sentinelHandleObtained": report["sentinelHandleObtained"],
                      "sentinelExited": report["sentinelExited"],
                      "cases": [{"name": item["name"], "passed": item["faultContractPassed"]}
                                for item in report["cases"]]}, ensure_ascii=False))
    return 0 if report["status"] == "Passed" else 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
