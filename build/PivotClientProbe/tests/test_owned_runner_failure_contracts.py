"""Verify owned-runner failure reporting with real task-owned Python processes."""

import contextlib
import importlib.util
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest import mock


_RUNNER_PATH = Path(__file__).parents[1] / "run-advanced-excel.py"
_SPEC = importlib.util.spec_from_file_location("pivot_client_runner", _RUNNER_PATH)
_RUNNER = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(_RUNNER)
_REAL_POPEN = subprocess.Popen


class OwnedRunnerFailureContractsTest(unittest.TestCase):
    """验证生命周期故障会保留真实进程清理及 JSON 报告证据。"""

    def setUp(self):
        self.temp_directory = tempfile.TemporaryDirectory(prefix="pivot-runner-contract-")
        self.root = Path(self.temp_directory.name)
        self.workbook = self.root / "fixture.xlsx"
        self.workbook.write_bytes(b"fixture")
        self.child_script = self.root / "child.py"
        self.child_script.write_text(
            "import json, pathlib, sys, time\n"
            "result, mode, exit_code = sys.argv[1], sys.argv[2], int(sys.argv[3])\n"
            "if mode == 'sleep': time.sleep(10)\n"
            "elif mode == 'invalid-json': pathlib.Path(result).write_text('{', encoding='utf-8')\n"
            "elif mode == 'non-object-json': pathlib.Path(result).write_text('[]', encoding='utf-8')\n"
            "else: pathlib.Path(result).write_text(json.dumps({'Status': mode}), encoding='utf-8')\n"
            "sys.exit(exit_code)\n",
            encoding="utf-8",
            newline="\n",
        )
        self.owned_processes = []
        self.child_processes = []
        self.sentinel = self.start_sleeper()

    def tearDown(self):
        for process in self.owned_processes + self.child_processes + [self.sentinel]:
            if process.poll() is None:
                process.kill()
            process.wait(timeout=10)
        self.temp_directory.cleanup()

    def start_sleeper(self):
        process = _REAL_POPEN([sys.executable, "-c", "import time; time.sleep(60)"],
                              stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                              stderr=subprocess.DEVNULL)
        return process

    def launch_owned_process(self, _workbook):
        process = self.start_sleeper()
        self.owned_processes.append(process)
        return process

    def describe_process(self, process):
        return {"ProcessId": process.pid, "StartTimeUtc": "2030-01-01T00:00:00+00:00",
                "Name": "EXCEL", "OwnershipPublishedBeforeCom": True}

    def invoke_runner(self, result_name, mode="Passed", exit_code=0, timeout_seconds=3,
                      child_start_error=None, launch_error=None, describe_error=None,
                      owner_write_error=False, basic=False):
        result = self.root / result_name
        harness = result.with_suffix(".harness.json")
        original_write_text = Path.write_text

        def child_process(command, **kwargs):
            if child_start_error is not None:
                raise child_start_error
            if len(command) > 4 and command[3] == "-EncodedCommand":
                child = _REAL_POPEN([sys.executable, str(self.child_script), str(result), mode,
                                     str(exit_code)], **kwargs)
                self.child_processes.append(child)
                return child
            return _REAL_POPEN(command, **kwargs)

        def write_text(path, *args, **kwargs):
            if owner_write_error and path == result.with_suffix(".owner.json"):
                raise OSError("simulated owner publication failure")
            return original_write_text(path, *args, **kwargs)

        launch = (mock.Mock(side_effect=launch_error) if launch_error is not None
                  else mock.Mock(side_effect=self.launch_owned_process))
        describe = (mock.Mock(side_effect=describe_error) if describe_error is not None
                    else mock.Mock(side_effect=self.describe_process))
        argv = [str(_RUNNER_PATH), str(self.workbook), str(result),
                "--timeout-seconds", str(timeout_seconds)]
        if basic:
            argv.append("--basic")
        with mock.patch.object(_RUNNER, "launch_owned_excel", launch), \
             mock.patch.object(_RUNNER, "describe_owned_excel", describe), \
             mock.patch.object(_RUNNER.shutil, "which", return_value=sys.executable), \
             mock.patch.object(_RUNNER.subprocess, "Popen", side_effect=child_process), \
             mock.patch.object(Path, "write_text", write_text), \
             mock.patch.object(sys, "argv", argv), \
             contextlib.redirect_stdout(io.StringIO()):
            exit_result = _RUNNER.main()
        return exit_result, result, harness, launch, describe

    def assert_sentinel_survived(self):
        self.assertIsNone(self.sentinel.poll(), "the unrelated sentinel process must remain alive")

    def test_launch_failure_without_handle_reports_unknown_exit(self):
        exit_code, _result, harness_path, launch, _describe = self.invoke_runner(
            "launch-failure.json", launch_error=OSError("simulated WinError 1312"))

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertFalse(report["ownedExcelHandleObtained"])
        self.assertIsNone(report["ownedExcelExited"])
        self.assertIn("1312", report["startupError"])
        self.assertFalse(report["childStarted"])
        launch.assert_called_once()
        self.assert_sentinel_survived()

    def test_identity_read_failure_cleans_only_owned_handle(self):
        exit_code, _result, harness_path, _launch, describe = self.invoke_runner(
            "identity-failure.json", describe_error=OSError("simulated process identity failure"))

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertTrue(report["ownedExcelHandleObtained"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertFalse(report["childStarted"])
        self.assertIn("identity failure", report["startupError"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        describe.assert_called_once()
        self.assert_sentinel_survived()

    def test_owner_publication_failure_cleans_only_owned_handle(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "owner-write-failure.json", owner_write_error=True)

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertTrue(report["ownedExcelHandleObtained"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertFalse(report["childStarted"])
        self.assertIn("owner publication", report["startupError"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()

    def test_child_process_creation_failure_cleans_owned_handle(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "child-start-failure.json", child_start_error=OSError("simulated child creation failure"))

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertFalse(report["childStarted"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertIn("child creation failure", report["childError"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()

    def test_nonzero_child_exit_cleans_owned_handle(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "child-nonzero.json", mode="Failed", exit_code=7)

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertEqual(7, report["childExitCode"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertEqual("Failed", report["probeStatus"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()

    def test_external_timeout_kills_child_and_owned_handle(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "child-timeout.json", mode="sleep", timeout_seconds=1)

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertTrue(report["timedOut"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assertIsNotNone(self.child_processes[0].poll())
        self.assert_sentinel_survived()

    def test_successful_terminal_status_cleans_owned_handle(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "success.json")

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(0, exit_code)
        self.assertEqual("Passed", report["status"])
        self.assertEqual("Passed", report["probeStatus"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()

    def test_invalid_terminal_json_should_write_failure_harness(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "invalid-terminal.json", mode="invalid-json")

        self.assertTrue(harness_path.exists())
        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertEqual("Failed", report["status"])
        self.assertEqual("Missing", report["probeStatus"])
        self.assertIn("property name", report["terminalError"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertIsNone(report["cleanupError"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()

    def test_non_object_terminal_json_should_write_failure_harness(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "non-object-terminal.json", mode="non-object-json")

        self.assertTrue(harness_path.exists())
        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertEqual("Failed", report["status"])
        self.assertEqual("Missing", report["probeStatus"])
        self.assertIn("JSON object", report["terminalError"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertIsNone(report["cleanupError"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()

    def test_basic_success_uses_owned_script_and_cleans_instance(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "basic-success.json", basic=True)

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        probe_arguments = report["probeArguments"]
        self.assertEqual(0, exit_code)
        self.assertEqual("Passed", report["status"])
        self.assertEqual("verify-excel.ps1", Path(probe_arguments[0]).name)
        self.assertIn("-OwnedProcessId", probe_arguments)
        self.assertIn("-OwnedStartTimeUtc", probe_arguments)
        self.assertTrue(report["ownedExcelExited"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()

    def test_basic_timeout_cleans_child_and_owned_instance(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "basic-timeout.json", mode="sleep", timeout_seconds=1, basic=True)

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertEqual("verify-excel.ps1", Path(report["probeArguments"][0]).name)
        self.assertTrue(report["timedOut"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assertIsNotNone(self.child_processes[0].poll())
        self.assert_sentinel_survived()

    def test_basic_child_failure_cleans_owned_instance(self):
        exit_code, _result, harness_path, _launch, _describe = self.invoke_runner(
            "basic-child-failure.json", mode="Failed", exit_code=9, basic=True)

        report = json.loads(harness_path.read_text(encoding="utf-8"))
        self.assertEqual(1, exit_code)
        self.assertEqual("verify-excel.ps1", Path(report["probeArguments"][0]).name)
        self.assertEqual(9, report["childExitCode"])
        self.assertTrue(report["ownedExcelExited"])
        self.assertIsNotNone(self.owned_processes[0].poll())
        self.assert_sentinel_survived()


if __name__ == "__main__":
    unittest.main()
