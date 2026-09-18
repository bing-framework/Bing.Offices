from __future__ import annotations

import re
import subprocess
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
TASK_DIR = Path(__file__).resolve().parent
HEAD_PROJECTS = (
    "tests/Bing.Offices.Tests",
    "tests/Bing.Offices.Tests.Integration",
)
CURRENT_PROJECTS = (
    "tests/Bing.Offices.Tests",
    "tests/Bing.Offices.Npoi.Tests",
    "tests/Bing.Offices.MiniExcel.Tests",
    "tests/Bing.Offices.Tests.Integration",
    "tests/Bing.Offices.Npoi.Tests.Integration",
    "tests/Bing.Offices.MiniExcel.Tests.Integration",
)
ATTRIBUTE = re.compile(r"^\s*\[(Fact|Theory)(?:\([^\r\n]*\))?\]\s*$")
CLASS = re.compile(r"^\s*(?:(?:public|private|protected|internal|sealed|abstract|static)\s+)*class\s+(\w+)")
METHOD = re.compile(
    r"^\s*(?:public|private|protected|internal)\s+"
    r"(?:(?:static|async|new|virtual|override|sealed)\s+)*"
    r"(?:[\w<>,.?\[\]]+\s+)+(?P<name>[A-Za-z_]\w*)\s*\("
)

# These are intentional class/method identity changes made during the split.
CLASS_ALIASES = {
    "ReviewFixRegressionTest": "ReviewFixCoreRegressionTest",
    "PublicExtensionCoverageTest": "PublicCoreExtensionCoverageTest",
}
METHOD_ALIASES = {
    "DataRowStartIndex_ShouldAlignRawDateSerialsAcrossProviders":
        "DataRowStartIndex_ShouldAlignRawDateSerials",
    "ReadColumnRange_ShouldReportAbsolutePhysicalColumnIndexAcrossProviders":
        "ReadColumnRange_ShouldReportAbsolutePhysicalColumnIndex",
}


def current_files(projects: tuple[str, ...]) -> list[Path]:
    files: list[Path] = []
    for project in projects:
        files.extend(sorted((ROOT / project).rglob("*.cs")))
    return files


def head_text(relative_path: str) -> str:
    result = subprocess.run(
        ["git", "show", f"HEAD:{relative_path}"],
        cwd=ROOT,
        check=True,
        capture_output=True,
    )
    return result.stdout.decode("utf-8")


def parse(text: str, relative_path: str) -> list[dict[str, str | int]]:
    lines = text.splitlines()
    records: list[dict[str, str | int]] = []
    class_name = "<unknown>"
    for index, line in enumerate(lines):
        class_match = CLASS.match(line)
        if class_match:
            class_name = class_match.group(1)
        attribute = ATTRIBUTE.match(line)
        if not attribute:
            continue
        method_name = "<unresolved>"
        method_line = index + 1
        for candidate_index in range(index + 1, min(len(lines), index + 64)):
            method_match = METHOD.match(lines[candidate_index])
            if method_match:
                method_name = method_match.group("name")
                method_line = candidate_index + 1
                break
        records.append(
            {
                "path": relative_path,
                "line": method_line,
                "attribute": attribute.group(1),
                "class": class_name,
                "method": method_name,
            }
        )
    return records


def read_head() -> list[dict[str, str | int]]:
    records: list[dict[str, str | int]] = []
    for project in HEAD_PROJECTS:
        paths = subprocess.run(
            ["git", "ls-tree", "-r", "--name-only", "HEAD", "--", project],
            cwd=ROOT,
            check=True,
            capture_output=True,
        ).stdout.decode("utf-8").splitlines()
        for path in paths:
            if path.endswith(".cs"):
                records.extend(parse(head_text(path), path))
    return records


def read_current() -> list[dict[str, str | int]]:
    records: list[dict[str, str | int]] = []
    for path in current_files(CURRENT_PROJECTS):
        relative_path = path.relative_to(ROOT).as_posix()
        records.extend(parse(path.read_text(encoding="utf-8"), relative_path))
    return records


def identity(record: dict[str, str | int]) -> str:
    return f"{record['class']}.{record['method']}"


def project_name(path: str) -> str:
    return path.split("/", 2)[1]


def match_records(head: list[dict[str, str | int]], current: list[dict[str, str | int]]):
    by_identity: dict[tuple[str, str], list[int]] = defaultdict(list)
    by_method: dict[str, list[int]] = defaultdict(list)
    for index, record in enumerate(current):
        by_identity[(str(record["class"]), str(record["method"]))].append(index)
        by_method[str(record["method"])].append(index)
    used: set[int] = set()
    results = []
    for old in head:
        candidates = by_identity.get((str(old["class"]), str(old["method"])), [])
        if not candidates:
            alias_class = CLASS_ALIASES.get(str(old["class"]))
            if alias_class:
                candidates = by_identity.get((alias_class, str(old["method"])), [])
        if not candidates:
            alias_method = METHOD_ALIASES.get(str(old["method"]))
            if alias_method:
                candidates = by_method.get(alias_method, [])
        if not candidates:
            candidates = by_method.get(str(old["method"]), [])
        candidates = [candidate for candidate in candidates if candidate not in used]
        if candidates:
            current_index = candidates[0]
            used.add(current_index)
            new = current[current_index]
            if old["path"] == new["path"] and old["class"] == new["class"] and old["method"] == new["method"]:
                status = "UNCHANGED"
            else:
                status = "MOVED"
            if str(old["class"]) != str(new["class"]) or str(old["method"]) != str(new["method"]):
                status = "RENAMED/MOVED"
            results.append((status, old, new))
        else:
            results.append(("REMOVED", old, None))
    additions = [("ADDED", None, record) for index, record in enumerate(current) if index not in used]
    return results, additions


def format_record(record: dict[str, str | int] | None) -> str:
    if record is None:
        return "-"
    return f"`{record['path']}:{record['line']}` `{record['class']}.{record['method']}` ({record['attribute']})"


def addition_reason(record: dict[str, str | int]) -> str:
    path = str(record["path"])
    method = str(record["method"])
    if "PublicNpoiExtensionCoverageTest" in path:
        return "NPOI provider extension direct-coverage gate split from the former mixed gate"
    if "MiniExcelProviderTest" in path and ("RawDate" in method or "Relations_" in method):
        return "MiniExcel hotspot or Relation delegate regression added for the RC hardening"
    if "Tests.Integration/MiniExcelProviderContractTest.cs" in path:
        return "Aggregate cross-provider DI registration-order contract"
    if "Npoi.Tests.Integration" in path:
        return "NPOI real-file acceptance, rich-workbook, or Large gate coverage"
    if "MiniExcel.Tests.Integration" in path:
        return "MiniExcel real-file IO, cancellation, or Large gate coverage"
    return "New direct coverage introduced after the baseline"


def main() -> None:
    head = read_head()
    current = read_current()
    results, additions = match_records(head, current)
    status_counts = Counter(status for status, _, _ in results)
    status_counts.update(status for status, _, _ in additions)
    lines = [
        "# Test Method Identity Audit",
        "",
        "This file is generated by `method-identity-audit.py` using UTF-8 source reads.",
        "The parser only accepts standalone Fact/Theory attributes and associates each attribute with its following method declaration.",
        "",
        f"- HEAD methods: {len(head)}",
        f"- Current six-role methods: {len(current)}",
        f"- Status totals: UNCHANGED={status_counts['UNCHANGED']}, MOVED={status_counts['MOVED']}, RENAMED/MOVED={status_counts['RENAMED/MOVED']}, ADDED={status_counts['ADDED']}, REMOVED={status_counts['REMOVED']}",
        "- Reconciliation formula: `After=Before+Added-Removed`; the status list below is the source of the totals.",
        "",
        "## HEAD identities and destinations",
        "",
        "| Status | HEAD identity | Current destination |",
        "| --- | --- | --- |",
    ]
    for status, old, new in results:
        lines.append(f"| {status} | {format_record(old)} | {format_record(new)} |")
    lines.extend(
        [
            "",
            "## Current identities without a HEAD match",
            "",
            "| Status | Current identity | Reason |",
            "| --- | --- | --- |",
        ]
    )
    for _, _, new in additions:
        lines.append(f"| ADDED | {format_record(new)} | {addition_reason(new)} |")
    lines.extend(
        [
            "",
            "## Current role totals",
            "",
            "| Project | Fact/Theory methods |",
            "| --- | ---: |",
        ]
    )
    counts = Counter(project_name(str(record["path"])) for record in current)
    for project in CURRENT_PROJECTS:
        lines.append(f"| `{project}` | {counts[project.split('/', 1)[1]]} |")
    (TASK_DIR / "method-identity-audit.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
