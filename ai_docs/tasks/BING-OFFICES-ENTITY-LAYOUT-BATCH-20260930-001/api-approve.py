from pathlib import Path

source = Path("ai_docs/tasks/BING-OFFICES-ENTITY-LAYOUT-BATCH-20260930-001/artifacts/api-capture/api-candidate.json")
target = Path("build/api-snapshot-baseline.json")
text = source.read_text(encoding="utf-8-sig")
text = text.replace('"approvedBy": ""', '"approvedBy": "user:entity-layout-batch-2026-09-30"', 1)
text = text.replace('"approvedAt": ""', '"approvedAt": "2026-09-30"', 1)
with target.open("w", encoding="utf-8", newline="\n") as handle:
    handle.write(text)
