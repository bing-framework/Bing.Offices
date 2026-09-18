# API Diff 报告

- Identity self-test：PASS（LF/CRLF checkout、source/assembly/nupkg/approval tamper 场景按预期拒绝）。
- API snapshot compare：NOT_VERIFIED。
- 执行命令：`dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release --no-build --no-restore -- --root output/release --baseline build/api-snapshot-baseline.json --packages artifacts/packages --repository . --output artifacts/api/rc-test-arch-compare`
- 阻断证据：`candidate identity candidate source manifest SHA-256 mismatch`（baseline expected `D78F11D98A5D5F9BF7FAA355072B1FB96920C16ACF53E77013BA2ACB3A28DBA1`，当前实际 `0D858201507D680386551CA755D137E0F3A6CEAFBE848461E1BD5734D2312CF9`）。
- 未修改 API baseline；需要维护者在候选提交上重新生成/批准 identity artifact 后才能判定 API Release PASS。

