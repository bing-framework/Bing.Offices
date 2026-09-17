# 基线快照

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 执行分支：`feat/miniexcel-provider`
- 快照时间：2026-09-16（Asia/Shanghai）
- 当前工作树：包含前一 MiniExcel Provider 任务的未提交变更；本任务不得覆盖或还原这些变更。
- 当前版本（只读）：`VersionPrefix=2.0.0`，`NpoiPackageVersion=[2.7.4]`，`MiniExcelPackageVersion=[1.46.0]`。
- `version.dev.props` SHA-256（任务开始）：`BABFE7703A15E1C11F46B45D34BC59D7913ECB1E3DCFDA901E72F4D60A8D134D`。
- `version.props` SHA-256（任务开始）：`77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1`。
- TFM：Abstractions/Core=`netstandard2.0`；NPOI/MiniExcel/Unit/Integration=`net6.0;net8.0`；Benchmark/Docs/ResourceProbe=`net8.0`。
- 生产 `InternalsVisibleTo`：Core 当前额外友元为 `Bing.Offices.Npoi`、`Bing.Offices.MiniExcel`；各生产程序集仍有测试友元。
- 已发现项目目录产物：`src/Bing.Offices.Abstractions/artifacts/packages-pack/`，包含前一任务生成的 2.0.0 nupkg/snupkg；需在本任务结束前迁移/清理并报告。
- Preflight 基线验证：Core `netstandard2.0` Release build 0 warning/0 error；NPOI XLSX preflight 双 TFM 27/27，通过 `luna_worker` 独立复核。
- Benchmark 基线：入口尚未显式设置 BenchmarkDotNet `ArtifactsPath`；已有 100K controlled probe 证据，500K/1M 未验证。
- Consumer 基线：net6/net8 Consumer 默认 `BingOfficesPackageVersion=2.0.0`，需要改为从本次 pack 实际包发现/注入，不得修改版本文件。
