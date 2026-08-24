# TASK-BING-OFFICES-20260821-MAPVAL-V2 验证证据

## 状态

- Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`
- 当前状态：`PARTIAL`

## 最新验证

- `dotnet restore Bing.Offices.sln --locked-mode`：`FAILED`，`NU1004`；使用 `dotnet restore Bing.Offices.sln --force-evaluate` 成功，未产生受跟踪 lock 文件差异。
- `dotnet build Bing.Offices.sln -c Release --no-restore`：通过，0 错误，180 个警告。
- Unit net6：通过，156/156。
- Unit net8：通过，156/156。
- Integration net6：通过，10/10。
- Integration net8：通过，10/10。
- Docs Consumer net8：通过，3/3。
- Loader 安全专项：通过，8/8；包含 v1/v2、XML DTD、未知字段、深度/字符串/文档大小、流所有权和 oversized file。
- Public API/IVT 定向验证：通过，Public API allowlist/hash 与 production IVT audit 均通过。
- 三包 local pack：通过；Abstractions、Core、Npoi 均输出到 `artifacts/packages`。
- Pack Consumer：通过；仅从本地包 restore/build/run，输出 `pack-consumer-ok`。
- Benchmark smoke：通过；BenchmarkDotNet ShortRun 结果见 `06-performance-gc.md`。

## 未完成验证

- 未新增计划要求的 10K/100K 行、1/5 Unique 列、Plan Build、租户缓存、JSON/XML parse 和注册扫描 Benchmark 矩阵。
- 未实现完整 round-trip writer、diagnostic API、model alias registry，因此对应验证只能覆盖当前 loader 能力。
- 外部 Office/LibreOffice 重开未执行，环境中没有纳入本 Task 的外部 Office runner。