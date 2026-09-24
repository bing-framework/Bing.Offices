# Benchmark 报告

## Materialization/Binding 真实探针

新增 `--materialization-binding-probe <artifact> <rows> <repetitions>`，使用固定 seed `20260920` 和六列显式映射，通过公开 NPOI/MiniExcel 导出与导入 API 分离执行 `export-only`、`import-only` 和 `roundtrip`，每种操作分别覆盖 `sync`、`async`。产物文档记录当前候选程序集 SHA-256、固定 dataset identity、每个 Provider 的输入 fixture identity，并记录导出/导入耗时、分配量、GC 次数、峰值工作集、输出字节、导入错误数、行数和字段级完整性结果；启动时还会重新解析 JSON 并校验 schema、样本数量、操作字段、fixture identity 和关键输出结构。该探针明确写入 `baselineStatus=NOT_APPLICABLE_BEFORE`：当前工作区没有能以同一公开 workload 运行且身份独立的 before candidate，因此只归因于隔离的当前候选 E2E，不伪造 before 数字。

Release smoke 命令：

```text
dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore --nologo
dotnet benchmarks/Bing.Offices.Benchmarks/bin/Release/net8.0/Bing.Offices.Benchmarks.dll --materialization-binding-probe artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/materialization-binding-smoke-v3.json 3 1
```

产物 `materialization-binding-smoke-v3.json` 是历史 3 行合并 roundtrip smoke，共 4 个样本；它只证明探针调用链和证据格式可运行，不代表容量或正式性能阈值。当前可审计的分离证据见下节；所有产物的 `thresholdStatus` 固定为 `UNAPPROVED`，不能据此标记性能 PASS。before/after 对比必须使用相同命令和独立候选程序集分别生成，并保留各自 identity。

## Materialization/Binding 分离证据（当前候选）

当前候选使用 Release Benchmark DLL 执行以下两次正式采样，每个文件均为固定 seed、固定六列 mapping、两 Provider、`sync/async`、`export-only/import-only/roundtrip`、每组 3 次重复，共 `36/36` 个 `measured` 样本，`0` 个失败；每个 `import-only` 样本均从该 Provider 预先生成且固定不变的 XLSX fixture 读取，fixture SHA-256 写入文档及样本。

| 行数 | 产物 | 样本 | measured | failed | dataset identity | baseline | threshold |
|---:|---|---:|---:|---:|---|---|---|
| 10,000 | `materialization-binding-current-10k-v4.json` | 36 | 36 | 0 | `C59A33C19CA9F35D5B20F14D599021B8BB0675885D43523D1552BF7AEA9E4B33` | `NOT_APPLICABLE_BEFORE` | `UNAPPROVED` |
| 100,000 | `materialization-binding-current-100k-v4.json` | 36 | 36 | 0 | `1F8469C53ECF45121C866368A07CFFC0DAAD01929764E8D4E5D09DB343A60FE6` | `NOT_APPLICABLE_BEFORE` | `UNAPPROVED` |

fixture identity：10K 为 NPOI `2C842F456442F18D16093F3258F6DAC6947A49E8A91F1CDCE2CE013096389284`、MiniExcel `02AAC909C04EEE2F7F29B238A41B8C4B5A8CFC527A9AABE0679A4D62E3845448`；100K 为 NPOI `1D32DD8226C5279A4986ED10DCE348D6BD494E4533DBA7CAC5DED0EEB1D6CA90`、MiniExcel `CC558D08E54F27C5BE1C7E4E21259B49C357DCC380E18340F1210A25E175EFFF`。

100K 中位耗时和分配量如下；这些是当前候选的隔离 E2E 观察，不是 before/after 或正式保留/回滚判定：

| Provider | Mode | Operation | Median ms | Median allocated bytes | Median output bytes |
|---|---|---|---:|---:|---:|
| NPOI | sync | export-only | 3712.14 | 614,206,528 | 3,462,805 |
| NPOI | sync | import-only | 9053.68 | 2,298,234,064 | 3,462,805 |
| NPOI | sync | roundtrip | 11666.26 | 2,912,410,760 | 3,462,805 |
| NPOI | async | export-only | 3356.60 | 614,172,176 | 3,462,804 |
| NPOI | async | import-only | 7374.49 | 2,298,215,232 | 3,462,805 |
| NPOI | async | roundtrip | 12282.01 | 2,912,392,632 | 3,462,805 |
| MiniExcel | sync | export-only | 827.36 | 516,097,936 | 4,646,229 |
| MiniExcel | sync | import-only | 4209.29 | 2,140,128,008 | 4,646,224 |
| MiniExcel | sync | roundtrip | 4521.25 | 2,656,216,912 | 4,646,230 |
| MiniExcel | async | export-only | 1966.23 | 516,215,816 | 4,646,225 |
| MiniExcel | async | import-only | 4049.97 | 2,140,144,088 | 4,646,224 |
| MiniExcel | async | roundtrip | 6227.25 | 2,656,350,320 | 4,646,225 |

10K 与 100K 产物均通过探针自验证：固定行数、固定列数、完整 XLSX 结构、`0` 导入错误和字段级 roundtrip 完整性。当前证据补齐了真实 materializer/binding 的分离 import/export 路径和重复统计；该隔离 E2E 项为 `FIX-002B=CLOSED`，before 明确为 `NOT_APPLICABLE_BEFORE`。同候选 before/after 与 A-I 正式阈值属于 `FIX-002C=BLOCKED_APPROVAL`，不是可继续执行的 OPEN_ACTIONABLE；`thresholdStatus` 仍为 `UNAPPROVED`。

正式命令（均使用上述已构建的 Release Benchmark DLL）：

```text
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --materialization-binding-probe artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/materialization-binding-current-10k-v4.json 10000 3
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --materialization-binding-probe artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/materialization-binding-current-100k-v4.json 100000 3
```

## 已执行

- MiniExcel 100K controlled probe（最终 Release 二进制）：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/miniexcel-100k-current.json`，sync `2475.82 ms`，async `2159.76 ms`，行数和输出完整性通过。
- NPOI/MiniExcel 100K provider comparison（最终 Release 二进制，3 repetitions）：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/provider-comparison-100k-current.json`，sync/async 均通过并记录最终运行程序集 SHA-256。
- RawDate/Relation hotspot after-only（最终 Release 二进制，3 repetitions）：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-after-current.json`，覆盖 RawDate 100K 行的 3/10/30 列和 Relation 1K/10K/100K 行；每个场景样本数正确，输出保留候选程序集 SHA-256。
- Real IO after-only（最终 Release 二进制，3 repetitions）：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/real-io-after-current.json`，24 个场景全部通过，覆盖 CSV/Excel 文件同步异步、延迟流和节流流的 1K/10K/100K 行；100K Excel 输出为约 `2,159,519-2,159,520` bytes，峰值工作集最高约 `382,246,912` bytes。
- Provider 100K before（从基线提交 `94bb52e84ffc70634067b04541857433ef7af9df` 提取源码后独立 Release 构建，3 repetitions）：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/provider-comparison-100k-before-final2.json`。该文件的 header、程序集物理 SHA-256 和基线提交一致，侧车说明见 `before-candidate-manifest.md`。
- Real IO before（同一基线提交、同一 .NET 8.0.30 工作负载，24 个场景通过）：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/real-io-before-final.json`。该 probe 的资源预算状态为 `UNAPPROVED`，因此只作为同机前后观察，不构成受限资源门禁。
- RawDate/Relation before hotspot：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-final.json`。RawDate 3/10/30 列与 Relation 1K/10K 有三次样本；旧关系实现的 100K 场景运行时间异常长，已在产生完整结果前停止，Relation 100K before 标记为 `NOT_VERIFIED`。
- Relation before 受限诊断：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-timeout-60s-20260921.json`。该文件使用诊断版 hotspot worker（每个 worker 60 秒上限）重新执行；RawDate 3/10/30 与 Relation 1K/10K 均各有 3 次完整样本，Relation 100K 在 `60089.3345 ms` 后被终止，记录 `completedRepetitions=0`、`status=not-verified`、退出码 `-1`，探针总退出码 `2`。该文件仅固化安全终止、部分进度和缺失结果，不替代冻结候选的正式 before/after，也不构成 Relation 100K 性能通过。
- Relation before 超时语义复核：`hotspot-before-timeout-20260921-v2.json`、`hotspot-before-timeout-60s-20260921-v2.json` 和 `hotspot-before-timeout-9s-20260921-v2.json`。超时记录现在包含 `killAttempted`、`killSucceeded`、`killError`、截断后的 stdout/stderr 与 `failureCategory=timeout-killed`；只要重复次数不足，summary 的五个正式统计字段全部为 `null`。100 ms 负向运行退出码为 `2`；60 秒运行保留 RawDate/Relation 1K/10K 完整 3 次样本，并在 Relation 100K 超时；9 秒边界运行也保持不完整场景统计为空。
- 最终诊断 DLL 复核：`hotspot-before-timeout-100ms-final-20260921.json`、`hotspot-before-timeout-60s-final-20260921.json` 和 `hotspot-after-final-diagnostic-20260921.json` 均由最终 Benchmark DLL `81249E5E1247272488B10D5FFA5B7CEB5ABC42AE92AEC118D5F014CA5F28D649` 生成。100 ms before 为 `6/6 timeout-killed`、总退出码 `2`；60 秒 before 的 RawDate 3/10/30、Relation 1K/10K 各 3 次通过，Relation 100K 在 `60061.3724 ms` 超时且五项统计为 `null`；after 六场景各 3 次通过、总退出码 `0`。
- Entity/Template after controlled probe：`entity-template-after-current.json`（1K 行、3 repetitions）、`entity-template-after-10k-current.json`（10K 行、2 repetitions）、`entity-template-100k-final.json`（100K 行、1 repetition）和 `entity-template-500k-final.json`（500K 行、1 repetition）。四种模式（Entity/Template sync/async）均完成真实导出→导入 roundtrip，并断言固定 Cell、Merge、List Region、模板公式/图片/明细行保真；100K 结果由 `windows-job-entity-template-100k-final.json` 在 2 CPU 等效/4 GiB Job Object 下复核，四种模式均导入 `100000` 行、错误为 `0`、Job 峰值 `622,891,008` bytes；500K 结果由 `windows-job-entity-template-500k-final.json` 复核，四种模式均导入 `500000` 行、错误为 `0`、Job 峰值 `2,396,901,376` bytes。
- 基线 Relation before 受限证据（Round 15）：从基线提交 `94bb52e84ffc70634067b04541857433ef7af9df` 导出隔离副本后，Benchmark Release 构建为 `0 warning/0 error`；独立构建日志为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-baseline-build-final.log`，命令为 `dotnet build .tmp-baseline-94bb52e/benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore --nologo --no-incremental`，`ExitCode=0`，日志明确 `0 个警告/0 个错误`。使用 Windows JobRunner 2 CPU 等效/4 GiB、60 秒超时和隔离 TEMP，产物为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-hotspot-before-baseline-60s-final.json` 与 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-baseline-60s-final.json`。header 为 `phase=before`、`repetitions=3`、`beforeSource=investigative numeric-cell replay; not a HEAD Reader baseline`，并记录 Benchmark/Abstractions/Core/MiniExcel 程序集物理 identity；RawDate 3/10/30 与 Relation 1K/10K 均各 3 次完成，Relation 100K 未完成，runner 为 `runner-timeout`/exit `-1`，隔离 TEMP 为 `0/0` 且 `deleted`。该结果仅是基线容量/性能边界的 `NOT_VERIFIED` 证据，不补填 Relation 100K 统计，也不构成门禁通过。

## 当前刷新候选 after 观察

当前 Benchmark DLL SHA-256 为 `6CBFFF14075B9544E83A65B95CFE9FB4D9ABCD913DB33437054B81B6CDD82DBD`，以下产物均由该 DLL 生成，并在 header 中绑定当前生产程序集 SHA：

- Provider comparison：`provider-comparison-100k-refresh-20260921-v2-npoi.json` 与 `provider-comparison-100k-refresh-20260921-v2-miniexcel.json`，各 100K 行、同步/异步、3 次重复，所有样本通过。NPOI 中位耗时为 sync `8563.48 ms`、async `4679.13 ms`；MiniExcel 为 sync `4685.51 ms`、async `4811.45 ms`。这些是 after-only 观察，不是正式阈值判定。
- Real IO：`real-io-refresh-20260921-v2.json`，24 场景 × 3 次重复全部通过；header 包含 Benchmark、Abstractions、Core、NPOI 和 MiniExcel 五个程序集 SHA，`budgetStatus=UNAPPROVED`。
- RawDate/Relation：`hotspot-after-refresh-20260921-v2.json`，RawDate 3/10/30 列和 Relation 1K/10K/100K 共六个场景全部通过。
- Entity/Template：`entity-template-after-refresh-20260921-v2.json`，10K 行、四种同步/异步模式、每种两次重复全部通过。

这些刷新产物补齐了当前候选的 after 追踪；same-candidate before/after、Relation 100K before 和正式阈值仍未获批准。真实 materializer/binding 的当前候选隔离 E2E 已由本报告开头的 10K/100K 分离证据闭合为 `FIX-002B=CLOSED`；Relation 100K 等基线边界为 `FIX-002A=VERIFIED_BOUNDARY`，同候选 before/after 与正式阈值为 `FIX-002C=BLOCKED_APPROVAL`。

上述诊断产物由新增超时处理逻辑的 Benchmark 程序生成，最终诊断 DLL SHA 为 `81249E5E1247272488B10D5FFA5B7CEB5ABC42AE92AEC118D5F014CA5F28D649`，与冻结的 `manifest-current-20260921.md` 中 `6CBFFF14075B9544E83A65B95CFE9FB4D9ABCD913DB33437054B81B6CDD82DBD` 不同；因此最终诊断文件是独立证据，不被静默归入当前候选的正式 after 产物。此前 `4CBB...` 与 `1156...` 文件保留为历史中间诊断，不用于本轮结论。正式发布判断仍以冻结候选及其 `NOT_VERIFIED` 缺口为准。

Entity/Template 的 before 记为 `NOT_APPLICABLE_BEFORE`：基线提交 `94bb52e84ffc70634067b04541857433ef7af9df` 不包含 `ExcelEntityLayout`、`IExcelEntity*` 或 `ExportEntity/ImportEntity` 公共契约，无法运行同一 workload；没有用 List/workbook 路径伪造 before 数字。该项仍不能替代正式 A-I 阈值审批。

Entity/Template after 中位数（同一候选，仅供容量和归因观察）：

| 行数 | 模式 | 样本数 | 中位数 ms | 平均分配 bytes | 平均输出 bytes |
|---:|---|---:|---:|---:|---:|
| 1K | entity-sync | 3 | 132.56 | 21,376,344 | 21,584 |
| 1K | entity-async | 3 | 118.29 | 21,203,024 | 21,584 |
| 1K | template-sync | 3 | 129.50 | 26,506,952 | 22,895 |
| 1K | template-async | 3 | 132.63 | 26,572,296 | 22,895 |
| 10K | entity-sync | 2 | 1,814.72 | 187,442,500 | 166,213 |
| 10K | entity-async | 2 | 1,665.56 | 184,992,396 | 166,212 |
| 10K | template-sync | 2 | 1,438.99 | 222,178,692 | 167,564 |
| 10K | template-async | 2 | 1,605.51 | 221,992,376 | 167,564 |
| 100K | entity-sync | 1 | 8,808.92 | 1,833,573,272 | 1,603,346 |
| 100K | entity-async | 1 | 6,513.30 | 1,821,339,480 | 1,603,347 |
| 100K | template-sync | 1 | 6,107.66 | 2,195,789,552 | 1,605,216 |
| 100K | template-async | 1 | 8,640.83 | 2,195,837,528 | 1,605,216 |
| 500K | entity-sync | 1 | 41,430.29 | 9,067,540,464 | 8,011,619 |
| 500K | entity-async | 1 | 33,025.76 | 9,053,904,248 | 8,011,619 |
| 500K | template-sync | 1 | 33,141.71 | 10,920,361,992 | 8,015,454 |
| 500K | template-async | 1 | 30,625.59 | 10,920,553,944 | 8,015,454 |

这些是单候选观察，不代表已批准阈值；100K/500K 样本均只有一次重复，不能据此做稳定性或回归结论。Entity/Template 1M 当前重跑已形成 `windows-job-entity-template-1m-current-20260921.json`，但在默认 `MaxWorksheetBytes=64 MiB` 的 ZIP 预检阶段以资源限制异常退出，未生成 roundtrip artifact，状态仍为 `NOT_VERIFIED`，没有填入伪造的性能数字。

资源补充：`windows-job-streaming-1m-timeout-isolated-final.json` 记录了一次设置 60 秒超时和隔离 TEMP 的强制终止演练；终止后观察到 2 个临时文件、373,260,288 字节，目录清理状态为 `deleted`。该证据只证明清理链路可观测，不构成 1M 性能或完整资源门禁通过。

候选说明：hotspot 文件中的 assembly SHA-256 是 Release 文件的物理哈希；API candidate manifest 的 assembly 表使用 API 工具定义的 canonical identity hash。两套值均已在 `artifacts/candidates/.../manifest.md` 中并列记录，不能互相替换。

after-only 中位数摘要：

| Provider | Mode | Median ms | Median allocated bytes | Median rows/s |
|---|---|---:|---:|---:|
| NPOI | sync | 7724.08 | 2164461424 | 12947 |
| NPOI | async | 8124.66 | 2164442288 | 12308 |
| MiniExcel | sync | 1785.57 | 1250071200 | 56004 |
| MiniExcel | async | 2287.67 | 1250202328 | 43713 |

## Provider 100K 前后观察

以下百分比只是同机、同 TFM、相同 100K 请求的观测差异，未套用正式阈值；`before` 是基线提交的独立构建，`after` 是当前候选 Release 构建。负值表示 after 用时/分配较低，不能直接解释为发布收益。

| Provider | Mode | Before ms | After ms | 时间变化 | Before allocated | After allocated | 分配变化 | Before peak | After peak |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| NPOI | sync | 7448.30 | 7724.08 | +3.70% | 2164454712 | 2164461424 | +0.00% | 674283520 | 673402880 |
| NPOI | async | 4734.42 | 8124.66 | +71.61% | 2164512192 | 2164442288 | -0.00% | 667807744 | 644648960 |
| MiniExcel | sync | 1423.58 | 1785.57 | +25.43% | 1383158752 | 1250071200 | -9.63% | 134995968 | 134635520 |
| MiniExcel | async | 2226.98 | 2287.67 | +2.73% | 1383280808 | 1250202328 | -9.63% | 156667904 | 160268288 |

Provider probe 的行数、请求形状和输出完整性断言均通过。NPOI async 的 elapsed/peak 观测出现回退，说明不能用单一 sync 结果代表整个 Provider；正式保留/回滚仍需计划中的 A-I 阈值和更稳定的重复运行。

## Real IO 100K 前后观察

下表只列 100K 行的中位数和峰值；CSV 输出字节在前后相同，Excel ZIP 输出在重复运行间有小幅字节波动，probe 本身的 round-trip/结构断言通过，因此不把 Excel 字节相等当作本表的结论。

| Scenario | Before ms | After ms | 时间变化 | Before allocated | After allocated | Before peak | After peak |
|---|---:|---:|---:|---:|---:|---:|---:|
| csv-file-sync | 331.26 | 230.82 | -30.32% | 56170693 | 56170376 | 74010624 | 73310208 |
| csv-file-async | 227.56 | 154.48 | -32.12% | 75417341 | 75433112 | 76939264 | 74928128 |
| excel-file-sync | 1586.69 | 3150.74 | +98.57% | 740394253 | 740394040 | 381460480 | 381095936 |
| excel-file-async | 2036.25 | 2915.24 | +43.17% | 740476776 | 740477115 | 383041536 | 382275584 |
| csv-delayed-async | 788.84 | 817.62 | +3.65% | 83551328 | 83551232 | 84750336 | 84615168 |
| csv-throttled-async | 806.16 | 784.86 | -2.64% | 83551232 | 83551232 | 84811776 | 84484096 |
| excel-delayed-async | 2835.35 | 2832.99 | -0.08% | 748671675 | 748671773 | 375984128 | 382840832 |
| excel-throttled-async | 2387.22 | 3359.43 | +40.73% | 748754077 | 748808179 | 379748352 | 380436480 |

Real IO before/after 的同机证据已存在，但预算仍为 `UNAPPROVED`，且没有生产机器或外部 CI 复核；不能把这些数字标记为正式性能门禁通过。

Hotspot after-only 中位数摘要：

| Workload | Size | Median ms | Median allocated bytes | Median index count |
|---|---:|---:|---:|---:|
| RawDate | 3 columns × 100K rows | 239.88 | 233365416 | 300000 |
| RawDate | 10 columns × 100K rows | 1082.71 | 794711400 | 1000000 |
| RawDate | 30 columns × 100K rows | 3924.33 | 2482397632 | 3000000 |
| Relation | 1K rows | 0.52 | 176600 | 1000 |
| Relation | 10K rows | 2.65 | 1664392 | 10000 |
| Relation | 100K rows | 28.12 | 15655296 | 100000 |

RawDate hotspot 的同候选前后样本可对照如下：

| Workload | Size | Before ms | After ms | 时间变化 | Before allocated | After allocated | Before index | After index |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| RawDate | 3 columns × 100K rows | 160.16 | 239.88 | +49.78% | 226130328 | 233365416 | 300000 | 300000 |
| RawDate | 10 columns × 100K rows | 731.89 | 1082.71 | +47.93% | 770691472 | 794711400 | 1000000 | 1000000 |
| RawDate | 30 columns × 100K rows | 7582.71 | 3924.33 | -48.25% | 2410412024 | 2482397632 | 3000000 | 3000000 |
| Relation | 1K rows | 230.23 | 0.52 | -99.77% | 16272328 | 176600 | 1000 | 1000 |
| Relation | 10K rows | 18834.92 | 2.65 | -99.99% | 1602866424 | 1664392 | 10000 | 10000 |
| Relation | 100K rows | NOT_VERIFIED | 50.56 | N/A | N/A | 15655272 | N/A |

## Cold-plan-build Tail Latency 前后观察

基线与当前候选均使用独立 Release 二进制，在同一 .NET 8.0.30 环境执行 100 次操作、64 次 warmup、5 次重复，覆盖并发 1/4/16/64。样本定义为从队列提交时间戳到 mapping plan 完成；结果不是 Provider 或 Entity/Template 文件端到端延迟。预算状态为 `UNAPPROVED`，因此以下仅作同机观察。

| Concurrency | Before p50 us | After p50 us | Before p95 us | After p95 us | Before p99 us | After p99 us | Before ops/s | After ops/s |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 1 | 1039 | 925 | 1967 | 1801 | 2181 | 1946 | 47380.79 | 52195.34 |
| 4 | 359 | 406 | 641 | 721 | 696 | 806 | 141530.80 | 130534.67 |
| 16 | 379 | 351 | 2726 | 2816 | 2820 | 2932 | 85919.51 | 42524.96 |
| 64 | 315 | 310 | 550 | 2290 | 599 | 2377 | 59861.60 | 65666.78 |

并发 1 的候选 p50/p95/p99 改善，并发 4 的 p50/p95/p99 与吞吐略有回退；并发 16/64 的尾延迟受单次抖动影响明显，不能据此归因于实现收益或回归。原始文件为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/tail-latency-before-100.json` 和 `tail-latency-after-100.json`，两者的 `gitHead`/`diffIdentity` 分别指向基线提交与当前工作树。

本轮首先尝试运行 `DynamicPlanBenchmarks` 的普通 BenchmarkDotNet `ShortRun`。自动生成的 BDN 项目在 restore 阶段因 NuGet SSL/凭据错误 `NU1301` 失败，报告仅含 `NA`；随后改用无需自动生成 restore 的 InProcess toolchain 重新执行，结果见下节。失败的普通模式不能替代成功的 InProcess 证据，也不构成性能门禁。

## DynamicPlan before/after 观察

随后使用同一 BenchmarkDotNet 版本和 `InProcessEmitToolchain`，分别对基线提交和当前工作树执行 8 个组合（工厂创建、冷构建、缓存命中、缓存未命中；计划数 100/500；ShortRun 三次迭代）。结构化结果见 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/dynamic-plan-before-after.json`。

| Method | Count | Before mean us | After mean us | Before KB | After KB |
|---|---:|---:|---:|---:|---:|
| FactoryCreation | 100 | 873.1 | 1207.0 | 8055.62 | 8055.62 |
| BuildCold | 100 | 4656.3 | 5977.0 | 754.36 | 754.36 |
| CacheHit | 100 | 1957.4 | 2762.0 | 436.74 | 436.74 |
| CacheMiss | 100 | 4279.9 | 7454.0 | 754.36 | 756.70 |
| FactoryCreation | 500 | 4973.6 | 7556.0 | 40278.09 | 40278.09 |
| BuildCold | 500 | 21668.9 | 22226.0 | 3716.88 | 3716.88 |
| CacheHit | 500 | 5371.6 | 5898.0 | 2177.37 | 2177.37 |
| CacheMiss | 500 | 13427.7 | 13542.0 | 3716.88 | 3716.88 |

该结果是 mapping plan microbenchmark，不是 Provider 文件端到端基准；短迭代、不可用 power-plan 和 `UNAPPROVED` 预算使其只能作为方向性 before/after 证据。它不能单独证明动态计划优化收益，也不能替代 A-I 正式阈值判定。

## PropertyAccessor before/after 观察

使用同一 InProcess ShortRun 对编译 getter/setter 与反射 getter/setter 各执行三次迭代。结构化结果见 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/property-accessor-before-after.json`。

| Method | Before ns | After ns | Before KB | After KB |
|---|---:|---:|---:|---:|
| CompiledGetter | 2.269 | 2.774 | 0 | 0 |
| ReflectionGetter | 12.475 | 8.534 | 0 | 0 |
| CompiledSetter | 2.599 | 2.523 | 0 | 0 |
| ReflectionSetter | 14.981 | 14.588 | 0 | 0 |

该结果仅说明 accessor microbenchmark 的方向，不能替代真实 materializer、Provider 或 Entity/Template E2E；未批准阈值和短迭代限制仍适用。

## GenericSheetDispatch before/after 观察

使用同一 InProcess ShortRun 对 NPOI 多 Sheet 导入/导出分派执行 1 Sheet 和 8 Sheet 两组 before/after。结构化结果见 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/generic-dispatch-before-after.json`。

| Method | Sheets | Before ms | After ms | Before KB | After KB |
|---|---:|---:|---:|---:|---:|
| Export | 1 | 4.905 | 4.263 | 1262.53 | 1262.89 |
| Import | 1 | 3.093 | 2.849 | 666.75 | 667.09 |
| Export | 8 | 4.518 | 4.586 | 1781.76 | 1782.43 |
| Import | 8 | 2.472 | 2.682 | 1280.86 | 1280.94 |

该结果覆盖真实 NPOI 多 Sheet 分派，但仍是小数据 microbenchmark，不代表 Entity/Template 或 100K E2E。BenchmarkDotNet 同时报告 1/8 Sheet Import 分别有 3/10 个 diagnostic exception events，且 before/after 均出现。最小公开基准探针确认异常类型为 `System.MissingMethodException`，消息为 `Attempted to access a missing method.`，堆栈落在 `NPOI.XSSF.UserModel.XSSFFactory.CreateDocumentPart(...)`；两种 Sheet 数量的导入都完成且结果行数正确。精确探针证据见 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/generic-dispatch-exception-probe.json`。该 NPOI 内部反射异常不是“无异常”通过，也未归因于本任务改动，因此不纳入正式门禁且不改生产逻辑。未批准阈值和短迭代限制仍适用。

## Relation 100K before 长时边界诊断

为记录 `FIX-002A=VERIFIED_BOUNDARY` 的基线边界，使用隔离 `.tmp-baseline-94bb52e` 的 Release Benchmark DLL，在 Windows Job Object `2 CPU`、`4 GiB`、隔离 TEMP、外层 `timeoutMs=900000`，Hotspot worker `600000ms`、`before`、3 repetitions 条件下执行：

- runner：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-hotspot-before-baseline-10m-relation100k-final.json`
- child：`artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/hotspot-before-baseline-10m-relation100k-final.json`
- child 身份为 `.tmp-baseline-94bb52e` Benchmark DLL，物理 SHA-256 为 `65AEA685D40829C81F3221B88DE813EC4B782521EAFE4E616873B5736B63BF25`；四个程序集 identity 与隔离基线实际文件一致。
- runner 最终为 `runner-timeout`、runner/child exit `-1/-1`、`runnerTimedOut=true`，Job 峰值 `389,754,880` bytes；隔离 TEMP 为 `0/0/deleted`，stdout/stderr 为空。
- child 已刷盘 15 个完整样本和 5 个 summary：RawDate 3/10/30 列各 3 次，Relation 1K/10K 各 3 次。Relation 1K 中位耗时 `379.3835 ms`、分配 `16,272,296` bytes；Relation 10K 中位耗时 `21,876.5469 ms`、分配 `1,602,866,424` bytes。Relation 100K 没有样本、没有 summary，输出在 Relation 10K summary 后结束。

该长时运行只强化了旧关系实现的资源/长尾边界证据，Relation 100K before 仍为 `NOT_VERIFIED` 边界，不填充伪造统计，也不构成 A-I 性能门禁通过；`beforeSource` 仍明确为 investigative numeric-cell replay，不冒充 HEAD Reader baseline。Sol Medium 已独立复核身份、字段、样本数和清理状态。该边界归类为 `FIX-002A=VERIFIED_BOUNDARY`；真实 materializer/binding 当前候选隔离 E2E 归类为 `FIX-002B=CLOSED`，同候选 before/after 与正式阈值归类为 `FIX-002C=BLOCKED_APPROVAL`。

## Round 43：Performance Budget v1 精准确认

用户批准 `Performance Budget v1` 后，仅对先前观察到明显超过 `25%` 的五个 `100K E2E` 场景执行一次 `5 repetitions` 确认；没有重跑完整 A-I、Relation 100K、RawDate hotspot、tail latency 或 c16/c64 同质实验。结构化总产物为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/targeted-performance-budget-v1-5rep.json`，原始场景产物为同目录下的 `targeted-provider-npoi-100k-5rep.json`、`targeted-provider-miniexcel-100k-5rep.json` 和三个 `targeted-realio-*-100k-5rep.json`。

| Target | Before median ms | After median ms | Delta | Samples | Budget result |
|---|---:|---:|---:|---:|---|
| NPOI async provider E2E | 4,734.42 | 13,784.323 | `+191.15%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| MiniExcel sync provider E2E | 1,423.58 | 7,599.0974 | `+433.80%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| Real IO `excel-file-sync` | 1,586.69 | 6,935.1346 | `+337.08%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| Real IO `excel-file-async` | 2,036.25 | 11,111.6705 | `+445.69%` | `5/5` | `REGRESSION_OVER_BUDGET` |
| Real IO `excel-throttled-async` | 2,387.22 | 5,636.8019 | `+136.12%` | `5/5` | `REGRESSION_OVER_BUDGET` |

所有五个目标均完成五次样本、输出字节非零且场景状态为 `passed`；当前候选程序集 hash 记录在各原始产物 header 中（Benchmark `419DC8A4F2D7907BEBFFE8351A6AD671FCDD552397CF0C4134F4B824DFC1202F`、NPOI `E87FA5777867FB1D8B22B824EB580373306C75B95F311447F554BE70390F010E`、MiniExcel `3FDDC6AF4EE936A44E783B1208CA6B33C59705419585D8A353110980C4D7DCEB`）。两个 Provider probe 顶层进程并发启动，原始产物保留该执行拓扑；Real IO 三场景串行执行。聚合 artifact 的 `beforeEvidence` 显式链接既有 `provider-comparison-100k-before-final2.json` 与 `real-io-before-final.json`，没有在本轮重跑或合成 before 样本。该轮因此确认五项回退超过批准预算，性能门禁状态为 `OPEN_ACTIONABLE`，不能继续沿用旧的 `BLOCKED_APPROVAL` 或宣称性能通过。

## Gate

Provider、Real IO、RawDate 3/10/30 列、Relation 1K/10K、DynamicPlan、PropertyAccessor、GenericSheetDispatch、cold-plan-build tail latency 和当前候选 materializer/binding 分离 E2E 已有可审计观测；Relation 100K before 与 Entity/Template before 的技术边界归类为 `FIX-002A=VERIFIED_BOUNDARY`，真实 materializer/binding 当前候选隔离 E2E 归类为 `FIX-002B=CLOSED`。本轮批准的五个 `100K E2E` 目标均超过 `Performance Budget v1`，因此 `FIX-002C=OPEN_ACTIONABLE`，需针对实际回退完成最小修复并再次由独立审查确认；本轮不是完整 A-I 结论。资源边界归类为 `FIX-003A/B=VERIFIED_BOUNDARY`，本地 staging/cleanup 和 100K/500K/1M 代表性 smoke 归类为 `FIX-003C=CLOSED`/`FIX-003D=LOCAL_PROFILE_EVIDENCE`；真实生产机器和外部 CI 仍归类为 `FIX-003D/E=BLOCKED_EXTERNAL`。因此当前 Performance=`OPEN_ACTIONABLE`、Resource=`VERIFIED_BOUNDARY`、External=`BLOCKED_EXTERNAL`、Release=`BLOCKED_APPROVAL`、Goal=`IN_PROGRESS`；本轮不能标记 `COMPLETED`。
