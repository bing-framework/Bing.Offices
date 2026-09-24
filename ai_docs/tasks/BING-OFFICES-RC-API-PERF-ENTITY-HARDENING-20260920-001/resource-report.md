# 资源与容量报告

NPOI `TempFile/excel-100k/c1` 受控场景已通过：

- 产物：`artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/npoi-tempfile-excel-100k.json`
- 100K rows，2 次 measurement，错误 `0`，临时文件清理后 `leftoverFiles=0`
- peak working set `1396236288` bytes；temporary disk after `0`
- 该进程未施加 OS 级 2CPU/4GiB 限制，不能作为受限资源门禁通过证据

MiniExcel 100K controlled probe 见 benchmark report。Entity 10,000 行边界和越界零输出由 NPOI 双 TFM 职责级测试覆盖。

## Windows Job Object 受限证据

本轮新增无第三方依赖的 `tests/Bing.Offices.WindowsJobRunner`，通过 Windows Job Object 设置作业内存上限和 CPU 等效上限。当前宿主机为 22 个逻辑处理器，`2 CPU` 映射为 `910/10000` CPU rate units，作业内存上限为 `4,294,967,296` bytes。runner smoke、child path、退出码和 Job Object 峰值均写入以下 JSON：

- smoke：`artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-runner-smoke.json`，Job 绑定成功，退出码 `0`。
- 100K：`windows-job-100k-runner.json`，Job 峰值 `1,389,539,328` bytes；child `errorCount=0`、`temporaryDiskBytesAfter=0`、`leftoverFiles=0`，退出码 `0`。
- 500K：`windows-job-500k-runner.json`，Job 峰值 `3,243,692,032` bytes；NPOI 导出发生 `OutOfMemoryException`，两次操作 `operationBytes=0`、`temporaryDiskBytesAfter=0`、`leftoverFiles` 未产生，退出码 `1`。
- 1M：`windows-job-1m-runner.json`，Job 峰值 `3,245,178,880` bytes；NPOI 导出发生 `OutOfMemoryException`，两次操作 `operationBytes=0`、`temporaryDiskBytesAfter=0`，退出码 `1`。

这证明了本机 Windows 作业限制确实生效，并明确暴露 NPOI 500K/1M XLSX DOM 导出的容量失败；它不是资源门禁 PASS。生产机器、外部 CI、不同容器运行时和可批准的容量阈值仍未验证，FIX-003 继续保持 `NOT_VERIFIED`。

此前无 OS 限制的 500K 受控前置探针未通过容量观察，不能标记 PASS：

- 命令：`dotnet tests/Bing.Offices.ResourceProbe/bin/Release/net8.0/Bing.Offices.ResourceProbe.dll --staging-scenario artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/npoi-tempfile-500k.json TempFile excel-100k 1 500000`
- 运行期间同一进程工作集约 `3.8 GiB`、Private Memory 约 `4.11 GiB`，超过本任务内部 2 GiB 子进程保护线，并接近/超过目标 4 GiB 约束；随后发送取消信号，进程以退出码 `1` 结束。
- 该 direct staging 模式只向 stdout 输出结果，不会自动写入目标 JSON；本次没有可宣称 PASS 的产物，也没有把中断结果伪装成完整性通过。
- 该次运行没有 OS 级限制，只作为历史容量风险记录；本轮已使用上节 Job Object runner 重新执行 500K/1M。两种证据均不能标记容量 PASS。

上面的 `windows-job-500k-runner.json` / `windows-job-1m-runner.json` 是流式路径接入前的 DOM 历史证据，保留用于说明容量风险，不能代表当前候选的最终结果。

## NPOI SXSSF 流式列表补充

为满足大列表资源门禁，在不改变高级布局能力边界的前提下，当前候选对无模板、XLSX、纯列表且数据源为已知集合的请求使用 NPOI `SXSSFWorkbook`。合并、自动列宽、图表、动态列、模板区域、`[MergeColumns]`/`[WrapText]` 等需要回溯或高级布局的请求仍走原有 DOM 路径，未静默扩大支持范围。

下表是流式路径接入后的首次复跑记录，均由 Windows Job Object runner 施加 `requestedCpuEquivalent=2`、`requestedMemoryLimitBytes=4,294,967,296`；每次独立执行 `TempFile`/`excel-100k`，`warmup=1`、`measurement=2`、并发 `1`。修复属性级 `[MergeColumns]` 误判和正文样式顺序后，当前候选的最终复跑见下一节。ResourceProbe 的完整 staging JSON 嵌入对应 runner JSON 的 `stdout` 字段，runner 本身和 stdout 均保留在 `artifacts/resource/...`。

| 场景 | Runner 证据 | 状态 | Job 峰值 | operationBytes | errors | temp/leftover |
|---|---|---|---:|---:|---:|---:|
| 100K | `windows-job-streaming-100k-rerun.json` | `PASS` | `396,312,576` | `84,820,165` | `0` | `0 / 0` |
| 500K | `windows-job-streaming-500k-rerun.json` | `PASS` | `1,479,540,736` | `424,133,350` | `0` | `0 / 0` |
| 1M | `windows-job-streaming-1m-rerun.json` | `PASS` | `2,462,220,288` | `848,270,803` | `0` | `0 / 0` |

100K 另有 `ExcelWorkbookRequestTest.Export_LargeSimpleXlsx_ShouldRoundTripRowsAndStyles` 双 TFM直接验证 100,000 行、首尾值、数值格式和表头样式。500K/1M 本轮证据验证了受限作业内完整双 measurement、无异常、非零输出和临时文件清理；没有再将未独立解析的全部单元格内容表述为逐单元格 round-trip 证据。

## 修复后当前候选复跑

修复属性级 `[MergeColumns]` 误判和 SXSSF 正文样式顺序后，使用当前 Release 二进制重新执行同一 Job Object 资源探针：

| 场景 | Runner 证据 | 状态 | Job 峰值 | operationBytes | errors | temp/leftover |
|---|---|---|---:|---:|---:|---:|
| 100K | `windows-job-streaming-100k-final.json` | `PASS` | `397,295,616` | `84,820,194` | `0` | `0 / 0` |
| 500K | `windows-job-streaming-500k-final.json` | `PASS` | `1,474,715,648` | `424,133,356` | `0` | `0 / 0` |
| 1M | `windows-job-streaming-1m-final.json` | `PASS` | `2,356,469,760` | `848,270,820` | `0` | `0 / 0` |

本轮新增 `ExcelWorkbookRequestTest.Export_LargeMergedList_ShouldPreserveMergedRegion`、`Export_LargeSimpleXlsx_CancellationShouldPreserveTargetAndCleanTemporaryFile`，并将大型样式测试加入 Formatter 与 BodyStyle 冲突合同；当前最终 NPOI 双 TFM 全量均为 `566/566 PASS`。上述大规模资源证据覆盖无模板纯列表 SXSSF，以及 Entity/Template 的 100K/500K 单次运行；完整的 SXSSF 强制终止矩阵、Entity/Template 1M、并发和生产/外部环境继续保持 `NOT_VERIFIED`。

### 强制终止与隔离 TEMP 演练

使用新增强制终止能力的 `WindowsJobRunner`，设置 `BING_OFFICES_JOB_TIMEOUT_MS=60000` 和 `BING_OFFICES_JOB_ISOLATE_TEMP=true`，对当前 Release `ResourceProbe` 的 1M `TempFile/excel-100k/c1` 场景执行一次可控超时演练。结构化产物为 `windows-job-streaming-1m-timeout-isolated-final.json`：`status=runner-timeout`、`runnerTimedOut=true`、`childExitCode=-1`，Job 峰值 `1,235,165,184` bytes；隔离 TEMP 在终止后仍观测到 `2` 个文件、`373,260,288` bytes，随后 `tempCleanupStatus=deleted`。该结果证明系统临时目录的强制终止观测和清理状态可追溯，但 1M 场景本身是超时而非通过，完整取消/失败提交矩阵、Entity/Template 1M、生产机器和外部 CI 仍为 `NOT_VERIFIED`。

## 公开入口与失败路径复核

当前 Release `ResourceProbe` 在 Windows Job Object（2 CPU 等效、4 GiB）下完成公开入口矩阵：

- `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entrypoints-final.json` 及其 `staging-entrypoints-final.json`：`Export`、`ExportAsync`、`ExportToFile`、`ExportToFileAsync` × 1K/10K/100K 行 × 3 次重复，共 `36/36 PASS`。
- 所有记录均完成输出重开、目标提交和临时文件清理；最大临时文件峰值为 `2`，测试结束后临时文件数为 `0`。
- `windows-job-entrypoint-failure-final.json` 及 `staging-entrypoint-failure-final.json` 覆盖 Setup、Sampler、Parser 三类注入失败；三类均按预期失败，目标/流清理均为成功，截断 ZIP 被拒绝且未留下残留文件。

### 500K 公开入口矩阵补充

使用参数化 `--staging-entrypoints` 对 `rowCount=500000` 做 `TempFile` 公开入口补充。child 命令参数为 `dotnet tests/Bing.Offices.ResourceProbe/bin/Release/net8.0/Bing.Offices.ResourceProbe.dll --staging-entrypoints artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-500k-parameter-final.json 500000`；runner 使用 2 CPU 等效、4 GiB、`timeoutMs=900000` 和隔离 TEMP。runner 产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entrypoints-500k-parameter-final.json`，入口矩阵产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-500k-parameter-final.json`。

产物记录 `Export`、`ExportAsync`、`ExportToFile`、`ExportToFileAsync` 四个入口各重复 3 次，共 `12/12 PASS`；header 的 `requestedRowCount`/`rowCounts` 为 `500000`/`[500000]`，场景数为 `12`。所有输出均完成工作簿重开与解析，文件入口的句柄重开也成功；`LeftoverFiles=0`、`TemporaryDiskBytesAfter=0`，runner 终止后 `tempFilesAfterTermination=0`、`tempBytesAfterTermination=0` 且 `tempCleanupStatus=deleted`。runner 记录 `peakJobMemoryBytes=2098900992`。

参数化的短 smoke 产物 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-1-parameter-final.json` 记录 `rowCount=1`，四个入口各重复 3 次，共 `12/12 PASS`；它只是参数化行为的本地 smoke，不是资源门禁证据。

旧的非正式产物 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-500k-parameter.json` 只有 header 和 3 条 `Export` 记录，预期场景数为 `12`、实际记录数为 `3`，应标记为 `3/12 PARTIAL, NOT_EVIDENCE`。正式 500K 证据仅指带 `-final.json` 后缀的 runner 与入口矩阵产物；使用 glob 或人工选取时必须排除该旧文件。

该证据只覆盖 `TempFile` 和四个公开入口的 500K 参数化补充，不是完整的 `Memory/TempFile/Hybrid × 场景 × 并发` 矩阵，也不覆盖生产环境、外部 CI 或正式容量/性能阈值判定。

### 1M 公开入口边界补充

使用参数化入口命令 `dotnet tests/Bing.Offices.ResourceProbe/bin/Release/net8.0/Bing.Offices.ResourceProbe.dll --staging-entrypoints artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-1m-parameter-final.json 1000000`，在 Windows Job Object 2 CPU 等效、4 GiB、`timeoutMs=1800000` 和隔离 TEMP 条件下运行。runner 产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entrypoints-1m-parameter-final.json`，入口产物为 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/staging-entrypoints-1m-parameter-final.json`。

入口产物完整记录四个公开入口各 3 次，共 12 条，全部为 `failed`；runner 状态为 `child-failed`、exit 为 `1`。所有失败均发生在 `TryReopenWorkbook` 的 NPOI DOM 工作簿重开阶段，异常为 `OutOfMemoryException`；最大输出约 `20,261,647` bytes，`peakJobMemoryBytes=3301044224`，`leftover/temp-after=0`，隔离 TEMP cleanup=`deleted`。

该结果是 1M 公开入口 roundtrip 的结构化 `NOT_VERIFIED`/容量失败证据，不是 PASS，不替代既有 SXSSF 单场景 1M PASS，也不构成产品支持决策；完整策略矩阵、生产环境、外部 CI 和正式阈值仍未验证。

## Staging 策略结构矩阵

`artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-staging-matrix-1k-final.json` 与 `staging-matrix-1k-final.json` 在同一 Job Object 预算下完成 `Memory`、`TempFile`、`Hybrid` × `excel-100k`、`failure-double-dom`、`template-image-style` × 并发 `1/4/16/64`，共 `36/36 PASS`。所有场景 `errorCount=0`、`leftoverFiles=0`、`temporaryDiskBytesAfter=0`，最大实际并发为 `1`，最大排队深度为 `63`，证明请求并发会被资源门控排队而不会突破 DOM 并发上限。

同一 Release 二进制在 2 CPU 等效/4 GiB Job Object 下完成 10K staging 矩阵：`windows-job-staging-matrix-10k-current.json` 的 runner 与 `staging-matrix-10k-current.json` 均记录 `36/36 PASS`，覆盖相同的三种策略、三类场景和 `1/4/16/64` 请求并发；`maxActualParallelism=1`、`maxQueuedRequests=63`、错误和残留均为 `0`。隔离 TEMP 终止后 `tempFilesAfterTermination=0`、`tempBytesAfterTermination=0`、`tempCleanupStatus=deleted`。审批人和时间仍为空，产物 `approvalStatus=BLOCKED`。

该矩阵的审批人和审批时间未提供，产物 `approvalStatus=BLOCKED`，所以只能作为结构与清理证据，不能作为已批准容量阈值的发布门禁。`100K` 矩阵仍只保留部分记录，不被表述为完整通过；10K 已形成完整 36 格，但不替代 100K/500K/1M 的容量、生产和外部环境门禁。100K/500K/1M 的完整受限证据仍以本报告上方的无模板纯列表 SXSSF 三个独立 Job Object 运行作为当前候选大数据证据。

随后使用相同 Release 二进制启动 100K staging 全矩阵，产物为 `windows-job-staging-matrix-100k-current.json` 与 `staging-matrix-100k-current.json`。runner 设置 2 CPU 等效、4 GiB、1 小时超时和隔离 TEMP；在超时前完成 `10/36` 格，均为 `passed`、错误 `0`、残留 `0`，已覆盖 `Memory` 策略下 Excel/失败/模板场景的部分并发组合。runner 最终为 `runner-timeout`（`timeoutMs=3600000`），峰值 Job 内存 `1,574,567,936` bytes，终止后隔离 TEMP `0` 文件/`0` 字节且清理状态为 `deleted`。该文件保留为原始矩阵与清理证据；后续当前候选独立 runner 已将合并覆盖提升到 `36/36`，两个 Excel c64 场景通过更长的串行观察窗口完成。

为隔离长矩阵超时影响，又单独执行了三个 100K `TempFile/c1` 场景：`excel-100k`、`failure-double-dom`、`template-image-style`。三个 runner 均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `396,922,880`、`1,443,917,824`、`1,247,277,056` bytes；每个隔离 TEMP 结束后均为 `0` 文件/`0` 字节且 `cleanup=deleted`。这组结果是独立的单并发补充证据，不单独改变 100K 全矩阵审批状态。

同样补跑了三个 100K `Hybrid/c1` 场景：`excel-100k`、`failure-double-dom`、`template-image-style`。三个 runner 均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `425,975,808`、`1,400,291,328`、`1,169,821,696` bytes；隔离 TEMP 均为 `0` 文件/`0` 字节且 `cleanup=deleted`。这些样本补充 Hybrid 单并发行为，但不单独改变完整矩阵审批状态。

继续补跑三个 100K `TempFile/c4` 场景：`excel-100k`、`failure-double-dom`、`template-image-style`。三个 runner 均 `passed`、child exit `0`、`errorCount=0`，Job 峰值分别为 `400,752,640`、`1,445,371,904`、`1,291,882,496` bytes；隔离 TEMP 均为 `0` 文件/`0` 字节且 `cleanup=deleted`。这些样本补充 TempFile 四请求并发行为，但不单独改变完整矩阵审批状态。

随后补跑 100K `Hybrid/c4` 与 `Hybrid/c16` 的 Excel、失败工作簿和模板图片样式场景；六项均 `passed`，c4 Job 峰值分别为 `419,799,040`、`1,453,576,192`、`1,210,265,600` bytes，c16 Job 峰值分别为 `432,259,072`、`1,511,641,088`、`1,210,667,008` bytes，错误均为 `0`，隔离 TEMP 均为 `0` 文件/`0` 字节且 `cleanup=deleted`。另补齐 `Memory/template-image-style/c16` 与 `c64`，均 `passed`，Job 峰值为 `1,384,951,808` 与 `1,374,752,768` bytes；`TempFile/c16` 三场景也均 `passed`，Job 峰值为 `398,258,176`、`1,492,705,280`、`1,451,864,064` bytes。上述通过场景均不单独改变完整矩阵审批状态。

继续执行的 `Hybrid/c64` 和 `TempFile/c64` 各三个场景曾在 `900000ms` runner 门限超时，状态为 `runner-timeout`、child exit `-1`；Hybrid Job 峰值分别为 `433,729,536`、`1,508,929,536`、`1,228,423,168` bytes，TempFile 分别为 `400,572,416`、`1,505,271,808`、`1,477,602,368` bytes。隔离 TEMP 均 `cleanup=deleted`；TempFile 三场景终止后分别观测 `3/1/1` 个文件及 `66,580,121/0/28,756,329` 字节。该组记录保留为旧门限的 timeout 证据，不作为成功或失败替代。

随后将观察窗口延长至 `3600000ms`，使用当前 Release ResourceProbe（NPOI 依赖 SHA-256 `D357F9A5EFA4D24327BBD8B856AAC7E4875AA0E4DEA7852FE7F50E378FFD75AD`）并行独立重跑六个 c64 场景。四个非 Excel 场景均 `passed`，每项 `128/128` 操作完成、`errorCount=0`、child exit `0`，隔离 TEMP 清理后为 `0`；两个 Excel 场景在并行窗口超时，旧结果保留为门限观察。随后以同一当前依赖串行、`7200000ms` 窗口重跑 `tempfile-excel-c64-serial-current.json` 和 `hybrid-excel-c64-serial-current.json`，两项均 `passed`，分别完成 `128/128`，耗时约 `1,218,066 ms` 和 `1,497,462 ms`，错误、遗留文件和终止后隔离 TEMP 均为 `0`。旧 `-rerun` 记录的 ResourceProbe 依赖 SHA-256 为 stale `7DC3CFDA...`，仅保留为历史观察；当前候选 100K staging 合并覆盖为 `36/36 PASS`，审批字段仍为空，所以产物仍是结构/清理证据而非已批准容量门禁。

## Entity/Template 大明细资源

`artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entity-template-100k-final.json` 运行当前 Release `EntityProbe`，在 2 CPU 等效/4 GiB Job Object 下完成 100K 行 Entity/Template 单次重复。四种公开模式（Entity sync/async、Template sync/async）均成功导出并重新导入 `100000` 行，结构化错误为 `0`，输出非零；Job 峰值工作集为 `622,891,008` bytes。对应 roundtrip JSON 为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/entity-template-100k-final.json`。

`artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entity-template-500k-final.json` 同样在 2 CPU 等效/4 GiB Job Object 下完成 500K 行单次重复，四种模式均导入 `500000` 行，结构化错误为 `0`，输出非零；Job 峰值工作集为 `2,396,901,376` bytes。对应 roundtrip JSON 为 `artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/entity-template-500k-final.json`，四种模式输出约 `8.0 MB`。

首次 1M Entity/Template 尝试曾运行超过 10 分钟无结构化进展；本轮使用相同 2 CPU 等效/4 GiB Job Object、当前 Release `EntityProbe`、隔离 TEMP 和 `1800000ms` 上限重新执行，形成 `artifacts/resource/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/windows-job-entity-template-1m-current-20260921.json`。子进程在 NPOI ZIP 预检阶段以 `BingOfficesResourceLimitException` 退出，原因是 `xl/worksheets/sheet2.xml` 超过默认 `ExcelResourceLimits.MaxWorksheetBytes=64 MiB`；Job 峰值约 `2,049,687,552` bytes，隔离 TEMP 为 `0` 文件/`0` 字节并成功清理。由于没有完成 Entity/Template roundtrip，`entity-template-1m-current-20260921.json` 未生成，该项继续为 `NOT_VERIFIED`，不能按通过处理；不得为通过 1M 而静默放宽默认资源边界。

上述 100K/500K 证据覆盖当前候选的大明细资源行为，但每个受限规模仅执行一次重复；100K staging 的两个 Excel c64 通过更长串行观察窗口完成，仍需按批准阈值解释该长尾耗时；Entity/Template 1M 被默认 worksheet XML 资源限制拒绝，完整取消/失败提交矩阵、生产/外部环境和批准阈值仍未验证。系统临时目录强制终止已有一次隔离 TEMP 演练，但不等价于完整资源门禁。

## 仍未验证

取消中途提交保护的完整资源矩阵、模板/图片/复杂布局在受限作业下的完整 500K/1M 矩阵、Entity/Template 1M roundtrip、生产机器和外部 CI 未执行或未形成可审计的通过产物，继续标记 `NOT_VERIFIED`。1M 当前已有资源限制拒绝证据，但不能替代成功容量证据；系统临时目录强制终止已有多次隔离 TEMP 证据，也不能替代这些门禁。Entity/Template 500K 和 100K staging `36/36` 仅代表当前本地受限样本，不替代完整矩阵或批准阈值。因此本报告不能单独把发布任务标记为 `COMPLETED`。

## 500K staging 单并发局部矩阵

当前 Release `ResourceProbe` 在 Windows Job Object 2 CPU 等效/4 GiB、`timeoutMs=1800000`、隔离 TEMP、500K 行、并发 `1`、预热 `1` 次/测量 `2` 次条件下，已完成全部 9 个 `strategy × scenario` 组合；每个组合均有独立 runner JSON，runner stdout 内嵌完整 staging 结果。下表各项均为两个 measurement 操作，`temporaryDiskBytesAfter=0`、`leftoverFiles=0`，runner 隔离 TEMP 最终均为 `0` 文件/`0` 字节且 `cleanup=deleted`。

| Strategy / scenario | Runner artifact | 结果 | Job 峰值 | 输出字节 | 临时磁盘峰值 | 错误 |
|---|---|---|---:|---:|---:|---:|
| Memory / excel-100k | `windows-job-staging-500k-memory-excel-100k-c1-final.json` | `PASS` | `1,422,483,456` | `424,133,356` | `0` | `0` |
| TempFile / excel-100k | `windows-job-staging-500k-tempfile-excel-100k-c1-final.json` | `PASS` | `1,480,302,592` | `424,133,356` | `212,066,678` | `0` |
| Hybrid / excel-100k | `windows-job-staging-500k-hybrid-excel-100k-c1-final.json` | `PASS` | `1,499,906,048` | `424,133,354` | `212,066,677` | `0` |
| Memory / failure-double-dom | `windows-job-staging-500k-memory-failure-double-dom-c1-final.json` | `FAILED` | `3,262,435,328` | `0` | `0` | `2` |
| TempFile / failure-double-dom | `windows-job-staging-500k-tempfile-failure-double-dom-c1-final.json` | `FAILED` | `3,254,591,488` | `0` | `0` | `2` |
| Memory / template-image-style | `windows-job-staging-500k-memory-template-image-style-c1-final.json` | `FAILED` | `3,247,964,160` | `0` | `0` | `2` |
| TempFile / template-image-style | `windows-job-staging-500k-tempfile-template-image-style-c1-final.json` | `FAILED` | `3,248,381,952` | `0` | `2,411` | `2` |
| Hybrid / failure-double-dom | `windows-job-staging-500k-hybrid-failure-double-dom-c1-final.json` | `FAILED` | `3,254,464,512` | `0` | `0` | `2` |
| Hybrid / template-image-style | `windows-job-staging-500k-hybrid-template-image-style-c1-final.json` | `FAILED` | `3,247,976,448` | `0` | `0` | `2` |

六个失败组合均为结构化 `OutOfMemoryException`：三个 `failure-double-dom` 策略均在 NPOI 失败工作簿注释/VML 序列化阶段失败；三个 `template-image-style` 策略均在 NPOI SharedStrings/ZIP 写出阶段失败。它们是容量边界失败，不是 runner timeout，也不单独构成产品支持结论。本次完整覆盖 500K c1 的 `3 策略 × 3 场景`，但不覆盖并发 `4/16/64`、取消/失败提交矩阵、1M staging、生产机器、外部 CI 或正式阈值；FIX-003 继续 `OPEN / MUST_FIX`，整体仍 `NOT_VERIFIED`。

## 500K staging 并发 4 Excel 局部补充

在同一当前 Release `ResourceProbe`、Windows Job Object 2 CPU 等效/4 GiB、`timeoutMs=1800000`、隔离 TEMP、500K 行、请求并发 `4`、预热 `1` 次/测量 `2` 次条件下，独立完成三个 `excel-100k` 策略。三个 runner 和 child 均 `passed`、退出码 `0`、`errorCount=0`；请求并发为 4，但 staging 内部 `maxActualParallelism=1`，不能把这次证据解释为四路同时 DOM 操作。

| Strategy | Runner artifact | Job 峰值 | 输出字节 | 临时磁盘峰值 | 实际并行度 | child elapsed ms | 清理 |
|---|---|---:|---:|---:|---:|---:|---|
| Memory | `windows-job-staging-500k-memory-excel-100k-c4-final.json` | `1,428,307,968` | `1,696,533,421` | `0` | `1` | `427,585.9499` | `0/0/deleted` |
| TempFile | `windows-job-staging-500k-tempfile-excel-100k-c4-final.json` | `1,194,053,632` | `1,696,533,420` | `212,066,678` | `1` | `431,069.3949` | `0/0/deleted` |
| Hybrid | `windows-job-staging-500k-hybrid-excel-100k-c4-final.json` | `1,507,565,568` | `1,696,533,421` | `212,066,678` | `1` | `438,236.5350` | `0/0/deleted` |

该补充将当前本地 Excel staging 观察扩展到 500K/c1 与请求并发 4；仍不构成完整 `3 策略 × 3 场景 × 4 并发` 矩阵，也不覆盖失败工作簿、模板图片样式、取消/失败提交、1M、生产机器、外部 CI 或正式阈值。FIX-003 继续 `OPEN / MUST_FIX`，资源门禁仍为 `NOT_VERIFIED`。

## 500K staging 请求并发 4 失败/模板局部补充

在同一当前 Release `ResourceProbe`、Windows Job Object 2 CPU 等效/4 GiB、`timeoutMs=1800000`、隔离 TEMP、500K 行、请求并发 `4`、预热 `1` 次/测量 `2` 次条件下，补齐六个 `failure-double-dom`/`template-image-style` 组合。首个错误的低优先级调用使用了小写策略参数 `memory`，child 以 CLI 参数错误退出 `2`，未计入下表；后续 `final2` 产物使用正确的 `Memory`/`TempFile`/`Hybrid` 枚举值。

| Strategy / scenario | Runner artifact | 结果 | Job 峰值 | 输出字节 | 错误 | 临时磁盘峰值 | 实际并行度 | 清理 |
|---|---|---|---:|---:|---:|---:|---:|---|
| Memory / failure-double-dom | `windows-job-staging-500k-memory-failure-double-dom-c4-final2.json` | `child-failed` / OOM | `3,270,365,184` | `0` | `8` | `0` | `1` | `0/0/deleted` |
| TempFile / failure-double-dom | `windows-job-staging-500k-tempfile-failure-double-dom-c4-final2.json` | `child-failed` / OOM | `3,276,005,376` | `0` | `8` | `0` | `1` | `0/0/deleted` |
| Hybrid / failure-double-dom | `windows-job-staging-500k-hybrid-failure-double-dom-c4-final2.json` | `child-failed` / OOM | `3,277,504,512` | `0` | `8` | `0` | `1` | `0/0/deleted` |
| Memory / template-image-style | `windows-job-staging-500k-memory-template-image-style-c4-final2.json` | `child-failed` / OOM | `3,249,311,744` | `0` | `8` | `0` | `1` | `0/0/deleted` |
| TempFile / template-image-style | `windows-job-staging-500k-tempfile-template-image-style-c4-final2.json` | `child-failed` / OOM | `3,249,360,896` | `0` | `8` | `2,459` | `1` | `0/0/deleted` |
| Hybrid / template-image-style | `windows-job-staging-500k-hybrid-template-image-style-c4-final2.json` | `child-failed` / OOM | `3,249,115,136` | `0` | `8` | `0` | `1` | `0/0/deleted` |

`failure-double-dom` 的错误栈落在 NPOI 注释/VML 失败工作簿序列化，`template-image-style` 落在 NPOI SharedStrings/ZIP 写出；六项均为 child failure 而非 runner timeout。该补充将 c4 的失败/模板局部证据补齐为 `6/6`，但仍不代表完整 500K `3 策略 × 3 场景 × 4 并发` 的发布通过，不覆盖 1M、取消/失败提交、生产机器、外部 CI 或正式阈值。FIX-003 继续 `OPEN / MUST_FIX`，资源门禁仍为 `NOT_VERIFIED`。

500K/c16 failure/template 试跑未形成可审计产物：首个有效组合在 JobRunner/child 生成 runner JSON 前进程消失，未观察到退出码、峰值或清理字段；未手动终止，也未启动其余五项。该尝试不计入矩阵结果，原因未被证据确定，继续标记 `NOT_VERIFIED`，不据此推断 c16 的容量结论。

随后补充一个明确 `60000ms` 上限的 500K/c16 `Memory/failure-double-dom` 诊断：`windows-job-staging-500k-memory-failure-double-dom-c16-timeout60-final.json` 的原始 runner 参数为 **16 GiB / 2 CPU**（不是 4 GiB），结果为 `runner-timeout`、runner/child exit `-1/-1`、`runnerTimedOut=true`，Job 峰值 `5,372,166,144` bytes；child 未生成 JSON/stdout，但隔离 TEMP 终止后为 `0` 文件/`0` 字节且 `cleanup=deleted`。该文件只证明 16 GiB 条件下的有界终止和清理可观测性，不替代 4 GiB c16 容量结果，也不改变完整资源门禁的 `NOT_VERIFIED` 状态。

随后按正确的 WindowsJobRunner 参数顺序（`memoryGiB=4`、`cpuCount=2`）重做 500K/c16 `Memory/failure-double-dom` 的 `60000ms` 有界诊断：`windows-job-staging-500k-memory-failure-double-dom-c16-timeout60-4g-final.json` 为 `runner-timeout`、runner/child exit `-1/-1`、`runnerTimedOut=true`，`requestedMemoryLimitBytes=4,294,967,296`，Job 峰值 `3,235,536,896` bytes；child 未生成 JSON/stdout，隔离 TEMP 终止后为 `0` 文件/`0` 字节且 `cleanup=deleted`。该结果仍只证明 4 GiB 条件下的有界超时和清理链路，不是 c16 容量 PASS/FAIL 或完整资源门禁。

另补充 500K/c16 `Memory/excel-100k` 的同条件（4 GiB、2 CPU、`60000ms`）有界诊断：`windows-job-staging-500k-memory-excel-100k-c16-timeout60-4g-final.json` 为 `runner-timeout`、runner/child exit `-1/-1`、`runnerTimedOut=true`，Job 峰值 `1,399,291,904` bytes，child 无 JSON/stdout；终止后清理前观察到 `1` 个临时文件、`29,843,456` bytes，最终 `cleanup=deleted`。该证据只说明超时期间存在可观测临时文件且清理成功，不是 c16 容量 PASS/FAIL，不能替代完整取消/失败提交矩阵。

继续补充 500K/c16 `TempFile/excel-100k` 与 `Hybrid/excel-100k` 的同条件（4 GiB、2 CPU、`60000ms`）有界诊断：两项均为 `runner-timeout`、runner/child exit `-1/-1`、`runnerTimedOut=true`，child 均无 JSON/stdout，最终隔离 TEMP 均 `cleanup=deleted`。TempFile Job 峰值 `1,191,579,648` bytes，终止后清理前为 `2` 个临时文件/`75,202,560` bytes；Hybrid Job 峰值 `1,208,631,296` bytes，终止后清理前为 `1` 个临时文件/`32,239,616` bytes。该证据只补充不同 staging 策略的超时残留与清理可观测性，不是 c16 容量 PASS/FAIL，也不替代完整取消/失败提交矩阵。

## 500K staging 请求并发 16 局部边界补充

当前 Release `ResourceProbe` 使用正确的 Windows JobRunner 入口（`childPath=dotnet`，首个 child 参数为 `ResourceProbe.dll`），在 `requestedCpuEquivalent=2`、`requestedMemoryLimitBytes=4,294,967,296`、隔离 TEMP、`timeoutMs=1800000`、500K 行、请求并发 `16`、预热 `1` 次/测量 `2` 次下完成六项 failure/template 运行。早期把 DLL 直接作为 child executable 的三份调用已隔离为 `*.invalid-launch.json`，stderr 为 `System.Runtime 8.0.0.0` 加载失败，不纳入正式统计。

| Strategy / scenario | Runner artifact | 结果 | Job 峰值 | operationBytes | errors | 实际并行度 | 临时磁盘峰值 | 清理 |
|---|---|---|---:|---:|---:|---:|---:|---|
| Memory / failure-double-dom | `windows-job-staging-500k-memory-failure-double-dom-c16-long-final2.json` | `child-failed` / OOM | `3,271,938,048` | `0` | `32` | `1` | `0` | `0/0/deleted` |
| TempFile / failure-double-dom | `windows-job-staging-500k-tempfile-failure-double-dom-c16-long-final2.json` | `child-failed` / OOM | `3,277,131,776` | `0` | `32` | `1` | `0` | `0/0/deleted` |
| Hybrid / failure-double-dom | `windows-job-staging-500k-hybrid-failure-double-dom-c16-long-final2.json` | `child-failed` / OOM | `3,277,570,048` | `0` | `32` | `1` | `0` | `0/0/deleted` |
| Memory / template-image-style | `windows-job-staging-500k-memory-template-image-style-c16-long-final2.json` | `child-failed` / OOM | `3,252,965,376` | `0` | `32` | `1` | `0` | `0/0/deleted` |
| TempFile / template-image-style | `windows-job-staging-500k-tempfile-template-image-style-c16-long-final2.json` | `child-failed` / OOM | `3,252,240,384` | `0` | `32` | `1` | `2,461` | `0/0/deleted` |
| Hybrid / template-image-style | `windows-job-staging-500k-hybrid-template-image-style-c16-long-final2.json` | `child-failed` / OOM | `3,252,502,528` | `0` | `32` | `1` | `0` | `0/0/deleted` |

六项均为 `runnerTimedOut=false`、runner/child exit `1/1`，每项 `submitted/completed/operation/errorCount=32/32/32/32`，最大排队深度为 `15`。failure 场景错误落在 NPOI Comments/VML 失败工作簿序列化，template 场景记录 `BingOfficesExportException` 包裹的 NPOI `System.OutOfMemoryException`；这些是 4 GiB 预算下的结构化容量失败边界，不是发布通过。三项 c16 Excel (`Memory/TempFile/Hybrid × excel-100k`) 的同一预算运行仍为 `runner-timeout`，分别在终止前观察到 `2/212,066,725`、`2/71,442,432`、`1/84,369,408` 个临时文件/字节，最终均 `cleanup=deleted` 且无 child summary。

本节只补充 500K/c16 局部边界和清理可观测性，不构成完整 `3 策略 × 3 场景 × 1/4/16/64` 资源矩阵，不替代取消中途提交保护、1M、生产机器、外部 CI 或批准阈值；FIX-003 仍为 `OPEN / MUST_FIX`，资源门禁继续 `NOT_VERIFIED`。

## 1M staging Excel-only c1 局部补充

为补充 1M staging 的可审计边界，使用当前工作树 Release `WindowsJobRunner` 和 `ResourceProbe`，在 Windows Job Object `2 CPU` 等效、`4 GiB`、隔离 TEMP、`timeoutMs=1800000`、请求并发 `1`、预热 `1` 次/测量 `2` 次条件下，分别执行 `Memory`、`TempFile`、`Hybrid` 的 `excel-100k` 场景，`rowCount=1000000`。三份 runner 产物均由正确的 `childPath=dotnet` 启动，首个 child 参数为 `tests/Bing.Offices.ResourceProbe/bin/Release/net8.0/Bing.Offices.ResourceProbe.dll`。

| Strategy | Runner artifact | Runner/child | submitted/completed/operation/errors | operationBytes | Job 峰值 | 临时峰值 | 终止后清理 |
|---|---|---|---:|---:|---:|---:|---|
| Memory | `windows-job-staging-1m-memory-excel-100k-c1-final.json` | `passed / 0 / 0` | `2 / 2 / 2 / 0` | `848,270,818` | `2,771,984,384` | `0` | `0/0/deleted` |
| TempFile | `windows-job-staging-1m-tempfile-excel-100k-c1-final.json` | `passed / 0 / 0` | `2 / 2 / 2 / 0` | `848,270,818` | `2,886,086,656` | `424,135,409` | `0/0/deleted` |
| Hybrid | `windows-job-staging-1m-hybrid-excel-100k-c1-final.json` | `passed / 0 / 0` | `2 / 2 / 2 / 0` | `848,270,819` | `2,334,699,520` | `424,135,410` | `0/0/deleted` |

三项 child 结果均以 `staging-scenario` JSON 嵌入 runner 的 `stdout` 字段；`RunScenario` 当前只输出结构化 child 结果，不单独写入命令参数中的 child 路径。三项共同字段为 `rowCount=1000000`、`scenario=excel-100k`、`concurrency=requestedConcurrency=1`、`maxActualParallelism=1`、`maxQueuedRequests=0`、`runnerTimedOut=false`、`assignedToJob=true`、`temporaryDiskBytesAfter=0`、`leftoverFiles=0`、`stderr=""`、child `status=passed`/exit `0`。

本轮运行时物理 identity（不回写或替代 `manifest-current-20260921.md`）为：`WindowsJobRunner.dll` SHA-256 `1FA17599E7F260191AEB56FED0F2496E39D6F53ECFFE1D75F5BFD2133D2473C8`，`ResourceProbe.dll` SHA-256 `28B34797B0CF47FCDA20615C9DA582B02457B59AE9520B72892B449A3F9CCD71`，依赖的 `Bing.Offices.Npoi.dll` SHA-256 `E87FA5777867FB1D8B22B824EB580373306C75B95F311447F554BE70390F010E`。这些 hash 只绑定本轮局部产物，不能把它们重新归属到其他冻结候选。

Sol Medium 只读复核确认三份 runner 字段、嵌入 child 结果、资源约束、退出码、输出、清理状态和场景参数一致，无新增 `MUST_FIX`/`SHOULD_FIX`。该结果仅是 `1M × Excel-only × c1 × 3 strategies` 的本地 PASS/清理证据；不覆盖 failure/template、并发 `4/16/64`、取消中途提交保护、Entity/Template 1M、生产机器、外部 CI 或正式阈值，FIX-003 继续 `OPEN / MUST_FIX`，资源门禁仍为 `NOT_VERIFIED`。

## 1M staging failure/template c1 边界补充

在同一当前工作树 Release `WindowsJobRunner`/`ResourceProbe`、Windows Job Object `2 CPU` 等效、`4 GiB`、隔离 TEMP、`timeoutMs=1800000`、请求并发 `1`、预热 `1` 次/测量 `2` 次条件下，补充执行 `rowCount=1000000` 的 `failure-double-dom` 与 `template-image-style` 六项组合。所有 runner 均由 `childPath=dotnet` 启动正确的 `ResourceProbe.dll`，`runnerTimedOut=false`、`assignedToJob=true`、实际并发为 `1`、排队为 `0`，child 结果嵌入 runner `stdout`。

| Strategy / Scenario | Runner artifact | Runner/child | submitted/completed/operation/errors | operationBytes | Job 峰值 | child 峰值 | child 临时峰值 | 终止后清理 |
|---|---|---|---:|---:|---:|---:|---:|---|
| Memory / failure-double-dom | `windows-job-staging-1m-memory-failure-double-dom-c1-final.json` | `child-failed / 1 / 1` | `2 / 2 / 2 / 2` | `0` | `303,833,088` | `345,890,816` | `0` | `0/0/deleted` |
| TempFile / failure-double-dom | `windows-job-staging-1m-tempfile-failure-double-dom-c1-final.json` | `child-failed / 1 / 1` | `2 / 2 / 2 / 2` | `0` | `309,362,688` | `346,001,408` | `0` | `0/0/deleted` |
| Hybrid / failure-double-dom | `windows-job-staging-1m-hybrid-failure-double-dom-c1-final.json` | `child-failed / 1 / 1` | `2 / 2 / 2 / 2` | `0` | `301,772,800` | `338,452,480` | `0` | `0/0/deleted` |
| Memory / template-image-style | `windows-job-staging-1m-memory-template-image-style-c1-final.json` | `child-failed / 1 / 1` | `2 / 2 / 2 / 2` | `0` | `3,247,915,008` | `3,277,959,168` | `0` | `0/0/deleted` |
| TempFile / template-image-style | `windows-job-staging-1m-tempfile-template-image-style-c1-final.json` | `child-failed / 1 / 1` | `2 / 2 / 2 / 2` | `0` | `3,248,037,888` | `3,278,245,888` | `0` | `0/0/deleted` |
| Hybrid / template-image-style | `windows-job-staging-1m-hybrid-template-image-style-c1-final.json` | `child-failed / 1 / 1` | `2 / 2 / 2 / 2` | `0` | `3,247,992,832` | `3,278,929,920` | `0` | `0/0/deleted` |

三项 `failure-double-dom` 均在 `ExcelXlsxZipPreflight` 对 `xl/worksheets/sheet1.xml` 解压大小进行默认资源限制拒绝，异常为 `Bing.Offices.Exceptions.BingOfficesResourceLimitException`，不是 OOM；三项 `template-image-style` 均为 NPOI 写出路径的 `System.OutOfMemoryException`，其中 Memory 产物还包含 `XSSFRichTextString` 相关栈。六项均为 `child-failed`、runner/child exit `1/1`、`submitted/completed/operation/errorCount=2/2/2/2`、`operationBytes=0`，child 的临时磁盘峰值、结束后文件和字节均为 `0`，runner 隔离 TEMP 均为 `0/0/deleted`。

上述结果只能作为 `1M × c1 × failure/template` 的边界失败与清理证据，不能表述为容量 PASS、产品支持或完整资源矩阵通过；`500K/1M` 其他并发、取消中途提交保护、Entity/Template 1M roundtrip、生产机器、外部 CI 和正式阈值仍为 `NOT_VERIFIED`。Sol Medium 对六份原始 JSON 只读复核通过，未新增 `MUST_FIX`/`SHOULD_FIX`；FIX-003 继续 `OPEN / MUST_FIX`。

## 1M staging Excel-only c4 局部补充

在同一当前 Release、Windows Job Object `2 CPU` 等效/`4 GiB`、隔离 TEMP、`timeoutMs=1800000`、`rowCount=1000000`、请求并发 `4`、预热 `1` 次/测量 `2` 次条件下，补充执行 `Memory`、`TempFile`、`Hybrid` 三项 `excel-100k` 场景。所有 runner 均使用 `childPath=dotnet` 启动正确的 `ResourceProbe.dll`，并将 child `staging-scenario` JSON 嵌入 runner `stdout`。

| Strategy | Runner artifact | Runner/child | submitted/completed/operation/errors | operationBytes | 实际并发/排队 | Job 峰值 | child 峰值 | child 临时峰值 | 终止后清理 |
|---|---|---|---:|---:|---:|---:|---:|---:|---|
| Memory | `windows-job-staging-1m-memory-excel-100k-c4-final.json` | `passed / 0 / 0` | `8 / 8 / 8 / 0` | `3,393,083,275` | `1 / 3` | `2,776,682,496` | `2,802,229,248` | `0` | `0/0/deleted` |
| TempFile | `windows-job-staging-1m-tempfile-excel-100k-c4-final.json` | `passed / 0 / 0` | `8 / 8 / 8 / 0` | `3,393,083,276` | `1 / 3` | `2,879,967,232` | `2,795,544,576` | `424,135,410` | `0/0/deleted` |
| Hybrid | `windows-job-staging-1m-hybrid-excel-100k-c4-final.json` | `passed / 0 / 0` | `8 / 8 / 8 / 0` | `3,393,083,277` | `1 / 3` | `2,900,910,080` | `2,808,381,440` | `424,135,410` | `0/0/deleted` |

三项均为 `runnerTimedOut=false`、`assignedToJob=true`、`operationBytes>0`，child `temporaryDiskBytesAfter=0`、`leftoverFiles=0`，runner 隔离 TEMP 结束后均为 `0` 文件/`0` 字节且 `cleanup=deleted`。`c4` 仅表示请求并发，资源门控下三项 `maxActualParallelism=1`、`maxQueuedRequests=3`，不能解释为四路 DOM 同时执行。

该结果仅是 `1M × Excel-only × c4 × 3 strategies` 的本地 PASS/清理证据；不覆盖 failure/template、其他并发、取消中途提交保护、Entity/Template 1M、生产机器、外部 CI 或正式阈值，不能单独改变 FIX-003 的 `OPEN / MUST_FIX` 和资源门禁 `NOT_VERIFIED`。Sol Medium 对三份原始 runner JSON 只读复核通过，未发现数据或执行链路问题。

## Round 43：最终候选生产等价 profile 代表性 smoke

针对用户批准的 FIX-003D，本轮使用当前重新构建的 Release `WindowsJobRunner`/`ResourceProbe`，在本机 Windows Job Object 生产等价资源 profile 下依次执行单一代表场景 `TempFile / excel-100k / concurrency=1`，每项预热 `1` 次、测量 `2` 次；约束固定为 `2 CPU` 等效、`4 GiB`、隔离 TEMP、100K/500K 使用 `900000/1800000 ms` 上限、1M 使用 `1800000 ms` 上限。runner 与 child 产物分别为：

- 100K：`windows-job-production-equivalent-smoke-100k.json`、`production-equivalent-smoke-100k.json`；runner/child `passed / 0`，`submitted/completed/operation/errorCount=2/2/2/0`，Job 峰值 `399,478,784` bytes，清理 `0/0/deleted`。
- 500K：`windows-job-production-equivalent-smoke-500k.json`、`production-equivalent-smoke-500k.json`；runner/child `passed / 0`，`submitted/completed/operation/errorCount=2/2/2/0`，Job 峰值 `1,224,552,448` bytes，清理 `0/0/deleted`。
- 1M：`windows-job-production-equivalent-smoke-1m.json`、`production-equivalent-smoke-1m.json`；runner/child `passed / 0`，`submitted/completed/operation/errorCount=2/2/2/0`，Job 峰值 `2,357,735,424` bytes，清理 `0/0/deleted`。

三项均绑定本轮当前工作树 Release 运行时，实际 DOM 并发为 `1`，排队为 `0`，输出字节非零，child 与 runner 均未超时。该证据证明本地可复核的生产等价资源 profile 在代表性 100K/500K/1M 单一路径上可完成；它不是真实生产机器或外部 CI 结果，不能关闭 FIX-003D/E。Entity/Template 1M 仍按用户批准的默认 `MaxWorksheetBytes=64 MiB` 记录为 `NOT_SUPPORTED_BY_DEFAULT / EXPECTED_RESOURCE_REJECTION`，本轮未重跑该 fixture。
