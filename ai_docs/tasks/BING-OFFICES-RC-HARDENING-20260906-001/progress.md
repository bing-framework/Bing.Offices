# 执行进度

## 当前状态

- 状态：`IN_PROGRESS`
- Task ID：`BING-OFFICES-RC-HARDENING-20260906-001`
- 已完成：任务状态登记、计划/规则复核、基线环境与 restore 阻断采集、`00-baseline.md`。
- 当前阶段：Phase 6 最终门禁复核；本地实现、测试、打包和文档验证已完成，外部门禁保持阻塞。

## 变更记录

| 时间 | 阶段 | 变更/证据 | 结果 | 下一步 |
| --- | --- | --- | --- | --- |
| 2026-09-06 | Phase 0 | `dotnet --info`、Git 状态、项目 TFM 与依赖扫描 | 环境可复现；locked restore 受 NU1004/NU1301 阻断 | 修改唯一文件导出路径、删除空 diagnostics API |
| 2026-09-06 | Phase 1 | exporter 正式 `ExportToFile`、`IFileExportCommitter`、Observer 同实例/单次测试、TryAddPicture 失败适配器 | 定向测试通过 | 继续完成 TFM/API/报告收口 |
| 2026-09-06 | Phase 1 | Failure Workbook 增加 `MaxCandidateErrorRows`、`MaxCopiedCells`、`MaxCopiedPictures` preflight budget | `ExcelP0RegressionTest.Import_FailureWorkbook_CandidateRowBudget_ShouldRejectBeforeMutation` 通过 | 完整资源矩阵仍待外部预算 |
| 2026-09-06 | Phase 2 | 删除 `ExcelMappingDiagnostic`、NPOI/Unit/Integration/CI 收敛 net8、成员级 API canonicalizer、基线 commit before snapshot | before/candidate diff 可审计；approvedBy/approvedAt 未填 | 保持 API 门禁 BLOCKED |
| 2026-09-06 | Phase 4 | Unit 477/478（历史快照，API baseline 明确 BLOCKED）、Integration 15/15、Docs 10/10、三包 pack、隔离 PackageReference consumer | package-only 证据通过；API 门禁真实失败而非 Skip | 保留 API/性能/独立 Review 外部门禁 |
| 2026-09-06 | Phase 5/6 | 基线/候选同参数资源场景、尾延迟 artifact、warning surface、Review/Readiness 报告 | 预算未批准；生产 CS1591 和独立 Review 未完成 | 保持 PARTIAL/No-Go，等待外部门禁 |

| 2026-09-07 | Phase 4/6 | 移除 TryAddPicture 可变静态测试注入；补齐 CreatePicture/Resize 后置失败测试；清理 Benchmark/测试 warning；重新 pack 并用独立 offline PackageReference consumer 验证（历史快照） | Unit `487/488`；Integration `15/15`；Docs `10/10`；Build `0/0`；consumer `excelBytes=4250` | 保留 API 审批、性能预算、跨平台和独立复审门禁 |

## 外部阻断

- NuGet 源 `https://api.nuget.org/v3/index.json` 当前 TLS/凭证失败，影响 restore、ApiSnapshot、Benchmark 和部分完整验证。
- 维护者 API diff 审批、性能预算和跨平台 runner 尚未提供；在最终 readiness 中保持 No-Go/Conditional。

## 2026-09-07 测试命名清理

- 删除已完成兼容层语义的测试名 `LegacyExcelAsyncBytes`、`ProviderSpiAndCompatibilityOverloads`；字节往返测试改为同步真实调用，不再用 `Task.Run` 伪造异步路径。
- 定向验证：Workbook 字节往返与 Provider SPI 边界测试 `2 passed / 0 skipped / 0 failed`。

## 2026-09-07 PackageConsumer contract expansion

- Consumer 新增 JSON/XML v2 loader、ExportMappingBuilder、`ExportToFile` 文件提交失败、Observer 同实例/单次及 `FileCommit/Commit` 分类断言；保持无 `ProjectReference`。
- 验证：独立 offline source/cache restore/build/run 通过（历史快照），首次断言错误已根据实际公开合同修正为 `BingOfficesOperation.FileCommit`。

## 2026-09-07 Final rebuild/package rerun

- 补齐 Abstractions public execution-detail XML 注释、ProfileFixtures 注释，并用 `#nullable enable annotations` 保留 Core nullable 注释而不改变 flow 语义；solution `Rebuild` 结果 `0 warning / 0 error`。
- 重新 pack 三个生产包并用 fresh local-source/cache consumer 验证（历史快照）；包哈希已同步到 `05-package-consumer-report.md`。
- API snapshot final-rerun-2 仍只有既定 Abstractions `7 added/5 removed`、Core `3 added/2 removed`；唯一失败是空审批字段。

## 2026-09-07 DefaultFileExportCommitter direct coverage

- 新增 `DefaultFileExportCommitterTest`，直接覆盖新目标移动、既有目标替换、内容异常原样传播、预取消和 staging 清理。
- 定向测试 `4/4`；随后完整 Unit 为 `487 passed / 0 skipped / 1 blocked failure`（历史快照），阻塞原因仍只有 API baseline 审批字段为空。

## 2026-09-07 职责文件拆分续跑

- CSV exporter/importer 从 `CsvEntityPipeline.cs` 拆为 `CsvEntityExporter.cs`、`CsvEntityImporter.cs`；CSV slice `39/39`。
- Failure Workbook 复制、序列化/受限流/临时文件系统适配器拆为 `NpoiFailureWorkbookCopier.cs`、`NpoiFailureWorkbookSerialization.cs`、`NpoiFailureWorkbookFileSystem.cs`；`ExcelP0RegressionTest` `138/138`。
- Mapping runtime model 拆为独立 internal 文件；Abstractions 策略、样式、异常派生类型拆为职责文件；solution Rebuild `0 warning / 0 error`，Mapping/Review slice `83/83`。
- NPOI importer runtime/resource context、exporter non-disposing stream 和 export column planner 拆为独立 internal 文件；Excel/stream slice `193/193`。
- 结构拆分后 Unit `490 passed / 0 skipped / 1 blocked failure`（预算 API 新增前历史快照）、Integration `15/15`、Docs `10/10`；API actual hash 与既有 candidate 一致，仍只受空审批字段阻塞。
- 结构拆分后重新 pack，独立 offline PackageReference consumer 已通过；包内 DLL 与 Release DLL SHA-256 已同步到 `05-package-consumer-report.md`。

## 2026-09-07 Benchmark 1M

- CSV 基准加入 `1_000_000` 行参数；隔离 BDN restore 因 NuGet TLS 阻断，使用 InProcessEmitToolchain 完成 8 场景 ShortRun。
- 结果：1M Import `945.480 ms / 1,984.1 MB / Gen2 1,000`；1M Export `1.216 s / 14,455.7 MB / Gen2 0`。不宣称低 GC，预算仍未批准；原始结果写入 `artifacts/benchmark/csv-pipeline-1m-inprocess-final`。

## 规则适用性

本任务不修改 `Bing.Data.Sql`，SQL 元数据测试规则不适用；不运行生产数据库测试。

## 2026-09-07 Final local gate check

- `git diff --check` 无空白错误；仅保留 `Bing.Offices.ProfileFixtures.xml` 的既有 CRLF/LF 转换提示。
- 未发现工作区内的 BDN 自动生成临时目录、consumer `bin/obj` 或测试缓存残留；`packages-rerun` 是为离线 PackageReference 验证保留的审计 artifact。
- 本地没有可继续安全解除的 P1 门禁；API 审批、性能预算、跨平台 runner 和独立 Review 仍需外部输入，故状态继续保持 `IN_PROGRESS` / `NO-GO / PARTIAL`。

## 2026-09-07 候选 Benchmark 矩阵

- 当前 Release 候选补跑 StreamPipeline 9 场景、Failure Workbook 3 场景、HeaderStyle 3 场景、ValidationRange 3 场景；原始结果保存在 `artifacts/benchmark/current-*`。
- 使用 InProcess toolchain 仅作为候选容量证据，不与隔离进程/历史结果计算回归比率；性能预算仍为 `UNAPPROVED`。

## 2026-09-07 Failure Workbook budget completion

- 新增 `MaxCopiedPictureBytes` 与 `MaxEstimatedTargetObjects`，预检分别按图片数据字节数、`Sheet + rows + cells + pictures` 保守估算对象数拒绝超限；直接测试覆盖 hit/miss 边界。
- 新增非正值配置校验、图片字节超限、目标对象超限直接测试；API candidate 更新为 Abstractions `9 added / 5 removed`，审批仍保持 BLOCKED。
- 预算 API 新增后重新 pack 和 fresh offline consumer 通过：`package-consumer-ok excelBytes=4251 csvBytes=20 npoiExtensions=ok`；最新包/DLL hash 已同步到 `05-package-consumer-report.md`。
