<!-- AI_EXECUTION_STATUS: IN_PROGRESS -->
AI_TASK_ID: BING-OFFICES-RC-HARDENING-20260906-001
AI_EXECUTION_STARTED_AT: 2026-09-06T13:16:05.515Z

# 实施执行报告

## 执行结论

本轮已完成大部分可在本地独立处理的 RC 加固实现，当前仍在执行中。核心代码、直接测试、net8 发布矩阵、包产物、隔离 package-only consumer、候选 API 快照和资源探针已落地；API 审批、性能预算、跨平台 runner 和本轮变更后的独立 Review 仍未满足，因此发布结论暂为 `NO-GO`。

## 任务信息

- Task ID：`BING-OFFICES-RC-HARDENING-20260906-001`
- 基线 commit：`958e5b4886c2fe8df80ece3218d9fab1c57a0ec6`
- 执行器：Codex
- 分支：`master`
- SQL 元数据规则：不适用，本轮未修改 `Bing.Data.Sql`。

## 计划执行情况

| 阶段 | 结果 |
| --- | --- |
| Phase 0 基线 | COMPLETED；restore 初始受 lock/TLS 阻断，详见 `00-baseline.md` |
| Phase 1 异常/NPOI/资源 | PARTIAL；文件提交和 TryAddPicture 合同完成，完整 Failure Workbook budget 矩阵未完成 |
| Phase 2 API/TFM | COMPLETED with approval block；诊断 API 删除、TFM 收敛、canonicalizer 完成，正式 baseline 待审批 |
| Phase 3 大类职责拆分 | PARTIAL；新增 `NpoiFailureWorkbookPreflight` 职责类和独立测试，四个大 orchestration 文件仍偏大 |
| Phase 4 测试/包消费 | COMPLETED locally；Unit/Integration/Docs/pack 和隔离 fresh-cache PackageReference consumer 通过 |
| Phase 5 Benchmark | PARTIAL；CSV 1M InProcess ShortRun 已补齐并保留原始 artifact，但 before/candidate 完整矩阵和预算审批仍缺失 |
| Phase 6 文档/Review/门禁 | PARTIAL；报告已生成，XML warning gate 已通过，独立 Review 和外部审批缺失 |

## 已完成事项

1. `IExcelExporter`、`ICsvExporter` 增加正式文件导出方法；扩展方法改为薄委托。
2. 新增 `IFileExportCommitter` SPI 和 Core 默认实现；Excel/CSV exporter 在自身异常观察边界内提交文件并观察文件提交异常。
3. 删除 public `ExcelMappingDiagnostic`、JSON/XML `out diagnostics` 重载和无作用 diagnostics 参数。
4. `TryAddPicture` 把 AddPicture 后失败转换为明确 `BingOfficesExportException`；内部适配器支持确定性 post-mutation 测试。
5. Failure Workbook 增加候选错误行、复制单元格估算和图片数量的 preflight budget，抽出 `NpoiFailureWorkbookPreflight` 与 `NpoiFailureWorkbookAnnotationWriter`，超限不污染 destination。
6. NPOI、Unit、Integration、CI 发布/测试目标收敛到 net8.0；Abstractions/Core 保持 netstandard2.0。
7. API snapshot canonicalizer 增加类型/成员修饰符、约束、参数名称/默认值、accessor、event 和 attributes；生成候选成员快照和 machine-readable diff。
8. 移除生产项目全局 `CS1591`/`NETSDK1138` 抑制，真实 warning surface 已记录。

## 修改文件

生产代码：

- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/IExcelExporter.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Csv/ICsvExporter.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/IO/IFileExportCommitter.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelMappingDiagnostic.cs`（删除）
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityExporter.cs`
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityImporter.cs`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
- `src/Bing.Offices.Core/Bing/Offices/Extensions/CsvStreamExtensions.cs`
- `src/Bing.Offices.Core/Bing/Offices/Extensions/ExcelStreamExtensions.cs`
- `src/Bing.Offices.Core/Bing/Offices/IO/DefaultFileExportCommitter.cs`
- `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactoryProvider.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookWriter.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookCopier.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookSerialization.cs`
- `src/Bing.Offices.Npoi/Imports/NpoiFailureWorkbookFileSystem.cs`
- `src/Bing.Offices.Npoi/Imports/ExcelImportRuntime.cs`
- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
- `src/Bing.Offices.Npoi/Exports/NpoiNonDisposingStream.cs`
- `src/Bing.Offices.Npoi/Exports/NpoiExportColumnPlanner.cs`
- `src/Bing.Offices.Npoi/Extensions/Extensions.Service.cs`
- `src/Bing.Offices.Npoi/Extensions/SheetExtensions.Picture.cs`
- `framework.props`、`common.props`、NPOI/测试 csproj、`.github/workflows/ci.yml`

测试、工具和文档：对应 Unit/Integration/Docs 测试文件、`build/ApiSnapshot/*`、日期/NPOI 文档和本任务全部报告。

## API/数据/配置变化

- 新增两个 exporter 文件导出成员和隐藏文件提交 SPI，属于 additive public API 变化。
- 删除空 diagnostics public surface，属于 breaking change，迁移方式为直接使用无 out 的 v2 API。
- 删除 netcoreapp3.1/net6.0 NPOI/测试目标；包只发布 net8.0 NPOI asset。
- 未修改数据库 schema、生产数据或 `Bing.Data.Sql`。

## 测试结果

- Unit net8.0：`492 passed / 0 skipped / 1 failed`；唯一失败是正式 API baseline 的 `approvedBy`/`approvedAt` 为空，明确标记 BLOCKED。
- Integration net8.0：`15/15`。
- Docs net8.0：`10/10`。
- 新增直接测试：文件提交 Observer 同实例/单次、TryAddPicture post-mutation throw、Failure Workbook preflight 的错误行/单元格/图片数量/图片字节/目标对象预算、diagnostics 删除、API canonicalizer golden lines。

## Build/Typecheck/Lint/Format

- `dotnet build Bing.Offices.sln -c Release --no-restore`：`0 errors / 0 warnings`。
- `dotnet pack` 三个生产包和 symbol 包：成功。
- `git diff --check`：最终收口时执行。
- 无单独 lint/formatter 工具配置；未伪造 PASS。

## 计划偏差

1. 计划建议的最小 SPI 已落地为 Abstractions `IFileExportCommitter` + Core `DefaultFileExportCommitter`，NPOI 通过构造器/DI 使用，避免生产 IVT。
2. 四个大类未完成完整职责拆分；本轮至少将 Failure Workbook 资源预检抽为独立职责类并补直接测试，剩余拆分风险已写入 `07-final-code-review.md`。
3. 跨平台测试没有用历史产物替代，保持 BLOCKED；fresh-cache package consumer 已通过。

## 基线问题

- 初始 locked restore 报 NU1004；外部 NuGet TLS/凭证失败报 NU1301。获得受控网络权限后 `--force-evaluate` restore 成功。
- 正式 `build/api-snapshot-baseline.json` 已从基线 commit 临时副本生成；审批字段保持为空，不自动批准。

## 已知问题

- API baseline 审批、性能预算、跨平台 runner 和本轮变更后的独立 Review 未完成。
- CSV 1M benchmark 已完成：Import `945.480 ms / 1,984.1 MB`，Export `1.216 s / 14,455.7 MB`；使用 InProcess toolchain，不能替代隔离进程/可比 baseline。
- `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` 不再 Skip；审批字段缺失时明确失败为 BLOCKED，而非更新旧 hash 常量。

## 风险与回归关注点

- 新增 exporter 文件方法是 public additive contract，第三方自定义 exporter 需要实现新成员。
- `TryAddPicture` 在 NPOI 已写入图片数据后失败会抛异常，调用方不能再依赖 post-mutation `false`。
- net6/netcoreapp3.1 包消费行为不再属于发布承诺。
- XML warning gate 已通过 solution Release build；跨平台和外部审批仍未解除。

## Reviewer 注意事项

先审 `02-api-breaking-changes.md` 的完整 candidate lines，再决定 baseline；再审 `05-package-consumer-report.md` 和 `06-benchmark-report.md` 的外部解除条件。独立 Review 不能由本执行器自审替代。

## Git 状态

- 未自动 `git add`、`git commit`、`git push`、创建 PR、tag 或 publish。
- `task-finish.mjs` 未执行：其实现会条件性向 Feishu 发送工作区元数据，当前未获得外部通知授权；runtime 状态仍由协议保持 `implementing`，需维护者在允许通知后手动收口。
- 当前工作区保留实现、测试、报告和 artifacts，供维护者审阅和继续执行。

## 2026-09-07 续跑补充

- `TryAddPicture` 测试注入改为内部重载，移除可变静态全局状态；后置 `CreatePicture` 与 `Resize` 失败测试均通过。
- warning 清理后 Release solution build 为 `0 warning / 0 error`；完整 net8 Unit 为 `492 passed / 0 skipped / 1 blocked failure`，Integration `15/15`，Docs `10/10`。
- 重新 pack 不使用 `--no-build`，确认 Abstractions/Core 仅含 `netstandard2.0`、NPOI 仅含 `net8.0`；预算 API 新增后 PackageReference-only consumer 使用相对本地源、离线独立缓存运行输出 `package-consumer-ok excelBytes=4251 csvBytes=20 npoiExtensions=ok`。
- 同步清理旧兼容语义测试名；API snapshot 复验无新增意外 public diff，仍仅因 baseline 审批字段为空退出 1。

## 2026-09-07 结构拆分与包复验

- CSV exporter/importer、Failure Workbook 复制/序列化/文件系统、Mapping runtime model、NPOI importer runtime/exporter column planner/stream、Abstractions 策略/样式/异常派生类型完成职责文件拆分；保持 namespace、签名和 public API canonical lines 不变。
- 拆分后重新 `dotnet pack` 三个生产包，fresh offline PackageReference consumer 运行输出 `package-consumer-ok excelBytes=4250 csvBytes=20 npoiExtensions=ok`（历史快照）；包内 DLL 与 Release 输出 SHA-256 一致。
- API snapshot `final-rerun-4` 为预算 API 新增前历史快照。
- 当前候选 Benchmark 已补齐 StreamPipeline 9 场景、Failure Workbook 3 场景、HeaderStyle 3 场景、ValidationRange 3 场景，原始结果位于 `artifacts/benchmark/current-*`；仍不替代同机 before/candidate 比较或预算审批。

## 2026-09-07 资源预算 API 最终复验

- 新增 `MaxCopiedPictureBytes`、`MaxEstimatedTargetObjects` 后 Unit 为 `492 passed / 0 skipped / 1 blocked`；API snapshot `final-rerun-5` 记录 Abstractions `9 added / 5 removed`，审批字段仍为空。
- 重新 pack 并运行 fresh offline PackageReference consumer：`package-consumer-ok excelBytes=4251 csvBytes=20 npoiExtensions=ok`；包内 DLL 与 Release DLL SHA-256 一致。

## 2026-09-07 格式化后最终包复验

- `NpoiFailureWorkbookCopier.cs` 仅做 whitespace formatting；solution Rebuild 与 Failure Workbook `138/138` 通过。
- 重新 pack 后 consumer 输出 `package-consumer-ok excelBytes=4250 csvBytes=20 npoiExtensions=ok`；最终包内 DLL 与 Release DLL hash 已同步到 `05-package-consumer-report.md`。

## 2026-09-07 最终本地门禁复核

- `git diff --check` 无空白错误；仅有生成 XML 文件的 CRLF/LF 转换提示。
- 已确认没有 BDN 自动生成临时目录、PackageConsumer `bin/obj` 或测试缓存残留；`packages-rerun` 为离线包消费验证的保留证据目录。
- 本地可处理项已收口；API 审批、性能/资源预算、跨平台 runner、独立 Reviewer 和通知授权仍未提供，因此不执行 `task-finish.mjs`，不把任务标记为完成或 Go。
