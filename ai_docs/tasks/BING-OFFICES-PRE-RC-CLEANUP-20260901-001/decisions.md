# 决策记录

## D-001：缺失评审输入不作为事实

用户指定的 `ai_docs/codebase-analysis/` 文件在当前工作区不存在。决定以当前源码、Git 状态、真实命令和测试结果为主证据；用户提供的完成度和 No-Go 只保留为待验证风险。

## D-002：保留工作区安全边界

初始工作树除本任务目录外无业务代码改动。后续只编辑计划要求范围，不执行 reset/clean/checkout，不覆盖用户未知修改，不自动提交、推送、Tag 或发布。

## D-003：先修 P0 正确性再做 API 删除和结构重构

Document Factory、方向迁移、关系异常、CSV 安全和 Probe 先于 API 收敛；避免在行为未稳定前扩大编译错误面。

## D-004：NPOI 资源限制表述不升级为 DOM 前硬上限

NPOI 主路径是 DOM 解析。可以增加输入/ZIP/OLE 预检和子进程隔离，但无法证明的解压、DOM、业务实体和失败工作簿峰值必须分别命名，不以 `MaxInputBytes` 宣称完整内存安全上限。

## D-005：Breaking Change 需以全仓引用和包消费证据确认

用户已授权未正式发布前合理 Breaking Change，但删除/重命名前必须完成生产、反射、DI、文档、测试、Benchmark、API contract 和 isolated nupkg consumer 搜索；存在真实使用价值的候选不能机械删除。

## D-006：本轮执行器

执行状态已通过 `task-state.mjs start ... --source copilot` 注册，当前模式为 `plan-execution`，task-id 固定为 `BING-OFFICES-PRE-RC-CLEANUP-20260901-001`。

## D-007：保留 Failure Workbook capability fallback

对 NPOI 2.7.4 net8.0 provider 做直接属性读写验证：`NPOI.HSSF.UserModel.HSSFWorkbook` 的 `IRow.Hidden` 与 `IRow.Collapsed` 抛出 `NotImplementedException`，`ZeroHeight` 可读写；`NPOI.XSSF.UserModel.XSSFWorkbook` 的三项均可读写。因此 `NpoiFailureWorkbookWriter.CopyRow` 的三个 `catch (NotImplementedException)` 仅包围对应 row capability，不是 catch-all，也不覆盖工作簿序列化、目标流复制或主异常路径。本轮保留该兼容 fallback，并在 Review 中将其列为已核实边界。

## D-008：P0-03 扫描以受控文本范围为准

P0-03 的最终扫描范围固定为 `src/`、`tests/`、`benchmarks/`、`docs/` 和根 `README.md`，使用 Python UTF-8 文本读取；排除 `bin/`、`obj/`、`artifacts/`、`output/`，防止生成 XML、旧二进制和历史报告污染引用计数。最终基线为：生产 C# 222 个文件、测试 C# 81 个文件、Benchmark C# 62 个文件、Markdown 14 个文件；`Obsolete` 14 处、生产 `EditorBrowsableState.Never` 71 处、生产 `NotImplementedException` 3 处、生产 `.Result/.Wait()/Task.Run/TODO` 分别为 `0/0/0/0`，生产 `InternalsVisibleTo` 6 处且只指向 Unit/Integration 测试友元。

## D-009：P0-03 完成不等于候选全部删除

本轮将“P0-03 VERIFIED”定义为：扫描口径、命中统计、逐符号定义、生产/测试/文档/Benchmark 引用、DI/反射检查、替代路径、删除风险和治理裁决均可复核；不将所有候选机械删除作为完成条件。`AddNpoi` 和 `MaxBytes` 已完成迁移并通过负向扫描；`ICellValueConverter`、DataTable/global CSV 状态、旧 validation attributes、`OfficeException` 层级、Settings、`UniqueTracker` 及 public execution detail 因仍有生产桥接、测试/文档/Benchmark 引用或兼容语义，保持 `BLOCKED/CANDIDATE/PARTIAL`，等待维护者逐项批准和迁移闭环。

## D-010：公开 API 分类账与正式 baseline 分离

`PublicApiContractTest` 已为当前公开顶层类型建立唯一分类：Abstractions 121（User API 70、Provider SPI 8、Execution detail 43）、Core 61（Compatibility 10、Execution detail 51）、NPOI 1（User API 1），合计 183；其中 `Execution detail` 94 个。分类账是治理证据，不是正式 API baseline；未获得维护者批准前，不更新旧 hash，也不以 `EditorBrowsable(Never)` 代替 internal 化或删除。

## D-011：P0-03 后续删除的统一前置条件

任何后续候选删除或 internal 化，必须先完成替代 User API/SPI、源码/字符串/`nameof`/反射/DI/程序集扫描、tests/Docs/Benchmark/package consumer 迁移，随后更新 API diff、Release build/test/pack 和负向扫描。若候选仍有外部使用价值，则在本任务保持 `BLOCKED` 或 `CANDIDATE`，并由后续变更记录批准人、后续 taskId、迁移期限和 No-Go 影响；本轮不修改 `review.md`，不伪造 Reviewer 通过。

## D-012：Round 5 已授权的 Obsolete/兼容 API 清理

- 授权来源：当前会话用户明确授权“删除已迁移的 `[Obsolete]` 相关文件”，本轮只按当前工作树已完成迁移的范围执行，不扩展到未授权的异常层级、Settings、public execution detail 或完整 P3 重构。
- 已覆盖范围：`ICellValueConverter`；`RequiredAttribute`、`RegexAttribute`、`RangeAttribute`、`MaxLengthAttribute`、`DateTimeAttribute`、`DuplicationAttribute`；`CsvSeparatorCharacter`、`CsvQuoteCharacter`；DataTable `CsvHelper` 的旧隐式 delimiter/quote 重载。
- 删除策略：不恢复 wrapper、forwarder、新 `[Obsolete]` 或仅 `EditorBrowsable(Never)` 的伪兼容层；显式 delimiter/quote 的 DataTable API 保留，避免把仍有明确行为测试的 API 一并误删。
- 验证要求：生产、测试、Docs、Benchmark、API contract、XML、DI/反射和 package-only consumer 均按 Round 5 重新扫描或运行；正式 API hash 不因本授权自动更新。

## D-013：未获具名批准的计划偏差、延期与 RC 影响

以下项目在本任务不继续实施，且不得写成已完成：

| 计划项 | 批准人 | 风险接受范围 | 后续 taskId | 迁移期限 | waiver | 当前裁决与 RC 影响 |
| --- | --- | --- | --- | --- | --- | --- |
| P2-01 剩余删除候选 | 无；当前会话未提供具名维护者批准 | 不接受剩余候选的 breaking/API 风险；仅接受已授权 Round 5 子集 | 未分配；不得伪造 | 未指定；不得伪造 | 无 | `BLOCKED/PARTIAL`；已授权的 Round 5 子集完成，其余候选需逐符号批准，RC `No-Go` |
| P2-02 public execution detail internal 化 | 无 | 不接受 source/binary breaking 风险；暂不 internal 化 | 未分配 | 未指定 | 无 | `BLOCKED/PARTIAL`；需 User API/SPI 替代和 breaking approval，RC `No-Go` |
| P3-01 NPOI import 拆分 | 无 | 不接受未有回归证明的大 diff/行为风险；暂缓 | 未分配 | 未指定 | 无 | `TODO/BLOCKED`；避免无批准的大 diff，RC `No-Go` |
| P3-02 CSV/Failure 拆分 | 无 | 不接受 CSV/Failure 行为回归风险；暂缓 | 未分配 | 未指定 | 无 | `TODO/BLOCKED`；避免扩大行为风险，RC `No-Go` |
| P3-03 Mapping hot-path 重构 | 无 | 不接受无完整 Benchmark 基线的性能回归风险；暂缓 | 未分配 | 未指定 | 无 | `TODO/BLOCKED`；需完整 Benchmark 基线，RC `No-Go` |
| P3-04 异步/所有权 ADR | 无 | 不接受无具名架构决策的 public async/ownership 变化；暂缓 | 未分配 | 未指定 | 无 | `TODO/BLOCKED`；需具名架构决策，RC `No-Go` |
| P5-01 完整 Benchmark/预算 | 无；预算未批准 | 不接受将 ShortRun 或未批准预算作为性能 Go 结论；暂缓 | 未分配 | 未指定 | 无 | `PARTIAL/BLOCKED`；ShortRun 仅作证据，RC `No-Go` |

“无”“未分配”“未指定”是当前真实治理状态，不是批准或 waiver。解除条件是由具名维护者/架构负责人补充批准人、后续 taskId、迁移期限或正式 waiver，并重新运行对应门禁；在此之前保留 `No-Go`。

### Round 6 用户批准实施记录

| 项目 | 批准范围 | 已实施内容 | 当前状态与限制 |
| --- | --- | --- | --- |
| P0-02 | 移除 net5.0/net7.0 | 清理共享 TFM、条件 PackageReference、API snapshot 条件和目标列表；受控源码配置无残留 | `netcoreapp3.1` runtime 仍缺失；历史 artifacts/报告未批量改写 |
| P2-01 | 继续逐符号删除 | 删除无生产调用的 Settings 与 Office 异常层级；类型映射/NPOI 错误改用标准异常 | DataTable 显式 helper 和其它候选仍保留，正式 API hash 待批准 |
| P2-02 | execution detail internal 化 | internal 化 Core plan/type-map/binding/loader/CSV concrete；DI 改为公开接口工厂注册 | `MappingConfigurationMerger`、provider 边界未强行 internal 化 |
| P3-01 | NPOI import 拆分 | 新增 `NpoiImportSheetExecutor`，Importer 保留 workbook/sheet orchestration | 需完整回归验证资源、取消、错误语义 |
| P3-02 | CSV 职责拆分 | 新增 `CsvPipelineSupport.cs`，分离记录读写、受限流和异常 | Failure Workbook 拆分尚未实施 |
| P3-03 | Mapping hot-path 拆分 | 新增 `ExcelMappingPlanCacheKey` 并从 Factory 提取缓存键 | rule-index/dynamic compiler 与正式性能基线仍待完成 |

### Round 8 P3-04 异步与所有权审计

- 生产代码扫描范围为 `src/`；未发现生产 `.Result`、`.Wait()`、`Task.Run` 或 `GetAwaiter().GetResult()`。Benchmark/测试中的同步等待仅属于基准协调和测试设施，不进入生产调用链。
- `IExcelImporter`、`IExcelExporter`、`ICsvImporter` 和 `ICsvExporter` 当前均为同步 API。NPOI `WorkbookFactory.Create`、Workbook 序列化和 DOM 遍历没有可由本项目控制的真实异步 I/O；本轮不增加 `async`/`Task` 公共包装，避免伪异步和额外线程池占用。
- 取消令牌在输入复制、Workbook/Sheet/Row/Picture/Relation/Failure Workbook 阶段按块或阶段检查；NPOI DOM 构造本身不可中断，取消只能在构造返回后生效，不能宣称任意时刻的即时取消。
- 导入器拥有 buffered source 和输入 Workbook，并在成功、异常、取消路径通过 `using` 释放；调用方输入流保持打开。导出器拥有创建的 Workbook，调用方目标流保持打开；File 扩展通过同目录临时文件和原子替换提交，异常/取消清理临时文件并保留旧目标。
- CSV reader/writer、limited stream 和 Failure Workbook limited stream 均不释放调用方底层流；File/Bytes 扩展只释放自身创建的 FileStream/MemoryStream。Failure Workbook 独立 Workbook 在 `finally` 关闭，临时输出流在复制完成或异常后释放。
- 结论：P3-04 的架构决策为“保持同步公共 API，明确 DOM 取消延迟和调用方流所有权；通过现有同步边界与回归测试验证，不新增伪异步抽象”。完整资源峰值和 Failure Workbook 双 DOM 的量化门禁仍属于 P1-06/P1-07/P5-03 未完成项。

### Round 6 验证补充

- `NpoiImportSheetExecutor` 使用共享 `NpoiSheetStructureException`，由外层 `NpoiExcelImporter` 统一归类工作表结构错误；固定列/别名/动态列/ReadColumnRange 的绑定行为保留，关键结构错误仍直接抛出既有 `InvalidOperationException` 边界。
- Core 的 internal Plan factory 不向 NPOI、Benchmark、ResourceProbe 暴露；这些跨程序集调用统一改用公开 `ExcelMappingPlanFactoryProvider.CreateDefault()` 和 `IExcelMappingPlanFactory`。
- 动态列文档示例不再实例化 internal `CsvEntityImporter`，改为公开 `ICsvImporter` + `AddBingOfficesNpoi()` DI 路径；Docs fence `11/11` 通过。
- Round 6 可运行 TFM 回归：StreamPipeline net6/net8 各 `90/90`，Integration net6/net8 各 `15/15`，Release solution build `0 error / 28 warning`。
- API contract 的类型清单、分类和 NPOI exact member 检查已同步；formal member hash 仍保持旧基线并真实失败，未经批准不更新，RC 继续 `No-Go`。
