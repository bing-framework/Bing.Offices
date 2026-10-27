<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BO-STREAM-FLUENT-REVIEW-20261026
AI_REVIEWED_AT: 2026-10-26T23:48:52+08:00

# Fluent 与 Streaming API 整体检查

## 结论与范围

尚不能认定两类“流式 API”全部完成。Fluent 链式配置的主要调用链已经接通，但有 2 项需要修正的问题；Stream/Async 读写已实现，逐行向调用方交付数据、异步序列输入以及 Excel 有界内存处理尚未实现。

本次是用户要求的现有能力审计，覆盖 Fluent 配置与 Streaming IO 两者，只检查、不修复。没有新功能实施计划，因此本报告不是一个新计划的正式验收，也不重开历史发布计划全部阶段。历史计划仅用于确认既有承诺和明确的设计边界：

- `ai_docs/tasks/BO-RC-20260907-001/plan.md` 及 `execution.md`：Phase 3–5 的文件提交、CSV Async、Excel 外围 Async IO，以及 Phase 6 的测试合同。
- `ai_docs/tasks/BO-API-USABILITY-20261025/plan.md` 及 `execution.md`：映射、导入边界和调用方结果语义。
- `docs/architecture/README.md`、`docs/excel/async-io.md`、`docs/excel/dates.md`：当前支持范围与公开使用说明。

本轮未修改生产代码、测试源码、API 基线或上述计划/执行记录。工作区原有包验证和 CI 变更保留。

## 能力矩阵

| 能力 | 当前情况 | 结论/边界 |
|---|---|---|
| Workbook/Sheet 链式构建 | 导入、导出、模板、工作表、样式、映射、校验等配置进入实际执行流程 | 主流程已实现；本轮未对每一种图表/样式组合重新穷举 |
| 方向专用列映射 | 导入/导出各自的构建器、Profile、配置快照与请求覆盖 | 已实现；属性表达式校验存在 FIX-001 |
| 配置合并 | Profile → Document → Request，类型映射继续应用属性配置 | 已接入 `ExcelMappingPlanFactory` 和 `ExcelTypeMapFactory`；相关回归通过 |
| 固定属性日期 Fluent 配置 | 无精确日期格式/offset 的专用导入 Fluent 入口；Attribute 与自定义转换器可用 | 文档宣称超出实际能力，见 FIX-002 |
| Excel Stream 导入/导出 | Sync/Async 入口、工作簿处理、失败工作簿输出均有实际实现 | 已实现 Stream 接口；内部仍使用完整 NPOI DOM |
| Excel Async | 输入复制、外围结果传输使用真实异步 IO | DOM 创建、处理、序列化仍同步；模板必须同步可读 |
| CSV 导入 | 内部逐记录解析，Sync/Async 共用处理语义 | 返回前累积实体集合，调用方不能逐行消费；结果内存随行数增加 |
| CSV 导出 | 遍历 `IEnumerable<T>`，逐记录写出；Async 使用异步输出 | 已具备增量写出，不预先将全部输入实体转成列表；不代表任意单条记录都有固定内存上限 |
| 异步序列 | 无公共 `IAsyncEnumerable<T>` 导入结果或导出输入入口 | 未实现，属于新能力 |
| 文件提交、取消和资源清理 | 临时文件提交、取消复查、失败保留既有目标、临时文件清理 | 本轮针对性测试通过；最终耐久性 flush/文件替换仍包含同步操作 |
| Excel 大文件有界内存 | 输入先复制到 MemoryStream，再加载 DOM；导出完整 DOM 后序列化 | 未实现；当前架构已明确排除 streaming 低内存承诺 |

主要源码证据：

- `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs:154`：合并 Profile、Document、Request 后编译映射。
- `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiExcelImporter.cs:203`、`:210`、`:356`：完整输入缓冲、异步复制、创建 NPOI DOM。
- `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExcelExporter.cs:227`：先执行同步导出到 staging，再异步传输。
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityImporter.cs:169`、`:289`、`:333`：异步导入创建列表、添加实体、返回汇总结果；同步路径同样累积列表。
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityExporter.cs:207`、`:269`：同步/异步导出均按输入序列遍历。
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvPipelineSupport.cs`：记录状态机、异步读写包装和内部回调，不是公共逐行消费者接口。
- `docs/architecture/README.md:29`、`:38`：低内存 Streaming 非目标和 NPOI 同步 DOM 边界。

## FIX-001：拒绝不指向当前实体直接属性的表达式

- 严重程度：MEDIUM（P2）
- 处理要求：SHOULD_FIX
- 状态：OPEN_ACTIONABLE
- 问题：两个方向的 `Property(...)` 只检查表达式末端是否为属性及其 `DeclaringType`，未检查接收者是否就是 lambda 参数。嵌套、捕获对象、新建对象和静态属性表达式均可能被接受。
- 证据：`ImportMappingBuilder.cs:46` 与 `ExportMappingBuilder.cs:25`。独立探针覆盖导入/导出两方向，每方向 1 个正常直接属性和 4 个非法表达式；8 个非法表达式全部被接受。原始结果为 `TestResults/stream-fluent-review/fluent-expression-probe.json`。
- 复现：当 `Row.Child` 也是 `Row` 类型，`Property(row => row.Child.Code)` 最终只存入 `PropertyName = "Code"`。`row => outside.Code`、`row => new Row().Code` 同样变成根属性 `Code`；`row => Row.StaticCode` 也被构建器接受。
- 影响：调用方表达式的对象/路径信息被丢弃，配置可能静默作用于根对象同名属性。静态属性则延后至映射编译时才被拒绝。`ExcelTypeMapFactory.cs:88` 按名称应用配置，`:177` 仅枚举公开实例属性，证实后续流程不会恢复原始表达式路径。
- 修复目标：在创建或修改列配置前准确验证“当前实体的直接实例属性”，让非法表达式立即得到明确异常。
- 修复要求：导入/导出规则一致；校验成员接收者与 lambda 参数的身份及实例属性约束；失败不得污染既有构建状态。不要顺带引入嵌套属性映射。继承属性是否支持需明确合同，不要仅凭本项擅自扩大能力。
- 验证方式：两个默认构建器分别直接测试上述四种非法表达式、正常直接属性、已有配置在失败后不变及独立快照。修复后运行映射构建/配置合并回归；若不改变成员签名，不应产生 API 基线差异。

## FIX-002：修正日期配置说明与实际 API 的差异

- 严重程度：MEDIUM（P2）
- 处理要求：SHOULD_FIX
- 状态：OPEN_ACTIONABLE
- 问题：`docs/excel/dates.md:3` 宣称可通过 Fluent 或 Mapping 文档显式配置精确日期格式和 Culture，没有区分固定 CLR 属性、动态列校验与导出格式。
- 证据：`ImportColumnMappingBuilder` 没有固定属性日期格式/offset 方法，`ExcelColumnConfiguration` 也没有相应日期字段。`ExportColumnMappingBuilder.HasFormatter` 是导出显示格式，不能设置导入日期解析。`ExcelMappingDynamicValidationConfiguration` 的 Format/Culture 属于动态日期校验，不等于固定属性日期解析配置。
- 已有能力：`ExcelDateAttribute` 提供 Format、CultureName、OffsetPolicy 和 OffsetMinutes；命名转换器也可承载自定义转换。Sheet 层 Culture 的存在并不补齐固定属性精确日期格式与 offset 的 Fluent 契约。
- 影响：按日期文档接入时找不到对应 Fluent 方法，或误将动态校验/导出格式配置用于固定属性导入，无法得到文档描述的结果。
- 修复目标：让日期文档准确对应可调用的公开 API。
- 修复要求：最小修复是明确固定属性使用 Attribute/自定义转换器，区分导出 formatter、Sheet Culture 与动态日期校验，不再声称存在专用日期 Fluent 入口。若后续确实新增该能力，应作为功能增量贯通配置字段、克隆、合并、缓存键、计划和执行，配套双 TFM API snapshot、追溯与迁移记录，不能只添加空入口。
- 验证方式：核对文档示例可编译并走实际转换路径；若仅修正文案，无须重新跑大规模资源测试。若新增运行时能力，直接覆盖 XLS/XLSX/CSV、Sync/Async、格式/Culture/offset、优先级及缓存隔离。

## 后续能力与可选优化

以下不计入本轮 OPEN_ACTIONABLE，也不作为既有 Stream/Async 计划的失败项。

| 项目 | 分类 | 当前可用路径 / 后续方向 |
|---|---|---|
| CSV 公共逐行导入与异步序列导出 | 新功能 | 增加行消费者/异步枚举合同前，先确定错误收集、提前停止、取消及流所有权；当前返回聚合结果 |
| Excel 有界内存读写 | 已明确非目标；若需要则新立项 | 需要 SAX/事件读取或专用写出 Provider，并定义与模板、图表、图片、多 Sheet 关系等功能的兼容范围；不能仅给现有 DOM 方法改名 |
| Fluent 清除/重置配置入口 | OPTIONAL | DTO 已支持 ClearTitle、ClearAliases、ResetColumnIndex、ClearFormatter、ResetDecimalScale、ClearValueMappings 等，链式入口尚未逐一覆盖；当前可直接构建 DTO |
| 继承属性的 Fluent 可用性 | OPTIONAL | `DeclaringType == typeof(T)` 会拒绝基类属性；底层实例属性映射可发现继承属性。应明确是否支持后再调整并补测试 |
| “流式导入器/导出器” XML 摘要措辞 | OPTIONAL | 接口摘要容易被理解为有界内存；建议统一为“基于 Stream 的导入/导出”，保留 DOM 说明 |
| CSV 异步输入所有权与读异常直接测试 | OPTIONAL / NOT_VERIFIED | 可进一步补充独立职责断言；本轮没有发现相应运行时缺陷，不将缺少断言等同实现失败 |

## 本轮验证

| 执行项 | 结果 | 原始证据 |
|---|---|---|
| net8.0 Release 映射/Profile、配置补丁、Workbook Request、Async 管线、模板边界、默认文件提交、CSV 记录读取专项 | 153/153 通过，0 失败，0 跳过 | `TestResults/stream-fluent-review/stream-fluent-unit-net8.trx` |
| net8.0 Release Excel/CSV 异步真实文件集成 | 23/23 通过，0 失败，0 跳过 | `TestResults/stream-fluent-review/stream-file-net8.trx` |
| Fluent 表达式独立最小复现 | 2 个直接属性正常；8 个非法表达式被错误接受 | `TestResults/stream-fluent-review/probe/Program.cs` 与 `fluent-expression-probe.json` |
| 生产代码伪异步扫描 | 未发现 `Task.Run`、`.Result`、`.Wait()` 匹配 | 对 `src` 的文本扫描，仅作为辅助证据 |

本轮共新运行 176 项自动化测试，全部通过。通过结果说明已有合同的相关回归仍稳定，不覆盖 FIX-001 的非法表达式，也不能证明不存在文档能力缺口。

执行命令：

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter 'FullyQualifiedName~MappingProfileV2Test|FullyQualifiedName~MappingConfigurationPatchTest|FullyQualifiedName~ExcelWorkbookRequestTest|FullyQualifiedName~AsyncPipelineTest|FullyQualifiedName~TemplateAsyncBoundaryTest|FullyQualifiedName~DefaultFileExportCommitterTest|FullyQualifiedName~CsvRecordReaderConfigurationTest' --logger 'trx;LogFileName=stream-fluent-unit-net8.trx' --results-directory TestResults/stream-fluent-review
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-restore --filter 'FullyQualifiedName~ExcelAsyncFileIntegrationTest|FullyQualifiedName~CsvAsyncFileIntegrationTest' --logger 'trx;LogFileName=stream-file-net8.trx' --results-directory TestResults/stream-fluent-review
```

验证边界：本轮没有重新执行 net6.0 全套、Benchmark、重型资源矩阵或远端 CI。原有双 TFM/包验证证据不计入这 176 项；本报告不对未重测的发布资源门槛作新结论。探针及 TRX 位于被忽略的 TestResults 目录，报告保留了复现条件与结果。

## 收口状态

- OPEN_ACTIONABLE：2（FIX-001、FIX-002）。
- 当前行为审查结论：NEEDS_FIX。
- 检查已完成；本轮没有自动进入 Fix 循环。
- 建议先修复表达式校验和日期文档，再分别决定是否补齐 Fluent 便捷入口、CSV 行消费者，以及需要独立设计的 Excel 低内存 Provider。
