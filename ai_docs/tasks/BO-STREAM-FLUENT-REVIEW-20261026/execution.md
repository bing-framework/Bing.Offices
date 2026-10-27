<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BO-STREAM-FLUENT-REVIEW-20261026
AI_EXECUTION_FINISHED_AT: 2026-10-27T00:44:59+08:00

# P2 Review 修复记录

## Round 1

- 模式：REVIEW_FIX；fixScope=recommended。
- 输入：同目录 `review.md` 的 FIX-001 与 FIX-002；用户于 2026-10-27 明确要求修复 P2。
- 原审查为能力审计，没有实施计划。本轮补充 `plan.md` 记录当前授权范围，没有回填或改写历史审批。
- 基准提交：`84958c1cc5eec5afbae51d76b0d07b4ab133b748`，工作区包含之前的包验证和 CI 修改；这些已有变更均保留。
- 执行结果：2/2 修复完成，相关构建和最终测试通过。
- 原 `review.md` 保持不变。COMPLETED 表示本轮 Executor 修复已完成，不代表已取得下一轮独立 Review 的 PASS。

### FIX-001：Fluent 属性表达式校验

- 严重程度：MEDIUM / P2；处理要求：SHOULD_FIX。
- 执行状态：COMPLETED；本轮 Finding 状态：CLOSED。
- 根因：只验证末端属性的声明类型，未验证属性的接收者，导致其他对象的同名属性路径被丢弃。
- 修复：`ImportMappingBuilder<T>.Property` 使用的 `ResolveProperty` 与 `ExportMappingBuilder<T>.Property` 都要求 `MemberExpression.Expression` 与 lambda 的唯一参数为同一对象。校验位于配置查找和新增之前，拒绝嵌套、捕获对象、新建对象、静态属性。
- 保留原有 `DeclaringType == typeof(T)` 合同，不扩大继承属性或嵌套映射支持。两个公开方法签名不变，补充参数与异常 XML 说明。
- 新增职责级测试：`tests/Bing.Offices.Tests/MappingBuilderPropertyTest.cs`，导入/导出分别验证四类非法接收者 × 空/已有配置，以及合法属性、重复配置、快照隔离、空表达式。非法输入断言准确异常类型与 `ParamName`，并比较完整配置序列化结果，确保失败后无状态变更。
- 验证：20 个直接测试用例在 net6.0 与 net8.0 均通过；映射/Profile/Workbook 请求回归和两套 API snapshot 合同测试通过。

### FIX-002：日期文档与实际入口对齐

- 严重程度：MEDIUM / P2；处理要求：SHOULD_FIX。
- 执行状态：COMPLETED；本轮 Finding 状态：CLOSED。
- 修改：`docs/excel/dates.md` 明确固定属性的精确格式/Culture 使用 `ExcelDateAttribute`，区分 Sheet Culture、导出 `HasFormatter`、动态列日期校验以及命名转换器；不再宣称存在专用的固定属性日期 Fluent 入口。
- 补充实际限制：`ExcelDateAttribute.OffsetMinutes` 为 `int?`，不能作为 C# 特性命名参数赋值。文档改为指引无 offset 文本使用自定义转换器显式处理偏移；带 offset 的 ISO 文本继续使用默认解析合同。本轮不改变该公共属性类型。
- 示例：保留可编译的 `DateTime` 格式/Culture 模型示例，使用 `[ExcelDate("dd/MM/yyyy", CultureName = "fr-FR")]`。
- 验证：扩展现有 Markdown 示例测试，将示例原文（包括 using 与模型声明）编译后通过 `AddBingOfficesCore` / `ICsvImporter` 导入真实 UTF-8 CSV。完整断言成功/完整标志、零错误、唯一结果、日期和 `DateTimeKind.Unspecified`。
- Docs 验证过程中修正了新增测试支架中的字符串闭合和注册方法名，最终运行通过。未修改生产日期解析或转换器实现。

## 成员行为与迁移记录

| 生产符号 | 变化 | 调用方处理 |
|---|---|---|
| `ImportMappingBuilder<T>.Property<TProperty>` | 不合法的属性接收者立即抛 `ArgumentException`，参数名为 `expression` | 使用 `row => row.Code` 等当前模型直接属性；嵌套数据先投影为平面模型，或使用既有自定义转换方案 |
| `ExportMappingBuilder<T>.Property<TProperty>` | 同上；不会再把其他对象的属性静默当成当前模型属性 | 同上 |
| `ExcelDateAttribute` 等日期 API | 无运行时和签名变化，仅纠正说明 | 固定日期格式/Culture 使用合法特性配置；固定 offset 不要复制不可编译的 nullable 特性命名参数写法 |

API shape 变化为 0：`PublicApiContractTest` 的 net6.0/net8.0 成员快照比较通过，未重写任何批准 baseline。

## 最终生产符号 → 测试方法

| 生产符号/关键行为 | 测试项目 | 方法 |
|---|---|---|
| `ImportMappingBuilder<T>.Property` / `ResolveProperty`：非法接收者、失败状态 | Bing.Offices.Tests | `MappingBuilderPropertyTest.ImportProperty_InvalidReceiver_ShouldRejectWithoutChangingConfiguration` |
| 导入直接属性、重复配置、独立快照 | Bing.Offices.Tests | `MappingBuilderPropertyTest.ImportProperty_DirectReceiver_ShouldConfigureAndBuildIndependentSnapshots` |
| 导入 null 参数 | Bing.Offices.Tests | `MappingBuilderPropertyTest.ImportProperty_NullExpression_ShouldRejectWithoutAddingColumn` |
| `ExportMappingBuilder<T>.Property`：非法接收者、失败状态 | Bing.Offices.Tests | `MappingBuilderPropertyTest.ExportProperty_InvalidReceiver_ShouldRejectWithoutChangingConfiguration` |
| 导出直接属性、重复配置、独立快照 | Bing.Offices.Tests | `MappingBuilderPropertyTest.ExportProperty_DirectReceiver_ShouldConfigureAndBuildIndependentSnapshots` |
| 导出 null 参数 | Bing.Offices.Tests | `MappingBuilderPropertyTest.ExportProperty_NullExpression_ShouldRejectWithoutAddingColumn` |
| 导入/导出 Fluent 配置进入 NPOI 真实流程 | Bing.Offices.Tests | `StreamPipelineTest.StreamPipeline_FluentMappingConfiguration_ShouldApplyToSingleRequest` |
| `ExcelDateAttribute` 文档示例 → `ICsvImporter.Import<T>` | Bing.Offices.Docs.Tests | `DocsConsumerTest.DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually` |
| Public API 双 TFM 成员形状 | Bing.Offices.Tests | `PublicApiContractTest.PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` |

## 最终验证证据

| 范围 | 结果 | TRX |
|---|---|---|
| net8.0 新增构建器直接测试 | 20/20 PASS | `TestResults/stream-fluent-fix/p2-direct-net8.trx` |
| net6.0 直接测试、Profile/映射补丁/Workbook 请求、API 合同 | 143/143 PASS | `TestResults/stream-fluent-fix/p2-regression-net6.trx` |
| net8.0 Profile/映射补丁/Workbook 请求、API 合同及相关 Excel/CSV 消费路径 | 127/127 PASS | `TestResults/stream-fluent-fix/p2-regression-net8.trx` |
| net8.0 Markdown 示例编译执行与方向配置外部消费者 | 2/2 PASS | `TestResults/stream-fluent-fix/p2-docs-net8.trx` |

最终运行合计 292/292，通过、0 失败、0 跳过；该数量按 TFM 统计，同一行为在不同 TFM 分别执行。Docs 的一个测试内部验证全部 11 个 C# 代码块，不额外叠加到 292 中。

所有运行均为 Release / `--no-restore`。编译覆盖 Abstractions、Core、Npoi、ProfileFixtures 和相应测试项目；net8 单元测试构建还生成 ResourceProbe 依赖，但未运行重型资源实验。

复现命令：

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter 'FullyQualifiedName~MappingBuilderPropertyTest' --logger 'trx;LogFileName=p2-direct-net8.trx' --results-directory TestResults/stream-fluent-fix
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore --filter 'FullyQualifiedName~MappingBuilderPropertyTest|FullyQualifiedName~MappingProfileV2Test|FullyQualifiedName~MappingConfigurationPatchTest|FullyQualifiedName~MappingProfileRegistryTest|FullyQualifiedName~MappingProfileServiceCollectionExtensionsTest|FullyQualifiedName~ExcelWorkbookRequestTest|FullyQualifiedName~PublicApiContractTest' --logger 'trx;LogFileName=p2-regression-net6.trx' --results-directory TestResults/stream-fluent-fix
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter 'FullyQualifiedName~MappingProfileV2Test|FullyQualifiedName~MappingConfigurationPatchTest|FullyQualifiedName~MappingProfileRegistryTest|FullyQualifiedName~MappingProfileServiceCollectionExtensionsTest|FullyQualifiedName~ExcelWorkbookRequestTest|FullyQualifiedName~PublicApiContractTest|FullyQualifiedName~StreamPipelineTest.StreamPipeline_FluentMappingConfiguration_ShouldApplyToSingleRequest|FullyQualifiedName~StreamPipelineTest.TypeMap_RequestConfiguration_ShouldOverrideProfile|FullyQualifiedName~StreamPipelineTest.TypeMap_Profile_ShouldKeepIndependentConfigurationSnapshot|FullyQualifiedName~CsvTest.EntityPipeline_NormalizedDocument_ShouldUseDirectionalMappings' --logger 'trx;LogFileName=p2-regression-net8.trx' --results-directory TestResults/stream-fluent-fix
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually|FullyQualifiedName~MappingV2_ExternalConsumer_ShouldBuildDirectionalRequests' --logger 'trx;LogFileName=p2-docs-net8.trx' --results-directory TestResults/stream-fluent-fix
```

## Round 1 汇总与停止条件

- Implementation：2/2（100%）；本次修复验证门槛全部通过，不作新的发布门槛结论。
- Changed Scope：两个构建器的输入校验、直接测试、日期说明与其可执行示例测试。
- L0：依赖构建、UTF-8 检查、`git diff --check`、最终差异检查通过。
- L1：双 TFM 构建器职责测试和映射相关回归通过。
- L2：采用受影响模块专项集，未重复运行无关完整单元项目。
- L3：公开 API 双 TFM snapshot 与文档消费者通过。
- L4/L5：未运行，无签名、Provider IO、DOM、缓存键、算法或资源路径变化。
- Reused Evidence：上一轮 IO/文件提交边界属于 REUSED_UNCHANGED_SCOPE，仅保留原结论，不计入本轮测试数量或声明新鲜全量通过。
- Completed：FIX-001、FIX-002。
- Open Actionable：无（本轮 Executor 范围内）。
- Blocked Approval / Blocked External：本轮修复无。
- Not Applicable：新的性能/容量比较及包发布。
- Accepted Limitations：既有 Excel DOM/低内存非目标不变；nullable OffsetMinutes 特性参数限制已在日期文档公开，并给出转换器替代路径。
- Verified Boundaries：本轮未重新测量资源边界。
- Deferred：原报告列出的可选 Fluent 入口、继承属性支持、CSV 异步序列与 Excel Streaming 新能力，不属于本次 P2 修复。
- No-Progress Check：CHANGED。
- Next Action：STOP；后续独立 Review 可依据本报告和原始 TRX 复核，不自动再进入 Fix。
- 状态：COMPLETED；任务使用 `task-finish --no-notify` 收口，不发送外部通知，不 commit/push。
