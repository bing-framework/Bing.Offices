# P2 修复范围与验证计划

- Task ID：BO-STREAM-FLUENT-REVIEW-20261026
- 模式：REVIEW_FIX，fixScope=recommended
- 授权：用户于 2026-10-27 明确要求修复上一轮列出的 P2 问题。
- 直接任务契约：同目录 `review.md` 中 FIX-001 与 FIX-002。
- 原检查任务没有实施计划；本文件仅记录本次已授权的修复边界，不回填为检查前已批准的历史计划。

## 修复项

1. FIX-001：导入、导出列映射构建器只接受 lambda 参数自身的直接实例属性。保留原有属性声明类型限制，非法表达式在修改配置前抛出 ArgumentException。补充两个实现的职责级测试，覆盖四类非法表达式、空/已有配置、正常访问、空参数和快照隔离。
2. FIX-002：修正日期文档，区分固定属性的 Attribute/自定义转换器、导出 formatter、Sheet Culture 和动态日期校验。采用报告要求的最小文档修复，不新增日期 Fluent API。

## Change Impact Analysis

| 项目 | 范围 |
|---|---|
| ChangedFiles | ImportMappingBuilder.cs、ExportMappingBuilder.cs、新增 MappingBuilderPropertyTest.cs、docs/excel/dates.md、DocsConsumerTest.cs、任务记录 |
| ChangedProjects | Bing.Offices.Abstractions、Bing.Offices.Tests、Bing.Offices.Docs.Tests |
| ChangedPublicContracts | 方法签名不变；落实已有“直接属性”参数合同，非法表达式提前拒绝 |
| ChangedRuntimePaths | Fluent 配置构建时的表达式检查 |
| ChangedProviders | 无 Provider 实现修改 |
| ChangedTFMs | 无；在 net6.0/net8.0 验证 |
| ChangedBuildPackaging | 无 |
| ChangedBenchmarkHarness | 无 |
| ChangedDocsOnly | FIX-002 |
| AffectedDependents | Profile 与请求映射配置消费者；合法配置的映射编译、CSV/NPOI 消费路径 |
| RiskLevel | LOW：修正非法输入；合法配置、缓存、转换和 IO 路径不变 |

## 验证

- L0：受影响项目及依赖编译、UTF-8/差异检查。
- L1：导入/导出构建器分别测试非法表达式及状态保持，正常属性及独立快照；先 net8.0 精准测试。
- 回归：net6.0/net8.0 的 MappingProfileV2Test、MappingConfigurationPatchTest、映射 Profile 解析、Workbook 请求相关测试及 PublicApiContractTest。
- API：用既有双 TFM snapshot 合同测试证明无成员形状变化，不覆盖批准基线。行为修正及迁移建议写入 execution.md，建立生产符号到测试方法的映射。
- FIX-002 为说明修正并添加合法日期模型示例，逐项核对真实源码；现有 Markdown 示例测试原文提取编译，并通过公开 CSV 服务验证实际日期结果，不添加文案字符串断言。
- 不重跑无关全量测试、包发布、重型 Benchmark/ResourceProbe；既有 IO/资源边界证据属于未变化范围。
- 汇总每个 FIX 的修复与验证到 execution.md；保留 review.md 原始证据，以 `task-finish --no-notify` 收口。

## 范围排除

不实现可选 Fluent 清除/重置方法、继承属性新支持、CSV 行消费者、IAsyncEnumerable 或 Excel 低内存 Provider；不改包版本、提交/推送代码或发送外部通知。
