# Continuation 14：分析器准确性、包消费与命名布局合同

## 范围与顺序

按上阶段 TODO 顺序实施三项：先修复静态分析器对命名实参、继承属性与内联定义的判定；再从独立 NuGet 包消费者验证分析器实际加载；最后扩展可复用 Entity Provider 合同，覆盖属性式固定单元格与命名模板锚点。合同发现的命名列表占位地址误判纳入本阶段修复。

## Change Impact

- ChangedProjects：Analyzers、Analyzers.Tests、新增 Analyzers.Consumer、Abstractions、Npoi、ClosedXml、Testing、ProviderContract.Tests、ThirdPartyProvider.Consumer、Npoi.Tests、Docs。
- ChangedPublicContracts：无新增或删除公共成员；新增测试合同入口仅位于独立源码模板。
- ChangedRuntimePaths：实体布局构建及 NPOI/ClosedXML 命名列表导入导出的冲突检查。
- ChangedBuildPackaging：分析器本地包消费并接入 CI 包验证；第三方消费者从本地生成的五个生产 NuGet 包运行。
- RiskLevel：MEDIUM；重点是分析器误报和命名地址解析后才可判定的布局冲突。

## 验证

1. 分析器双 TFM 直接回归；本地打包后消费者正常构建及 `BOE_NEGATIVE` 预期编译失败。
2. 可复用模板在两个 Provider、双 TFM 验证成功往返和冲突负例；NPOI 追加 HSSF/XSSF 成功、失败矩阵。
3. Provider Contract、NPOI、ClosedXML 的 `Category!=Large` 常规测试；第三方 NuGet 包消费者双 TFM。
4. Docs、Release 构建、diff 与严格 UTF-8/BOM/EOL 检查。无公共 API 差异，复用现有双 TFM API 快照审批证据；不运行无关性能矩阵。

## 后续顺序

独立 Provider 合同测试包与版本/发布流程另立阶段；随后评估 Footer 公式或签字区、旧版迁移分析器。Source Generator 和 DOM/Streaming 策略须先有性能或容量证据。本阶段不自动提交、推送或发布。
