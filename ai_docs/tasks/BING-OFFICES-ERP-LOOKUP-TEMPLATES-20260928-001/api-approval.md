# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001
- approvedBy: user (明确要求实施本任务计划)
- approvedAt: 2026-09-29 (Asia/Shanghai)
- 范围：仅增加 ERP 文本字典、名称列表来源、Entity 原生校验选项及 Provider 能力标志。

双 TFM 候选快照见 `api-snapshot-net6.0.json`、`api-snapshot-net8.0.json`；成员比较见 `api-diff.json`。两个 TFM 均只在 `Bing.Offices.Abstractions` 增加 14 个成员，删除 0 个。现有公共构造、接口和 Fluent 签名保持不变。成员差异核对及审批后，已更新 `build/api-snapshot-baseline.json`；公共 API 合同测试通过。

## 批准的新增成员

- `ExcelWorkbookExportBuilder.LookupList`、`ExcelEntityLayoutBuilder<TEntity>.LookupList`、`ExcelEntityLayoutBuilder<TEntity>.DataValidation`。
- `ExcelDataValidationDefinition.ListSourceName`；列表规则与 `Values` 二选一。
- `ExcelEntityImportOptions` 四参数构造，以及 `ValidationMode`、`UnsupportedFeaturePolicy` 属性。旧构造继续默认 `ConfiguredRules`。
- `ExcelProviderFeatures.NamedListValidation`。
- 隐藏于 IntelliSense 的 Provider SPI：`ExcelLookupListDefinition` 类型及其 `Name`、`Values`，`ExcelWorkbookExportRequest.LookupLists`，`ExcelEntityLayout<TEntity>.LookupLists`、`DataValidations`。

## 迁移

现有内联列表和仅配置校验的导入调用无需修改。超过内联长度或包含分隔符的文本列表改用 `LookupList(name, values)`，下拉声明改为 `ListSourceName = name`。Entity 导入若需检查工作簿原生规则，显式传入 `ExcelEntityImportOptions(..., ExcelImportValidationMode.ConfiguredAndWorkbook, ExcelUnsupportedFeaturePolicy.Fail)`；业务规则仍由应用配置维护。
