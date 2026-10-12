# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-CASCADE-LOOKUPS-20260929-001
- approvedBy: user（要求按建议顺序继续实现 TODO 能力）
- approvedAt: 2026-09-29 (Asia/Shanghai)
- 范围：仅增加两级联动文本字典、Workbook/Entity 配置入口、原生列表规则属性和 Provider SPI；不删除公共成员，不修改既有签名。

## 批准的新增成员

- 隐藏于 IntelliSense 的 `ExcelCascadeListDefinition`、`ExcelCascadeListEntry` 类型及 `Name`、`Entries`、`Parent`、`Values` 属性。
- `ExcelWorkbookExportBuilder.CascadeList(string, IReadOnlyDictionary<string, IReadOnlyList<string>>)`。
- `ExcelEntityLayoutBuilder<TEntity>.CascadeList(string, IReadOnlyDictionary<string, IReadOnlyList<string>>)`。
- `ExcelDataValidationDefinition.CascadeSourceName`、`ParentColumnIndex`。
- 隐藏于 IntelliSense 的 `ExcelWorkbookExportRequest.CascadeLists`、`ExcelEntityLayout<TEntity>.CascadeLists`。
- `ExcelProviderFeatures.CascadeListValidation`。

## 兼容性与迁移

现有内联与静态名称列表配置不变。需要联动填写时，声明 `CascadeList`，在上级列使用已有 `ListSourceName`，下级单列区域设置 `CascadeSourceName` 与同一行上级列的 `ParentColumnIndex`。仅 NPOI XLS/XLSX 和 ClosedXML XLSX 提供该导出能力。级联公式只服务于 Excel 填写；导入业务合法性仍需配置服务端规则，原生公式按既有 Unsupported 策略处理。

双 TFM 各新增 13 个成员、删除 0 个；对照 `api-diff.json` 核对后更新 `build/api-snapshot-baseline.json`。
