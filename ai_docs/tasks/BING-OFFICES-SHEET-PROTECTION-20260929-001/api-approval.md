# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-SHEET-PROTECTION-20260929-001
- approvedBy: user（要求按建议顺序继续实现 TODO 能力）
- approvedAt: 2026-09-29 (Asia/Shanghai)
- 范围：仅增加工作表保护定义、Workbook/Entity 配置入口、Provider SPI 属性及细粒度能力位；不删除公共成员，不改变既有签名。

## 批准的新增成员

- `ExcelSheetProtectionDefinition` 及其构造函数、`EditableRanges`、`Password`。
- `ExcelSheetExportBuilder<T>.Protection(ExcelSheetProtectionDefinition)`。
- `ExcelEntityLayoutBuilder<TEntity>.Protection(string, ExcelSheetProtectionDefinition)`。
- 隐藏于 IntelliSense 的 Provider SPI：`ExcelSheetExportRequest.Protection`、`ExcelEntityLayout<TEntity>.Protections`。
- `ExcelProviderFeatures.SheetProtection`。

## 兼容性与迁移

现有 Workbook、Entity、模板调用无需修改。需要填写区保护时，新增 `Protection` 配置；区域采用零基闭区间。Excel 工作表保护不代替导入校验或文件加密。MiniExcel、SpreadCheetah 对此请求返回结构化 Unsupported；NPOI 支持 XLS/XLSX，ClosedXML 支持 XLSX。

双 TFM 成员级差异见 `api-diff.json`，候选快照见 `api-snapshot-net6.0.json`、`api-snapshot-net8.0.json`。
两个目标框架均仅在 `Bing.Offices.Abstractions` 增加 9 个成员，删除 0 个；对照审批清单核对后更新 `build/api-snapshot-baseline.json`。
