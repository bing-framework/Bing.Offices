# 生产符号到测试方法

| 生产符号 / 行为 | 测试项目 | 测试方法 |
| --- | --- | --- |
| `ExcelSheetProtectionDefinition` 区域快照、参数校验 | Npoi.Tests | `SheetProtectionModelTest.Protection_ShouldSnapshotEditableRanges`；`Protection_ShouldRejectInvalidDefinitions` |
| `ExcelSheetExportBuilder<T>.Protection`、`ExcelSheetExportRequest.Protection` | Npoi.Tests；ClosedXml.Tests | `WorkbookSheetProtectionTest.Export_Protection_ShouldLockSheetAndPreserveEditableTemplateStyle` |
| `ExcelEntityLayoutBuilder<TEntity>.Protection`、`ExcelEntityLayout<TEntity>.Protections` | Npoi.Tests；ClosedXml.Tests | `WorkbookSheetProtectionTest.EntityLayout_Protection_ShouldApplyAcrossFormatsAndTemplates`；`EntityLayout_Protection_ShouldApplyToTemplateAndPreserveEditableStyle` |
| NPOI XLS/XLSX 保护、样式与区域边界 | Npoi.Tests | `Export_Protection_ShouldLockSheetAndPreserveEditableTemplateStyle`；`Export_ProtectionOutsideBounds_ShouldPreserveDestination` |
| ClosedXML XLSX 保护、样式与区域边界 | ClosedXml.Tests | `Export_Protection_ShouldLockSheetAndPreserveEditableTemplateStyle`；`Export_ProtectionOutsideXlsxBounds_ShouldPreserveDestination` |
| 未配置保护的模板保持原状态 | Npoi.Tests；ClosedXml.Tests | `EntityLayout_WithoutProtection_ShouldPreserveTemplateProtection` |
| 同步/异步保护和预取消目标完整性 | Npoi.Tests；ClosedXml.Tests | `ExportAsync_Protection_ShouldWriteAndHonorPreCancellation` |
| MiniExcel / SpreadCheetah 保护请求预检拒绝 | MiniExcel.Tests；SpreadCheetah.Tests | `SheetProtectionUnsupportedTest.Export_ShouldRejectProtectionBeforeWriting` |
| `ExcelProviderFeatures.SheetProtection`、双 TFM 公共 API | Tests.Integration | `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`；`PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline` |
| NuGet 包公开 API 消费 | Consumer.Net6；Consumer.Net8 | `Program.cs` 的 Workbook/Entity 保护模板导出与真实工作簿读回 |
