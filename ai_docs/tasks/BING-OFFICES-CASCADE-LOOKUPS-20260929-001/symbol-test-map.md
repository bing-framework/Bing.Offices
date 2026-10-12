# 生产符号到测试方法

| 生产符号或行为 | 测试项目 | 测试方法 |
| --- | --- | --- |
| `ExcelCascadeListDefinition`、`ExcelCascadeListEntry`、两个 `CascadeList` 构建入口与快照 | `Bing.Offices.Npoi.Tests` | `CascadeListExportTest.CascadeModel_RejectsInvalidSourcesAndSnapshotsValues` |
| `ExcelDataValidationDefinition.CascadeSourceName`、`ParentColumnIndex` 及来源互斥 | `Bing.Offices.Npoi.Tests` | `CascadeListExportTest.CascadeModel_RejectsInvalidSourcesAndSnapshotsValues` |
| `ExcelWorkbookExportRequest.CascadeLists`、NPOI XLS/XLSX 隐藏字典、名称、公式、文本 | `Bing.Offices.Npoi.Tests` | `CascadeListExportTest.Export_CascadeList_WritesHiddenMapAndFormula` |
| `ExcelEntityLayout<TEntity>.CascadeLists` 与 Entity 原生下拉 | `Bing.Offices.Npoi.Tests` | `CascadeListExportTest.ExportEntity_CascadeList_WritesValidation` |
| NPOI XLS/XLSX 模板名称冲突及目标流保全 | `Bing.Offices.Npoi.Tests` | `CascadeListExportTest.Export_CascadeNameCollision_LeavesOutputEmpty` |
| ClosedXML Workbook 隐藏映射及填写公式 | `Bing.Offices.ClosedXml.Tests` | `CascadeListExportTest.Export_CascadeList_ShouldWriteHiddenMapAndDependentValidation` |
| ClosedXML Entity 模板、异步和流所有权 | `Bing.Offices.ClosedXml.Tests` | `CascadeListExportTest.ExportEntityAsync_CascadeList_ShouldKeepTemplateAndStreamOpen` |
| ClosedXML 模板名称冲突 | `Bing.Offices.ClosedXml.Tests` | `CascadeListExportTest.Export_CascadeList_TemplateNameCollision_ShouldLeaveOutputEmpty` |
| 级联公式导入的 Unsupported 边界 | `Bing.Offices.Npoi.Tests`、`Bing.Offices.ClosedXml.Tests` | `CascadeListExportTest.Import_CascadeFormula_ReportsUnsupportedWorkbookRule`、`CascadeListExportTest.Import_CascadeFormula_ShouldReportUnsupportedWorkbookRule` |
| MiniExcel 与 SpreadCheetah 不支持能力和写入前拒绝 | `Bing.Offices.MiniExcel.Tests`、`Bing.Offices.SpreadCheetah.Tests` | `CascadeListUnsupportedTest.Export_CascadeList_RejectsWithoutWriting` |
| 第三方公开 API 与包运行 | `Bing.Offices.Consumer.Net6`、`Bing.Offices.Consumer.Net8` | `Program` 的 `cascadeRequest` 验证 |
