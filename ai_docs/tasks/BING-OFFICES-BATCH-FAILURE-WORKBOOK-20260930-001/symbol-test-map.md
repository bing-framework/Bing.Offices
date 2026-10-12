# 生产符号到测试方法追溯

| 生产符号 | 关键行为 | 测试项目 | 测试方法 |
| --- | --- | --- | --- |
| `NpoiWorkbookBatchImporter.ImportWorkbookBatches` | XLS/HSSF 与 XLSX/XSSF 按物理工作表顺序交付；错误导入完成后生成 Failure Workbook；输入流保持开放 | `Bing.Offices.Npoi.Tests` | `NpoiWorkbookBatchFailureWorkbookTest.ImportWorkbookBatches_ShouldDeliverPhysicalOrderAndFailureWorkbook` |
| `NpoiWorkbookBatchImporter.ImportWorkbookBatchesAsync` | 异步回调取消停止后续交付，并保持输入流所有权 | `Bing.Offices.Npoi.Tests` | `NpoiWorkbookBatchFailureWorkbookTest.ImportWorkbookBatchesAsync_CallbackCancellation_ShouldPreserveSource` |
| `ClosedXmlWorkbookBatchImporter.ImportWorkbookBatches` | XLSX 按物理工作表顺序交付；错误导入完成后生成 Failure Workbook；输入流保持开放 | `Bing.Offices.ClosedXml.Tests` | `ClosedXmlWorkbookBatchFailureWorkbookTest.ImportWorkbookBatches_ShouldDeliverPhysicalOrderAndFailureWorkbook` |
| `ExcelWorkbookBatchProjection` | 将完整导入的成功实体、原始行号和错误投影为顺序稳定的 Workbook 批次与摘要 | `Bing.Offices.Npoi.Tests`、`Bing.Offices.ClosedXml.Tests` | 两个 `...ShouldDeliverPhysicalOrderAndFailureWorkbook` 方法 |
| NPOI/ClosedXML `IExcelWorkbookBatchImporter` DI 注册与 `MultiSheetBatchImport` 能力声明 | Provider 通过容器解析批次入口，并声明 Workbook 多 Sheet 批次能力 | `Bing.Offices.Npoi.Tests`、`Bing.Offices.ClosedXml.Tests` | 两个 `...ShouldDeliverPhysicalOrderAndFailureWorkbook` 方法中的能力断言 |
| Failure Workbook 无错误分支 | 没有错误时不写出失败工作簿 | `Bing.Offices.Npoi.Tests`、`Bing.Offices.ClosedXml.Tests` | `...WithoutErrors_ShouldNotWriteFailureWorkbook` |
