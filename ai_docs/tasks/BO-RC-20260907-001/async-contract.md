# Async 合同

## 公共签名

- `Task<ExcelWorkbookImportResult<T>> IExcelImporter.ImportAsync<TWorkbook>(Stream, ExcelWorkbookImportRequest<TWorkbook>, CancellationToken)`
- `Task IExcelExporter.ExportAsync(ExcelWorkbookExportRequest, Stream, CancellationToken)`
- `Task IExcelExporter.ExportToFileAsync(ExcelWorkbookExportRequest, string, CancellationToken)`
- `Task<CsvImportResult<T>> ICsvImporter.ImportAsync<T>(Stream, CsvImportOptions<T>, CancellationToken)`
- `Task ICsvExporter.ExportAsync<T>(IEnumerable<T>, Stream, CsvExportOptions<T>, CancellationToken)`
- `Task ICsvExporter.ExportToFileAsync<T>(IEnumerable<T>, string, CsvExportOptions<T>, CancellationToken)`
- `Task IFileExportCommitter.CommitAsync(string, Func<Stream,CancellationToken,Task>, CancellationToken, string)`

## 不变量

- 输入/输出 Stream 由调用方拥有，Async API 不关闭它们。
- 参数错误、取消、致命异常保留原生语义；公共异常 Observer 只接收一次同一实例。
- CSV Parser/Writer、输入复制、输出复制、文件 flush 使用真实 Async API。
- NPOI DOM 创建/操作/序列化保留同步边界；不使用 `Task.Run`、`.Result`、`.Wait()`。
- Excel Export/Failure Workbook 在 NPOI 同步 staging 完成后异步复制到调用方目标。
- Sync/Async 共享 Mapping、Validation、Conversion、错误结果和资源策略。
