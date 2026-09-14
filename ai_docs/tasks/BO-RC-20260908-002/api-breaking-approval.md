# BO-RC-20260908-002 API Breaking Approval

- Task-ID：`BO-RC-20260908-002`
- Fix：`FIX-003`
- 批准人：`jian玄冰`
- 批准时间：`2026-09-09T16:54:21+08:00`
- 批准范围：删除下列四个与接口实例方法重复的公开扩展声明，并以对应接口实例方法作为迁移入口。

## 批准删除的成员

1. `Bing.Offices.Extensions.CsvStreamExtensions.ExportToFile<T>(ICsvExporter exporter, IEnumerable<T> data, string path, CsvExportOptions<T> options = null, CancellationToken cancellationToken = default)`
2. `Bing.Offices.Extensions.CsvStreamExtensions.ExportToFileAsync<T>(ICsvExporter exporter, IEnumerable<T> data, string path, CsvExportOptions<T> options = null, CancellationToken cancellationToken = default)`
3. `Bing.Offices.Extensions.ExcelStreamExtensions.ExportToFile(IExcelExporter exporter, ExcelWorkbookExportRequest request, string path, CancellationToken cancellationToken = default)`
4. `Bing.Offices.Extensions.ExcelStreamExtensions.ExportToFileAsync(IExcelExporter exporter, ExcelWorkbookExportRequest request, string path, CancellationToken cancellationToken = default)`

## 迁移合同

删除的是扩展声明，不是接口实例成员。调用方应使用：

```csharp
exporter.ExportToFile(data, path, options, cancellationToken);
await exporter.ExportToFileAsync(data, path, options, cancellationToken);
exporter.ExportToFile(request, path, cancellationToken);
await exporter.ExportToFileAsync(request, path, cancellationToken);
```

`ICsvExporter.ExportToFile*` 和 `IExcelExporter.ExportToFile*` 的实例成员保持不变。该批准同时授权更新双 TFM API baseline、覆盖清单、消费者调用和迁移文档；不授权修改其他公共成员或生产运行时语义。

候选身份与机器化语义 diff 证据由 CI API snapshot 步骤生成并作为临时 Actions artifact 上传，不作为版本控制文件；后者过滤编译器生成的 `AsyncStateMachineAttribute` 后，在 net6/net8 均未发现新增成员，唯一删除集合对应本节四个批准成员。
