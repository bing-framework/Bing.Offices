# Async IO 合同

## API

Excel 提供 `IExcelImporter.ImportAsync`、`IExcelExporter.ExportAsync` 和 `IExcelExporter.ExportToFileAsync`；CSV 提供 `ICsvImporter.ImportAsync`、`ICsvExporter.ExportAsync` 和 `ICsvExporter.ExportToFileAsync`。同步 API 保持不变，异步 API 返回与同步 API 相同的结果、错误分类和映射行为。

## 真异步边界

CSV Parser/Writer、输入复制、输出复制、FileStream flush 和原子文件提交中的内容 IO 使用 `ReadAsync`、`WriteAsync`、`CopyToAsync` 或 `FlushAsync`。`CancellationToken` 会传递到这些 IO 调用；取消继续抛出 `OperationCanceledException`，不会包装成普通导入/导出异常。

NPOI 没有异步的 `WorkbookFactory.Create` 与 Workbook DOM 序列化 API。Excel Async 会在这些阶段保持同步，并只把外围流/文件 IO 异步化；实现不使用 `Task.Run`、`.Result` 或 `.Wait()` 伪装异步。Failure Workbook 的 NPOI 序列化完成后，再通过异步 Stream 复制到调用方目标。

## 流所有权与文件提交

调用方提供的输入和输出 Stream 始终由调用方拥有，API 不会关闭它们。`ExportToFileAsync` 使用同目录临时文件，写入、flush 和取消检查完成后再 replace/move；写委托失败原样传播，真正的文件系统提交失败才分类为 FileCommit。失败或取消不会截断已有目标文件。

## 资源预期

Excel DOM 仍会产生与工作簿规模相关的内存分配；Async 不代表零 GC。大文件部署应同时设置 `ExcelResourceLimits` 和进程级内存/CPU 限额，并使用发布任务的 ResourceProbe 结果评估 staging 峰值。
