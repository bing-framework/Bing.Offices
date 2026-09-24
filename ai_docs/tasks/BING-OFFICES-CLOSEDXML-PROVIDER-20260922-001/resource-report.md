# Resource and Safety Report

## Verified

- 输入先复制到 `MemoryStream`，并使用默认或请求级 `ExcelResourceLimits` 检查 `MaxInputBytes`。
- 共享 `ExcelXlsxZipPreflight` 在 `XLWorkbook` 创建前检查 ZIP/XML；ClosedXML 显式启用图片数量、单图大小和总图片大小检查，NPOI/MiniExcel 保留原有“不绑定图片不扫描”语义。
- `ClosedXmlRowBudgetPreflight` 在创建 `XLWorkbook` 前读取选定 worksheet XML 的物理行上界；超限返回结构化 `ExcelImportErrorCode.ResourceLimit`，不创建 DOM、不导入部分实体。
- 取消在缓冲、Sheet/Row/Cell 处理、保存和文件提交边界传播。
- ClosedXML 导入在 DOM 创建前响应已取消令牌；直接测试验证取消不会进入 Workbook 加载路径。
- 文件导出沿用公共 `IFileExportCommitter`，取消/失败不覆盖已有目标文件。
- 直接测试验证 `MaxInputBytes`、图片数量/单项大小/总大小会在 `BingOfficesStage.Preflight` 拒绝；`MaxRows` 超限在 DOM 前返回结构化 `ResourceLimit` 错误，Workbook 为空。
- `MaxInputBytes` 直接覆盖 seekable 恰好命中、seekable 超限、non-seekable 超限和取消边界；non-seekable 超限只读取 `MaxInputBytes + 1` 字节，调用方 source 保持打开。
- ClosedXML 本地资源矩阵覆盖 ZIP entries、单 entry 解压大小、总解压大小、压缩比、SharedStrings、Styles、Worksheet、总 Worksheet、XML 字符数、XML 深度、图片数量、单图片大小、总图片大小、Rows/MaxRows、Unique、MaxErrors 和非法 ZIP；均在 `Preflight` 或结构化导入结果边界拒绝。
- ClosedXML 预检现在在 `XLWorkbook` 创建前流式扫描 `workbook.xml` 与 worksheet `<c>` 元素，覆盖 `MaxSheets`、`MaxColumnsPerSheet` 和 `MaxCells` 的 exact-limit/over-limit；不信任 worksheet dimension。
- ClosedXML DOM admission 使用每个 ServiceProvider 共享的 `ClosedXmlProviderOptions` gate，默认单 Workbook、最大排队 64；队列超限在 DOM 创建前返回 `ResourceLimit`，排队取消释放计数且不写目标。
- 观察器直接覆盖同步/异步 Unsupported、ResourceLimit、FileCommit 和 observer 自身抛错；每个主异常最多通知一次，observer 错误写入 `ExceptionObserverFailure`，不覆盖主异常。

## Not run / blocked

Style/SharedString/Unique/Error 的大容量容量矩阵、专门 ZIP bomb 样本、多平台字体和 AutoFit 峰值探针尚未运行，状态为 `BLOCKED_APPROVAL` 或 `BLOCKED_EXTERNAL`；已完成的本地限制矩阵不替代这些发布门禁。
