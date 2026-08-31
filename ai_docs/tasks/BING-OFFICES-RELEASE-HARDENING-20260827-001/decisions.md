# Review Fix 决策记录

## 任务

- Task ID：`BING-OFFICES-RELEASE-HARDENING-20260827-001`
- 本轮范围：仅处理 Reviewer 标记的 `MUST_FIX`，跳过 FIX-004/FIX-007。
- 版本决策：当前包版本由用户控制并保持为 `2.0.0`。移除 `ExcelSetting.Default` 的 breaking 行为已在 API diff 和迁移文档中明确记录，不在本轮修改版本号。

## 输出提交合同

- Excel/CSV File API 使用目标同目录随机 `CreateNew` staging。
- exporter 返回、目标 staging `Flush(true)` 完成后，以及 `File.Replace`/`File.Move` 前再次检查取消。
- 取消、序列化、写入或提交失败不覆盖既有目标；成功新建和成功替换均以目标可重开为验收条件。
- Excel 与 CSV 共用 Core 内部 `AtomicFileCommitter`，不在扩展方法中各自复制提交逻辑。
- 临时文件清理失败不替换主异常；主异常的 `Exception.Data` 保留清理异常，成功路径无主异常时抛出清理失败。

## Failure Workbook

- `TemporaryDirectory` 为请求级配置；为空使用系统临时目录，并使用随机文件名与独占文件共享。
- 目录创建、临时文件创建、序列化、目标复制、取消、大小超限和删除分别位于不同边界；主异常优先。
- `DiagnosticSink` 存在时接收 `FailureWorkbookTemporaryCleanupFailed` 结构化诊断；sink 自身异常不能覆盖主异常。
- 无 sink 时：失败路径把清理异常放入主异常 `Exception.Data`；无主异常的清理失败直接抛出。诊断不写入工作簿内容或原始字段值。

## 模板 metadata

- 默认策略为 `preserve`：只使用 `UseTemplate` 时保留模板六个 metadata 字段。
- 显式调用 `Metadata(...)` 为 `override`：在实际加载的 XLS/XLSX Workbook 上覆盖 Author、Company、Title、Subject、Category、Description 六个字段。
- `ExcelWorkbookExportRequest` 在构建时复制 options，并用 `MetadataSpecified` 区分未指定与显式默认值。

## 资源与解析边界

- CSV 新增 `MaxInputBytes`、`MaxRows`、`MaxErrors`、`MaxFieldLength`、`MaxColumns`；超限结果带 `CsvImportErrorCode.ResourceLimit`、`IsTruncated` 和 `MaxErrors`。
- Excel `MaxInputBytes` 只限制进入 NPOI 前复制的输入字节；NPOI `WorkbookFactory` 建立 DOM 后的 shared strings、styles、drawings 和解压后峰值不由该属性完全限制。
- 当前不声称可阻止所有 ZIP/OLE 压缩放大；不受信任输入必须配合进程内存/CPU/容器隔离。图片限制只扫描实际绑定图片列。
- 公式导入读取 NPOI 缓存结果并保留公式文本，不执行重算；NPOI WorkbookFactory/Workbook.Write 的不可中断阶段只在边界前后检查取消。
- 本库暂不新增伪异步 API，不使用 `Task.Run`、`.Result` 或 `.Wait()`。
