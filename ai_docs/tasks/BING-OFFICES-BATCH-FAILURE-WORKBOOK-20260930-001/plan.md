# Workbook 分批导入接入 Failure Workbook

## 目标

为具有 DOM 写入能力的 NPOI 和 ClosedXML Provider 实现已有的
`IExcelWorkbookBatchImporter`，使 Workbook 分批请求可以复用完整导入的动态列、关系、原生校验和
Failure Workbook 语义。ExcelDataReader 继续保持只读边界，不伪造失败工作簿输出。

## 实施范围

1. 新增 Provider 共用的内部批次投影工具：从完整导入结果和请求目标集合提取成功实体、源行和按工作表归属的错误，按 `BatchSize` 生成全局批次号和 `ExcelWorkbookBatchImportSummary`。
2. NPOI 和 ClosedXML 增加 Workbook 批次适配器。适配器先暂存输入并读取物理 Sheet 顺序，再调用已有同步/异步完整 Workbook 导入；完整导入完成后交付批次。
3. 批次请求复用既有 Failure Workbook：失败输出在批次回调前完成，路径使用原子提交器，调用方流不关闭；回调取消或异常不回滚已经提交的失败产物。
4. 两个 Provider 通过 DI 注册 `IExcelWorkbookBatchImporter`，声明 `MultiSheetBatchImport` 和 `SupportsBatchImport`；ExcelDataReader 的能力和 Failure Workbook Unsupported 边界保持不变。
5. 增加 NPOI XLS/XLSX、ClosedXML XLSX 的同步/异步职责测试，覆盖失败工作簿、批次内容、物理 Sheet 顺序、原始错误坐标、预取消、回调取消、调用方流所有权和无错误时不输出。
6. 更新 Provider 文档、任务执行记录和生产符号到测试方法映射；不改版本、不提交、不发布。

## 语义约定

- DOM Provider 的 Workbook 批次不是前向解析：为保证 Failure Workbook 可复制原始工作簿，输入和 DOM 在完整导入期间保留，批次回调在失败产物生成后串行执行。
- 失败工作簿只由完整导入的错误集合驱动；成功批次仍然交付，调用方回调抛错或取消时已生成的失败工作簿不回滚。
- 批次按源文件物理工作表顺序排列；缺失工作表的结构化错误追加在物理工作表之后，并保留请求的工作表名称。
- 现有完整导入错误、资源限制、取消、异常观察、文件提交和流所有权契约保持原样。

## 验收

- NPOI XLS/HSSF、XLSX/XSSF 和 ClosedXML XLSX 均能通过 Workbook 批次接口交付实体及错误。
- `AnnotatedOriginal` 与 `ErrorRowsOnly` 失败工作簿均在有错误时生成；无错误、禁用模式或调用方未配置输出时不调用提交器。
- 同步只调用同步完整导入和提交路径，异步使用 `ImportAsync` 及真实异步回调等待，不使用 `Task.Run`、`.Result` 或 `.Wait()`。
- 双 TFM 受影响项目构建、Provider 职责测试、API 差异、文档检查、`git diff --check` 和编码检查通过。
