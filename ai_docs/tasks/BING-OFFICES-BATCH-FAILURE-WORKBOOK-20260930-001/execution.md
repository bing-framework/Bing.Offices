<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-BATCH-FAILURE-WORKBOOK-20260930-001
AI_EXECUTION_FINISHED_AT: 2026-09-30T22:18:00+08:00
AI_EXECUTION_STARTED_AT: 2026-09-30T14:00:50.709Z

# 执行记录

## 状态

已完成。

## Change Impact Analysis

- L0 公共契约：实现已有 `IExcelWorkbookBatchImporter`，不新增或删除公共成员；Provider 能力位和 DI 注册行为改变。
- L1 Provider：NPOI、ClosedXML 增加 DOM 批次桥接；完整导入和 Failure Workbook 写入算法保持复用。
- L2 资源与异常：批次桥接增加一次输入暂存和物理 Sheet 顺序读取，沿用已有输入上限、取消、文件提交和异常观察机制。
- L3 测试与文档：补双 TFM 职责测试、Provider 能力说明、符号追溯和验证证据。

## 已完成

- 已核对现有 Failure Workbook writer、Workbook 批次模型和 ExcelDataReader Unsupported 边界。
- 新增共用 `ExcelWorkbookBatchProjection`，将完整导入结果按物理工作表顺序投影为同步/异步批次，保留实体、原始行号、错误和摘要状态。
- NPOI 新增 XLS/HSSF、XLSX/XSSF Workbook 批次适配器；ClosedXML 新增 XLSX Workbook 批次适配器。
- 两个 Provider 均注册 `IExcelWorkbookBatchImporter` 并声明 `MultiSheetBatchImport`；ExcelDataReader 保持 Failure Workbook Unsupported。
- 批次回调在完整导入和失败产物提交之后执行；回调取消或异常停止后续交付，不回滚已提交失败工作簿。
- 增加 NPOI XLS/XLSX、ClosedXML XLSX 的 Failure Workbook、物理 Sheet 顺序、无错误不输出和异步取消职责测试。
- 更新 Provider 能力矩阵、总览说明和生产符号到测试方法追溯表。

## 验证结果

- `dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-restore -m:1`：net6.0/net8.0 通过，0 警告，0 错误。
- `dotnet build src/Bing.Offices.ClosedXml/Bing.Offices.ClosedXml.csproj -c Release --no-restore -m:1`：net6.0/net8.0 通过，0 警告，0 错误。
- NPOI `NpoiWorkbookBatchFailureWorkbookTest`：net6.0 4/4、net8.0 4/4 通过。
- ClosedXML `ClosedXmlWorkbookBatchFailureWorkbookTest`：net6.0 3/3、net8.0 3/3 通过。
- NPOI `Category!=Large`：net6.0 800/800、net8.0 800/800 通过；ClosedXML：net6.0 250/250、net8.0 250/250 通过。
- `Bing.Offices.sln` Release 构建通过，0 错误；存在仓库既有的 Microsoft.Bcl.Memory net6.0 兼容性警告。
- 全量 `Category!=Large`：Core 204/204（双 TFM）、公共集成 50/50（双 TFM）、文档 19/19、Provider 合同 298/298（双 TFM）、ExcelDataReader 60/60（双 TFM）、MiniExcel 56/56（双 TFM）、NPOI 集成 28/28（双 TFM）、MiniExcel 集成 7/7（双 TFM）、ClosedXML 集成 4/4（双 TFM）、SpreadCheetah 69/69（双 TFM）、Aspose 2/2（双 TFM）、分析器 20/20（双 TFM）通过。
- 公共 API 合同 net8.0 11/11 通过；本任务没有新增、删除或修改公共成员签名。
- `git diff --check` 通过；本任务新增 C# 文件为 UTF-8 BOM + LF，项目文件为 UTF-8 无 BOM + CRLF，Markdown 为 UTF-8 无 BOM + LF。

## 资源与边界

- NPOI/ClosedXML 批次桥接会先暂存输入、执行一次完整 DOM 导入，再读取物理 Sheet 顺序；它保留 Failure Workbook 语义，不是低内存批处理实现。
- 批次回调不提供事务回滚；调用方应按批次摘要和已交付批次自行持久化幂等。
- ExcelDataReader 仍为前向只读 Provider，未伪造 Failure Workbook 能力。

## 追溯

- 成员级映射见 `symbol-test-map.md`。
