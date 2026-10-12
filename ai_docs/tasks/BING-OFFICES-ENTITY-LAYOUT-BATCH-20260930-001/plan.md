# Entity Layout 分批导入

## 目标

为 NPOI（XLS/HSSF、XLSX/XSSF）和 ClosedXML（XLSX）的实体布局增加批次入口。复用现有完整实体导入，保留固定单元格、多个列表区域、动态列、Footer、分页/小计、关系、模板和原生校验语义；完整 DOM 导入完成后按列表区域投影批次，不宣称前向低内存。

## 公共契约

- 在 `Bing.Offices.Entities` 增加 `[EditorBrowsable(Never)]` 的 `IExcelEntityBatchImporter`。
- 增加 `ExcelEntityImportBatch<TEntity>`、`ExcelEntityBatchRegionSummary` 和 `ExcelEntityBatchImportSummary<TEntity>`，提供区域名称、Sheet、Item 类型、批次号、首行、成功项、区域错误、实体级诊断错误及完整实体导入结果。
- 同步/异步支持普通布局和既有模板布局，并保留 `ExcelEntityImportOptions` 重载。
- 新增 `ExcelProviderFeatures.EntityLayoutBatchImport`；仅 NPOI 和 ClosedXML 声明并通过 DI 注册。MiniExcel、SpreadCheetah、ExcelDataReader 不伪造该能力。
- 不修改既有实体导入接口和构造函数；不在本轮加入 Entity Failure Workbook 选项。

## 批次语义

- 完整导入先完成；区域成功项按 `ExcelSheetImportResult.SourceRows` 与目标集合顺序配对。
- 每个列表区域按 `BatchSize`（默认 1000）分批，跨区域使用全局批次号；同一源行上的成功项和错误合并到同一批次。
- 固定单元格、关系及其他无法归属列表区域的错误先通过一个诊断批次交付；诊断批次不包含成功项。
- 空列表不触发空批次；只有实体级错误时仍交付诊断批次。无列表且无错误时不触发回调。
- 同步回调只调用同步完整导入；异步回调等待真实异步完整导入和回调任务。回调取消/异常不回滚已完成的 DOM 导入。
- 调用方输入流由 Provider 保持开放；模板 `LeaveOpen` 继续由既有实体导入处理。

## 实现与测试

1. 增加公共模型和共用内部投影器。
2. 实现两个 Provider 适配器、DI 注册和能力位。
3. 用 NPOI XLS/XLSX、ClosedXML XLSX 覆盖双列表区域、固定字段错误、动态列/关系结果、模板、同步/异步、取消、空列表及流所有权。
4. 更新 Provider 矩阵、任务执行记录和生产符号到测试方法映射；维护 API snapshot 和类型分类。

## 验收

- 双 Provider 双 TFM 构建和职责测试通过，批次内容/行号/错误归属稳定。
- 既有 Entity Layout 测试、Provider 合同、公共 API 合同、`Category!=Large` 和 Release 构建通过。
- 新 C# 文件 UTF-8 BOM + LF，项目文件 UTF-8 无 BOM + CRLF，Markdown UTF-8 无 BOM + LF；`git diff --check` 通过。
- Entity Failure Workbook 继续属于后续扩展，当前通过完整实体导入结果和诊断批次暴露错误。
