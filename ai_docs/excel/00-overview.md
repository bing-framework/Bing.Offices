# Excel 00：范围与契约

本模块以 Workbook Request 作为唯一生产入口。导出使用 `ExcelExport.Workbook(...)`，导入使用 `ExcelImport.Workbook<TWorkbook>(...)`。

生产接口只接受 Workbook Request，旧的单 Sheet Options、`MultiSheet` 和路径兼容包装器已删除。所有输入输出流均由调用方拥有；模板流通过 `UseTemplate(stream, leaveOpen)` 明确所有权。

支持矩阵：NPOI 2.7.4；基础读写支持 Xlsx/Xls；ClosedXML 0.105.1 作为富 XLSX Provider 支持 Workbook/List、模板、样式、合并、公式保存/读回、行高和 Entity Layout（固定 Cell、多个 List Region、Relations）；图表当前只支持 NPOI Xlsx。ClosedXML 不承诺 XLS、Chart、PivotTable、XLSM 宏保留、Entity List Region 动态列或完整公式计算。
