# Excel 00：范围与契约

本模块以 Workbook Request 作为唯一生产入口。导出使用 `ExcelExport.Workbook(...)`，导入使用 `ExcelImport.Workbook<TWorkbook>(...)`。

生产接口只接受 Workbook Request，旧的单 Sheet Options、`MultiSheet` 和路径兼容包装器已删除。所有输入输出流均由调用方拥有；模板流通过 `UseTemplate(stream, leaveOpen)` 明确所有权。

支持矩阵：NPOI 2.7.4；基础读写支持 Xlsx/Xls；图表当前只支持 Xlsx。
