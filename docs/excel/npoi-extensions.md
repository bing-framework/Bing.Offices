# NPOI 用户扩展

`Bing.Offices.Npoi` 保持以下七个公开扩展容器：`CellExtensions`、`CellStyleExtensions`、`FontExtensions`、`RowExtensions`、`SheetExtensions`、`WorkbookExtensions` 和 `ExcelNpoiServiceCollectionExtensions`。这些是 Provider-specific 用户 API；业务主流程仍建议依赖 provider-neutral 的 `IExcelImporter`、`IExcelExporter`、Workbook Request 和 Mapping API。

Cell、Style、Font、Row、Sheet 和 Workbook 扩展可用于调用方直接操作 NPOI 对象。`TryAddPicture` 会先校验 null、行列范围、图片字节和枚举值，只对合同明确的可恢复 NPOI 参数/状态失败返回 `false`；取消和致命异常不会被吞掉。需要异常详情时使用对应 Throw API。

图片移动和其他依赖具体 Workbook 实现的操作支持 HSSF/XSSF。传入未知 `ISheet` 实现时会抛 `BingOfficesUnsupportedFeatureException`，不会静默无操作。图片资源限制只统计进入映射读取链的图片；未映射图片不应被视为完整工作簿扫描结果。

样式、图片、批注、公式和 Data Validation 的支持范围受 XLS/HSSF 与 XLSX/XSSF 能力差异约束。不支持特性应按请求的 `ExcelUnsupportedFeaturePolicy` 报告或拒绝，不应通过 catch-all 伪装成功。

本轮 package-only 消费者从实际 nupkg 编译并运行这些扩展容器，覆盖 netcoreapp3.1、net6.0 和 net8.0；Npoi 包不提供 netstandard2.0 资产。
