# Excel 05：样式

样式使用 `ExcelCellStyle`、`ExcelColor` 和边框/对齐枚举，不直接暴露 NPOI 类型。请求级样式按 Workbook 缓存，重复定义复用同一 `ICellStyle` 和 `IFont`。

优先级：动态列样式 > Sheet Header/Body 样式 > Sheet 默认样式。XLSX 支持自定义 RGB；XLS 对不支持的自定义颜色明确抛出 `NotSupportedException`，不静默改变颜色。

数值格式保持 Numeric Cell，通过 DataFormat 显示格式；日期格式同样使用样式，不把数值强制写成文本。
