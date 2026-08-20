# Excel 06：图表

图表模型由 `ExcelChartDefinition`、`ExcelChartSeries`、`ExcelChartRange` 和 `ExcelChartAnchor` 组成，范围引用当前 Sheet 的列 Key。

NPOI 适配器当前支持 XLSX 的 `Column`、`Line` 和单系列 `Pie`。图表在数据写入后创建，默认范围使用当前数据行；显式范围可以指定起止行。XLS 或无效列引用按能力边界显式失败。
