# Excel 01：导出请求

```csharp
var request = ExcelExport.Workbook(workbook =>
{
    workbook.AddSheet("订单", orders, sheet => sheet
        .HeaderRowIndex(1)
        .DataRowStartIndex(2));
    workbook.AddSheet("客户", customers);
});
exporter.Export(request, destination);
```

`AddSheet` 保存每个 Sheet 的数据、表头行、数据起始行、映射、区域性、样式、模板区域和图表配置。`AddNavigationSheet` 用于父集合导航集合，父集合只被枚举一次。
