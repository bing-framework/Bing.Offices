# Excel 03：导入与关系

```csharp
var request = ExcelImport.Workbook<OrderWorkbook>(workbook =>
{
    workbook.Sheet("订单", root => root.Orders, sheet => sheet
        .HeaderRowIndex(2).DataRowStartIndex(3));
    workbook.Sheet("明细", root => root.Details);
    workbook.HasMany(root => root.Orders, root => root.Details,
        order => order.OrderNo, detail => detail.OrderNo,
        order => order.DetailItems);
});
```

每个 Sheet 独立解析表头和数据行。缺失/隐藏 Sheet、表头错误、值转换错误和关系错误都进入结构化结果；关系绑定使用显式父键和子键，不根据属性名猜测。
