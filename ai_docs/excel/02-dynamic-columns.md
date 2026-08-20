# Excel 02：动态列

动态属性使用 `[DynamicColumn] IDictionary<string, object>`。请求级定义必须提供稳定 `Key` 和展示 `Title`，可选 `Aliases`、`DataType`、`NumberFormat`、`Order`、`Placement`。

- 导出按 `Key` 读取字典值。
- 导入按 `Title/Alias` 匹配表头，再按 `Key` 写回字典。
- `DataType` 决定数字、日期、枚举等转换。
- `Order` 提供稳定排序；`Before/After/PhysicalColumnIndex` 提供确定的物理布局。
- 重复 Key、Title、Alias、固定列标题冲突和未知动态值按请求策略失败。
