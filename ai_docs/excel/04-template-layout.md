# Excel 04：模板与布局

`UseTemplate(stream, leaveOpen)` 从现有 Workbook 加载模板。Sheet Request 可使用 `UseTemplateRegion("Name")`，命名区域的首单元格作为当前 Sheet 写入原点。

模板缺少请求 Sheet、命名区域缺失、命名区域跨 Sheet、地址非法时立即失败。模板中的其它 Sheet 和未覆盖单元格保留。自定义表头使用 `HeaderRows`，重叠、越界和覆盖属性表头的布局在构建阶段失败。
