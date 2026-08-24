# Import Validation

校验按以下阶段执行：原始单元格空白规范化、Required/Regex、类型转换、Date/MaxValue/Range/MaxLength、Unique、成功后 Setter。

可用 v2 Attribute：

- `ExcelRequired`
- `ExcelRegex`
- `ExcelDate`
- `ExcelMaxValue`
- `ExcelRange`
- `ExcelMaxLength`
- `ExcelUnique`

Regex 在绑定时使用带 timeout 的缓存。Date 支持格式、区域性、DateTime/DateTimeOffset 和 provider-neutral Excel serial。Range 在构建阶段拒绝无效边界。MaxLength 和 MaxValue 使用独立错误码。

Excel 与 CSV 都使用同一组配置字段。Unique 状态按 Sheet 或 CSV 输入范围隔离，当前行只进入 pending journal；整行成功后提交，转换、校验、Setter 失败时回滚。可通过 `MaxTrackedUniqueValues` 设置资源上限。

示例：

```csharp
public sealed class OrderRow
{
    [ExcelRequired]
    [ExcelRegex(@"^ORD-\\d+$")]
    [ExcelUnique]
    public string Code { get; set; }

    [ExcelDate("yyyy-MM-dd")]
    public DateTime CreatedAt { get; set; }
}
```

不要把用户输入当作 CLR 类型名或动态反射目标；命名校验器和转换器必须来自已注册集合。

CSV 与 XLSX 共用同一个方向化 Mapping Plan 的列元数据；CSV 只负责记录读写，XLSX 只负责 Workbook/Cell 适配。两者仍分别负责自己的行号和单元格坐标格式化。

ASP.NET Core 上传时保持文件流由调用方拥有：

```csharp
public IResult Upload(IFormFile file, IExcelImporter importer)
{
    using var input = file.OpenReadStream();
    var request = ExcelImport.Workbook<UploadWorkbook>(builder =>
        builder.Sheet("Data", workbook => workbook.Rows));
    var result = importer.Import(input, request);
    return Results.Ok(result);
}
```
