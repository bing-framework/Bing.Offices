# Import Validation

校验按以下阶段执行：原始单元格空白规范化、Required/Regex、类型转换、Date/MaxValue/Range/MaxLength、Unique、成功后 Setter。

`ExcelImportValidationMode` 控制配置/属性规则和 Workbook 原生规则的来源：

- `Disabled`：禁用配置/属性规则和 Workbook 原生规则。
- `ConfiguredRules`：只执行配置/属性规则。
- `WorkbookRules`：只执行 Workbook 原生规则。
- `ConfiguredAndWorkbook`：按 Workbook 原生规则在前、配置/属性规则在后的顺序执行两类规则。

在 `ValidateMode.Continue` 下，两个来源的错误都会收集；Workbook 校验失败的实体不会进入成功结果，Unique pending journal 也会回滚。

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
