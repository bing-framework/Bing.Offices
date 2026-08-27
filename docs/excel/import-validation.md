# Import Validation

校验按以下阶段执行：原始单元格空白规范化、Required/Regex、类型转换、Date/MaxValue/Range/MaxLength、Unique、成功后 Setter。

`ExcelImportValidationMode` 控制配置/属性规则和 Workbook 原生规则的来源：

- `Disabled`：禁用配置/属性规则和 Workbook 原生规则。
- `ConfiguredRules`：只执行配置/属性规则。
- `WorkbookRules`：只执行 Workbook 原生规则。
- `ConfiguredAndWorkbook`：按 Workbook 原生规则在前、配置/属性规则在后的顺序执行两类规则。

在 `ValidateMode.Continue` 下，同一行的 Workbook 原生校验失败后会停止该行后续物化，避免继续转换并产生次生错误；Workbook 校验失败的实体不会进入成功结果，Unique pending journal 也会回滚。配置校验失败但 Workbook 校验通过时，仍按 Continue 规则收集该行可继续发现的配置错误。

可用 v2 Attribute：

- `ExcelRequired`
- `ExcelRegex`
- `ExcelDate`
- `ExcelMaxValue`
- `ExcelRange`
- `ExcelMaxLength`
- `ExcelUnique`

Regex 在绑定时使用带 1 秒 timeout 的进程级缓存，最多保留 256 个不同 pattern；超出容量时按 FIFO 淘汰最早加入的 pattern。缓存命中、淘汰和并发访问不改变校验结果；恶意 pattern 仍受 Regex timeout 限制并按规则异常契约处理。Date 支持格式、区域性、DateTime/DateTimeOffset 和 provider-neutral Excel serial。Range 在构建阶段拒绝无效边界。MaxLength 和 MaxValue 使用独立错误码。

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

## 失败工作簿能力边界

失败工作簿输出目前支持以下对象：

| 对象 | AnnotatedOriginal | ErrorRowsOnly |
| --- | --- | --- |
| 单元格值、公式、错误值、布尔值 | 保留 | 复制 |
| 单元格样式、字体、数字格式 | 保留 | 复制 |
| 行高、隐藏、列宽、列隐藏 | 保留 | 复制 |
| 合并区域 | 保留 | 仅复制完全落在保留行集合中的区域 |
| 数据验证及提示/错误显示属性 | 保留 | 仅复制完全落在保留行集合中的区域 |
| 图片、超链接、批注 | 保留 | 复制保留行中可映射的对象 |
| 富文本格式运行 | 保留 | 复制 |
| 冻结窗格和基础显示设置 | 保留 | 复制并按连续输出行重新映射冻结行 |
| 条件格式、命名区域、打印设置/分页、页眉页脚、完整 drawing relationship | 保留 | 当前不承诺复制 |

ErrorRowsOnly 对不能完整映射的跨行对象会跳过该对象，不会伪造完整保真。需要完整保留这些对象时，应使用 `AnnotatedOriginal`。失败批注冲突由 `ExcelImportFailureOptions.CommentConflictPolicy` 控制：`Preserve` 保留已有批注，`Append`（默认）追加失败信息，`Replace` 替换已有批注，`Fail` 在冲突时抛出异常。

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
