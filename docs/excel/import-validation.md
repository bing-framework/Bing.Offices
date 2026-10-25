# Import Validation

校验按以下阶段执行：原始单元格空白规范化、Required/Regex、类型转换、Date/MaxValue/Range/MaxLength、Unique、成功后 Setter。

`ExcelImportValidationMode` 控制配置/属性规则和 Workbook 原生规则的来源：

- `Disabled`：禁用配置/属性规则和 Workbook 原生规则。
- `ConfiguredRules`：只执行配置/属性规则。
- `WorkbookRules`：只执行 Workbook 原生规则。
- `ConfiguredAndWorkbook`：按 Workbook 原生规则在前、配置/属性规则在后的顺序执行两类规则。

在 `ValidateMode.Continue` 下，同一行的 Workbook 原生校验失败后会停止该行后续物化，避免继续转换并产生次生错误；Workbook 校验失败的实体不会进入成功结果，Unique pending journal 也会回滚。配置校验失败但 Workbook 校验通过时，仍按 Continue 规则收集该行可继续发现的配置错误。

`ValidateMode.StopOnFirstFailure` 是默认值，表示遇到当前行的第一个校验错误后停止该行的后续校验；下一行仍会继续处理，不是整个文件的 fail-fast 开关。需要限制总错误收集量时，设置 `ExcelResourceLimits.MaxErrors`；达到限制后，结果的 `ErrorsTruncated` 为 true。CSV 同样跳过失败行并继续后续记录，可用 `CsvImportOptions.MaxErrors` 限制错误收集。

常用默认值与坐标约定：

| 配置 | 默认值 | 调整方式 |
| --- | --- | --- |
| 表头行 / 正文起始行 | 零基 0 / 1 | `HeaderRowIndex` / `DataRowStartIndex` |
| 最大表头列数 | 100 | 宽表显式设置 `MaxReadColumns(实际允许列数)`，超限返回 InvalidHeader |
| Excel 原始输入大小 | 128 MiB，省略资源配置也生效 | `ExcelResourceLimits.MaxInputBytes`；显式 null 关闭，超限抛资源异常（Open 阶段） |
| 必需固定表头 | 全部存在 | `RequireExpectedHeaders(false)` 允许缺列 |
| 表头比较 / 空白处理 | 忽略大小写 / Trim | `HeaderComparison` / `HeaderWhitespace` |
| 正文空白处理 | Trim | `BodyWhitespace` 或列级配置 |
| 空行 | 跳过并继续 | `ReportEmptyRows` / `StopAtFirstEmptyRow` |

Sheet 索引、读取范围和 `SourceRows` 均从 0 开始；错误结果的行号、列号从 1 开始，0 表示无法定位到具体行或列。将 `SourceRows` 展示给用户时应加 1。

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
    [ExcelRegex(@"^ORD-\d+$")]
    [ExcelUnique]
    public string Code { get; set; }

    [ExcelDate("yyyy-MM-dd")]
    public DateTime CreatedAt { get; set; }
}
```

不要把用户输入当作 CLR 类型名或动态反射目标；命名校验器和转换器必须来自已注册集合。

CSV 与 XLSX 共用同一个方向化 Mapping Plan 的列元数据；CSV 只负责记录读写，XLSX 只负责 Workbook/Cell 适配。两者仍分别负责自己的行号和单元格坐标格式化。

## 成功、完成与停止原因

CSV、Workbook 和单表便利入口的结果均提供 `IsSuccess`、`IsComplete`、`StopReason`。只有 `IsSuccess` 为 true 才能按全部成功处理；`IsComplete` 表示没有结构拒绝或限制导致的不完整结果，普通行错误不会将它改为 false。

| 情况 | IsSuccess | IsComplete | StopReason |
| --- | --- | --- | --- |
| 配置范围内全部行有效 | true | true | None |
| 部分行转换/校验失败，后续行已处理 | false | true | None |
| 缺少 Sheet、缺少必需表头等结构拒绝 | false | false | InvalidHeader |
| 结果中存在资源限制错误 | false | false | ResourceLimit |
| 错误收集达到上限 | false | false | ErrorLimit |
| 手工构造的 CSV 结果仅标记截断，无法判定原因 | false | false | Unspecified |

多原因同时存在时按 ResourceLimit、ErrorLimit、InvalidHeader、Unspecified 的顺序归纳，`Errors` 保留详细原因。`StopAtFirstEmptyRow` 或读取范围是调用方主动设置的终点，按该范围处理完属于完整，不表示底层文件一定读到 EOF。发生取消、IO 或自定义扩展异常时仍抛出原有异常，不会转为一个成功结果。

旧属性保持原义：Excel `ErrorsTruncated` 只表示错误收集截断；CSV `IsTruncated` 还涵盖输入、行数等资源限制。不要将它们互相重命名。Excel 恰在一个 Sheet 末行耗尽 MaxRows 且还有请求 Sheet 时，结果现在明确包含 ResourceLimit，已成功导入的实体仍保留。若恰好处理完最后一个请求 Sheet，则不额外报告超限。手工构造旧 Workbook 结果时，原有 `IsSuccess` 仍只检查 Errors，因此还应检查 `IsComplete`，避免人为构造的空错误/截断矛盾状态。

Excel 与 CSV 导入都会检查表头绑定：两个不同表头（正式标题、属性名或别名）若指向同一固定属性或已声明动态 Key，会返回结构化 `InvalidHeader` 错误，避免静默覆盖。CSV 的 `MaxFieldLength` 按解码后的 UTF-16 字符串长度约束每个字段，包括表头和未映射列；解析阶段另有原始大小上限，避免为超长字段读取完整内容。超限返回 `ResourceLimit` 并设置 `IsTruncated`；若解析器提前中止时无法确定列号，错误的 `ColumnIndex` 为 `0`。

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
