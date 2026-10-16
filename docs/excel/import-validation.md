# 导入校验

校验按以下阶段执行：原始单元格空白规范化、Required/Regex、类型转换、Date/MaxValue/Range/MaxLength、Unique、成功后 Setter。

`ExcelImportValidationMode` 控制配置/属性规则和 Workbook 原生规则的来源：

- `Disabled`：禁用配置/属性规则和 Workbook 原生规则。
- `ConfiguredRules`：只执行配置/属性规则。
- `WorkbookRules`：只执行 Workbook 原生规则。
- `ConfiguredAndWorkbook`：按 Workbook 原生规则在前、配置/属性规则在后的顺序执行两类规则。

在 `ExcelValidationFailureMode.Continue` 下，同一行的 Workbook 原生校验失败后会停止该行后续物化，避免继续转换并产生次生错误；Workbook 校验失败的实体不会进入成功结果，Unique pending journal 也会回滚。配置校验失败但 Workbook 校验通过时，仍按 Continue 规则收集该行可继续发现的配置错误。

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

## Workbook 字典列表与原生下拉

`LookupList` 会把规范化后的文本快照写入隐藏辅助工作表，并为该区域建立工作簿级名称；`DataValidation` 通过 `ListSourceName` 将名称绑定到固定单元格或 Entity Layout 明细区域。下面的 XLSX 示例用采购类别、仓库和采购状态建立模板下拉，随后启用 Workbook 原生规则导入：

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

var categories = new[] { "原材料", "包装材料", "成品" };
var warehouses = new[] { "华东仓", "华南仓" };
var statuses = new[] { "草稿", "待审批", "已批准" };

var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
    .Cell("采购单", "B2", order => order.ProductCategory)
    .Cell("采购单", "B3", order => order.Warehouse)
    .Cell("采购单", "B4", order => order.Status)
    .ListRegion("采购单", "A7", order => order.Lines, region => region.Header())
    .LookupList("ProductCategories", categories)
    .LookupList("Warehouses", warehouses)
    .LookupList("PurchaseStatuses", statuses)
    .DataValidation("采购单", NamedList(1, 1, 1, "ProductCategories"))
    .DataValidation("采购单", NamedList(2, 2, 1, "Warehouses"))
    .DataValidation("采购单", NamedList(3, 3, 1, "PurchaseStatuses"))
    .DataValidation("采购单", NamedList(7, 206, 0, "ProductCategories")));

var order = new PurchaseOrder
{
    ProductCategory = "原材料",
    Warehouse = "华东仓",
    Status = "待审批",
    Lines = new List<PurchaseLine> { new() { ProductCategory = "包装材料" } }
};

using var template = new MemoryStream();
using (IWorkbook workbook = new XSSFWorkbook())
{
    workbook.CreateSheet("采购单");
    workbook.Write(template, true);
}
template.Position = 0;

using var output = new MemoryStream();
new NpoiExcelExporter().ExportForTemplate(order, layout,
    new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

using var input = new MemoryStream(output.ToArray());
var importOptions = new ExcelEntityImportOptions(null,
    ExcelValidationFailureMode.Continue,
    ExcelImportValidationMode.ConfiguredAndWorkbook,
    ExcelUnsupportedFeaturePolicy.Fail);
var result = new NpoiExcelImporter().ImportEntity(input, layout, importOptions);
if (!result.IsSuccess)
    throw new InvalidOperationException(string.Join("; ", result.Errors));

static ExcelDataValidationDefinition NamedList(int row, int endRow, int column, string name) => new()
{
    Range = new ExcelRangeDefinition
    {
        StartRow = row, EndRow = endRow, StartColumn = column, EndColumn = column
    },
    Type = ExcelDataValidationType.List,
    ListSourceName = name,
    IgnoreBlanks = false,
    ShowErrorMessage = true,
    ErrorTitle = "值无效",
    ErrorMessage = "请从下拉列表中选择一个值。"
};

public sealed class PurchaseOrder
{
    public string ProductCategory { get; set; }
    public string Warehouse { get; set; }
    public string Status { get; set; }
    public List<PurchaseLine> Lines { get; set; } = new();
}

public sealed class PurchaseLine
{
    public string ProductCategory { get; set; }
}
```

示例中的实体下拉适用于 NPOI XLS/HSSF、NPOI XLSX/XSSF 和 ClosedXML XLSX；ClosedXML 仅支持 XLSX。普通 Workbook Request 也可通过 `LookupList(...)` 与 Sheet 上的 `DataValidation(... ListSourceName = ...)` 使用相同字典机制。ExcelDataReader 在 XLSX 批量路径读取有限的原生列表、名称列表、数值、日期和文本长度规则，固定列和动态列均保留原始坐标；NPOI XLSX/XSSF 与 ClosedXML XLSX 的 Forward Entity 路径也执行同一有限规则，固定单元格和明细行均在值转换前校验；NPOI HSSF/XLS Forward 因投影不保留原生规则而明确拒绝。完整 Workbook、单 Sheet 分批和 Workbook 分批在 XLSX 输入上支持 `ErrorRowsOnly` 与 `AnnotatedOriginal` 失败工作簿，后者保留源部件并追加错误批注；XLS/XLSB 的原生校验和失败工作簿 AnnotatedOriginal 仍返回 Unsupported。Excel 下拉和 `IgnoreBlanks` 是工作簿交互规则，不能替代服务端业务校验；调用方仍应从应用配置维护类别、仓库权限、状态转换等权威规则，并在导入时选择 `ConfiguredRules` 或 `ConfiguredAndWorkbook`。默认导入不会执行 Workbook 原生规则；显式配置 `WorkbookRules` 时才会执行原生列表校验，非法值会进入结构化导入错误。

后续能力按依赖顺序扩展：先由应用在每次请求时提供当前租户/业务上下文的权威字典和服务端规则；再根据实际需求增加字典版本或模板刷新策略；最后评估更丰富的原生校验类型及各 Provider 的差异。字典选择权和业务授权始终留在应用层。

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

`ErrorRowsOnly` 的低保留 XLSX 输出不会复制完整源工作簿部件，但会保留捕获到的单元格字体、对齐、索引色纯色填充、标准四边框、数字格式和保护状态，并保留源表显式列宽和隐藏列状态，便于回看失败数据。图片、合并、公式、完整样式和其他部件仍不属于该契约。对有界同表区域中的字面量列表、整数/小数/日期/时间/文本长度数值边界规则，会按保留失败行重映射 `sqref` 并保留提示属性；引用已捕获 Sheet 的单行或单列绝对区域、且区域中的行全部保留的定义名称也会同步重映射。公式、外部引用、二维区域、隐藏字典和其他无法安全重映射的名称或规则仍跳过，因为压缩后的 Sheet、行列坐标和公式引用需要完整保真才能保持有效。XLSX 条件格式无法在压缩行集合中安全重映射，遇到该规则时 `ErrorRowsOnly` 在预检阶段返回结构化 `UnsupportedFeature`；需要保留条件格式时使用 `AnnotatedOriginal`。

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
| 条件格式、命名区域、打印设置/分页、页眉页脚、完整 drawing relationship | 保留 | 条件格式存在时预检拒绝，其余仅按受支持子集处理 |

ErrorRowsOnly 对不能完整映射的跨行对象会跳过该对象，不会伪造完整保真。需要完整保留这些对象时，应使用 `AnnotatedOriginal`。失败批注冲突由 `ExcelImportFailureOptions.CommentConflictPolicy` 控制：`Preserve` 保留已有批注，`Append`（默认）追加失败信息，`Replace` 替换已有批注，`Fail` 在冲突时抛出异常。

`Bing.Offices.ExcelDataReader` 的完整 Workbook、单 Sheet 分批和 Workbook 分批入口支持 `ErrorRowsOnly`，在前向读取期间暂存错误行和原始值；分批回调完成后输出包含 `__BingErrors` 摘要及来源表头的最小 XLSX。该输出保留上述受支持的单元格样式子集，并只保留失败来源表数据列范围内的 XLSX `<cols>` 显式列宽和隐藏状态；新加的 `__Errors` 列不继承源列属性。列定义若包含列样式、分组/折叠等其他元数据或无效范围，会在目标写入前返回结构化 `UnsupportedFeature`；范围对象计入 `MaxEstimatedTargetObjects`。该输出不复制图片、合并、公式或其他完整工作簿部件。XLSX 条件格式会在预检阶段返回结构化 `UnsupportedFeature`，需要保留条件格式时使用 `AnnotatedOriginal`。XLSX 输入还支持 `AnnotatedOriginal`：暂存的原始 ZIP 部件保持不变，在原坐标写入批注并追加 `__BingErrors` 摘要；已有批注按 `CommentConflictPolicy` 处理，并在序列化前校验图片资源预算和所有关系目标。该模式只接受普通 XML XLSX；宏部件、外部关系和复杂 VML 会返回结构化导入错误。XLS/XLSB 输入不会伪造源文件保真，返回结构化 `UnsupportedFeature`。`MaxCandidateErrorRows`、`MaxCopiedCells`、`MaxCopiedPictures`、`MaxCopiedPictureBytes`、`MaxEstimatedTargetObjects` 和 `MaxSerializedBytes` 在捕获或序列化前生效，取消或提交失败不会替换已有目标文件。低保留输出的资源异常保留实际 Provider 名称，便于区分 NPOI、ClosedXML 与 ExcelDataReader 的资源边界。
NPOI 与 ClosedXML 的完整 Entity 导入及 `Buffered` Entity 批次也接受 `ExcelEntityImportOptions.FailureOptions`。`AnnotatedOriginal` 在原始实体坐标写入批注并追加 `_ImportErrors`，`ErrorRowsOnly` 生成包含表头和失败行的最小工作簿；同步、异步和模板入口共用同一错误集合、资源限制、提交器和流所有权契约。XLSX `Forward` Entity 批次在解析和回调完成后保留暂存源，因此有错误时支持同样的 `AnnotatedOriginal` 和 `ErrorRowsOnly` 输出；`SummaryOnly + ErrorRowsOnly` 只捕获表头与失败行，并写出保留受支持单元格样式子集的最小 XLSX，安全的一维字面量列表和数值边界校验会按保留失败行重映射，其余名称/公式规则继续跳过；XLSX 条件格式在该模式下于预检阶段返回结构化 `UnsupportedFeature`，需要完整保真时使用 `AnnotatedOriginal`；NPOI HSSF/XLS 在同一组合下从投影阶段捕获快照并构造稀疏 HSSFWorkbook，无法由快照完整表示的固定字段错误回退原有 DOM 路径；其他组合仍按原路径额外打开一次 DOM；成功导入或无错误时不生成失败产物。`SummaryOnly` 的最终 `Result.Entity` 是空聚合对象，不保留固定单元格值、成功明细或关系导航集合；关系声明只使用有界父键跟踪器验证关系并报告缺失父项或资源上限错误，批次回调仍提供读取期间的完整实体，需要聚合结果时使用 `Complete`。NPOI 与 ClosedXML XLSX `Forward` 均支持直接 A1、工作簿级或局部单格命名锚点固定单元格、单格命名锚点列表区域及同一工作表中的多个非重叠列表区域，固定字段转换或校验错误会先以 `RegionIndex = -1` 的诊断批次交付。NPOI 与 ClosedXML XLSX `Forward` 会预读布局声明的 worksheet 合并区域；源文件存在完全相同的合并时允许继续读取，缺失或损坏的合并会在 Plan/Preflight 阶段返回结构化错误，HSSF/XLS Forward 继续拒绝合并；两者还执行有限的工作簿原生 Data Validation，固定单元格和明细行均在值转换前校验并保留源坐标。HSSF Forward 仍不接受工作簿原生校验规则。HSSF Forward 也支持同一工作表的多个非重叠列表、工作簿级和局部单格命名锚点，以及与列表位于同一工作表或其他工作表的直接 A1 固定单元格；模板入口现在接受 HSSF/XLS 与 XLSX 并先验证目标 Sheet、列表边界和命名锚点，HSSF 模板仅投影为用于预检的最小 XLSX 结构，不复制样式/合并；最终 Footer 已支持单列表 marker 截断；完整结果模式下 NPOI 与 ClosedXML Forward 会在所有列表读取完成后执行公共 `HasMany` 关系绑定，关系错误保留可用的源行坐标。HSSF/XLS 原生校验及其他复杂布局仍返回结构化 `UnsupportedFeature`。XLSX Forward 的 NPOI 与 ClosedXML 以及 NPOI HSSF/XLS Forward 均支持独立实际重复页眉，但必须配置 `PageBreak` 和明确 `End`，且不能与分页/分组小计或 Footer 组合；HSSF/XLS 输入先投影为受限 XLSX 行流，不复制复杂模板对象。

HSSF/XLS 的 `SummaryOnly + ErrorRowsOnly` 稀疏失败工作簿会保留投影阶段捕获的显式列宽、隐藏列、行高和隐藏行状态，并复制完全位于保留表头或失败行集合内的合并区域；跨越未保留成功行的合并会被跳过。低保留路径还保留捕获到的字体、对齐、索引色纯色填充、标准四边框、数字格式和保护状态；复杂字体、图案填充、斜线或复杂边框、公式和图片仍不属于该契约，固定字段错误继续按可安全重建性选择稀疏或 DOM 路径。
HSSF/XLS 低保留输出还会重建投影阶段捕获的安全一维绝对定义名称：单格名称以及行或列范围在目标行全部保留时按错误行压缩映射，并保留全局或局部 Sheet 作用域；公式名称、外部引用、隐藏名称、二维区域和未完整保留范围继续跳过。若投影检测到 HSSF 原生 Data Validation，稀疏路径会主动回退已有 HSSF DOM 错误行复制器；复制区域能够完整映射时保留校验及提示属性，无法完整映射时按跨行对象规则跳过，不生成失效规则。HSSF/XLS Forward 的导入校验模式仍不执行 Workbook 原生规则，`WorkbookRules` 和 `ConfiguredAndWorkbook` 继续在预检阶段拒绝。

Forward 同一工作表的多个列表区域必须声明结束地址且互不重叠；导入器按区域边界隔离相同表头，无法确定列归属时在 Plan 阶段返回配置错误，不会将整行表头扫描结果混入其他区域。

NPOI HSSF Forward 还支持按布局顺序读取同一或不同工作表的多个列表区域；每个区域继续保留原始 Sheet、行列坐标和全局批次号，单格命名锚点和 HSSF/XLS、XLSX 模板入口已支持，模板入口只执行布局预检，复杂布局仍返回结构化不支持错误。

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

## 核心实现与依赖配置

规则定义、Required/Regex/Range/Unique 等内置实现位于 `src/Bing.Offices.Core/Bing/Offices/Validations`；绑定与计划创建在同项目 `Mappings`。Workbook 原生规则由 NPOI、ClosedXML、ExcelDataReader 的导入管道各自解析，不能按公共枚举推断所有格式均支持。

`ExcelResourceLimits` 位于 `src/Bing.Offices.Abstractions/Bing/Offices/Imports/Resources`，失败策略和诊断位于 `Imports/Failures`。Provider 注册时提供默认规则；宿主可注册公开 `IExcelValidationRule`/`INamedExcelValidationRule`，并通过映射绑定。错误数截断需同时检查结果的截断标记，不应只按 `Errors.Count` 判断完整性。批次回调与失败产物提交顺序详见[批量导入与失败工作簿](batch-import.md)。
