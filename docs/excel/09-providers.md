# Excel Provider 能力与边界

## Provider 选择

`Bing.Offices.Abstractions` 和 `Bing.Offices.Core` 提供公共契约及 Mapping Plan。`Bing.Offices.Npoi`、`Bing.Offices.MiniExcel`、`Bing.Offices.ClosedXml` 与只读的 `Bing.Offices.ExcelDataReader` 是核心数据 Provider；`Bing.Offices.SpreadCheetah` 只提供前向流式导出，`Bing.Offices.AsposeCells` 只提供商业渲染、转换和公式等独立服务。应用启动时按职责选择注册入口：

```csharp
services.AddBingOfficesNpoi();
// 或
services.AddBingOfficesMiniExcel();
// 或
services.AddBingOfficesClosedXml();
// 只读导入（不注册 IExcelExporter）
services.AddBingOfficesExcelDataReader();
// 可与完整 Provider 并存的前向流式导出
services.AddBingOfficesSpreadCheetah();
// 可选商业渲染、转换和公式服务
services.AddBingOfficesAsposeCells(options => { /* 许可证与字体配置 */ });
```

同一个 `IServiceCollection` 中的公共服务使用 `TryAdd`，核心数据 Provider 仍按先注册优先解析。SpreadCheetah 注册 `IExcelStreamingExporter`，Aspose.Cells 注册渲染、转换、公式和能力描述接口，不会伪装成完整的 `IExcelImporter`/`IExcelExporter`。

## Provider 实现与依赖

以下版本和包引用以项目文件为准；NPOI、MiniExcel、ClosedXML 的版本属性定义在根目录 `version.dev.props`，其余版本写在各 Provider 项目文件中。

| Provider | 注册入口与公共职责 | 第三方依赖 | 关键实现位置 |
| --- | --- | --- | --- |
| NPOI | `AddBingOfficesNpoi`；完整导入/导出、批次与 Entity 批次导入、公式接口 | NPOI 2.7.4 | `src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj`、`src/Bing.Offices.Npoi/Bing/Offices/Extensions/Extensions.Service.cs`、`src/Bing.Offices.Npoi/Bing/Offices` |
| MiniExcel | `AddBingOfficesMiniExcel`；完整导入/导出 | MiniExcel 1.46.0 | `src/Bing.Offices.MiniExcel/Bing.Offices.MiniExcel.csproj`、`src/Bing.Offices.MiniExcel/Bing/Offices/Extensions/Extensions.Service.cs`、`src/Bing.Offices.MiniExcel/Bing/Offices` |
| ClosedXML | `AddBingOfficesClosedXml`；完整导入/导出与 Entity Layout | ClosedXML 0.105.1 | `src/Bing.Offices.ClosedXml/Bing.Offices.ClosedXml.csproj`、`src/Bing.Offices.ClosedXml/Bing/Offices/Extensions/Extensions.Service.cs`、`src/Bing.Offices.ClosedXml/Bing/Offices` |
| ExcelDataReader | `AddBingOfficesExcelDataReader`；只读导入与批次接口 | ExcelDataReader 3.9.0、System.Text.Encoding.CodePages 6.0.0 | `src/Bing.Offices.ExcelDataReader/Bing.Offices.ExcelDataReader.csproj`、`src/Bing.Offices.ExcelDataReader/Bing/Offices/Extensions/ExcelDataReaderServiceCollectionExtensions.cs`、`src/Bing.Offices.ExcelDataReader/Bing/Offices/ExcelDataReader` |
| SpreadCheetah | `AddBingOfficesSpreadCheetah`；前向流式导出 | SpreadCheetah 1.28.0 | `src/Bing.Offices.SpreadCheetah/Bing.Offices.SpreadCheetah.csproj`、`src/Bing.Offices.SpreadCheetah/Bing/Offices/Extensions/ExcelSpreadCheetahServiceCollectionExtensions.cs`、`src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports` |
| Aspose.Cells | `AddBingOfficesAsposeCells`；渲染、页面图像、格式转换与公式服务 | Aspose.Cells 26.8.0 | `src/Bing.Offices.AsposeCells/Bing.Offices.AsposeCells.csproj`、`src/Bing.Offices.AsposeCells/Bing/Offices/Extensions.cs`、`src/Bing.Offices.AsposeCells/Bing/Offices/AsposeCellsEngine.cs` |

完整 Provider 与 ExcelDataReader、SpreadCheetah 的注册入口会调用 Core 的默认 Mapping Plan 和 CSV 服务注册；Aspose.Cells 项目只引用 Abstractions，不注册 Core 或 CSV 默认实现。ClosedXML 可通过 `ClosedXmlProviderOptions.MaxConcurrentWorkbooks` 和 `MaxQueuedOperations` 限制 DOM 准入，默认分别为 1 和 64；Aspose.Cells 的 `AsposeCellsProviderOptions` 提供 `LicensePath`、`FontDirectories`、`FontSubstitutions` 和 `StrictFonts`。这些设置由宿主在注册扩展中提供，未配置许可证时商业能力不可用。

NPOI 与 ClosedXML 都注册 `IExcelEntityBatchImporter`。NPOI 的完整 Entity Layout 能力由解析得到的 `IExcelImporter`/`IExcelExporter` 对象实现，但 DI 不额外注册 `IExcelEntityImporter`/`IExcelEntityExporter` 别名；可直接使用 Core `Bing.Offices.Extensions` 上的 Entity 扩展。ClosedXML 同时注册这两个别名。

### NPOI XLSX 图表

NPOI 的 XLSX 导出已实现柱状图、折线图和单系列饼图。类别与数值系列通过列 Key、标题或属性名解析，`StartRow`/`EndRow` 可限制使用的数据行，图表锚点使用零基行列坐标。饼图只允许一个系列；列不存在、范围无效或 XLS/HSSF 请求会在导出时明确失败。ClosedXML、MiniExcel 与 SpreadCheetah 当前不支持公共图表请求，不能静默忽略。

### 按报表选择完整工作簿或前向流式导出

Core 的 `ExcelWorkbookExportStrategy` 允许在同一应用中按报表显式选择 `CompleteWorkbook` 或 `ForwardStreaming`。组合根可把两种可用接口传入策略；每次导出仍须明确选择模式：

```csharp
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.SpreadCheetah.Extensions;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using System.Threading;

var services = new ServiceCollection();
services.AddBingOfficesNpoi();
services.AddBingOfficesSpreadCheetah();
using var provider = services.BuildServiceProvider();
var strategy = new ExcelWorkbookExportStrategy(
    provider.GetService<IExcelExporter>(),
    provider.GetService<IExcelStreamingExporter>());

var orders = new[] { new OrderRow { Number = "A-001", Amount = 12m } };
var request = ExcelExport.Workbook(builder => builder.AddSheet("订单", orders));
using var output = new MemoryStream();
await strategy.ExportAsync(request, output, ExcelWorkbookExportMode.ForwardStreaming,
    new ExcelStreamingExportOptions { BatchSize = 500 }, CancellationToken.None);

public sealed class OrderRow
{
    public string Number { get; set; }
    public decimal Amount { get; set; }
}
```

也可将模式改为 `CompleteWorkbook`，此时调用完整 `IExcelExporter`；NPOI 对符合条件的大型纯列表可能在 Provider 内部使用 SXSSF，MiniExcel 的完整工作簿契约也不意味着使用 DOM。`ForwardStreaming` 只调用 `IExcelStreamingExporter`，不接受模板；所选 Provider 未声明请求格式或流式创建能力时，会在写入前拒绝。完整工作簿模式要求完整导出器及其格式/能力声明。两种模式均由调用方选择，不会按数据规模自动切换，也不会在失败后回退到另一种模式。直接写入调用方 Stream 失败时可能已有部分输出；文件入口沿用被选 Provider 的文件提交边界。

### SpreadCheetah 前向流式导出

`Bing.Offices.SpreadCheetah` 只实现 `IExcelStreamingExporter`，不会注册或伪造 `IExcelImporter`、完整 `IExcelExporter`，也不会打开模板。它使用 SpreadCheetah 1.28.0 创建新的 XLSX，按 `ExcelWorkbookExportRequest` 中的数据序列前向写入；`ExcelStreamingExportOptions.BatchSize` 约束批次缓存和取消检查，异步文件入口使用真实异步外围 IO。

| 能力 | SpreadCheetah 首版 | 说明 |
| --- | --- | --- |
| 新建 XLSX、多 Sheet、固定列、动态列 | 已实现 | 只允许前向数据；输出不支持回溯修改 |
| 样式、数字格式、公式文本 | 已实现 | 以公共 Workbook Request 可表达且已验证的子集为准 |
| 表格、筛选、冻结 | 有条件支持 | 仅在前向 API 可直接表达时启用；不静默丢弃 |
| PNG/JPEG 图片、整数/小数/日期/文本长度比较、显式列表校验 | 已实现 | Sheet 级公共定义；图片请求磁盘暂存，输出 96 DPI 像素锚点 |
| 异步写入、取消、串行背压、文件原子提交 | 已实现 | 调用方拥有输入/输出 Stream；文件目标复用公共提交器 |
| 读取、模板编辑、AutoFit、透视表、图表、批注、宏、加密 | 不支持 | 预检拒绝，不自动换用其他 Provider |

### 图片和原生校验写入示例

以下 Sheet 定义适用于 NPOI、ClosedXML 和 SpreadCheetah；MiniExcel 在输出前结构化拒绝。锚点行列为零基，尺寸及偏移为像素。图片字节由调用方提供，不发生网络请求或隐式文件读取。

```csharp
var request = ExcelExport.Workbook(book => book.AddSheet("订单", orders, sheet =>
{
    sheet.Image(new ExcelSheetImageDefinition
    {
        Content = logoBytes, Row = 0, Column = 5,
        Width = 120, Height = 48, OffsetX = 4, OffsetY = 4
    });
    sheet.DataValidation(new ExcelDataValidationDefinition
    {
        Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1000, StartColumn = 2, EndColumn = 2 },
        Type = ExcelDataValidationType.List,
        Values = new[] { "待处理", "已完成" },
        IgnoreBlanks = false,
        InputTitle = "订单状态", InputMessage = "请选择列表中的状态",
        ErrorTitle = "状态无效", ErrorMessage = "仅允许待处理或已完成"
    });
}));
await streamingExporter.ExportBatchesAsync(request, output, cancellationToken: cancellationToken);
```

比较校验支持整数、小数、日期、文本长度及八种比较运算符。日期使用 `Date1`/`Date2`，数字使用 `Value1`/`Value2`；显式列表项不能包含逗号、引号或换行，列表文本总长度（含 Excel 引号）不超过 255。规则写入 Excel，不会替代导入配置校验。复杂公式、名称列表、外部引用及跨 Sheet 列表不在此模型中。

无图片的 SpreadCheetah 请求直接前向输出；含图片请求先写磁盘暂存，再适配图片部件后异步复制输出。批次缓存有界不等于整个引擎恒定内存；图片字节、样式、ZIP 元数据和图片修正资源需要计入容量预算。NPOI XLSX 显式列表采用磁盘暂存和流式 XML 修正，规避锁定版本的引号序列化问题。所有目标 Stream 由调用方拥有，失败后允许部分输出；需要旧文件保护时使用文件扩展入口。

报表预检拒绝无效样式、名称冲突、越界或倒序范围、互斥筛选和打印缩放冲突。同范围 Table 与 AutoFilter 合并；`ShowTotals` 仅显示汇总行，不生成求和公式。NPOI XLSX 与 ClosedXML 提供双色最小/最大色阶（默认红/绿）、数据条（默认蓝）和三交通灯（0/33/67）；NPOI XLS 对不能表达的高级规则提前拒绝。打印区域使用有界 A1 区域，重复标题行使用 `1:2`，重复标题列使用 `A:B`，允许 `$` 绝对引用标记。

表格样式建议使用公共名称，例如 `TableStyleMedium2`；SpreadCheetah 同时保留原有 `Medium2` 简写。SpreadCheetah 的表格必须覆盖从 A1 起的全部映射列，不含汇总行；不为检查配置提前枚举业务数据。

NPOI XLSX/ClosedXML 的 `ShowTotals=true` 在声明的数据范围之后追加空白汇总行，不占用最后一条业务数据，筛选仍覆盖原数据范围；超过工作表末行时预检拒绝。名称范围允许工作簿级与各 Sheet 的局部名称同名，只拒绝相同作用域内重名。

SpreadCheetah 表格区域的结束行须与实际数据行数一致，单次枚举中发现不符时返回结构化写入错误，不能静默扩缩区域；使用文件目标可保全旧文件。NPOI XLS 普通条件格式的颜色沿用 HSSF 有限调色板最近色语义，不承诺任意 RGB 与 XLSX 真彩完全一致。NPOI/ClosedXML 的打印缩放共同支持 10–400%，与适应页数选项互斥。

### Aspose.Cells 商业扩展

本轮开源 Provider 验收不包含 Aspose；以下商业能力描述不代表已完成部署许可证、真实渲染或格式保真验收。

`Bing.Offices.AsposeCells` 是独立的可选包，Core 不传递商业依赖。宿主通过 `AddBingOfficesAsposeCells` 配置许可证文件、允许的字体目录和替代映射；没有有效许可证时所有商业能力在预检阶段拒绝。密码只进入引擎调用，不写入日志、异常消息或 API 快照。

| 能力 | Aspose.Cells 首版 | 说明 |
| --- | --- | --- |
| 工作簿到 PDF | 已实现 | `IExcelDocumentRenderer`；支持 Sheet/页范围请求，结果返回页数和字体警告 |
| 页面 PNG/JPEG/TIFF | 已实现 | `IExcelPageRenderer` 串行回调，不累计全部图片 |
| 公式文本、缓存值、重新计算 | 已实现 | `IExcelFormulaProcessor`；宿主负责许可证和函数/外部链接边界 |
| XLSM 宏保留/剥离 | 有条件支持 | 只保留宏二进制，不执行或编辑 VBA；XLSM 输出需显式选择策略 |
| ODS 与 XLSX/XLSM 转换 | 有条件支持 | 返回可能的格式损失警告，不承诺完全无损 |
| 加密读取/写入 | 有条件支持 | 密码通过 `ExcelWorkbookOpenOptions`/`ExcelWorkbookSaveOptions` 传递 |
| 宏执行、签名修改、VBA 编辑、缺许可证结果 | 不支持 | 预检失败，不产生带水印或受限结果 |

### ClosedXML 导入的名称列表校验

在 `WorkbookRules` 或 `ConfiguredAndWorkbook` 模式下，ClosedXML 可将工作簿已有的名称列表规则解析为允许值。适用于“订单 Sheet 的状态列引用字典 Sheet”这类场景：例如工作簿名称 `AllowedStatus` 指向 `'字典'!$A$1:$A$3`，校验公式为 `=AllowedStatus`。

当前 Sheet 的局部同名定义优先于工作簿定义；无效局部定义不会回退到全局同名定义。名称只支持一个绝对、有界、单行或单列区域（含单单元格），允许跨 Sheet。别名链、动态公式、循环名称、联合区域、整行整列和二维名称区域均报告不支持，不执行公式求值。首版名称引用区域中的允许值也必须是普通值，包含公式单元格时报告不支持，即使其他单元格已经匹配。外部工作簿名称定义会在 ClosedXML 名称解析阶段被拒绝，不能承诺进入此校验流程的 `Report`/`Fail` 策略。

`UnsupportedFeaturePolicy.Report` 保留不支持规则对应的行并记录错误，`Fail` 拒绝该行；已支持规则中的非法值始终属于校验失败，不会因 Report 而放行。空值仍按 `IgnoreBlanks` 处理。此能力读取 Excel 已有规则，不扩展上文仅支持显式列表的公共导出定义，也不提供重新计算。

### 公式读写与计算边界

`IExcelFormulaProcessor` 将公式文本、缓存值和重新计算拆成 `ExcelFormulaReadMode` 与 `ExcelFormulaCalculationMode`。NPOI 和 ClosedXML 当前只声明并测试公式文本/缓存读取；对重新计算、循环引用、外部链接或 UDF 请求返回结构化 `UnsupportedFeature`。Aspose 可在商业 Provider 中执行更完整的计算，但函数支持和许可证应在部署环境中单独验证。

业务层继续注入 `IExcelImporter`、`IExcelExporter`，不应引用 MiniExcel 的原生 Attribute、Workbook 类型或第二套请求模型。两个扩展都通过 `TryAdd` 注册公共服务和 Excel 实现；如果同一个 `IServiceCollection` 同时调用两者，先注册的实现保留。需要切换 Provider 时应在组合根选择注册分支，而不是依赖注册顺序或按请求解析。

## ClosedXML 第一阶段矩阵

ClosedXML 作为富 XLSX Provider 使用同一套 Workbook Request 和 Core Mapping Plan。它适合模板、样式、合并和公式保存场景；DOM 操作仍是单实例串行边界，不是 MiniExcel 的低内存替代品。

| 能力 | ClosedXML 第一阶段 | 说明 |
| --- | --- | --- |
| 常规 XLSX List/Workbook 导入/导出 | 已实现 | 真实 `XLWorkbook` 和 XLSX 读回测试；支持多 Sheet、名称/索引选择 |
| Mapping/Profile/JSON/XML、Converter、ValueMap、Validation、Unique、Relations | 已实现 | 由 Core 生成计划；固定列、动态列和 Entity 多 List Region 关系绑定已接入 |
| 基本字体、填充、对齐、数字格式、列宽、行高 | 已实现 | 行高单位为 point；null 保留模板高度，显式值覆盖 |
| Merge、Formula 保存/读回、模板 Stream/File | 已实现 | 公式契约是保存、读取和缓存值，不是完整 Excel 计算引擎 |
| 原子文件提交、取消、流所有权、输入大小/ZIP 预检 | 已实现 | 使用公共文件提交器和共享 XLSX 预检；DOM 创建前拒绝超限输入 |
| Entity Layout | 已实现 | 支持固定 Cell、属性式 Cell、命名锚点、多个同 Sheet/跨 Sheet List Region、单字典/多分组动态列、导出计算列、连续分组小计、分页小计、水平分页符、Footer、Merge、Relations 和模板布局读写 |
| Failure Workbook、原生 Workbook Validation | 已实现 | 支持 AnnotatedOriginal/ErrorRowsOnly、同步/异步和结构化不支持边界；校验顺序固定为原生规则、值转换、配置规则 |
| Sheet PNG/JPEG 图片、原生数据校验写入 | 已实现 | 字节图片、零基锚点；四类比较和显式列表。复杂 Custom、名称列表及跨 Sheet 列表不在新增写入子集中 |
| Chart、PivotTable、XLSM 宏保留、XLS | 不支持 | 不静默丢弃结构；请求在 `XLWorkbook` 创建前返回 `UnsupportedFeature` |

注册方式：

```csharp
services.AddBingOfficesClosedXml();
```

该扩展与其他 Provider 一样使用 `TryAdd`。需要限制 ClosedXML DOM 准入时可传入 `ClosedXmlProviderOptions`：

```csharp
services.AddBingOfficesClosedXml(options =>
{
    options.MaxConcurrentWorkbooks = 1;
    options.MaxQueuedOperations = 64;
});
```

同时注册多个 Provider 时保留 first-registration-wins；只读 Provider 只注册 `IExcelImporter` 和 `IExcelBatchImporter`，不会伪造 `IExcelExporter`。

## ExcelDataReader 只读 Provider

`Bing.Offices.ExcelDataReader` 使用 ExcelDataReader 3.9.0，面向低内存、前向读取和 XLS/XLSX/XLSB 输入。它复用公共 Mapping Plan、Converter、ValueMap、配置校验和错误模型，支持完整 Workbook 导入以及单 Sheet 的回调式 `IExcelBatchImporter`。

```csharp
services.AddBingOfficesExcelDataReader();

var importer = serviceProvider.GetRequiredService<IExcelImporter>();
var batches = serviceProvider.GetRequiredService<IExcelBatchImporter>();
var workbookBatches = serviceProvider.GetRequiredService<IExcelWorkbookBatchImporter>();
```

完整 Workbook 导入支持固定列；`IExcelBatchImporter` 用于单 Sheet 批次，`IExcelWorkbookBatchImporter` 用于 Workbook 多 Sheet 批次。两种批次接口都按请求定义读取工作表，并支持动态列；Workbook 批次还支持跨 Sheet `HasMany` 关系。XLSX 的单 Sheet 与 Workbook 批次路径可执行有限的原生列表/名称列表、数值、日期和文本长度校验；完整 Workbook 导入不执行 Workbook 原生校验，XLS/XLSB 批次请求这些规则会返回结构化 `UnsupportedFeature`。三个导入入口均支持 `ErrorRowsOnly` 并生成最小 XLSX；该模式只复制受支持的单元格样式子集和满足边界的条件格式，不复制图片、合并、公式或其他完整工作簿部件。条件格式仅在表头和失败行保持源坐标、规则范围完整保留且规则不依赖差异样式时复制；行压缩会在预检阶段拒绝。`AnnotatedOriginal` 仅支持普通 XML XLSX，在原坐标追加错误批注与摘要；宏、外部关系及复杂 VML 等不支持，XLS/XLSB 输入会被拒绝。Entity 布局和图片不支持，Provider 不提供写入或模板操作。批次回调串行等待，默认每批 1,000 行；已交付批次不回滚，调用方负责持久化幂等和事务。公式只读取引擎提供的缓存值，不执行重新计算；输入流由调用方拥有，Provider 会使用临时文件暂存并在成功、失败、取消或回调异常后清理。

| 能力 | ExcelDataReader 首版 | 说明 |
| --- | --- | --- |
| XLS/XLSX/XLSB 固定列导入 | 已实现 | XLSB 只在此 Provider 声明支持；旧 Provider 不自动降级 |
| 完整 Workbook 导入 | 已实现 | 固定列；跨 Sheet `MaxRows` 共享预算；超限返回空 Workbook 和 `ResourceLimit` |
| 单 Sheet / Workbook 多 Sheet 分批 | 已实现 | 同步/异步、顺序交付；动态列可用；Workbook 批次支持跨 Sheet `HasMany`；已交付批次保留 |
| Workbook 原生校验 | 有条件支持 | XLSX 单 Sheet 与 Workbook 批次支持有限列表、名称列表、数值、日期和文本长度规则；完整 Workbook 与 XLS/XLSB 批次不支持 |
| Failure Workbook | 有条件支持 | 三个导入入口支持 `ErrorRowsOnly` 最小 XLSX；`AnnotatedOriginal` 只支持普通 XML XLSX |
| 资源限制 | 有条件支持 | XLSX 的 `MaxCells`/`MaxColumnsPerSheet` 使用 ZIP 物理 `<c>` 预检；XLS/XLSB 设置这两项限制时明确返回 `UnsupportedFeature`，不以 `FieldCount` 估算物理 Cell |
| 写入、模板、图片、Entity 布局 | 不支持 | 预检拒绝，不注册导出器或伪造结果 |

XLSB 行为证据使用 ExcelDataReader 上游仓库 `src/TestData/Issue635.xlsb` 的固定夹具，来源提交为 `dc860de020d1cd40af587c3636c766e2ea829d86`，SHA-256 为 `1D2932ED8424D5278BC7500EB5A3C195FF81BB8403ED4CEC5BFC46AE90F5842F`；测试同时覆盖完整导入、单 Sheet 分批导入、日期/布尔/文本值和输入流所有权。该夹具来源于 [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)，按其开源许可随测试资源使用。

## MiniExcel 第一阶段矩阵

| 能力 | MiniExcel 第一阶段 | 说明 |
| --- | --- | --- |
| 常规 XLSX 导入/导出 | 已实现 | Stream、文件、字节扩展和多 Sheet；使用公共 Workbook Request |
| Mapping/Profile/JSON/XML | 已实现 | 由 Core 编译 Mapping Plan，Provider 只消费计划 |
| Converter、ValueMap、Validation、Unique | 已实现 | 复用公共 converter、validation binding 和 `UniqueTracker` |
| 固定列、别名、忽略列、日期/数字/枚举/bool | 已实现 | 保持现有请求和结果语义 |
| 动态列 | 已实现 | 保持延迟枚举；不得默认把整个输入 `ToList()` |
| Sheet name/index、header、trim、空行、缺列/未知列策略 | 已实现 | 通过现有 `ExcelSheetPolicies` 和 `ExcelImportPolicies` 表达 |
| 基本表头和基本数字格式 | 已实现 | 不承诺 NPOI 的完整样式 DOM |
| 模板导出 | 有条件支持 | 仅按已验证的占位符、循环和多 Sheet 语义使用 |
| 简单宽度、对齐、字体、填充、边框 | 有条件支持 | 不能满足的配置必须显式报告不支持 |
| 复杂样式、复杂多级表头、图片、批注 | 不支持 | 不得静默丢弃；按 Provider 能力预检 |
| XLS | 不支持 | MiniExcel 请求 XLS 必须快速失败，不得偷偷转换为 XLSX |
| Chart、NPOI Failure Workbook、完整图片复制 | 不支持 | 不复制 NPOI 专属 DOM 能力；需要时继续使用 NPOI |

NPOI 继续承担 XLS/XLSX、复杂模板、图片、批注、图表、复杂样式和完整 Workbook DOM 能力。Provider 不支持的请求必须在输出或导入前返回现有结构化 UnsupportedFeature 异常/错误，不能返回一个悄悄缺少内容的成功文件。

## 流、异步和文件边界

输入和输出 Stream 始终由调用方拥有，Provider 不负责关闭调用方流。直接写入调用方 Stream 没有回滚保证；文件 API 继续使用公共 `IFileExportCommitter`，通过同目录暂存、flush、取消检查和最终提交保护已有目标文件。

MiniExcel 的 Async API 必须把真实的输入复制、输出复制和 FileStream flush 保持为异步 IO，并在枚举、行处理、复制和提交边界检查 `CancellationToken`。如果底层 MiniExcel 操作本身是同步的，Async 只描述外围 IO，不得使用 `Task.Run`、`.Result` 或 `.Wait()` 伪装异步。取消或失败必须清理暂存文件，并保持已有目标文件不变。

非定位输入流是否需要暂存、模板流是否要求同步读取、输出流的当前位置和 `leaveOpen` 行为，必须由 Provider 测试固定；不能因为第三方 API 接受 Stream 就推断公共 Stream Contract 已满足。

### MiniExcel 日期和验证边界

MiniExcel 导入在第三方解析前读取 `xl/workbook.xml` 的 `workbookPr/@date1904`，并将该标志传入 Core 日期转换合同；日期单元格的原始 numeric serial 也从 worksheet XML 保留后交给 `ExcelDateParser`，因此 1900 闰年兼容边界、负数线性语义和 1904-01-01 起点由同一规则处理。非 numeric 文本继续按显式日期格式解析。因此 1904 parity 由真实日期 fixture 和双 Provider 结构化合同验证，不把单独的 ZIP 标志读取误称为完整能力。

MiniExcel 的 `ExcelImportValidationMode.ConfiguredRules` 和 `Disabled` 已有明确执行语义：前者执行公共 Mapping/Attribute/Unique 规则，后者跳过这些配置校验但仍执行值转换。MiniExcel 无法读取并执行 NPOI Workbook 原生 Data Validation，因此 `WorkbookRules` 和 `ConfiguredAndWorkbook` 在预检阶段返回结构化 `BingOfficesUnsupportedFeatureException`，不会静默忽略原生规则；需要原生校验时应选择 NPOI Provider。

MiniExcel 历史 100K 受控往返探针写入根 `artifacts/benchmarks/`；该 Provider 的 500K/1M、生产 2C4G 机器和外部 CI 尚未执行，相关结果待实现/待确认。SpreadCheetah 的 100K/500K/1M 流式写出矩阵独立记录于任务 `resource-report.md`，不能据此推断 MiniExcel 容量。所有测试、包、Consumer、资源矩阵及 API 日志统一写入根 `artifacts/`。

## 资源与并存部署

MiniExcel 的目标是降低常规 XLSX 的 Workbook DOM 和峰值内存，但不等于零分配或无限数据规模。仍需应用级内存/CPU 限制，并使用 `ExcelResourceLimits`、最大错误数、最大唯一值和文件提交策略。跨行 Unique 的 Dictionary/HashSet 生命周期应随一次导入结束，不能成为进程级无限增长状态。

NPOI 和 MiniExcel 不应在同一个请求上同时构建 Workbook。发布入口如需并发限制，应在调用 Provider 前共享进程级 `SemaphoreSlim`；Provider 选择和并发门禁属于应用组合根职责。性能与资源结论必须来自真实文件/Stream 测试和 Benchmark，不能以第三方库的宣传吞吐替代证据。

## Forward Entity 增量能力

NPOI HSSF/XLS 的 `SummaryOnly + ErrorRowsOnly` 稀疏失败工作簿会保留投影阶段捕获的显式列宽、隐藏列、行高和隐藏行状态，并复制完全位于保留表头或失败行集合内的合并区域；跨越未保留成功行的合并会被跳过。低保留路径还保留捕获到的字体、对齐、索引色纯色填充、标准四边框、数字格式和保护状态；复杂字体、图案填充、斜线或复杂边框、公式和图片仍不属于该契约。固定单元格错误也会按投影快照写入稀疏工作簿，只有快照不完整或元数据无法安全重建时才回退 DOM。
HSSF/XLS 低保留输出还会重建投影阶段捕获的安全一维绝对定义名称：单格名称以及行或列范围在目标行全部保留时按错误行压缩映射，并保留全局或局部 Sheet 作用域；公式名称、外部引用、隐藏名称、二维区域和未完整保留范围继续跳过。若投影检测到 HSSF 原生 Data Validation，稀疏路径会主动回退已有 HSSF DOM 错误行复制器；复制区域能够完整映射时保留校验及提示属性，无法完整映射时按跨行对象规则跳过，不生成失效规则。HSSF/XLS Forward 的导入校验模式仍不执行 Workbook 原生规则，`WorkbookRules` 和 `ConfiguredAndWorkbook` 继续在预检阶段拒绝。

XLSX `SummaryOnly + ErrorRowsOnly` 的低保留失败工作簿会保留来源表的显式列宽和隐藏列状态，以及捕获到的字体、对齐、索引色纯色填充、标准四边框、数字格式和保护状态；有界同表区域中的字面量列表和数值边界校验会随保留失败行重映射，引用已捕获 Sheet 的单行或单列绝对定义名称在目标行全部保留时也会重映射。单元格/外部/动态公式、二维区域、隐藏字典及其他无法安全重映射的名称或规则继续跳过，避免产生失效的 Sheet、行列和公式引用。条件格式不会被转换为静态样式；当表头和失败行保持源行坐标、每个 `sqref` 范围完整落在已保留行列内且规则不依赖差异样式时，写出器直接复制原始条件格式 XML；样式表没有差异样式定义时，默认 `dxfId="0"` 会先规范化移除。行压缩、范围超出保留区域或规则元数据不完整时，在预检阶段返回结构化 `UnsupportedFeature`；需要完整模板元数据和条件格式时继续使用 `AnnotatedOriginal` 或完整结果路径。

NPOI 的 XLS/HSSF、XLSX/XSSF 和 ClosedXML 的 XLSX Forward 实体批次均支持单列表最终 `Footer`。读取器按 Footer 的 marker 截断明细，跳过 `GapRows` 和尾部内容，并保留成功明细的原始物理行号；缺失、重复、过早 marker 以及显式区域结束地址溢出会在计划阶段返回结构化配置错误。Footer 后的物理行只用于发现重复 marker，不会被物化为明细。NPOI 的 HSSF 路径仍通过受限 XLSX 投影复用相同语义。

同一 Forward 列表还支持 `PageBreak(rowsPerPage).PageSubtotal(...)`。读取器按物理行号校验每个完整页的明细容量，跳过分页小计的 `GapRows` 和完整小计占用行，再继续读取下一页；最终 `Footer` 可以同时配置，末页不要求分页小计。分页小计缺失、错位、重复、位于末页之后、超出显式 `End` 或占用区域与最终尾部冲突时，读取前返回结构化 Plan 配置错误。NPOI 的 XLS/HSSF、XLSX/XSSF 和 ClosedXML 的 XLSX 均复用该边界语义，批次中的源行号仍对应原工作表物理行。

不分页的单列表还支持连续 `GroupSubtotal(...)` 与最终 `Footer` 组合。分组小计 marker 前必须存在至少一行物理明细，读取器跳过分组小计的 `GapRows` 和完整占用行；小计体内重复 marker、marker 位于明细起始处或占用区域超出显式 `End` 时返回 Plan 配置错误。Forward 允许 `GroupSubtotal + PageFooter`，组可以跨逻辑页，读取器按页容量先裁剪分组小计再裁剪每页尾部；`GroupSubtotal + PageSubtotal` 以及没有 `PageFooter` 的 `GroupSubtotal + PageBreak` 仍返回结构化不支持错误。

Forward 的最终 Footer、分页小计与分组小计直测覆盖 NPOI XLS/XLSX、ClosedXML XLSX、同步/异步、两行小计区域、源行号和 marker 边界。NPOI XLSX 与 ClosedXML XLSX Forward 还会在读取前扫描 worksheet 的合并元数据，允许布局声明且源文件确实存在的合并区域；缺失合并或损坏的合并地址返回结构化 Plan/Preflight 错误。该支持只用于布局预检和固定单元格的已完成布局归一，不复制模板样式或未声明合并；NPOI HSSF/XLS 仍明确拒绝合并。NPOI XLSX/XSSF 与 ClosedXML XLSX Forward 还执行有限的工作簿原生数据校验，顺序为原生校验、值转换、配置校验，固定单元格和明细行均保留源坐标；HSSF/XLS 因投影不保留原生规则而明确拒绝。Forward 完整结果模式已在所有受支持的 NPOI/ClosedXML 路径执行公共 `HasMany` 关系绑定，关系错误沿用错误上限和源行坐标；`SummaryOnly` 使用有界父键跟踪器验证关系，不保留子实体或导航集合，缺失关系和关系键资源上限仍返回结构化错误。XLSX Forward 的 `SummaryOnly + ErrorRowsOnly` 会只保留来源表头和失败行，并序列化保留受支持单元格样式子集的最小错误工作簿；条件格式在表头与失败行保持源坐标且规则范围完整落在保留区域、规则不依赖差异样式时会原样复制，行压缩或范围不完整时在预检阶段返回结构化 `UnsupportedFeature`；NPOI HSSF/XLS 在同一组合下会从投影快照构造稀疏 HSSFWorkbook，保留列宽、隐藏列、行高、隐藏行、字体、对齐、索引色纯色填充、标准四边框、数字格式、保护状态及完全位于保留行中的合并区域，无法覆盖的固定字段错误回退原有 DOM；复杂字体、图案填充、斜线或复杂边框、公式、图片和复杂模板对象仍不复制。`Complete`、`AnnotatedOriginal` 和其他组合仍使用原有 DOM 失败输出路径。低保留错误输出的资源异常保留实际 Provider 名称，便于定位 NPOI、ClosedXML 与 ExcelDataReader 的边界。分页与分组小计组合、模板样式/合并复制及其他复杂布局仍在预检阶段拒绝，不会隐式退回 Buffered DOM。

Forward 现支持受限的实际重复页眉导入。列表区域必须同时声明 `PageBreak(rowsPerPage)` 和 `End`，页眉不得与分页小计、分组小计、最终 Footer 或逐页 Footer 组合；读取器按页容量跳过后续页复制的页眉，并保留明细原始物理行号。首页 marker、复制页 marker 和复制页列表表头会在行流中校验，缺失、错位或标题被改动时返回原始工作表坐标。NPOI XLSX/XSSF、NPOI HSSF/XLS 与 ClosedXML XLSX 均支持同步和异步路径；HSSF/XLS 输入先通过受限事件转换投影为 XLSX，页眉只保留单元格值、行号和必要格式，复杂模板对象仍按既有边界拒绝。打印标题行仍属于工作表元数据，不能通过该受限页眉入口替代。实际页眉与 `GroupSubtotal + PageFooter` 仍互斥。

Forward 还支持命名固定单元格和命名列表起点；NPOI XLS/XLSX 与 ClosedXML 的直接测试覆盖缺失、非法名称和同步/异步导入。逐页 `PageHeaderNamed`、`PageFooterNamed` 在 Forward 布局预检中拒绝。Forward 读取器中未发现 `PageSubtotalAnchorPrefix` 的解析或预检拒绝路径，因此无法确认 `PageSubtotalNamed` 的命名锚点行为，待实现/待确认。源码证据位于 `NpoiForwardXlsxEntityBatchImporter.ValidateLayout`、`ClosedXmlForwardXlsxEntityBatchImporter.ValidateLayout` 及对应的 `EntityLayoutBatchImportTest`。

Forward `SummaryOnly` 只保留批次统计、结构化错误和工作表诊断；最终 `Result.Entity` 是新的空聚合对象，不再留存固定单元格值或成功明细。批次回调仍接收读取期间填充了固定字段的实体，因此需要业务对象时应在回调中消费或复制；需要完整聚合结果时继续使用 `Complete`。
