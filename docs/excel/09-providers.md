 Excel Provider 能力与边界

## Provider 选择

`Bing.Offices.Abstractions` 和 `Bing.Offices.Core` 提供公共契约及 Mapping Plan；`Bing.Offices.Npoi`、`Bing.Offices.MiniExcel`、`Bing.Offices.ClosedXml` 与只读的 `Bing.Offices.ExcelDataReader` 是并列 Provider。应用启动时选择一个 Provider：

```csharp
services.AddBingOfficesNpoi();
// 或
services.AddBingOfficesMiniExcel();
// 或
services.AddBingOfficesClosedXml();
// 只读导入（不注册 IExcelExporter）
services.AddBingOfficesExcelDataReader();
```

### SpreadCheetah 前向流式导出

`Bing.Offices.SpreadCheetah` 只实现 `IExcelStreamingExporter`，不会注册或伪造 `IExcelImporter`、完整 `IExcelExporter`，也不会打开模板。它使用 SpreadCheetah 1.28.0 创建新的 XLSX，按 `ExcelWorkbookExportRequest` 中的数据序列前向写入；`ExcelStreamingExportOptions.BatchSize` 约束批次缓存和取消检查，异步文件入口使用真实异步外围 IO。

| 能力 | SpreadCheetah 首版 | 说明 |
| --- | --- | --- |
| 新建 XLSX、多 Sheet、固定列、动态列 | Supported | 只允许前向数据；输出不支持回溯修改 |
| 样式、数字格式、公式文本 | Supported | 以公共 Workbook Request 可表达且已验证的子集为准 |
| 表格、筛选、冻结 | Conditional | 仅在前向 API 可直接表达时启用；不静默丢弃 |
| PNG/JPEG 图片、整数/小数/日期/文本长度比较、显式列表校验 | Supported | Sheet 级公共定义；图片请求磁盘暂存，输出 96 DPI 像素锚点 |
| 异步写入、取消、串行背压、文件原子提交 | Supported | 调用方拥有输入/输出 Stream；文件目标复用公共提交器 |
| 读取、模板编辑、AutoFit、透视表、图表、批注、宏、加密 | Unsupported | 预检拒绝，不自动换用其他 Provider |

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
| 工作簿到 PDF | Supported | `IExcelDocumentRenderer`；支持 Sheet/页范围请求，结果返回页数和字体 warning |
| 页面 PNG/JPEG/TIFF | Supported | `IExcelPageRenderer` 串行回调，不累计全部图片 |
| 公式文本、缓存值、重新计算 | Supported | `IExcelFormulaProcessor`；宿主负责许可证和函数/外部链接边界 |
| XLSM 宏保留/剥离 | Conditional | 只保留宏二进制，不执行或编辑 VBA；XLSM 输出需显式选择策略 |
| ODS 与 XLSX/XLSM 转换 | Conditional | 返回可能的格式损失 warning，不承诺完全无损 |
| 加密读取/写入 | Conditional | 密码通过 `ExcelWorkbookOpenOptions`/`ExcelWorkbookSaveOptions` 传递 |
| 宏执行、签名修改、VBA 编辑、缺许可证结果 | Unsupported | 预检失败，不产生带水印或受限结果 |

### ClosedXML 导入的名称列表校验

在 `WorkbookRules` 或 `ConfiguredAndWorkbook` 模式下，ClosedXML 可将工作簿已有的名称列表规则解析为允许值。适用于“订单 Sheet 的状态列引用字典 Sheet”这类场景：例如工作簿名称 `AllowedStatus` 指向 `'字典'!$A$1:$A$3`，校验公式为 `=AllowedStatus`。

当前 Sheet 的局部同名定义优先于工作簿定义；无效局部定义不会回退到全局同名定义。名称只支持一个绝对、有界、单行或单列区域（含单单元格），允许跨 Sheet。别名链、动态公式、循环名称、联合区域、整行整列和二维名称区域均返回 Unsupported，不执行公式求值。首版名称引用区域中的允许值也必须是普通值，包含公式单元格时返回 Unsupported，即使其他单元格已经匹配。外部工作簿名称定义会在 ClosedXML 名称解析阶段被拒绝，不能承诺进入此校验流程的 Report/Fail 策略。

`UnsupportedFeaturePolicy.Report` 保留不支持规则对应的行并记录错误，`Fail` 拒绝该行；已支持规则中的非法值始终属于校验失败，不会因 Report 而放行。空值仍按 `IgnoreBlanks` 处理。此能力读取 Excel 已有规则，不扩展上文仅支持显式列表的公共导出定义，也不提供重新计算。

### 公式读写与计算边界

`IExcelFormulaProcessor` 将公式文本、缓存值和重新计算拆成 `ExcelFormulaReadMode` 与 `ExcelFormulaCalculationMode`。NPOI 和 ClosedXML 当前只声明并测试公式文本/缓存读取；对重新计算、循环引用、外部链接或 UDF 请求返回结构化 `UnsupportedFeature`。Aspose 可在商业 Provider 中执行更完整的计算，但函数支持和许可证应在部署环境中单独验证。

业务层继续注入 `IExcelImporter`、`IExcelExporter`，不应引用 MiniExcel 的原生 Attribute、Workbook 类型或第二套请求模型。两个扩展都通过 `TryAdd` 注册公共服务和 Excel 实现；如果同一个 `IServiceCollection` 同时调用两者，先注册的实现保留。需要切换 Provider 时应在组合根选择注册分支，而不是依赖注册顺序或按请求解析。

## ClosedXML 第一阶段矩阵

ClosedXML 作为富 XLSX Provider 使用同一套 Workbook Request 和 Core Mapping Plan。它适合模板、样式、合并和公式保存场景；DOM 操作仍是单实例串行边界，不是 MiniExcel 的低内存替代品。

| 能力 | ClosedXML 第一阶段 | 说明 |
| --- | --- | --- |
| 常规 XLSX List/Workbook 导入/导出 | P0 | 真实 `XLWorkbook` 和 XLSX 读回测试；支持多 Sheet、名称/索引选择 |
| Mapping/Profile/JSON/XML、Converter、ValueMap、Validation、Unique、Relations | P0/P1 | 由 Core 生成计划；固定列、动态列和 Entity 多 List Region 关系绑定已接入 |
| 基本字体、填充、对齐、数字格式、列宽、行高 | P1 | 行高单位为 point；null 保留模板高度，显式值覆盖 |
| Merge、Formula 保存/读回、模板 Stream/File | P1 | 公式契约是保存、读取和缓存值，不是完整 Excel 计算引擎 |
| 原子文件提交、取消、流所有权、输入大小/ZIP 预检 | P1 | 使用公共文件提交器和共享 XLSX 预检；DOM 创建前拒绝超限输入 |
| Entity Layout | P1 | 支持固定 Cell、多个同 Sheet/跨 Sheet List Region、Merge、Relations 和模板布局读写；Entity List Region 动态列仍显式不支持 |
| Failure Workbook、原生 Workbook Validation | P1 | 支持 AnnotatedOriginal/ErrorRowsOnly、同步/异步和结构化 Unsupported 边界；校验顺序固定为原生规则、值转换、配置规则 |
| Sheet PNG/JPEG 图片、原生数据校验写入 | P1 | 字节图片、零基锚点；四类比较和显式列表。复杂 Custom、名称列表及跨 Sheet 列表不在新增写入子集中 |
| Chart、PivotTable、XLSM 宏保留、XLS | Unsupported | 不静默丢弃结构；请求在 `XLWorkbook` 创建前返回 UnsupportedFeature |

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
```

批量导入的批次回调串行等待，默认每批 1,000 行；已交付批次不回滚，调用方负责持久化幂等和事务。请求必须明确 Sheet 名称，跨 Sheet 关系、动态列、Entity 布局、Failure Workbook 和 Workbook 原生校验会在预检阶段返回结构化 `UnsupportedFeature`。公式只读取引擎提供的缓存值，不执行重新计算；输入流由调用方拥有，Provider 会使用临时文件暂存并在成功、失败、取消或回调异常后清理。

| 能力 | ExcelDataReader 首版 | 说明 |
| --- | --- | --- |
| XLS/XLSX/XLSB 固定列导入 | Supported | Xlsb 只在此 Provider 声明支持；旧 Provider 不自动降级 |
| 完整 Workbook 导入 | Supported | 跨 Sheet `MaxRows` 共享预算；超限返回空 Workbook 和 ResourceLimit |
| 分批导入 | Supported | 单 Sheet、同步/异步回调、批次顺序稳定、已交付批次保留 |
| 资源限制 | Conditional | XLSX 的 `MaxCells`/`MaxColumnsPerSheet` 使用 ZIP 物理 `<c>` 预检；XLS/XLSB 设置这两项限制时明确返回 `UnsupportedFeature`，不以 `FieldCount` 估算物理 Cell |
| 写入、模板、图片、动态列、关系 | Unsupported | 预检拒绝，不注册导出器或伪造结果 |
| Workbook 原生校验、Failure Workbook | Unsupported | 使用 NPOI 或 ClosedXML Provider |

XLSB 行为证据使用 ExcelDataReader 上游仓库 `src/TestData/Issue635.xlsb` 的固定夹具，来源提交为 `dc860de020d1cd40af587c3636c766e2ea829d86`，SHA-256 为 `1D2932ED8424D5278BC7500EB5A3C195FF81BB8403ED4CEC5BFC46AE90F5842F`；测试同时覆盖完整导入、单 Sheet 分批导入、日期/布尔/文本值和输入流所有权。该夹具来源于 [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)，按其开源许可随测试资源使用。

## MiniExcel 第一阶段矩阵

| 能力 | MiniExcel 第一阶段 | 说明 |
| --- | --- | --- |
| 常规 XLSX 导入/导出 | P0 | Stream、文件、字节扩展和多 Sheet；使用公共 Workbook Request |
| Mapping/Profile/JSON/XML | P0 | 由 Core 编译 Mapping Plan，Provider 只消费计划 |
| Converter、ValueMap、Validation、Unique | P0 | 复用公共 converter、validation binding 和 `UniqueTracker` |
| 固定列、别名、忽略列、日期/数字/枚举/bool | P0 | 保持现有请求和结果语义 |
| 动态列 | P0 | 保持延迟枚举；不得默认把整个输入 `ToList()` |
| Sheet name/index、header、trim、空行、缺列/未知列策略 | P0 | 通过现有 `ExcelSheetPolicies` 和 `ExcelImportPolicies` 表达 |
| 基本表头和基本数字格式 | P0 | 不承诺 NPOI 的完整样式 DOM |
| 模板导出 | P1 | 只有实际验证通过的占位符、循环和多 Sheet 语义才可标记支持 |
| 简单宽度、对齐、字体、填充、边框 | P1 | 不能满足的配置必须显式报告不支持 |
| 复杂样式、复杂多级表头、图片、批注 | P2/Unsupported | 不得静默丢弃；按 Provider 能力预检 |
| XLS | Unsupported | MiniExcel 请求 XLS 必须快速失败，不得偷偷转换为 XLSX |
| Chart、NPOI Failure Workbook、完整图片复制 | Unsupported/P2 | 第一阶段不复制 NPOI 专属 DOM 能力；需要时继续使用 NPOI |

NPOI 继续承担 XLS/XLSX、复杂模板、图片、批注、图表、复杂样式和完整 Workbook DOM 能力。Provider 不支持的请求必须在输出或导入前返回现有结构化 UnsupportedFeature 异常/错误，不能返回一个悄悄缺少内容的成功文件。

## 流、异步和文件边界

输入和输出 Stream 始终由调用方拥有，Provider 不负责关闭调用方流。直接写入调用方 Stream 没有回滚保证；文件 API 继续使用公共 `IFileExportCommitter`，通过同目录暂存、flush、取消检查和最终提交保护已有目标文件。

MiniExcel 的 Async API 必须把真实的输入复制、输出复制和 FileStream flush 保持为异步 IO，并在枚举、行处理、复制和提交边界检查 `CancellationToken`。如果底层 MiniExcel 操作本身是同步的，Async 只描述外围 IO，不得使用 `Task.Run`、`.Result` 或 `.Wait()` 伪装异步。取消或失败必须清理暂存文件，并保持已有目标文件不变。

非定位输入流是否需要暂存、模板流是否要求同步读取、输出流的当前位置和 `leaveOpen` 行为，必须由 Provider 测试固定；不能因为第三方 API 接受 Stream 就推断公共 Stream Contract 已满足。

### MiniExcel 日期和验证边界

MiniExcel 导入在第三方解析前读取 `xl/workbook.xml` 的 `workbookPr/@date1904`，并将该标志传入 Core 日期转换合同；日期单元格的原始 numeric serial 也从 worksheet XML 保留后交给 `ExcelDateParser`，因此 1900 闰年兼容边界、负数线性语义和 1904-01-01 起点由同一规则处理。非 numeric 文本继续按显式日期格式解析。因此 1904 parity 由真实日期 fixture 和双 Provider 结构化合同验证，不把单独的 ZIP 标志读取误称为完整能力。

MiniExcel 的 `ExcelImportValidationMode.ConfiguredRules` 和 `Disabled` 已有明确执行语义：前者执行公共 Mapping/Attribute/Unique 规则，后者跳过这些配置校验但仍执行值转换。MiniExcel 无法读取并执行 NPOI Workbook 原生 Data Validation，因此 `WorkbookRules` 和 `ConfiguredAndWorkbook` 在预检阶段返回结构化 `BingOfficesUnsupportedFeatureException`，不会静默忽略原生规则；需要原生校验时应选择 NPOI Provider。

MiniExcel 历史 100K 受控往返探针写入根 `artifacts/benchmarks/`；该 Provider 的 500K/1M、生产 2C4G 机器和外部 CI 未执行时仍为 `NOT_VERIFIED`。SpreadCheetah 本轮单独运行 100K/500K/1M 流式写出矩阵，见任务 `resource-report.md`，不能据此推断 MiniExcel 容量。所有测试、包、Consumer、资源矩阵及 API 日志统一写入根 `artifacts/`。

## 资源与并存部署

MiniExcel 的目标是降低常规 XLSX 的 Workbook DOM 和峰值内存，但不等于零分配或无限数据规模。仍需应用级内存/CPU 限制，并使用 `ExcelResourceLimits`、最大错误数、最大唯一值和文件提交策略。跨行 Unique 的 Dictionary/HashSet 生命周期应随一次导入结束，不能成为进程级无限增长状态。

NPOI 和 MiniExcel 不应在同一个请求上同时构建 Workbook。发布入口如需并发限制，应在调用 Provider 前共享进程级 `SemaphoreSlim`；Provider 选择和并发门禁属于应用组合根职责。性能与资源结论必须来自真实文件/Stream 测试和 Benchmark，不能以第三方库的宣传吞吐替代证据。
