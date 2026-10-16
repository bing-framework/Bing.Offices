# 公式、渲染与格式转换

## 功能与接口

这些能力独立于业务行模型导入导出，不通过 `IExcelExporter` 自动触发。

| 能力 | 接口 | 实际 Provider |
| --- | --- | --- |
| 读取公式文本及已有缓存 | `IExcelFormulaProcessor.Process*` | NPOI（XLS/XLSX）、ClosedXML（XLSX）、Aspose.Cells |
| 重新计算支持的公式 | 同上，显式配置计算模式 | Aspose.Cells；NPOI/ClosedXML 拒绝重算请求 |
| PDF 渲染 | `IExcelDocumentRenderer.Render*` | Aspose.Cells |
| 页面 PNG/JPEG/TIFF | `IExcelPageRenderer.RenderPages*` | Aspose.Cells |
| 格式、密码、宏策略转换 | `IExcelDocumentConverter.Convert*` | Aspose.Cells（XLS/XLSX/XLSM/ODS 范围内） |

开源 Provider 的完整 Excel 公式计算为**待实现/待确认（待实现）**。底层引擎拥有计算 API，不表示 Bing.Offices 将其作为完整兼容引擎公开。

## 公式用法与语义

以下片段使用已注册 NPOI 的 `provider` 和调用方拥有的 `source`；命名空间为 `Bing.Offices`、`Bing.Offices.Formula`、`Microsoft.Extensions.DependencyInjection`。

```csharp
var processor = provider.GetRequiredService<IExcelFormulaProcessor>();
var result = processor.Process(source, ExcelFormat.Xlsx, new ExcelFormulaRequest
{
    ReadMode = ExcelFormulaReadMode.FormulaAndCachedValue,
    CalculationMode = ExcelFormulaCalculationMode.DoNotCalculate,
    SheetName = "订单"
});
```

`ReadMode` 可选 `CachedValue`（默认）、`FormulaText`、`FormulaAndCachedValue`。结果按单元格提供公式、缓存、位置及错误；行列索引为零基。`CalculationCompleted` 表示本次请求已完成，不是“缓存一定新鲜”的证明；NPOI 的不计算路径同样返回完成。

`CalculationMode` 默认为 `DoNotCalculate`。Aspose 支持 `RecalculateSupported` 与更严格的 `MustRecalculate`，后者遇到不支持引用或计算错误会拒绝。`CalculateOnOpen` 在当前三个公式处理器中均拒绝：该接口只返回分析报告，不返回可保存的工作簿。公式导出保存和透视缓存刷新是另外的合同。

## 商业扩展配置与用法

安装 `Bing.Offices.AsposeCells` 后，使用 `Bing.Offices.AsposeCells` 命名空间注册：

```csharp
services.AddBingOfficesAsposeCells(options =>
{
    options.LicensePath = "licenses/Aspose.Cells.lic";
    options.FontDirectories.Add("fonts");
    options.StrictFonts = true;
});
```

`services` 是宿主的 `IServiceCollection`；许可证和字体路径由宿主准备。注册配置被复制，之后修改原集合不会替换已注册快照。此扩展不注册 Excel/CSV 行模型导入导出服务；需要这些能力时另注册数据 Provider。

从容器解析 `IExcelDocumentRenderer` 后调用 `Render(source, destination, new ExcelRenderRequest { InputFormat = ExcelFormat.Xlsx })` 即请求 PDF 输出。请求可配置 Sheet、一基起止页、PDF 兼容级别、区域和字体策略。页面接口通过串行 `pageSink` 返回页号、图片流、扩展名；图片流只在当前回调完成前有效，保留图片需在回调内复制。`Dpi` 默认 150，图片格式默认 PNG。

格式转换使用 `Convert(source, destination, sourceFormat, targetFormat, openOptions, saveOptions)`。`ExcelWorkbookOpenOptions` 和 `ExcelWorkbookSaveOptions` 分别携带密码和宏策略，默认拒绝宏；打开 XLSM 必须显式选择 `Preserve` 或 `Strip`。保留宏只允许输出 XLSM；不要将工作表保护密码当作工作簿加密密码。

## 核心实现与依赖

- `src/Bing.Offices.Abstractions/Bing/Offices/{Formula,Rendering,Conversions}`：公共请求、结果和接口。
- `src/Bing.Offices.Npoi/Bing/Offices/Formula/NpoiExcelFormulaProcessor.cs`：按公式单元格读取文本与缓存，不求值。
- `src/Bing.Offices.ClosedXml/Bing/Offices/Formula/ClosedXmlExcelFormulaProcessor.cs`：XLSX 缓存/文本适配。
- `src/Bing.Offices.AsposeCells/Bing/Offices/{Extensions.cs,AsposeCellsProviderOptions.cs,AsposeCellsFontConfiguration.cs,AsposeCellsEngine.cs}`：许可证检查、字体设置、反射适配引擎、渲染/计算/转换及流输出。

Aspose 项目显式依赖 `Aspose.Cells 26.8.0`；该依赖不进入 Core 的依赖图。NPOI 和 ClosedXML 分别使用自己的引擎包。与数据 Provider 同时注册时，`IExcelFormulaProcessor` 使用 `TryAdd`，先注册的默认实现优先；需要不同公式引擎时应显式绑定所需实现，不能假定后注册会覆盖。

## 已知限制与注意事项

- 有效商业许可证、字体存在性、系统图片编码与格式保真度为**待实现/待确认（待确认）**的部署项；代码中存在配置不能代替目标机器验证。
- `StrictFonts` 检查配置目录；非严格模式可返回已配置替代字体警告，不代表已逐字形检测缺字。`FontSubstitutions` 键为原字体名，值为替代字体名。
- 引擎缺少所请求的渲染/编码/密码等入口时返回结构化 `UnsupportedFeature`。ODS 与其他格式互转不承诺无损，需检查结果警告。
- Async 入口异步复制输入、输出并等待页面回调，引擎计算和渲染仍同步；内部可能使用内存暂存，不承诺恒定内存。
- 调用方拥有输入/输出流；本专题的转换与渲染接口不提供文件原子替换。密码不应进入日志。
