# 第三方 Excel 引擎选型

下表描述底层组件的能力，不代表 Bing.Offices 已经暴露对应公共合同。新增 Provider 必须先声明“读取、写入、复杂工作簿、计算/渲染”职责，再补职责级合同和双 TFM 包消费者测试；不要求每个引擎实现全部职责。

| 组件／常用包 | 主要职责 | 授权与边界 | 在 Bing.Offices 中的定位 |
| --- | --- | --- | --- |
| [NPOI](https://www.nuget.org/packages/NPOI/2.7.4) | XLS/XLSX 读写、复杂 Workbook DOM | 当前锁定 2.7.4，Apache-2.0；后续版本授权需重新核实 | 已接入，复杂工作簿和 XLS 基准 |
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) | 高层 XLSX/XLSM 操作 | MIT；DOM 模型需要关注内存 | 已接入，富报表和模板 |
| [MiniExcel](https://github.com/mini-software/MiniExcel) | 轻量数据读写、模板 | 当前锁定 1.46.0，Apache-2.0；侧重数据处理 | 已接入，常规大数据读取/写入 |
| [ExcelDataReader](https://www.nuget.org/packages/ExcelDataReader/3.9.0) | XLS/XLSX/XLSB 前向读取 | MIT；只读 | 已接入，固定/动态列、单表/多表批次、关系与失败工作簿，边界见 Provider 矩阵 |
| [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [SpreadCheetah](https://github.com/sveinungf/spreadcheetah) | 流式 XLSX 生成 | MIT；生成型组件 | 已接入 `Bing.Offices.SpreadCheetah`，只创建新工作簿 |
| [LargeXlsx](https://github.com/salvois/LargeXlsx) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [EPPlus](https://www.epplussoftware.com/en/Home/GetEPPlus) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [Aspose.Cells](https://docs.aspose.com/cells/net/) | 多格式处理、公式计算、Excel 到 PDF/图片 | 商业授权；许可证由宿主配置 | 已接入独立 `Bing.Offices.AsposeCells`，不进入 Core |
| [Syncfusion XlsIO](https://www.syncfusion.com/document-processing/excel-framework/net) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [GemBox.Spreadsheet](https://www.gemboxsoftware.com/spreadsheet) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [DevExpress Spreadsheet Document API](https://docs.devexpress.com/OfficeFileAPI/14912/spreadsheet-document-api) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [Telerik SpreadProcessing](https://www.telerik.com/document-processing-libraries/documentation/libraries/radspreadstreamprocessing/overview) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [MESCIUS DS.Documents.Excel](https://developer.mescius.com/document-solutions/dot-net-excel-api/docs/online/getting-started) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |
| [Spire.XLS](https://cdn.e-iceblue.com/Introduce/free-xls-component.html) | 待确认 | 待实现/待确认：未接入；授权以选定版本官方许可为准 | 未接入 |

## 已接入实现与依赖

下列包版本以项目文件为准；其公共接口和能力边界见[Provider 能力与边界](09-providers.md)。

| Provider | 实际包版本 | 注册入口 | 关键源码路径 |
| --- | --- | --- | --- |
| NPOI | NPOI 2.7.4 | `AddBingOfficesNpoi` | `src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj`、`src/Bing.Offices.Npoi/Bing/Offices` |
| ClosedXML | ClosedXML 0.105.1 | `AddBingOfficesClosedXml` | `src/Bing.Offices.ClosedXml/Bing.Offices.ClosedXml.csproj`、`src/Bing.Offices.ClosedXml/Bing/Offices` |
| MiniExcel | MiniExcel 1.46.0 | `AddBingOfficesMiniExcel` | `src/Bing.Offices.MiniExcel/Bing.Offices.MiniExcel.csproj`、`src/Bing.Offices.MiniExcel/Bing/Offices` |
| ExcelDataReader | ExcelDataReader 3.9.0 | `AddBingOfficesExcelDataReader` | `src/Bing.Offices.ExcelDataReader/Bing.Offices.ExcelDataReader.csproj`、`src/Bing.Offices.ExcelDataReader/Bing/Offices/ExcelDataReader` |
| SpreadCheetah | SpreadCheetah 1.28.0 | `AddBingOfficesSpreadCheetah` | `src/Bing.Offices.SpreadCheetah/Bing.Offices.SpreadCheetah.csproj`、`src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports` |
| Aspose.Cells | Aspose.Cells 26.8.0 | `AddBingOfficesAsposeCells` | `src/Bing.Offices.AsposeCells/Bing.Offices.AsposeCells.csproj`、`src/Bing.Offices.AsposeCells/Bing/Offices` |

项目注册代码位于各 Provider 的 `Bing/Offices/Extensions` 子目录。ClosedXML 的准入配置为 `ClosedXmlProviderOptions`；Aspose.Cells 的许可证与字体配置为 `AsposeCellsProviderOptions`。尚未接入的组件没有固定项目版本或宿主配置，其能力与授权均为“待实现/待确认”；接入前应以选定版本的官方文档和许可为准。

## Provider 设计边界

- Core 只依赖 Abstractions，不依赖商业引擎；商业 Provider 由宿主单独引用并配置授权。
- 数据读取、数据写出、复杂工作簿、计算和渲染是独立职责。只读 Provider 可以只注册已声明的 `IExcelImporter`、`IExcelBatchImporter` 或 Workbook 批次接口。
- 公式文本、缓存值和重新计算必须分别声明；Excel 到 PDF/图片应采用独立渲染接口，不把导出接口伪装成渲染接口。
- XLSM、ODS、加密文件、宏保留、工作表保护和文件加密分别定义合同；宏保留不等于执行，工作表保护不等于文件加密。
- 接入新包时固定包版本、许可证、支持格式、内存模型、同步/异步边界和 Unsupported 行为，并补充真实文件、取消、资源限制、所有权和 package consumer 测试。
- 当前 Provider 版本由项目文件和 `version.dev.props` 固定：NPOI 2.7.4、MiniExcel 1.46.0、ClosedXML 0.105.1、ExcelDataReader 3.9.0、SpreadCheetah 1.28.0、Aspose.Cells 26.8.0。Aspose 的 PDF/图片、公式、XLSM、ODS 和加密能力必须在已配置许可证与固定容器字体中验证，不能把本机无许可证测试当作生产渲染证据。
