# 第三方 Excel 引擎选型

下表描述底层组件的能力，不代表 Bing.Offices 已经暴露对应公共合同。新增 Provider 必须先声明“读取、写入、复杂工作簿、计算/渲染”职责，再补职责级合同和双 TFM 包消费者测试；不要求每个引擎实现全部职责。

| 组件／常用包 | 主要职责 | 授权与边界 | 在 Bing.Offices 中的定位 |
| --- | --- | --- | --- |
| [NPOI](https://www.nuget.org/packages/NPOI/2.7.4) | XLS/XLSX 读写、复杂 Workbook DOM | 当前锁定 2.7.4，Apache-2.0；后续版本授权需重新核实 | 已接入，复杂工作簿和 XLS 基准 |
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) | 高层 XLSX/XLSM 操作 | MIT；DOM 模型需要关注内存 | 已接入，富报表和模板 |
| [MiniExcel](https://github.com/mini-software/MiniExcel) | 轻量数据读写、模板 | Apache-2.0；侧重数据处理 | 已接入，常规大数据读取/写入 |
| [ExcelDataReader](https://www.nuget.org/packages/ExcelDataReader/3.9.0) | XLS/XLSX/XLSB 前向读取 | MIT；只读 | 已接入，XLSB 与单 Sheet 分批导入 |
| [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel) | DbDataReader 风格读取；平面 XLSX/XLSB 写入 | MIT；不编辑已有文件或复杂格式 | 数据库流式读取候选 |
| [SpreadCheetah](https://github.com/sveinung/spreadcheetah) | 流式 XLSX 生成 | MIT；生成型组件 | 已接入 `Bing.Offices.SpreadCheetah`，只创建新工作簿 |
| [LargeXlsx](https://github.com/salvois/LargeXlsx) | 低内存 XLSX 写入 | BSD-2-Clause；侧重生成 | 与 SpreadCheetah 对照后择一 |
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) | OOXML 底层结构读写 | MIT；不是完整计算/渲染引擎 | 结构保真和底层工具候选 |
| [EPPlus](https://www.epplussoftware.com/en/Home/GetEPPlus) | 丰富 XLSX 报表 | 5+ 为非商业/商业双许可 | 高级报表可选 Provider，需授权配置 |
| [Aspose.Cells](https://docs.aspose.com/cells/net/) | 多格式处理、公式计算、Excel 到 PDF/图片 | 商业授权；许可证由宿主配置 | 已接入独立 `Bing.Offices.AsposeCells`，不进入 Core |
| [Syncfusion XlsIO](https://www.syncfusion.com/document-processing/excel-framework/net) | Excel 报表和文档处理 | 商业授权；符合条件可申请社区许可 | 已使用 Syncfusion 技术栈时评估 |
| [GemBox.Spreadsheet](https://www.gemboxsoftware.com/spreadsheet) | 表格读写和转换 | 商业版、受限免费版 | 中等规模报表候选 |
| [DevExpress Spreadsheet Document API](https://docs.devexpress.com/OfficeFileAPI/14912/spreadsheet-document-api) | 工作簿编辑、打印、PDF/HTML | 商业订阅 | 已有 DevExpress 授权时评估 |
| [Telerik SpreadProcessing](https://www.telerik.com/document-processing-libraries/documentation/libraries/radspreadstreamprocessing/overview) | 工作簿和流式表格处理 | 商业/试用授权 | 已有 Telerik 技术栈时评估 |
| [MESCIUS DS.Documents.Excel](https://developer.mescius.com/document-solutions/dot-net-excel-api/docs/online/getting-started) | 跨平台 Excel 文档和模板 | 商业授权 | 复杂模板候选 |
| [Spire.XLS](https://cdn.e-iceblue.com/Introduce/free-xls-component.html) | Excel 操作和格式转换 | 商业版；免费版有格式、规模或转换限制 | 转换候选，需独立版本合同 |

## Provider 设计边界

- Core 只依赖 Abstractions，不依赖商业引擎；商业 Provider 由宿主单独引用并配置授权。
- 数据读取、数据写出、复杂工作簿、计算和渲染是独立职责。只读 Provider 可以只注册 `IExcelImporter`/`IExcelBatchImporter`。
- 公式文本、缓存值和重新计算必须分别声明；Excel 到 PDF/图片应采用独立渲染接口，不把导出接口伪装成渲染接口。
- XLSM、ODS、加密文件、宏保留、工作表保护和文件加密分别定义合同；宏保留不等于执行，工作表保护不等于文件加密。
- 接入新包时固定包版本、许可证、支持格式、内存模型、同步/异步边界和 Unsupported 行为，并补充真实文件、取消、资源限制、所有权和 package consumer 测试。
- 当前新增包版本固定为 SpreadCheetah 1.28.0 和 Aspose.Cells 26.8.0；Aspose 的 PDF/图片、公式、XLSM、ODS 和加密能力必须在已配置许可证与固定容器字体中验证，不能把本机无许可证测试当作生产渲染证据。
