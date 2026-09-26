# Bing.Offices
[![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)](https://mit-license.org/)

Bing.Offices是Bing应用框架的 Excel 导入导出类库。
当前开源主程序集提供 Excel/CSV 相关能力；PDF/图片渲染仅由独立的可选 Aspose.Cells 商业扩展提供，Core 不依赖该组件。

## Runtime 与异步边界

- `Bing.Offices.Npoi`、`Bing.Offices.MiniExcel`、`Bing.Offices.ClosedXml`、`Bing.Offices.ExcelDataReader`、`Bing.Offices.SpreadCheetah`、`Bing.Offices.AsposeCells`、Unit Tests 和 Integration Tests 支持 `.NET 6` 与 `.NET 8`；Abstractions/Core 继续提供 `netstandard2.0` 资产。
- Excel 和 CSV 同时提供 Sync 与 Async API。CSV 的 Reader/Writer、Excel 的文件/Stream 复制和文件 flush 使用真实异步 IO，并继续传递 `CancellationToken`。
- NPOI `WorkbookFactory.Create`、Workbook DOM 操作和 `workbook.Write` 没有异步 API，因此 Excel Async 不承诺 DOM 阶段完全异步，也不使用 `Task.Run` 伪装异步。
- NPOI-specific 扩展位于 `Bing.Offices.Npoi.Extensions`；provider-neutral 文件/字节扩展位于 `Bing.Offices.Extensions`。
- MiniExcel-specific 扩展位于 `Bing.Offices.MiniExcel.Extensions`；应用启动时选择一个 Excel Provider，业务代码继续依赖 `IExcelImporter`/`IExcelExporter`。
- ClosedXML-specific 注册扩展位于 `Bing.Offices.ClosedXml.Extensions`；它面向富 XLSX 报表、模板、样式、合并、公式保存和基础 Entity Layout，不承诺 XLS、图表、PivotTable 或复杂 Entity/模板结构。

## Nuget Packages
|Nuget|版本号|说明|
|---|---|---|
|Bing.Offices.Abstractions|[![NuGet Badge](https://buildstats.info/nuget/Bing.Offices.Abstractions?includePreReleases=true)](https://www.nuget.org/packages/Bing.Offices.Abstractions)|
|Bing.Offices.Core|[![NuGet Badge](https://buildstats.info/nuget/Bing.Offices.Core?includePreReleases=true)](https://www.nuget.org/packages/Bing.Offices.Core)|
|Bing.Offices.Npoi|[![NuGet Badge](https://buildstats.info/nuget/Bing.Offices.Npoi?includePreReleases=true)](https://www.nuget.org/packages/Bing.Offices.Npoi)|
|Bing.Offices.MiniExcel|独立的 XLSX 流式 Provider；发布后按版本选择使用|
|Bing.Offices.ClosedXml|ClosedXML 富 XLSX Provider；发布后按版本选择使用|
|Bing.Offices.ExcelDataReader|只读 XLS/XLSX/XLSB Provider；支持完整导入和单 Sheet 分批导入，不提供导出|
|Bing.Offices.SpreadCheetah|只创建新 XLSX 的前向流式导出 Provider；不读取、不编辑模板|
|Bing.Offices.AsposeCells|可选商业 Provider；提供 PDF/图片渲染、公式计算、XLSM/ODS/加密转换；宿主负责许可证|

## 实现功能
- Excel 导入
- Excel 导出
- Workbook Request 异构多 Sheet、动态列、模板、样式和结构化错误
- XLSX 柱状图、折线图和饼图
- ClosedXML 基础 XLSX 导入导出、公式保存/读回、模板样式保留、合并和真实文件提交
- MiniExcel Provider 的常规 XLSX 导入导出；100K 行受控探针已验证，500K/1M 和生产机器峰值仍为 `NOT_VERIFIED`
- 公共报表模型（表格、筛选、冻结、条件格式、名称范围和打印布局）已接入 NPOI/ClosedXML；SpreadCheetah 仅实现前向模型可表达的子集
- Sheet 级 PNG/JPEG 图片和原生数据校验已接入 NPOI、ClosedXML、SpreadCheetah；MiniExcel 在预检阶段拒绝。图片仅接受调用方字节，不下载或自动读取文件。
- SpreadCheetah 固定列和动态列复用 Mapping Plan、ValueMap、命名 Converter；批次串行写入，默认 1,000 行，不预取下一批。图片请求使用磁盘暂存以适配 JPEG 和精确像素锚点。
- 公式文本、缓存值和重新计算已拆分为独立合同；NPOI/ClosedXML 只读取已验证缓存，Aspose.Cells 的完整计算与渲染属于独立商业扩展

## Excel 文档

- [高级 Excel 导入导出](docs/excel/README.md)
- [Excel Provider 能力与边界](docs/excel/09-providers.md)
- [第三方 Excel 引擎选型](docs/excel/10-third-party-components.md)

## 依赖类库
- [Bing.Utils](https://github.com/bing-framework/Bing.NetCore)
- [NPOI](https://github.com/tonyqus/npoi)
- [MiniExcel](https://github.com/mini-software/MiniExcel)
- [ClosedXML](https://github.com/ClosedXML/ClosedXML)
- [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)
- [SpreadCheetah](https://github.com/sveinungf/spreadcheetah)
- [Aspose.Cells](https://docs.aspose.com/cells/net/)

## Demo

## 作者

简玄冰

## 贡献与反馈

> 如果你在阅读或使用Bing中任意一个代码片断时发现Bug，或有更佳实现方式，请通知我们。

> 当前发布范围已覆盖 Excel/CSV 的导入导出、Workbook Request、映射与校验、动态列、模板、结构化错误、文件原子提交以及 Sync/Async IO。NPOI DOM 构建和序列化仍受其同步 API 约束，资源限制与并发入口配置应按 Excel 文档中的运行时边界部署。

> 你可以通过github的Issue或Pull Request向我们提交问题和代码，如果你更喜欢使用QQ进行交流，请加入我们的交流QQ群。

> 对于你提交的代码，如果我们决定采纳，可能会进行相应重构，以统一代码风格。

> 对于热心的同学，将会把你的名字放到**贡献者**名单中。

## 免责声明
- 虽然我们对代码已经进行高度审查，并用于自己的项目中，但依然可能存在某些未知的BUG，如果你的生产系统蒙受损失，Bing 团队不会对此负责。
- 出于成本的考虑，我们不会对已发布的API保持兼容，每当更新代码时，请注意该问题。

## 开源地址
[https://github.com/bing-framework/Bing.Offices](https://github.com/bing-framework/Bing.Offices)

## License

**MIT**

> 这意味着你可以在任意场景下使用 Bing 应用框架而不会有人找你要钱。

> Bing 会尽量引入开源免费的第三方技术框架，如有意外，还请自行了解。
