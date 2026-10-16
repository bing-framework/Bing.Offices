# Bing.Offices
[![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)](https://mit-license.org/)

Bing.Offices是Bing应用框架的 Excel 导入导出类库。
当前开源主程序集提供 Excel/CSV 相关能力；PDF/图片渲染仅由独立的可选 Aspose.Cells 商业扩展提供，Core 不依赖该组件。

## 运行时与异步边界

- 六个 Provider 及主要职责测试支持 `.NET 6` 与 `.NET 8`；Abstractions/Core 提供 `netstandard2.0` 资产。文档示例测试项目自身为 `net8.0`，具体目标以各项目文件为准。
- Excel 和 CSV 同时提供同步与异步 API。CSV 的读取/写入、Excel 的文件/流复制和文件刷新使用真实异步 IO，并继续传递 `CancellationToken`。
- NPOI `WorkbookFactory.Create`、Workbook DOM 操作和 `workbook.Write` 没有异步 API，因此 Excel Async 不承诺 DOM 阶段完全异步，也不使用 `Task.Run` 伪装异步。
- NPOI 扩展位于 `Bing.Offices.Npoi.Extensions`；通用文件/字节扩展位于 `Bing.Offices.Extensions`。
- MiniExcel 扩展位于 `Bing.Offices.MiniExcel.Extensions`；应用启动时选择一个 Excel Provider，业务代码继续依赖 `IExcelImporter`/`IExcelExporter`。
- ClosedXML 注册扩展位于 `Bing.Offices.ClosedXml.Extensions`；它支持富 XLSX 报表、Entity Layout、基础/高级透视及已有 XLSX 透视模板保留/更新，不承诺 XLS、图表或客户端刷新结果。具体范围见 [Provider 能力与边界](docs/excel/09-providers.md)。

## NuGet 包
|NuGet 包|版本信息|说明|
|---|---|---|
|Bing.Offices.Abstractions|[NuGet](https://www.nuget.org/packages/Bing.Offices.Abstractions)|公共接口与请求模型|
|Bing.Offices.Core|[NuGet](https://www.nuget.org/packages/Bing.Offices.Core)|映射、CSV、校验、文件提交与公共扩展|
|Bing.Offices.Npoi|[NuGet](https://www.nuget.org/packages/Bing.Offices.Npoi)|XLS/XLSX 主 Provider|
|Bing.Offices.MiniExcel|按包源中可用版本选择|轻量 XLSX 导入导出|
|Bing.Offices.ClosedXml|按包源中可用版本选择|富 XLSX 报表、模板和实体布局|
|Bing.Offices.ExcelDataReader|按包源中可用版本选择|只读 XLS/XLSX/XLSB，完整/单表/多表批次；支持失败行 XLSX 和受限原文件批注输出|
|Bing.Offices.SpreadCheetah|按包源中可用版本选择|新建 XLSX 前向导出，不读取或编辑模板|
|Bing.Offices.AsposeCells|按包源中可用版本选择|独立商业渲染、公式和格式/加密转换；宿主提供许可证|

仓库版本配置不代表线上发布状态；构建、分析器及合同包另见[模块总览](docs/project-modules.md)。

## 实现功能
- Excel 导入
- Excel 导出
- Workbook Request 异构多 Sheet、动态列、模板、样式和结构化错误
- XLSX 柱状图、折线图和饼图
- ClosedXML 基础 XLSX 导入导出、公式保存/读回、模板样式保留、合并和真实文件提交
- MiniExcel Provider 的常规 XLSX 导入导出；历史 100K 行受控探针不代表任意容量，500K/1M 和生产机器峰值为待实现/待确认（待确认）
- 公共报表模型（表格、筛选、冻结、条件格式、名称范围和打印布局）已接入 NPOI/ClosedXML；SpreadCheetah 仅实现前向模型可表达的子集
- ClosedXML XLSX 透视表支持基础及高级定义和完整缓存；历史本机 Excel 首开、刷新、重开验证只覆盖基础夹具，高级聚合、WPS/LibreOffice 与其他客户端组合为待实现/待确认（待确认）。详见 [透视表合同](docs/excel/pivot-tables.md)。
- Sheet 级 PNG/JPEG 图片和原生数据校验已接入 NPOI、ClosedXML、SpreadCheetah；MiniExcel 在预检阶段拒绝。图片仅接受调用方字节，不下载或自动读取文件。
- SpreadCheetah 固定列和动态列复用 Mapping Plan、ValueMap、命名 Converter；批次串行写入，默认 1,000 行，不预取下一批。图片请求使用磁盘暂存以适配 JPEG 和精确像素锚点。
- 公式文本、缓存值和重新计算已拆分为独立合同；NPOI/ClosedXML 只读取已有缓存，不证明缓存新鲜；Aspose.Cells 的计算与渲染属于独立商业扩展

## 文档入口

- [全部文档索引](docs/README.md)
- [快速开始](docs/getting-started.md)
- [功能与文档对照清单](docs/documentation-inventory.md)
- [CSV 导入导出](docs/csv/README.md)
- [高级 Excel 导入导出](docs/excel/README.md)
- [Excel Provider 能力与边界](docs/excel/09-providers.md)
- [第三方 Excel 引擎选型](docs/excel/10-third-party-components.md)
- [源码目录与命名空间约定](docs/source-layout.md)

## 依赖类库
- [Bing.Utils](https://github.com/bing-framework/Bing.NetCore)
- [NPOI](https://github.com/tonyqus/npoi)
- [MiniExcel](https://github.com/mini-software/MiniExcel)
- [ClosedXML](https://github.com/ClosedXML/ClosedXML)
- [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)
- [SpreadCheetah](https://github.com/sveinungf/spreadcheetah)
- [Aspose.Cells](https://docs.aspose.com/cells/net/)

## 示例

可执行的 Excel 往返示例见[快速开始](docs/getting-started.md)，批次、报表和 Entity 示例按[索引](docs/README.md)进入。

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

## 许可证

**MIT**

> 本仓库代码使用 MIT 许可证；这不授予第三方商业组件的免费使用权，Aspose.Cells 的许可证由宿主另行准备。

> Bing 会尽量引入开源免费的第三方技术框架，如有意外，还请自行了解。
