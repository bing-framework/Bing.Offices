# Bing.Offices
[![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)](https://mit-license.org/)

Bing.Offices是Bing应用框架的 Excel/CSV 导入导出类库。
当前发布程序集提供 Excel/CSV 相关能力；Word、PDF 不属于当前交付范围。

## Runtime 与异步边界

- `Bing.Offices.Npoi`、Unit Tests 和 Integration Tests 支持 `.NET 6` 与 `.NET 8`；Abstractions/Core 继续提供 `netstandard2.0` 资产。
- Excel 和 CSV 同时提供 Sync 与 Async API。CSV 的 Reader/Writer、Excel 的文件/Stream 复制和文件 flush 使用真实异步 IO，并继续传递 `CancellationToken`。
- NPOI `WorkbookFactory.Create`、Workbook DOM 操作和 `workbook.Write` 没有异步 API，因此 Excel Async 不承诺 DOM 阶段完全异步，也不使用 `Task.Run` 伪装异步。
- NPOI-specific 扩展位于 `Bing.Offices.Npoi.Extensions`；provider-neutral 文件/字节扩展位于 `Bing.Offices.Extensions`。

## Nuget Packages
|Nuget|版本号|说明|
|---|---|---|
|Bing.Offices.Abstractions|[![NuGet Badge](https://buildstats.info/nuget/Bing.Offices.Abstractions?includePreReleases=true)](https://www.nuget.org/packages/Bing.Offices.Abstractions)|
|Bing.Offices.Core|[![NuGet Badge](https://buildstats.info/nuget/Bing.Offices.Core?includePreReleases=true)](https://www.nuget.org/packages/Bing.Offices.Core)|
|Bing.Offices.Npoi|[![NuGet Badge](https://buildstats.info/nuget/Bing.Offices.Npoi?includePreReleases=true)](https://www.nuget.org/packages/Bing.Offices.Npoi)|

## 实现功能
- CSV 流式导入导出（只使用 CSV 时依赖 `Bing.Offices.Core`）
- Excel 导入
- Excel 导出
- 无需声明 Workbook 根模型的单 Sheet 导入
- Workbook Request 异构多 Sheet、动态列、模板、样式和结构化错误
- XLSX 柱状图、折线图和饼图

## Excel 文档

- [首次接入快速开始（CSV / Excel）](docs/excel/quick-start.md)
- [高级 Excel 导入导出](docs/excel/README.md)

## 架构文档

- [系统架构与设计文档](docs/architecture/README.md)
- [系统架构图册](docs/architecture/architecture.html)

## 文档自动化

- [文档自动化方案设计](docs/doc-automation/README.md)
- [docgen 使用指南](tools/docgen/README.md)
- 本地生成目录：`docs/generated/`（Markdown 和图表）、`artifacts/docs/`（DOCX / XLSX / PPTX / PDF），均由 Git 忽略
- 一键生成：`python tools/docgen/run.py all`（5 类文档 × 5 种格式 = 25 个产物）

本地校验命令（未通过返回非 0；文档 CI 工作流也会执行这些门禁）：

| 命令 | 作用 |
|---|---|
| `python tools/docgen/run.py gate` | 质量门禁：XML 注释覆盖率、提交规范符合率 |
| `python tools/docgen/run.py docs-check` | 文档腐化防护：文档示例必须来自 `Bing.Offices.Docs.Tests` 中可编译的源码且逐字一致 |

## 依赖类库
- [Bing.Utils](https://github.com/bing-framework/Bing.NetCore)
- [NPOI](https://github.com/tonyqus/npoi)

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

## XLSX 低内存处理

新增显式注册的 SXSSF 导出与 ZIP/XML 逐行导入服务，支持有界临时存储和逐行回调。参见 [低内存 API 与迁移说明](docs/excel/streaming.md)。
