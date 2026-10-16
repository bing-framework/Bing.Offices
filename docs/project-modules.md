# 项目模块与实现边界

本文面向后续开发者，按程序集说明当前已实现能力、核心入口、依赖与限制。更细的行为以链接的专题文档和公开 API 为准。

## 公共契约与核心服务

### Bing.Offices.Abstractions

- **功能**：定义 Excel/CSV 导入导出、Workbook Request、Entity Layout、映射配置、校验、异常、文件提交、渲染与 Provider 能力契约。
- **核心位置**：`src/Bing.Offices.Abstractions/Bing/Offices/{Exports,Imports,Entities,Csv,Providers,Rendering}`。
- **主要入口**：`ExcelExport.Workbook(...)`、`ExcelImport.Workbook<T>(...)`、`ExcelEntity.Layout<T>(...)`、`IExcelExporter`、`IExcelImporter`、`ICsvExporter`、`ICsvImporter`。
- **依赖与配置**：目标框架为 `netstandard2.0`，不绑定具体 Excel 引擎。
- **限制**：契约声明不代表每个 Provider 都支持全部能力；运行前应结合 `IExcelProviderFeatureDescriptor` 和 [Provider 矩阵](excel/09-providers.md) 判断。

### Bing.Offices.Core

- **功能**：编译 Mapping Plan，合并 Attribute、Profile、JSON/XML 与请求级配置；实现内置校验、CSV 管道、异常分发、文件原子提交和完整/流式导出策略。
- **核心位置**：`src/Bing.Offices.Core/Bing/Offices/{Mappings,Configurations,Validations,Csv,IO,Exports}`。
- **主要入口**：`AddMappingProfile<TProfile>()`、`ExcelMappingConfigurationLoader`、`ExcelWorkbookExportStrategy`；官方数据 Provider 和 SpreadCheetah 注册时会通过 `ExcelMappingPlanFactoryProvider.RegisterDefault` 注册 Core 服务和 CSV 服务，Aspose.Cells 商业扩展除外。
- **依赖与配置**：依赖 `Bing.Utils`、`CsvHelper` 与依赖注入抽象；Mapping Plan 默认缓存容量为 256。
- **限制**：Core 不创建具体 Excel 文件，也不承担商业渲染；多 Provider 同时注册时公共服务遵循 `TryAdd` 的先注册优先语义。

## Excel Provider

| 项目 | 已实现用途 | 关键入口与文件 | 主要依赖 | 已知限制 |
| --- | --- | --- | --- | --- |
| `Bing.Offices.Npoi` | XLS/XLSX 导入导出、模板、Entity Layout、批量导入、图表、透视和公式读取 | `AddBingOfficesNpoi`；`NpoiExcelImporter`、`NpoiExcelExporter` | NPOI | 图表仅 XLSX；完整公式重算与 XLS 高级透视不支持 |
| `Bing.Offices.ClosedXml` | 富 XLSX、模板、Entity Layout、批量导入、透视和名称列表校验 | `AddBingOfficesClosedXml`；`ClosedXmlProviderOptions` | ClosedXML | 仅 XLSX；DOM 并发受准入配置约束；图表不支持 |
| `Bing.Offices.MiniExcel` | 轻量 XLSX 列表/Workbook 导入导出 | `AddBingOfficesMiniExcel`；`MiniExcelExcelImporter`、`MiniExcelExcelExporter` | MiniExcel | 不提供 Entity Layout、XLS、复杂 DOM 功能 |
| `Bing.Offices.ExcelDataReader` | XLS/XLSX/XLSB 只读导入和批量导入 | `AddBingOfficesExcelDataReader`；`ExcelDataReaderExcelImporter` | ExcelDataReader、CodePages | 不注册导出器；不支持模板编辑/Entity Layout；失败工作簿支持 ErrorRowsOnly，AnnotatedOriginal 限 XLSX |
| `Bing.Offices.SpreadCheetah` | 新建 XLSX 的前向流式导出 | `AddBingOfficesSpreadCheetah`；`SpreadCheetahStreamingExcelExporter` | SpreadCheetah | 不读取、不编辑模板、不提供完整工作簿导出器 |
| `Bing.Offices.AsposeCells` | PDF/页面图片渲染、格式转换、公式处理、宏保留与加密 | `AddBingOfficesAsposeCells`；`AsposeCellsEngine` | Aspose.Cells | 必须配置有效商业许可证；部署字体与格式保真度待环境验证 |

各 Provider 的格式、流所有权、异步边界和细粒度能力见 [Excel Provider 能力与边界](excel/09-providers.md)。第三方引擎许可证和选型见 [第三方 Excel 引擎选型](excel/10-third-party-components.md)。

## CSV

CSV 实现在 Core 中，不是独立 Provider。任一官方数据 Provider 或 SpreadCheetah 的依赖注入扩展注册默认 Core 服务后，都可解析 `ICsvImporter` 与 `ICsvExporter`；仅注册 Aspose.Cells 不会获得 CSV 服务。CSV 复用 Excel Mapping Plan、值转换器、校验规则、异常 Observer 和文件提交器；用法、选项及限制见 [CSV 导入导出](csv/README.md)。

## 编译期分析器

`Bing.Offices.Analyzers` 是 `netstandard2.0` Roslyn 分析器包，检查可静态确定的 Entity Layout 地址、冲突、动态列重复和旧 API 使用。入口位于 `src/Bing.Offices.Analyzers`，依赖 `Microsoft.CodeAnalysis.CSharp`，以 `PrivateAssets="all"` 方式使用。它不解析工作簿模板或运行时变量，完整规则见 [Entity Layout 编译期诊断](excel/entity-layout-analyzers.md)。

## Provider 扩展与验证工具

| 项目/目录 | 用途 | 注意事项 |
| --- | --- | --- |
| `tests/Bing.Offices.EntityProviderContracts` | 可打包的 Entity Layout Provider 合同程序集 | 仅依赖 Abstractions，不依赖 xUnit；当前只在本地和 CI 生成包 |
| `tests/Bing.Offices.Testing` | 仓库内部合同数据、请求、快照和流测试辅助代码 | 供仓库测试引用，不作为业务运行时包说明 |
| `tests/Bing.Offices.ProviderContract.Tests` | 跨 Provider 的格式、流、取消、资源和能力一致性验证 | 通用驱动器覆盖 NPOI、MiniExcel、ClosedXML 和只读 ExcelDataReader；不替代引擎特有测试 |
| `tests/Bing.Offices.ResourceProbe` | NPOI 暂存路径和资源矩阵探针 | 测量结果依赖运行机器，不应直接当作生产容量承诺 |
| `tests/Bing.Offices.Docs.Tests` | 编译并执行指定 Excel Markdown 中的 C# 示例 | net8.0 项目；只覆盖 `DocsConsumerTest.cs` 固定清单中的 C# 代码块 |
| `templates/excel-provider-starter` | 第三方基础 XLSX Provider 起始模板 | **待实现/待确认（待实现）** Entity Layout，完成前不得启用对应合同声明 |

SpreadCheetah 只实现流式导出，Aspose.Cells 只实现商业扩展接口，因此不进入完整导入导出通用驱动器；它们由各自的职责级测试覆盖。`Bing.Offices.ThirdPartyProvider.Consumer` 验证开源 Provider 与合同包的 NuGet 消费，完整 Provider 包消费范围由 Net6/Net8 Consumer 项目承担。

第三方 Provider 的接入方式和合同范围见 [Entity Layout Provider 合同](excel/entity-provider-contracts.md)。

## 构建与配置

- 解决方案入口为 `Bing.Offices.sln`；公共版本由 `version.props` 管理。
- Abstractions、Core 和 Analyzer 目标为 `netstandard2.0`；官方 Provider 使用 `framework.props` 中的 `net6.0;net8.0`。
- Provider 包通过项目引用依赖 Core，Core 再依赖 Abstractions；商业 Aspose 依赖不会由 Core 传递。
- 配置文件、模板、输入流和输出流的所有权规则按各专题执行，不要从第三方库原生 API 推断公共契约。

## 待实现/待确认

- **待确认**：高级透视表部分客户端组合未完成最终验收，不能把包结构测试等同于 Excel/WPS 呈现成功。
- **待确认**：Aspose.Cells 的许可证、字体目录、替代字体与目标格式保真度需要部署环境验证。
- **待实现**：Starter Provider 的 Entity Layout 和对应合同接入。
- **待确认**：测试项目和资源探针是质量保障工具，不构成新的公开运行时 API；对外发布范围以包配置和 API 快照为准。

## 模块用法入口

| 模块 | 对外用法、配置和限制 |
| --- | --- |
| 公共接口、Core、依赖注入 | [快速开始](getting-started.md)、[Profile](excel/mapping-profile.md)、[v2 文档](excel/mapping-json-xml.md) |
| 工作簿与报表 | [工作簿、模板、报表内容与图表](excel/workbook-reports.md) |
| Entity Layout | [基础布局与模板](excel/entity-layout.md)、[动态列与分页](excel/dynamic-columns.md) |
| 各数据 Provider | [Provider 矩阵及实际注册/调用示例](excel/09-providers.md) |
| 分批、失败产物 | [批量导入与失败工作簿](excel/batch-import.md) |
| CSV | [CSV 接口与选项](csv/README.md) |
| 公式和商业扩展 | [公式、渲染与格式转换](excel/formula-rendering-conversion.md) |
| 编译器与维护工具 | [分析器](excel/entity-layout-analyzers.md)、[构建与验证](development/build-validation.md) |

`src/provider-shared` 通过各项目的 Compile Link 复用源码，没有单独的运行时程序集；包含前向 XLSX 解析、Entity 行计划/批次投影、报表预检及透视模板部件处理。不要把这些内部类型作为业务 API 使用。
