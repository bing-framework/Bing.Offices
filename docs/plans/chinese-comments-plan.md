# 中文 XML 注释补全计划

## 扫描范围

- `src/Bing.Offices.Abstractions`
- `src/Bing.Offices.Core`
- `src/Bing.Offices.Npoi`

共扫描 177 个 C# 源文件。排除 `bin`、`obj`、`*.g.cs`、`*.generated.cs`、`*.Designer.cs`、EF Core Migration、ModelSnapshot、自动生成客户端/代理代码及第三方源码。

`src/YourProject` 不存在；经确认，本计划覆盖上述三个实际生产项目。

## 问题统计

| 类别 | 估算缺口 | 主要位置 |
| --- | ---: | --- |
| 类型摘要 | 约 30 | Npoi 内部辅助类、Core CSV/映射类型 |
| 构造函数 | 约 25 | Npoi 运行时/资源限制类型、CSV 异常与流包装器 |
| 方法 | 约 105 | Npoi 导入导出私有流程、Core CSV 私有流程 |
| 属性 | 约 70 | 内部计划/运行时对象、旧模型 |
| 字段、常量、`readonly`、`static readonly` | 约 80 | 导入导出器、CSV 管道、运行时状态 |
| DTO、Entity、Options、枚举 | 约 20 | Core Metadata/特性属性、内部枚举 |
| 可用 `inheritdoc` 的成员 | 约 18 | 接口实现、重写成员、`Stream` 包装器 |
| 标签不完整或不一致 | 约 40 | Core/Npoi 的早期方法 |
| 机械化或过时说明 | 约 45 | “获取/设置”式摘要、遗留序列化说明 |

未发现需要补充文档的索引器或事件。统计用于排序；实施前按当前签名逐项复核，不以文本扫描结果直接替代语义判断。

## 分批实施清单

### 1. 接口、抽象类和基类

- 完善 `Bing.Offices.Abstractions` 中接口、抽象类、公共模型基类和全部契约成员。
- 上游契约缺失时先补齐契约，之后由实现使用 `/// <inheritdoc />`。
- 逐项核验 `<param>`、`<typeparam>` 和 `<returns>` 与当前签名、返回类型一致。
- 重点目录或文件：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Providers/`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Csv/`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/FilterAttributeBase.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/DecoratorAttributeBase.cs`

### 2. DTO、Entity、Options 和枚举

- 补齐配置模型、特性参数、错误结果、运行时选项和所有枚举成员。
- 说明默认值、范围、单位、可空性、生命周期及配置作用域；不以字段或属性名称的简单翻译代替语义。
- 重点目录或文件：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Imports/`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Csv/`
  - `src/Bing.Offices.Npoi/ExcelColumnPlan.cs`
  - `src/Bing.Offices.Npoi/Imports/ExcelImportExecutionOptions.cs`
  - `src/Bing.Offices.Npoi/Imports/ExcelImportErrorCollector.cs`

### 3. 实现类和重写成员

- 接口实现、显式接口实现、抽象实现和 `override` 在存在有效上游注释时使用 `/// <inheritdoc />`。
- 删除或避免复制接口、抽象类和基类已有的 `<summary>`、`<param>`、`<returns>`。
- 实现具备额外资源所有权、缓存、线程安全或副作用时，在 inheritdoc 后添加准确的 `<remarks>`。
- 重点目录或文件：
  - `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
  - `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs`

### 4. private、internal、static 辅助方法

- 为所有具名 private、internal、protected、static、实例、异步和泛型辅助方法补齐说明。
- 摘要应描述实际职责，例如工作表定位、列计划、值转换、XML 安全设置、资源限制和流所有权。
- 复杂本地函数仅按需使用简洁普通注释，不强制 XML 文档。
- 重点目录或文件：
  - `src/Bing.Offices.Npoi/Imports/`
  - `src/Bing.Offices.Npoi/Exports/`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingDocumentValidator.cs`

### 5. 字段、常量、缓存键和配置键

- 为字段、常量、`readonly` 和 `static readonly` 成员补充中文 XML 注释。
- 说明依赖用途、缓存隔离范围、容量淘汰、并发保护、默认值、单位、输入上限和资源所有权。
- 重点目录或文件：
  - `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`
  - `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`
  - `src/Bing.Offices.Npoi/Exports/NpoiStyleCache.cs`
  - `src/Bing.Offices.Npoi/Extensions/CellExtensions.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvEntityPipeline.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlanFactory.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingDocumentValidator.cs`

### 6. 最终审计和构建验证

- 复扫全部纳入文件，确认所有目标成员均有准确注释或可审计例外。
- 校验 `inheritdoc` 的上游成员存在有效说明，并删除重复契约描述。
- 检查 XML 标签与签名一致，无机械化、错误、过时或无业务价值说明。

## 风险和待确认项

- 历史序列化构造函数和少数旧 Npoi 扩展的业务语义可能无法从局部代码可靠判断；此类成员应标记为待维护者确认，不编造说明。
- 需要复核各项目 Release 配置是否启用 XML 文档输出。未经额外授权，不调整构建配置或警告等级。
- 文档注释变更不得更改业务逻辑、公开 API、签名、命名空间、可见性、测试或项目配置。
- 持续 PowerShell 进程若锁定默认 Release 输出，使用唯一 `-p:OutputPath=...` 做隔离构建，不清理现有输出目录。

## 验证命令

```powershell
dotnet restore Bing.Offices.sln --locked-mode
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-build
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build
```

按 CI 环境可用性继续运行 `Bing.Offices.Tests.Integration` 的 `net6.0` 与 `net8.0`。同时检查包或构建产物仍包含 XML 文档文件。

## 验收标准

- 本计划包含扫描范围、问题统计、六个批次、涉及目录或文件、风险和待确认项、验证命令及验收标准。
- 排除范围外的类型、构造函数、具名方法、属性、索引器、事件、字段、常量、`readonly`、`static readonly`、DTO、Entity、Options、枚举及枚举成员均有准确中文 XML 文档，或具有明确可审计例外。
- 当前项目维护的接口、抽象类和基类契约优先完整；可继承成员使用 `/// <inheritdoc />`，不复制上游契约说明。
- 所有 `<param>`、`<typeparam>`、`<returns>` 与当前签名、泛型参数和返回类型一致。
- 不发生业务逻辑、公开 API 或项目配置变更；还原、Release 构建和核心 `net6.0`、`net8.0` 测试通过，且不引入 XML 文档格式错误或警告。
