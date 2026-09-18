# 中文 XML 注释补全计划（tests/*）

## 1. 目标和范围

扫描日期：2026-09-18。规范依据：`C:/Users/jianx/.agents/skills/chinese-comments/SKILL.md`，版本 1.2.1。

本轮只规划 `tests/**` 下当前存在的 C# 文件。规划阶段只扫描和分析，不修改任何 C# 文件；后续实施按批次推进，每次只执行一个未完成批次并在验证后停止。扫描以当前工作区文件为基线，保留仓库已有的测试架构改动、删除项和未跟踪项目，不恢复或重排它们。

扫描对象包括类型（class、abstract class、record、struct、enum、delegate 以及嵌套类型）、构造函数、具名方法、扩展方法、属性、索引器、事件、所有访问级别字段、常量和枚举成员；局部函数只作为语义上下文，不强制生成 XML 文档。

当前范围共 73 个候选 C# 文件、2,643 个声明，分布如下：

| 项目/目录 | 文件数 | 声明数 | 目标框架或职责 |
| --- | ---: | ---: | --- |
| `tests/Bing.Offices.Tests` | 37 | 728 | Core/Abstractions 单元测试，net6.0 / net8.0 |
| `tests/Bing.Offices.Tests.Integration` | 3 | 138 | Provider/API 集成测试，net6.0 / net8.0 |
| `tests/Bing.Offices.Npoi.Tests` | 21 | 1,144 | NPOI 单元和边界测试，net6.0 / net8.0 |
| `tests/Bing.Offices.Npoi.Tests.Integration` | 2 | 135 | NPOI 文件集成测试，net6.0 / net8.0 |
| `tests/Bing.Offices.MiniExcel.Tests` | 1 | 207 | MiniExcel 单元测试，net6.0 / net8.0 |
| `tests/Bing.Offices.MiniExcel.Tests.Integration` | 1 | 38 | MiniExcel 文件集成测试，net6.0 / net8.0 |
| `tests/Bing.Offices.Docs.Tests` | 2 | 54 | 文档消费者和示例验证，net8.0 |
| `tests/Bing.Offices.ProfileFixtures` | 1 | 6 | 外部 Profile/模型夹具，netstandard2.0 |
| `tests/Bing.Offices.ResourceProbe` | 3 | 183 | staging 资源/入口矩阵，net8.0 |
| `tests/Bing.Offices.Consumer.Net6` | 1 | 5 | 外部消费者示例，net6.0 |
| `tests/Bing.Offices.Consumer.Net8` | 1 | 5 | 外部消费者示例，net8.0 |
| **合计** | **73** | **2,643** | 当前工作区可见测试源码 |

## 2. 排除范围

- `bin/**`、`obj/**`。
- `*.g.cs`、`*.generated.cs`、`*.Designer.cs`；EF Core Migration/Migrations、ModelSnapshot。
- 自动生成客户端、代理代码、第三方源码和 NuGet 包源码；本轮未发现测试目录中有命中上述规则的现存候选文件。
- 删除但当前磁盘不存在的旧测试文件不恢复、不纳入统计；未跟踪测试项目若已存在于磁盘则按当前源码纳入。
- `DocsExamples.cs`、Consumer 程序和测试中的字符串代码片段只审查实际 C# 声明；不把字符串内的示例代码当作独立声明，也不修改测试数据、断言、资源路径或运行逻辑。
- 只读扫描脚本、临时 XML、测试输出和性能矩阵产物不写入仓库；本阶段不修改任何 C#、`.csproj`、快照或测试资源文件。

## 3. 扫描统计

### 3.1 声明类型统计

| 声明类型 | 总数 | 缺失 summary/inheritdoc |
| --- | ---: | ---: |
| 类型 | 418 | 281 |
| 构造函数 | 48 | 39 |
| 方法 | 1,223 | 725 |
| 属性/索引器 | 818 | 555 |
| 字段 | 98 | 92 |
| 常量 | 12 | 12 |
| 枚举成员 | 26 | 20 |
| **合计** | **2,643** | **1,724** |

### 3.2 注释结构、标签和继承统计

| 检查项 | 数量 | 说明 |
| --- | ---: | --- |
| 现有 `inheritdoc` | 100 | 主要集中在 Stream/文件适配测试替身 |
| 单行 summary | 18 | 与非三行 summary 完全重合 |
| 非三行 summary | 18 | 需统一为独立起始、正文、结束三行 |
| `<param>` 与签名不一致 | 647 | 其中 587 个同时缺少 summary，60 个为已有文档标签错误 |
| `<typeparam>` 与签名不一致 | 50 | 其中 49 个同时缺少 summary，1 个为已有文档标签错误 |
| `<returns>` 与返回类型不一致 | 327 | 其中 321 个同时缺少 summary，6 个为已有文档标签错误 |
| 构造函数句式/cref 异常 | 41 | 39 个缺失，2 个已有句式或 cref 不符合规则 |
| 属性访问前缀异常 | 774 | 555 个缺失，219 个已有摘要未按访问方式起句 |
| `override` 缺少 `inheritdoc` | 175 | 全部为缺失注释的覆盖成员；显式接口实现候选为 0 |
| XML 解析错误 | 0 | 当前文本可解析 |
| 单数 `<remark>` | 0 | 未发现非法标签 |
| summary 超过 120 字 | 0 | 仍需人工检查多句测试契约是否应下沉到 remarks |

缺失声明按访问级别分布为：public 1,124、private 546、protected 16、internal 18、implicit 20；其中 static 缺失 234 个。字段/常量缺失包括 private 非 readonly 字段 25 个、private readonly 字段 55 个、private static readonly 字段 12 个以及 private 常量 12 个。

### 3.3 高密度文件

缺失声明最多的文件为：`MiniExcelProviderTest.cs` 199 个、`Npoi.Tests/AsyncPipelineTest.cs` 165 个、`Npoi.Tests/ExcelWorkbookRequestTest.cs` 138 个、`Npoi.Tests/ExcelP0RegressionTest.cs` 116 个、`ResourceProbe/StagingResourceMatrix.cs` 93 个、`ResourceProbe/StagingEntrypointMatrix.cs` 70 个、`Tests/ExcelStreamExtensionsTest.cs` 58 个、`Tests.Integration/MiniExcelProviderContractTest.cs` 58 个、`Tests/CsvStreamExtensionsTest.cs` 56 个、`Npoi.Tests/StreamPipelineTest.cs` 52 个、`Npoi.Tests/TemplateAsyncBoundaryTest.cs` 50 个、`Tests/AsyncPipelineCsvTest.cs` 46 个、`Tests/ReviewFixCoreRegressionTest.cs` 42 个、`Npoi.Tests/NpoiProviderBoundaryRegressionTest.cs` 40 个、`Npoi.Tests/NpoiSheetPictureExtensionsTest.cs` 39 个。

单行摘要集中在 `Npoi.Tests/{AsyncPipelineTest,ExcelMappingConfigurationLoaderTest,ExcelP0RegressionTest,NpoiFailureWorkbookPreflightTest}.cs`、`Npoi.Tests.Integration/ExcelImporterIntegrationTest.cs`、`ProfileFixtures/ExternalMappingProfile.cs`、`Tests/{AsyncPipelineCsvTest,ExcelMappingConfigurationLoaderTest,ReviewFixCoreRegressionTest}.cs`，共 18 个声明。

## 4. 问题分类

### 4.1 类型、DTO、Entity、Options 和枚举

测试模型和嵌套工作簿/行模型占类型缺口主体，包括 `tests/Bing.Offices.Tests/Models/**/*.cs`、Consumer 的 `ConsumerRow/ConsumerWorkbook`、Docs 的 `Docs*`/`Order*` 模型、MiniExcel 的多组 Workbook/Row/Converter 类型，以及 NPOI 的异步、关系、图片、校验和文件系统测试替身。当前类型缺失 281 个、枚举成员缺失 20 个。测试夹具即使是 private nested type，也要按真实职责写摘要；不要把测试名称或字段名称直接复制为类型说明。

### 4.2 方法、属性和构造函数

公共测试方法、内部示例方法和大量 private helper/stream 方法均有缺口：方法缺失 725 个，属性缺失 555 个，构造函数缺失 39 个。属性访问前缀检查另有 219 个已有摘要需要人工核对；测试替身的 `Dispose`、异步流边界、转换器和验证器成员需要根据真实行为决定 `inheritdoc` 或独立摘要。`Task<T>`、`ValueTask<T>`、`bool`、nullable 和 `out` 参数的返回语义不能由测试名称臆测。

### 4.3 字段、常量、缓存键和配置键

ResourceProbe 的任务 ID、策略、阈值、并发级别、矩阵场景和报告字段，以及各测试类的静态 fixture、流状态和临时路径字段均需检查。当前字段缺失 92 个、常量缺失 12 个；需说明实际用途、作用域、单位、默认值和边界，不把断言数据误写成生产配置或缓存策略。

### 4.4 标签与签名

已有文档中存在 647 个参数标签、50 个泛型参数标签和 327 个返回标签不匹配。问题集中于超长测试文件、测试替身的 `Stream`/Provider 接口成员、Docs 示例 helper、Public API 快照 helper 和 ResourceProbe 矩阵方法。应按实际签名重建标签，删除不存在的参数，补充 `Task<T>` 最终结果、`bool` 双分支、`IEnumerable<T>`/nullable 和异常传播说明；不得为了消除计数而虚构业务约束。

### 4.5 机械化、冗长和过时文案

现有测试摘要大量使用“`测试 -` + 完整断言条件”的模板，虽然当前均未超过 120 字，但其中部分同时描述输入、实现路径、异常优先级和清理结果，适合把调用方需要关注的边界放入 `<remarks>`。单行 summary 需统一三行；不应把 `Should...` 方法名逐字翻译成无意义摘要，也不应删除对测试契约有价值的取消、流所有权、资源清理和错误分类信息。

## 5. 文案规范问题

- 类型 summary 用一句话说明测试夹具、模型、测试类或矩阵 runner 的核心职责；测试场景的条件、来源和门禁写入 remarks。
- 构造函数统一使用 `初始化一个 <see cref="当前类型" /> 类型的实例。`；若是编译器生成或无显式构造函数的类型，不机械添加文档。
- 属性按访问器使用“获取”“设置”“获取或设置”“获取或初始化”；测试模型的 DTO 属性只说明数据含义和空值/集合边界。
- 方法 summary 只说明测试验证或 helper 执行的核心用途；重载保持相同核心摘要，差异放入 param、returns 或 remarks。
- `Stream`、`DispatchProxy`、Provider 接口和 Profile 实现不得复制上游契约；实现额外记录的事件、取消、异常或流所有权只追加 remarks。
- 字段/常量直接说明任务标识、阈值、状态、集合或临时资源的实际用途；不使用“获取或设置”，不把测试输入值描述成生产缓存或配置。
- `DocsExamples.cs` 与 Consumer 示例中的实际 helper/模型可以有文档，但不修改字符串代码块、测试数据和断言逻辑。

## 6. inheritdoc 使用问题

- 当前已存在 100 个 `inheritdoc`，主要在 `Npoi.Tests/StreamPipelineTest.cs`（44）、`Npoi.Tests.Integration/ExcelImporterIntegrationTest.cs`（32）和 `Tests/CsvTest.cs`（24）；后续只校验其上游契约和额外 remarks，不复制 summary、param 或 returns。
- Roslyn 初筛发现 175 个 override 缺少 `inheritdoc`，集中在 MiniExcel/NPOI/CSV 的异步流、不可寻址流、DispatchProxy、失败文件系统和测试替身。它们应逐成员核对基类真实契约后补 `/// <inheritdoc />`，不能用统一摘要掩盖覆盖差异。
- 未发现显式接口实现缺少 inheritdoc，但存在大量隐式实现 `Stream`、`IExcelExporter`、`ICsvImporter`、`IMappingProfile`、`IExcelValidationRule`、`INamedExcelValueConverter` 等接口的测试替身；执行批次 1 先阅读 `src/*` 上游契约，再在批次 4 处理实现成员。
- 若上游契约缺少必要注释，记录为上游依赖问题并只在允许的契约范围内补齐；本测试计划不借机修改生产逻辑、签名或 API。

## 7. 实施批次

### 第 1 批：接口、抽象类和公共契约

范围：`ProfileFixtures/ExternalMappingProfile.cs`、`Docs.Tests/{DocsConsumerTest,DocsExamples}.cs`、`Tests.Integration/{PublicApiContractTest,MiniExcelProviderContractTest,CsvAsyncFileIntegrationTest}.cs`、`Tests/{ApiSnapshotCandidateIdentityTest,PublicCoreExtensionCoverageTest}.cs`、`Npoi.Tests/{PublicNpoiExtensionCoverageTest,NpoiServiceCollectionExtensionsTest}.cs`，以及 Consumer 的公开示例类型。

任务：补齐公共测试类、外部 Profile、消费者示例和契约检查入口的类型/方法摘要；确认隐式接口实现的上游注释可供后续 inheritdoc 使用；修正现有契约标签但不修改快照、断言或项目配置。

### 第 2 批：DTO、Entity、Options 和枚举

范围：`tests/Bing.Offices.Tests/Models/**/*.cs`、Consumer `Program.cs` 中的行/工作簿模型、Docs 的 `Docs*`/`Order*` 模型、MiniExcel/NPOI 单元和集成文件中所有 Workbook/Row/Result/Converter/Validation fixture，以及 `CsvStreamExtensionsTest.cs`、`ExcelStreamExtensionsTest.cs`、`NpoiCellExtensionsTest.cs`、`ExcelP0RegressionTest.cs`、`StreamPipelineTest.cs`、`MiniExcelProviderContractTest.cs`、ResourceProbe 枚举。

任务：补齐 281 个类型和 20 个枚举成员缺口，明确测试数据职责、集合关系、可空值、转换器/校验器作用；枚举成员逐项说明测试分支含义，不把枚举值名称直接翻译为无意义摘要。

### 第 3 批：构造函数和属性文案统一

范围：全量 73 个候选文件，优先处理 `MiniExcelProviderTest.cs`、`ExcelWorkbookRequestTest.cs`、`ExcelP0RegressionTest.cs`、`StreamPipelineTest.cs`、`StagingResourceMatrix.cs`、`DocsConsumerTest.cs`、`DocsExamples.cs` 和 `tests/Models/**/*.cs`。

任务：处理 48 个构造函数和 818 个属性；补齐 39 个构造函数缺口，统一 41 个构造函数句式/cref 异常，补齐并修订 774 个属性访问文案；不为属性访问器重复生成方法注释。

### 第 4 批：实现类、重写成员和重载方法

范围：所有实现 `Stream`、`DispatchProxy`、`IExcelExporter`、`IExcelImporter`、`ICsvExporter`、`ICsvImporter`、`IMappingProfile`、`IExcelValidationRule`、`INamedExcelValidationRule`、`INamedExcelValueConverter`、`IExcelMappingPlanFactory`、`IFailureWorkbookFileSystem` 等契约的测试替身；重点文件为 `Npoi.Tests/{AsyncPipelineTest,ExcelP0RegressionTest,ExcelWorkbookRequestTest,StreamPipelineTest,NpoiSheetPictureExtensionsTest,NpoiSheetExtensionsTest,CsvTest}.cs`、两个 NPOI 集成文件、MiniExcel Provider 测试和 `Tests/{AsyncPipelineCsvTest,ExcelStreamExtensionsTest,CsvStreamExtensionsTest}.cs`。

任务：为 175 个 override 候选补 `inheritdoc`；检查隐式接口实现和现有 100 个 inheritdoc；统一同步/异步、流/文件、XLS/XLSX、CSV/Excel 和重载方法的核心 summary，差异写入 param、returns、remarks 或必要 exception。

### 第 5 批：private、internal、static 辅助方法

范围：`Npoi.Tests/ExcelP0RegressionTest.cs`、`ExcelWorkbookRequestTest.cs`、`StreamPipelineTest.cs`、`AsyncPipelineTest.cs`、`TemplateAsyncBoundaryTest.cs`，`MiniExcelProviderTest.cs`，`Tests/CsvTest.cs`、`ReviewFixCoreRegressionTest.cs`、`AsyncPipelineCsvTest.cs`，`ResourceProbe/{StagingEntrypointMatrix,StagingResourceMatrix}.cs`，`Docs.Tests/{DocsConsumerTest,DocsExamples}.cs` 及其余测试文件中的 helper。

任务：补齐 725 个方法缺口，重点覆盖 private/internal/static 的流控制、资源矩阵、快照 canonicalizer、文档 fence、转换器和异常构造 helper；返回 `bool`、nullable、集合和异步最终结果必须按实现确认，不把测试名称当作业务语义。

### 第 6 批：private 字段、其他字段、常量、缓存键和配置键

范围：`ResourceProbe/{StagingEntrypointMatrix,StagingResourceMatrix}.cs`、`Npoi.Tests/{DependencyContainer,TestBase,AsyncPipelineTest,ExcelP0RegressionTest,ExcelWorkbookRequestTest,StreamPipelineTest}.cs`、`MiniExcelProviderTest.cs`、`Tests` 和集成测试中的 fixture/状态字段；覆盖所有 98 个字段与 12 个常量。

任务：补齐 92 个字段和 12 个常量摘要，说明测试状态、任务 ID、策略、阈值、并发、临时文件、缓存/字典键和对象生命周期；不把测试缓存描述成生产缓存淘汰策略，不修改常量值、测试数据或资源清理逻辑。

### 第 7 批：最终审计和构建验证

范围：全量 73 个候选文件。

任务：复扫缺失/三行格式、XML/cref、参数/泛型/返回标签、构造函数、属性、bool/null/异步语义、重载一致性、隐式/显式接口实现和 override；构建解决方案未覆盖的 Consumer/ResourceProbe 项目；运行双 TFM 单元/集成测试、Docs 测试和必要的 API/示例验证；确认差异只包含注释和本计划记录。

## 8. 风险和待确认项

- 测试项目当前工作区存在大量修改、删除和未跟踪新项目；计划基线按磁盘现状统计，实施不得恢复或覆盖这些改动。
- `Bing.Offices.Consumer.Net6/Net8` 和 `Bing.Offices.ResourceProbe` 不在解决方案项目列表中，最终验证必须直接构建；不能只依赖 `dotnet build Bing.Offices.sln`。
- 测试替身同时继承 .NET 基类并实现生产接口，接口/override 的 inheritdoc 目标需按实际编译符号和 `src/*` 契约确认；不能把所有缺失成员机械标成 inheritdoc。
- `DocsExamples.cs` 和 Consumer 文件含有文档代码或顶层程序；字符串内容、样例数据、断言和入口行为不属于 XML 注释修改范围。
- ResourceProbe 会生成性能/资源报告并包含阈值和并发语义；只能记录代码明确的单位、默认值和边界，不推断未实现的性能保证。
- 测试类的大量 `测试 -` 摘要是已有文案，需区分真正冗余和测试契约；无法确认时保留客观描述并列为待确认，不编造业务规则。
- `common.tests.props` 当前关闭 CS1591 警告；实施不得通过继续抑制 XML 文档警告代替注释修复。
- 全部文本按 UTF-8 处理并保留换行约定；`git diff --check` 的既有 LF/CRLF 提示不应被误判为代码错误。

## 9. 验证命令

以下命令供实施批次使用；规划阶段不执行构建、测试或 C# 修改。命令中的临时文档目录必须位于 `%TEMP%`，不删除现有构建输出。

```powershell
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

# 使用 Roslyn 语法树扫描 tests，应用本计划的排除规则并输出 UTF-8 JSON 统计。
& "$env:TEMP\chinese-comments-tests-audit.ps1" -Root (Get-Location).Path

dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false
dotnet build tests/Bing.Offices.Consumer.Net6/Bing.Offices.Consumer.Net6.csproj -c Release --no-restore -v:minimal -p:UseSharedCompilation=false
dotnet build tests/Bing.Offices.Consumer.Net8/Bing.Offices.Consumer.Net8.csproj -c Release --no-restore -v:minimal -p:UseSharedCompilation=false
dotnet build tests/Bing.Offices.ResourceProbe/Bing.Offices.ResourceProbe.csproj -c Release --no-restore -v:minimal -p:UseSharedCompilation=false

dotnet test Bing.Offices.sln -c Release --no-build --no-restore -v:minimal --nologo
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release --no-build --no-restore -v:minimal --nologo
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release --no-build --no-restore -v:minimal --nologo
dotnet test tests/Bing.Offices.Npoi.Tests/Bing.Offices.Npoi.Tests.csproj -c Release --no-build --no-restore -v:minimal --nologo
dotnet test tests/Bing.Offices.Npoi.Tests.Integration/Bing.Offices.Npoi.Tests.Integration.csproj -c Release --no-build --no-restore -v:minimal --nologo
dotnet test tests/Bing.Offices.MiniExcel.Tests/Bing.Offices.MiniExcel.Tests.csproj -c Release --no-build --no-restore -v:minimal --nologo
dotnet test tests/Bing.Offices.MiniExcel.Tests.Integration/Bing.Offices.MiniExcel.Tests.Integration.csproj -c Release --no-build --no-restore -v:minimal --nologo
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release --no-build --no-restore -v:minimal --nologo

# 各测试项目以隔离 DocumentationFile 生成 XML，再按 UTF-8 解析全部文件。
Get-ChildItem "$env:TEMP\bing-offices-tests-docs" -Filter '*.xml' -Recurse |
    ForEach-Object { [xml](Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8) | Out-Null }

git diff --check
git diff --unified=0 -- tests docs/plans/chinese-comments-plan.md
```

扫描命令需同时验证 `cref` 目标、summary 三行结构、隐式接口映射和条件编译；文本检索只能作为初筛。构建/测试失败必须记录真实项目、目标框架和错误，不通过删除测试或修改测试逻辑规避。

## 10. 验收标准

- 覆盖当前基线 73 个候选文件和 2,643 个声明；缺失 summary/inheritdoc、单行/非三行摘要清零，或为明确的合法排除项记录证据。
- 48 个构造函数全部使用合法当前类型 `cref` 和固定句式；818 个属性按访问方式描述；98 个字段和 12 个常量无遗漏。
- 类型、测试模型、Profile、Workbook/Row fixture 和 26 个枚举成员职责准确；测试数据、枚举值和矩阵阈值不被改写。
- 1,223 个方法的 `<param>`、`<typeparam>`、`<returns>` 与签名一致；`bool` 双结果、nullable、异步最终结果和异常传播有真实依据；重载核心摘要一致。
- 175 个 override 和所有确认的隐式接口实现使用正确 `inheritdoc`；不复制生产接口/基类的 summary、param 或 returns；额外测试行为只写 remarks/exception。
- XML 可解析、`cref` 合法、无非法 `<remark>`、无机械化无意义文案；多句测试条件下沉到 remarks，摘要保持核心用途。
- 解决方案及其未收录的 Consumer/ResourceProbe 项目构建通过；双 TFM 单元/集成、Docs 和 API/示例验证通过，失败原因有记录。
- `git diff --check` 通过；实施差异只包含 XML 注释、必要邻近格式和计划执行记录，不修改测试行为、成员签名、公开 API、项目配置、快照或资源。
- 每批记录日期、文件清单、注释/inheritdoc 增量、复扫统计、待确认项和验证结果；第 7 批完成后停止，不自动开启下一轮。

## 11. 执行记录

### 本轮计划状态（2026-09-18 / Skill 1.2.1）

当前状态：**第 1-7 批已完成**。

| 批次 | 状态 | 备注 |
| --- | --- | --- |
| 1. 接口、抽象类和公共契约 | 已完成 | 已补齐公共测试入口、外部 Profile、Consumer 示例和现有契约标签；未处理后续批次的模型、实现成员、private helper 或字段 |
| 2. DTO、Entity、Options 和枚举 | 已完成 | 已补齐测试类型、Workbook/Row/Result/Converter/Validation 夹具及枚举成员；构造函数、属性、方法、字段和实现成员留待后续批次 |
| 3. 构造函数和属性文案统一 | 已完成 | 已统一全部构造函数和属性摘要，并补齐属性实现/接口属性的继承注释；方法、字段和常量留待后续批次 |
| 4. 实现类、重写成员和重载方法 | 已完成 | 已补齐 override 与已确认隐式接口实现的继承注释，并完成本批复扫和构建验证 |
| 5. private、internal、static 辅助方法 | 已完成 | 已补齐测试范围内的方法 XML 注释并完成复扫、构建和差异验证 |
| 6. private 字段、其他字段、常量、缓存键和配置键 | 已完成 | 已补齐 92 个字段和 12 个常量摘要，并完成复扫、构建和差异验证 |
| 7. 最终审计和构建验证 | 已完成 | 已完成全量复扫、双 TFM 测试、构建、XML 解析和差异检查 |

### 只读扫描记录（2026-09-18）

- **扫描范围**：当前 `tests/**` 下 73 个 C# 文件、2,643 个声明；排除项未命中现存候选文件；无 C# 文件修改。
- **统计结果**：缺失 1,724；单行/非三行 18；已有 `inheritdoc` 100；参数、泛型参数、返回标签异常分别为 647、50、327；构造函数异常 41；属性前缀/缺失异常 774；override 缺少 inheritdoc 175；XML 解析错误和非法单数 remark 均为 0。
- **工作区约束**：当前仓库存在既有测试文件修改、删除和未跟踪测试项目；统计和后续批次必须以当前磁盘基线为准，不能恢复、删除或格式化无关文件。
- **待确认项**：无须在规划阶段确认的业务语义；执行时需逐项确认测试替身的上游契约、Docs/Consumer 字符串代码片段边界以及 ResourceProbe 阈值的实际单位。

### 第 1 批执行记录（2026-09-18）

- **实际修改文件（8 个）**：
  - `tests/Bing.Offices.ProfileFixtures/ExternalMappingProfile.cs`
  - `tests/Bing.Offices.Consumer.Net6/Program.cs`
  - `tests/Bing.Offices.Consumer.Net8/Program.cs`
  - `tests/Bing.Offices.Docs.Tests/{DocsConsumerTest,DocsExamples}.cs`
  - `tests/Bing.Offices.Tests.Integration/{MiniExcelProviderContractTest,CsvAsyncFileIntegrationTest}.cs`
  - `tests/Bing.Offices.Tests/ApiSnapshotCandidateIdentityTest.cs`
- **注释增量**：新增或统一 26 个三行 `<summary>`，新增 3 个 `<param>`，新增 1 个 `inheritdoc` 和 1 个描述 Profile 额外映射行为的 `<remarks>`。补齐外部 Profile、Consumer 公开模型、API 快照测试及 Provider 合同入口；没有改动快照、断言、成员签名、项目配置或测试逻辑。`MiniExcelProviderContractTest.cs` 在本批开始前已是未跟踪文件，保留其既有工作区状态并仅补充公开合同测试摘要。
- **继承与文案处理**：`ExternalMappingProfile.Configure` 使用上游 `IMappingProfile<TImport,TExport>.Configure` 的 `inheritdoc`，仅追加输入/输出标题映射差异；没有复制接口的 summary、param 或 returns。Consumer 属性按访问方式使用“获取或设置”或“获取”，测试方法摘要只保留合同验证核心用途。Docs 的嵌套模型、Profile 实现、private helper、NPOI 替身成员和字段缺口按第 2、4、5、6 批保留。
- **范围复扫**：本批涉及的公共测试类和公开测试入口均有摘要；Consumer 与 Profile 文件缺失摘要、单行摘要、XML 错误均为 0。全量测试扫描由 1,724 个缺失声明降为 1,703 个，参数标签异常由 647 个降为 643 个（其中 3 个为新增 `<param>`，1 个由 `inheritdoc` 继承上游标签）；剩余缺口均落在后续批次范围（DTO/枚举、实现成员、private/internal/static helper、字段/常量）或合法的嵌套测试模型范围。`DocsConsumerTest.WorkbookMetadata_ExternalConsumer_ShouldRoundTripAllFields`、`CsvAsyncFileIntegrationTest.ExportToFileAsync_ImportFromFileAsync_ShouldRoundTripUtf8AndReplaceTarget` 和 API 快照哈希测试的参数标签已与签名一致。
- **构建与静态验证**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；Consumer net6.0/net8.0 项目使用 `-p:BingOfficesPackageVersion=2.0.0` 构建通过，net6.0 仅有 `NETSDK1138` 生命周期警告；`git diff --check` 退出码为 0；Roslyn 临时扫描语法错误为 0。工作区中 `CsvAsyncFileIntegrationTest.cs` 原有的 Provider 注册代码差异和多个既有测试文件改动均未回滚，本批自身新增内容仅为 XML 注释及既有 BOM 保留。
- **待确认项**：无本批必须确认的业务语义。Consumer/Profile 的公开模型已完成；Docs 与 Provider 测试替身的 DTO、接口实现/override、构造函数、private helper、字段和常量仍按后续批次执行。本轮完成后停止，不自动执行第 2 批。

### 第 2 批执行记录（2026-09-18）

- **实际修改文件（44 个 C# 文件）**：
  - `tests/Bing.Offices.Tests/Models/Bugs/Issue{1..9}.cs`、`tests/Bing.Offices.Tests/Models/ExportTestDataAnnotations.cs`；
  - `tests/Bing.Offices.Docs.Tests/{DocsConsumerTest,DocsExamples}.cs`；
  - `tests/Bing.Offices.MiniExcel.Tests/{MiniExcelProviderTest}.cs`、`tests/Bing.Offices.MiniExcel.Tests.Integration/MiniExcelProviderIntegrationTest.cs`；
  - `tests/Bing.Offices.Npoi.Tests.Integration/{ExcelAsyncFileIntegrationTest,ExcelImporterIntegrationTest}.cs`；
  - `tests/Bing.Offices.Npoi.Tests/{AsyncPipelineTest,ExcelMappingConfigurationLoaderTest,ExcelP0RegressionTest,ExcelWorkbookRequestTest,NpoiCellExtensionsTest,NpoiProviderBoundaryRegressionTest,NpoiServiceCollectionExtensionsTest,NpoiSheetPictureExtensionsTest,ReviewFixRegressionTest,StreamPipelineTest,TemplateAsyncBoundaryTest}.cs`；
  - `tests/Bing.Offices.ResourceProbe/{Program,StagingEntrypointMatrix,StagingResourceMatrix}.cs`；
  - `tests/Bing.Offices.Tests.Integration/{CsvAsyncFileIntegrationTest,MiniExcelProviderContractTest,PublicApiContractTest}.cs`；
  - `tests/Bing.Offices.Tests/{AsyncPipelineCsvTest,CsvStreamExtensionsTest,CsvTest,ExcelMappingConfigurationLoaderTest,ExcelMappingPlanCacheKeyTest,ExcelStreamExtensionsTest,MappingConfigurationPatchTest,MappingProfileRegistryTest,MappingProfileServiceCollectionExtensionsTest,MappingProfileV2Test,ReviewFixCoreRegressionTest}.cs`。
- **注释增量**：新增 295 个类型/枚举成员三行 `<summary>`，其中 275 个类型摘要和 20 个枚举成员摘要；另将 5 个已有类型摘要统一为三行结构。类型/枚举补全依据测试夹具真实职责编写，覆盖模型、Workbook/Row、转换器、校验器、结果、流/文件系统替身和 ResourceProbe 矩阵分支；未修改测试数据、枚举值、断言或成员签名。
- **继承与文案处理**：本批只处理类型和枚举成员，不新增 `inheritdoc`，不复制接口或基类成员文案；摘要保持一句话核心职责，枚举成员说明具体测试分支。未将构造函数、属性、方法、override、private helper、字段和常量提前纳入本批。
- **范围复扫**：418 个类型和 26 个枚举成员的缺失摘要均为 0；类型摘要单行/非三行均为 0；Roslyn 语法错误为 0。全量剩余缺口为 1,408 个，集中在 39 个构造函数、716 个方法、549 个属性、92 个字段和 12 个常量，留待第 3-6 批；当前已有 `inheritdoc` 为 101 个。
- **构建与静态验证**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；Consumer net6.0/net8.0 直接构建通过，net6.0 仅有已知 `NETSDK1138` 生命周期警告；ResourceProbe Release 直接构建通过，0 警告、0 错误；测试目录现有 1 个 XML 文件按 UTF-8 解析通过；`git diff --check` 通过，仅报告既有 XML 换行提示。
- **差异与待确认项**：本批通过 `apply_patch` 仅添加或调整 XML 注释；工作区中既有业务、测试架构、删除项和未跟踪项目差异均保留，未被本批回滚或格式化。无本批必须确认的业务语义；实现成员的 `inheritdoc`、构造函数/属性文案、private 方法和字段/常量仍按原计划待后续批次处理。本批完成后停止，不自动执行第 3 批。

### 第 3 批执行记录（2026-09-18）

- **实际修改文件（55 个 C# 文件）**：`tests/Bing.Offices.Tests/Models/**/*.cs` 下 21 个模型文件；以及 `tests/Bing.Offices.Docs.Tests`、`tests/Bing.Offices.MiniExcel.Tests`、`tests/Bing.Offices.MiniExcel.Tests.Integration`、`tests/Bing.Offices.Npoi.Tests`、`tests/Bing.Offices.Npoi.Tests.Integration`、`tests/Bing.Offices.Tests`、`tests/Bing.Offices.Tests.Integration` 和 `tests/Bing.Offices.ResourceProbe` 中涉及构造函数/属性的 34 个文件。重点覆盖 `MiniExcelProviderTest.cs`、`ExcelWorkbookRequestTest.cs`、`ExcelP0RegressionTest.cs`、`AsyncPipelineTest.cs`、`StreamPipelineTest.cs`、`AsyncPipelineCsvTest.cs`、`StagingResourceMatrix.cs`、Docs 示例和 Consumer/模型夹具。
- **注释增量**：48 个构造函数目标全部符合固定句式，其中补齐 39 个缺失摘要、修正 2 个异常句式/cref，其余 7 个已有文案完成核对；774 个属性访问文案问题清零，其中补齐 549 个缺失摘要、统一 225 个已有摘要的访问前缀或核心表述。属性实现/接口属性另将 89 个独立摘要改为 `/// <inheritdoc />`（57 个 `override` 属性、32 个 `INamedExcelValueConverter`/`INamedExcelValidationRule` 属性），没有为属性访问器生成方法注释。
- **继承与文案处理**：已读取并确认 `INamedExcelValueConverter.Name`、`INamedExcelValidationRule.Name/ErrorMessage`、`FilterAttributeBase.ErrorMsg` 和 `Stream`/`Assembly` 上游契约；实现属性不再复制上游 summary。属性摘要按访问方式使用“获取”“获取或设置”等简洁句式，构造函数统一使用 `初始化一个 <see cref="当前类型" /> 类型的实例。`；未修改成员体、特性、签名、访问级别、命名空间或公开 API。
- **范围复扫**：48 个构造函数和 818 个属性的缺失摘要均为 0；构造函数异常、属性前缀异常、属性 `override`/显式接口继承遗漏、非三行摘要、超长摘要和 XML 错误均为 0。全量剩余缺口为 820 个，集中在 716 个方法、92 个字段和 12 个常量；418 个类型、26 个枚举成员、48 个构造函数和 818 个属性已完成。测试源码现有 `inheritdoc` 为 190 个，较第 2 批增加 89 个；Roslyn 语法错误为 0。
- **构建与静态验证**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；Consumer net8.0 直接构建通过，Consumer net6.0 通过并仅报告已知 `NETSDK1138` 生命周期警告；ResourceProbe Release 直接构建通过，0 警告、0 错误；`git diff --check` 退出码为 0，仅报告工作区既有 Profile XML 的 CRLF 转换提示。当前批次源码差异仅为 XML 注释及必要的注释结构调整，工作区中既有业务、测试架构、删除项和未跟踪项目差异均保留。
- **待确认项**：无本批必须确认的业务语义。第 4 批仍需处理方法级实现/重写、重载和其余接口成员的 `inheritdoc` 与标签；private/internal/static 方法、字段和常量按原计划留待后续批次。本批完成后停止，不自动执行第 4 批。

### 第 4 批执行记录（2026-09-18）

- **实际修改文件（18 个 C# 文件）**：
  - `tests/Bing.Offices.MiniExcel.Tests/MiniExcelProviderTest.cs`；
  - `tests/Bing.Offices.Npoi.Tests/{AsyncPipelineTest,ExcelP0RegressionTest,ExcelWorkbookRequestTest,NpoiProviderBoundaryRegressionTest,NpoiServiceCollectionExtensionsTest,NpoiSheetPictureExtensionsTest,ReviewFixRegressionTest,StreamPipelineTest,TemplateAsyncBoundaryTest}.cs`；
  - `tests/Bing.Offices.Tests.Integration/MiniExcelProviderContractTest.cs`；
  - `tests/Bing.Offices.Tests/{CsvStreamExtensionsTest,CsvTest,ExcelMappingConfigurationLoaderTest,ExcelStreamExtensionsTest,MappingProfileRegistryTest,MappingProfileServiceCollectionExtensionsTest,ReviewFixCoreRegressionTest}.cs`。
- **注释增量**：第 3 批已完成的 175 个 override 候选中，本批补齐剩余 118 个方法级 `/// <inheritdoc />`；隐式接口实现扫描基线中的 168 个主线缺失成员全部补齐，`AsyncPipelineTest.cs` 另有 12 个隐式实现成员由协作单元补齐，隐式实现缺口复扫为 0。`AsyncPipelineTest.cs` 中取消、同步 IO 拒绝、staging 失败和清理行为保留为必要 `<remarks>`，没有复制上游 summary、param 或 returns。当前全量测试审计包含 488 个 `inheritdoc`，其中方法 358 个、属性 130 个。
- **继承与重载处理**：复核 `Stream`、Excel/CSV Provider、转换器、验证器、映射 Profile、映射计划工厂、图片突变适配器、异常观察器和文件系统契约；重写成员、隐式接口实现和同语义重载均使用 `inheritdoc`，未发现实现成员重复保留上游 summary/param/returns。实现的取消、异常和资源生命周期差异仅保留在必要 remarks 中。
- **范围复扫**：`override` 缺失为 0，显式接口实现缺失为 0，隐式接口实现扫描剩余为 0；构造函数 48、属性 818、类型 418、枚举成员 26 的缺口仍为 0。全量剩余缺口为 522 个，其中普通方法 418 个、字段 92 个、常量 12 个，均属于第 5、6 批范围；XML 语法错误为 0。当前重载实现未发现 summary 不一致或应下沉而未下沉的新增说明。
- **构建与静态验证**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；Consumer net8.0 直接构建通过，Consumer net6.0 通过并仅报告已知 `NETSDK1138` 生命周期警告；ResourceProbe Release 直接构建通过，0 警告、0 错误；tests 下现有 XML 按 UTF-8 解析通过；`git diff --check` 退出码为 0，仅报告既有 Profile XML 的 CRLF/LF 换行提示。
- **差异与待确认项**：本批通过 `apply_patch` 仅新增 XML 注释和必要的 `<remarks>`；未修改业务逻辑、成员签名、访问级别、命名空间、公开 API、项目配置、测试数据或断言，工作区既有修改、删除和未跟踪文件均保留。无本批必须确认的业务语义；普通 private/internal/static 方法、字段和常量仍按第 5、6 批处理。本批完成后停止，不自动执行第 5 批。

### 第 5 批执行记录（2026-09-18）

- **实际修改文件（38 个 C# 文件）**：
  - `tests/Bing.Offices.Docs.Tests/{DocsConsumerTest,DocsExamples}.cs`；
  - `tests/Bing.Offices.MiniExcel.Tests/MiniExcelProviderTest.cs`、`tests/Bing.Offices.MiniExcel.Tests.Integration/MiniExcelProviderIntegrationTest.cs`；
  - `tests/Bing.Offices.Npoi.Tests.Integration/{ExcelAsyncFileIntegrationTest,ExcelImporterIntegrationTest}.cs`；
  - `tests/Bing.Offices.Npoi.Tests/{AsyncPipelineTest,ExcelMappingConfigurationLoaderTest,ExcelP0RegressionTest,ExcelWorkbookRequestTest,NpoiCellExtensionsTest,NpoiCellStyleExtensionsTest,NpoiFailureWorkbookPreflightTest,NpoiFontExtensionsTest,NpoiProviderBoundaryRegressionTest,NpoiRowExtensionsTest,NpoiSheetExtensionsTest,NpoiSheetPictureExtensionsTest,NpoiWorkbookExtensionsTest,NpoiXlsxZipPreflightTest,PublicNpoiExtensionCoverageTest,StreamPipelineTest,TemplateAsyncBoundaryTest}.cs`；
  - `tests/Bing.Offices.ResourceProbe/{StagingEntrypointMatrix,StagingResourceMatrix}.cs`；
  - `tests/Bing.Offices.Tests.Integration/{CsvAsyncFileIntegrationTest,MiniExcelProviderContractTest,PublicApiContractTest}.cs`；
  - `tests/Bing.Offices.Tests/{AsyncPipelineCsvTest,CsvStreamExtensionsTest,DefaultFileExportCommitterTest,ExcelDateParserTest,ExcelMappingPlanCacheKeyTest,ExcelStreamExtensionsTest,MappingProfileRegistryTest,MappingProfileServiceCollectionExtensionsTest,PublicCoreExtensionCoverageTest,ReviewFixCoreRegressionTest}.cs`。
- **注释增量**：补齐 418 个方法 XML 注释（主线 415 个、协作单元 3 个），覆盖 public、internal、private、static、异步、泛型和测试替身辅助方法；同步修正生成文案中的 XML 泛型转义、机械化参数说明和明显的中英文拼接。方法摘要保持简洁，`Task`/`ValueTask` 返回标签按实际最终返回类型处理；没有修改方法体、成员签名、访问级别或公开 API。
- **继承与文案处理**：本批没有新增或删除 `inheritdoc`；全量 Roslyn 审计识别的 `inheritdoc` 保持 488 个，override、显式接口实现和隐式接口实现缺口均为 0。新增方法没有复制上游契约文案，也没有把测试条件扩写进 summary；无需要追加 remarks 的新方法说明。
- **范围复扫**：方法缺失摘要为 0；构造函数、属性、类型和枚举成员缺失仍为 0；字段缺口 92 个、常量缺口 12 个，全部保留给第 6 批。XML/语法错误为 0，测试目录 XML 文件按 UTF-8 解析通过；本批新增方法未产生新的单行或非三行 summary。隐式接口扫描无遗漏。
- **构建与静态验证**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；Consumer net6.0/net8.0（含 `-p:BingOfficesPackageVersion=2.0.0`）和 ResourceProbe Release 直接构建通过，net6.0 仅报告既有 `NETSDK1138` 生命周期警告；`git diff --check` 退出码为 0，仅报告既有 `tests/Bing.Offices.ProfileFixtures/Bing.Offices.ProfileFixtures.xml` 的 CRLF/LF 换行提示。
- **差异与待确认项**：本批通过 `apply_patch` 只新增或修正 XML 注释及必要的注释结构，没有业务逻辑、测试断言、测试数据、项目配置或文件生成行为变更。工作区中既有修改、删除和未跟踪测试项目均保留。无本批必须确认的业务语义；92 个字段和 12 个常量/缓存键/配置键仍按第 6 批处理。本批完成后停止，不自动执行第 6 批。

### 第 6 批执行记录（2026-09-18）

- **实际修改文件（21 个 C# 文件）**：
  - `tests/Bing.Offices.MiniExcel.Tests/{MiniExcelProviderTest}.cs`、`tests/Bing.Offices.MiniExcel.Tests.Integration/MiniExcelProviderIntegrationTest.cs`；
  - `tests/Bing.Offices.Npoi.Tests.Integration/ExcelImporterIntegrationTest.cs`；
  - `tests/Bing.Offices.Npoi.Tests/{AsyncPipelineTest,ExcelP0RegressionTest,ExcelWorkbookRequestTest,NpoiCellExtensionsTest,NpoiSheetPictureExtensionsTest,PublicNpoiExtensionCoverageTest,ReviewFixRegressionTest,StreamPipelineTest,TemplateAsyncBoundaryTest}.cs`；
  - `tests/Bing.Offices.ResourceProbe/{StagingEntrypointMatrix,StagingResourceMatrix}.cs`；
  - `tests/Bing.Offices.Tests.Integration/{CsvAsyncFileIntegrationTest,PublicApiContractTest}.cs`；
  - `tests/Bing.Offices.Tests/{AsyncPipelineCsvTest,CsvTest,MappingProfileRegistryTest,MappingProfileServiceCollectionExtensionsTest,PublicCoreExtensionCoverageTest}.cs`。
- **注释增量**：补齐 92 个字段和 12 个常量的三行中文 `<summary>`；覆盖异步流缓冲、取消信号、异常注入、反射操作码字典、映射记录、ResourceProbe 任务标识、字节阈值、并发计数、峰值指标和采样状态。字段和常量摘要均直接说明实际用途，未使用“获取或设置”，未修改常量值、测试数据、缓存行为或资源清理逻辑；新增 `inheritdoc` 为 0。
- **范围复扫**：Roslyn 审计的 98 个字段和 12 个常量缺失均为 0，单行/非三行字段和常量均为 0；私有字段没有遗漏，XML 与 C# 语法错误均为 0。隐式接口实现扫描无输出，`inheritdoc` 总数保持 488。字段/常量摘要均为简短单句，无需下沉到 `<remarks>`；本批没有修改方法，故未引入重载 summary 不一致。全范围尚有 7 个既存方法单行摘要，留待第 7 批最终审计。
- **构建与静态验证**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；Consumer net8.0 和 ResourceProbe Release 直接构建通过，均为 0 警告、0 错误；Consumer net6.0 构建通过，只有已知 `NETSDK1138` 生命周期警告。tests 下现有 1 个 XML 文件按 UTF-8 解析通过；`git diff --check` 退出码为 0，仅报告既有 `tests/Bing.Offices.ProfileFixtures/Bing.Offices.ProfileFixtures.xml` 的 CRLF/LF 换行提示。
- **差异与待确认项**：本批新增差异仅为 XML 注释和本执行记录；当前工作区存在既有业务/测试架构修改、删除和未跟踪测试项目，均已保留且未作为本批变更。无本批必须确认的业务语义；第 7 批仍需完成全量文案、标签、继承与构建终审。本批完成后停止，不自动执行第 7 批。

### 第 7 批执行记录（2026-09-18）

- **实际修改文件（34 个 C# 文件）**：覆盖 Docs、MiniExcel、NPOI 单元/集成测试、ResourceProbe、Core 测试和公共 API 合同测试；具体文件以本批起点基线比较为准。修改内容仅包括单行摘要三行化、构造函数/方法参数与泛型参数标签、返回值语义、重载摘要和必要的 `remarks`。
- **注释增量**：全量审计目标 2,643 个声明；补齐并校正本批终审发现的 7 个单行摘要、参数/泛型参数/返回标签以及过于笼统的集合、布尔和异步结果说明；新增 `inheritdoc` 为 0，最终 `inheritdoc` 数量保持 488。NPOI 子集另修正 2 个单行摘要和流 staging 开始信号说明。
- **全量复扫**：73 个测试文件、2,643 个声明通过 Roslyn 扫描；缺失、单行/非三行摘要、参数/泛型参数/返回标签异常、构造函数异常、属性访问前缀异常、override/显式接口继承遗漏和 XML/C# 语法错误均为 0。字段 98、常量 12、方法 1,223、属性 818、构造函数 48、类型 418、枚举成员 26 均无缺失。隐式接口扫描无输出。
- **文案与继承复核**：方法 summary 仅保留核心用途；集合、布尔、可空和异步返回值按实现说明；冗长测试条件保留在 `remarks` 或参数标签；重载核心摘要一致；实现成员未复制上游 summary、param 或 returns，既有 `inheritdoc` 保持不变。
- **构建与测试**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；Consumer net8.0 和 ResourceProbe Release 通过且 0 警告/错误；Consumer net6.0 通过，仅有已知 `NETSDK1138` 生命周期警告。解决方案测试通过：Core 单元测试双 TFM 各 196/196，MiniExcel 单元双 TFM 各 39/39，Core 集成双 TFM 各 29/29，NPOI 单元双 TFM 各 539/539，Docs net8.0 为 10/10；NPOI 集成双 TFM 各 30/30。tests 下 XML 文件按严格 UTF-8 解析通过；`git diff --check` 退出码为 0，仅报告既有 ProfileFixtures XML 的 CRLF/LF 提示。
- **差异与待确认项**：相对第 7 批起点基线的 34 个 C# 文件差异均为 XML 注释及必要邻近格式，未修改业务逻辑、签名、访问级别、命名空间、公开 API、项目配置、测试数据或断言；工作区原有修改、删除和未跟踪项目均保留。无待确认业务语义。本轮第 1-7 批全部完成，按计划停止。

---

# 历史记录：src/* 中文 XML 注释补全计划（2026-09-18）

## 1. 目标和范围

扫描日期：2026-09-18。规范依据：`C:/Users/jianx/.agents/skills/chinese-comments/SKILL.md`，版本 1.2.1。

本轮重新扫描当前工作区的 `src/*`，形成分批实施基线；实施按批次逐项推进，每次只完成一个未完成批次并在验证后停止。旧版计划的 2026-09-15 执行记录保留在第 11 节作为历史，不代表本轮新增范围或新版格式要求已经完成。

| 范围 | 候选文件 | 实际职责 / 目标框架 |
| --- | ---: | --- |
| `src/Bing.Offices.Abstractions` | 86 | 公共契约、请求/结果、配置、枚举、Builder；netstandard2.0 |
| `src/Bing.Offices.Core` | 58 | CSV、映射、配置、校验、文件提交、元数据；netstandard2.0 |
| `src/Bing.Offices.MiniExcel` | 8 | MiniExcel XLSX 导入导出、值适配、计划调度、日期 serial 读取；net6.0 / net8.0 |
| `src/Bing.Offices.Npoi` | 45 | NPOI XLS/XLSX、DOM 扩展、失败工作簿、staging；net6.0 / net8.0 |
| `src/provider-shared` | 1 | 由两个 Provider 链接编译的项目自有 XLSX ZIP/XML 预检源码 |
| **合计** | **198** | 四个项目和一个共享源码目录；共享文件按物理路径只计一次 |

扫描包括类型（含嵌套类型、委托）、构造函数、方法、扩展方法、属性、索引器、事件、所有访问级别字段、常量及枚举成员。字段多变量声明按变量计数；局部函数不强制 XML 文档，属性访问器不重复计为方法。四个 AssemblyInfo.cs 纳入文件清单，但无目标成员。

本轮统计以当前未提交工作区为准。仓库已有业务代码、测试架构和项目配置改动，均属于原有工作，实施时必须保留，并以批次开始时的内容为差异基线。

## 2. 排除范围

- `bin/**`、`obj/**`。
- `*.g.cs`、`*.generated.cs`、`*.Designer.cs`；实际命中 `src/Bing.Offices.Core/Bing/Offices/Resources.Designer.cs`。
- EF Core Migration/Migrations、ModelSnapshot、标记自动生成的客户端/代理、第三方源码。本轮非构建目录未发现其他此类文件。
- `tests/**`、`benchmarks/**`、`build/**` 和其他源码目录不纳入注释修改；可只读参考测试和构建配置。
- NuGet 的 NPOI、MiniExcel、CsvHelper 源码不纳入；本项目对它们的适配器和 `provider-shared` 不属于第三方源码。

构建目录之外共 199 个 C# 文件，排除 1 个 Designer 后纳入 198 个。
`src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs` 是未标记生成、位于仓库内的兼容桩：本轮列入扫描并单列 1 个缺失，不沿用历史的默认排除。实施前确认归属；若确属第三方来源，记录证据后排除，不能仅凭 System 命名空间判断。

## 3. 扫描统计

### 3.1 按目录统计

| 项目 / 目录 | 目标声明 | 缺少有效 summary/inheritdoc | 单行 summary | 非三行 summary（含单行） | inheritdoc |
| --- | ---: | ---: | ---: | ---: | ---: |
| Abstractions | 996 | 1 | 519 | 519 | 5 |
| Core | 654 | 0 | 334 | 334 | 124 |
| MiniExcel | 125 | 106 | 3 | 4 | 6 |
| Npoi | 620 | 1 | 322 | 322 | 57 |
| provider-shared | 9 | 7 | 1 | 1 | 0 |
| **合计** | **2,404** | **115** | **1,179** | **1,180** | **192** |

115 个缺失分布在 9 个文件；已有单行摘要分布在 134 个文件。另 1 个非三行摘要是 MiniExcel 导入器的 ImportTypedSheetRows：起止标签分行，但正文有两行，应将行号推进说明移入 remarks。

共有 2,097 个有效 summary 声明、192 个 inheritdoc 声明、115 个缺失声明。存在文档不等于内容或继承目标正确；本轮不能用旧版“全部覆盖”替代复扫。

### 3.2 按成员种类统计

| 对象 | 目标声明 | 缺失 | 单行 summary |
| --- | ---: | ---: | ---: |
| 类型（含委托） | 311 | 9 | 94 |
| 构造函数 | 118 | 4 | 114 |
| 方法 | 826 | 79 | 273 |
| 属性 | 720 | 4 | 395 |
| 索引器 | 0 | 0 | 0 |
| 事件 | 0 | 0 | 0 |
| 字段（不含常量） | 201 | 16 | 172 |
| 常量 | 45 | 3 | 32 |
| 枚举成员 | 183 | 0 | 99 |
| **合计** | **2,404** | **115** | **1,179** |

缺失方法：private 64、internal 5、public 10、protected 0；其中 static 67，static 数量与访问级别数量重叠，不重复相加。
197 个 private 字段中有 14 个缺失；另 2 个缺失字段为私有泛型嵌套缓存类中的 internal static readonly Values。16 个字段均在 MiniExcel。

### 3.3 标签、文案和继承检查

| 检查项 | 结果 | 解释 |
| --- | ---: | --- |
| param 与签名名称/数量/顺序不一致 | 85 个声明 | MiniExcel 77、Npoi 1、shared 7；主要是缺失标签，需同时检查错名、重复、多余标签 |
| typeparam 与签名不一致 | 23 个声明 | 均在 MiniExcel，含委托、嵌套泛型缓存和辅助方法 |
| returns 与类型规则不一致 | 55 个声明 | MiniExcel 52、Npoi 1、shared 2；包含非 void 返回缺口 |
| 构造函数 summary 不符合规定句式 | 6 / 118 | 4 个未注释、2 个已有摘要但句式错误；均在 MiniExcel |
| 已注释属性访问前缀不匹配 | 0 | 另有 4 个未注释只读属性，均在 MiniExcel |
| override / 显式接口成员未 inheritdoc | 0 / 0 | 语法候选检查；不等于隐式实现无问题 |
| 已确认隐式接口属性重复独立摘要 | 3 | Core 的 ExcelValidationBinding.Kind、IsRaw、ErrorMessage |
| 已注释同名重载组存在不同摘要 | 34 个候选组 | 仅字符串筛选，含不同泛型类型/不同职责；不得把全部认定为缺陷 |
| summary 纯文本超过 120 字符 | 0 | 短摘要仍可能包含应下沉的说明 |
| XML 文档不可解析 / 单数 remark 标签 | 0 / 0 | 文档块语法检查；cref 合法性留待编译验证 |
| 源码语法错误 | 0 | Roslyn 语法树检查，不代表本轮构建通过 |

标签检查对 inheritdoc 成员不要求重复本地标签；泛型类型和委托也纳入签名检查。上述数量互相重叠，不能相加为总问题数。

隐式接口补充检查使用 net8 条件和现有依赖建立合并源码符号模型，得到 123 个接口实现映射，发现 3 个未使用 inheritdoc 的候选并经源码对照确认。模型仍有 6 个诊断（跨程序集同名类型及依赖/全局上下文差异），因此只用于定位，不作为四个项目、双 TFM 继承审计的最终结论；最终逐项目验证，不把模型诊断认定为业务构建失败。

### 3.4 确定缺失文件清单

| 文件 | 缺失声明 | 主要对象 |
| --- | ---: | --- |
| `src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs` | 1 | 兼容桩类型，归属待确认 |
| `src/Bing.Offices.MiniExcel/Bing/Offices/Exports/MiniExcelExcelExporter.cs` | 14 | 依赖字段和导出辅助方法 |
| `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs` | 61 | 委托、泛型缓存、嵌套模型、异常、辅助方法、字段、属性 |
| `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelMappingPlanBuilder.cs` | 10 | 委托、构造函数、缓存字段、计划调度方法 |
| `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelRawDateSerialReader.cs` | 9 | XML 命名空间常量、日期 serial/坐标/路径辅助方法 |
| `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelValueAdapter.cs` | 11 | 日期规则字段、导入/导出/动态值转换方法 |
| `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelXlsxPreflight.cs` | 1 | Validate |
| `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiXlsxZipPreflight.cs` | 1 | Validate |
| `src/provider-shared/ExcelXlsxZipPreflight.cs` | 7 | Validate、ZIP/XML 安全和资源预检辅助方法 |

## 4. 问题分类

1. 确定缺失：115 个声明，以 MiniExcel 内部辅助成员为主；委托和私有嵌套缓存同样需要完整文档。
2. 新版格式缺口：1,180 个 summary 不符合独立起始、单行正文、独立结束的三行格式；应按成员所属批次处理，而非一次格式化全仓库。
3. 契约继承重复：3 个接口实现属性存在独立摘要，先完善上游，再替换为 inheritdoc。
4. DTO/配置文档：已注释模型访问前缀无明显缺口，但需统一三行格式、复核默认值/单位/null；ColumnBinding、MiniExcelRowException.Error 等新增内部载体仍未注释。
5. 类型职责失真：ExcelImportRuntime 误用导入器文案；MiniExcelXlsxPreflight 的“Core 共享实现”已经不符合当前共享源码的编译归属。
6. summary 承载过多说明：Provider 实现策略、流所有权、请求作用域、提交/缓存策略等，应移入 remarks，不能只依据字符长度判断。
7. 标签语义不足：结构完整不等于 true/false、null 条件、异步结果、参数格式/范围/单位明确。
8. 重载差异候选：保留相同核心用途，输入载体、依赖注入、范围筛选、格式和同步/异步差异由 param/returns/remarks 表达；真实职责不同的不机械统一。

## 5. 文案规范问题

### 5.1 统一规则

- 每个独立 summary 恰为三行：起始标签、核心用途正文、结束标签；inheritdoc 仍保持独立一行，不添加空 summary。
- 类型摘要只用一句话说明核心职责。额外适用范围、生命周期、策略、兼容边界放入合法的 remarks；不添加重复、空泛 remarks。
- 构造函数固定为：`初始化一个 <see cref="当前类型" /> 类型的实例。`，泛型 cref 必须合法。
- 属性按实际访问器使用“获取……”“设置……”“获取或设置……”“获取或初始化……”。表达式体属性按只读处理；不在中文词语内部人为插入空格。
- 字段直接说明实际用途，禁止“获取或设置”；缓存/键说明实际隔离维度、作用域和生命周期，不捏造过期或淘汰机制。
- 方法 summary 只说明做什么；重载保持相同或高度一致核心摘要，差异写入标签或 remarks。
- 非 void、非裸 Task/ValueTask 方法（含返回结果的委托）补 returns；bool 说明成功和失败，nullable 说明 null 条件；只记录真实异常。

### 5.2 具体复核项

| 位置 | 问题 / 处理建议 | 批次 |
| --- | --- | --- |
| `MiniExcel/Exports/MiniExcelExcelExporter.cs:30`、`MiniExcel/Imports/MiniExcelExcelImporter.cs:55` | 构造函数“初始化 MiniExcel ……”不符合固定句式 | 3 |
| `MiniExcel/Imports/MiniExcelExcelImporter.cs:306-308` | ImportTypedSheetRows 正文两行，行号推进解释应下沉 remarks | 5 |
| `Npoi/Imports/ExcelImportRuntime.cs:3-6` | 实际跟踪行数/图片资源预算，却写成“Excel 导入器”；重写类型核心职责 | 2 |
| `MiniExcel/Internals/MiniExcelXlsxPreflight.cs:5-9` | “Core 共享实现”过时；改为本地链接共享预检源码，并将来源放 remarks | 5 |
| `Npoi/Imports/NpoiXlsxZipPreflight.cs:5-8` | 同样残留“Core 共享实现”文案；按 provider-shared 链接编译来源修正 | 5 |
| `Npoi/Exports/NpoiExportPlanBuilder.cs:18` | CreatePlanInvoker 返回工作簿映射计划，但缺少 returns 标签 | 5 |
| `MiniExcel/Internals/MiniExcelMappingPlanBuilder.cs:9-13` | 区分本类反射委托缓存与 Core 工厂计划缓存，不宣称本类缓存完整计划 | 5 |
| `MiniExcel/Internals/MiniExcelRawDateSerialReader.cs:7-11` | 类型摘要混合读取顺序/XML 实现细节；保留日期 serial 读取职责 | 5 |
| `MiniExcel/Internals/MiniExcelRawDateSerialReader.cs:170` | CreateKey 是物理行列坐标编码，不应机械写成带生命周期的缓存键 | 5 |
| `Abstractions/Configurations/ExcelMappingDocument.cs:3-6`、`Exports/ExcelWorkbookMetadataOptions.cs:3-6` | 方向隔离、请求实例归属下沉 remarks | 2 |
| `Abstractions/IO/IFileExportCommitter.cs:5-8` | exporter 异常观察边界下沉 remarks | 1 |
| `Abstractions/Providers/UniqueTracker.cs:5-8` | pending journal/提交策略下沉 remarks | 4 |
| `Core/IO/DefaultFileExportCommitter.cs:4-6` | exporter 负责异常观察的分工下沉 remarks | 4 |
| `Core/Exceptions/BingOfficesExceptionDispatcher.cs:22-24,49-52` | Observe 的同实例幂等和辅助方法异常隔离说明下沉 remarks | 1 / 5 |
| `Npoi/Imports/ExcelImportErrorCollector.cs:3-6`、`Npoi/IO/NpoiAsyncStaging.cs:74-75` | 子收集器边界、内部流所有权下沉 remarks | 2 / 4 |
| `MiniExcel/Exports/MiniExcelExcelExporter.cs:20-22`、`MiniExcel/Imports/MiniExcelExcelImporter.cs:22-24` | 延迟序列、DOM、Core 分工属于实现策略，移入 remarks | 4 |
| `Npoi/Exports/NpoiExcelExporter.cs:12-15`、`Npoi/Imports/NpoiExcelImporter.cs:15-18` | 内存 DOM/输入复制策略下沉 remarks | 4 |
| `Core/Dates/ExcelDateParser.cs:36,168,202,217,278` | returns 只写 true；明确 false 条件和 out 值，按代码逐项确认 | 5 |
| `provider-shared/ExcelXlsxZipPreflight.cs:13-17`、`MiniExcel/Internals/MiniExcelXlsxPreflight.cs:15-19` | GetDate1904 的 false 含义、流位置恢复说明不足，读取约束写 param/remarks | 5 |

表内项目缩写都相对于 `src/Bing.Offices.*`；行号来自本次工作区，执行时可能移动。上述仅为具体语义发现，不代替对 198 个候选文件中其他文档的成员级审计。

重载重点：Core AtomicFileCommitter.Commit/CommitAsync 的默认与指定文件系统重载核心用途相同；Core 配置加载器文本/流/别名重载的流所有权和校验差异下沉；Npoi ExcelHelper.PrepareWorkbook、图片和合并区域扩展族统一核心措辞。MappingConfigurationCloner.Clone、ExcelTypeMapFactory.Get 等不同对象职责可保留差异；同名不同泛型类型不误认作重载。

## 6. inheritdoc 使用问题

已确认 `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelValidationBinding.cs:56-61` 的 Kind、IsRaw、ErrorMessage 实现 `IExcelValidationBinding` 对应只读属性，却重复书写独立摘要。上游位于 `src/Bing.Offices.Abstractions/Bing/Offices/Validations/IExcelValidationBinding.cs:8-15`，已有文档，但三行格式及 IsRaw 的 true/false 和失败消息语义需先完善。第 4 批改用 `/// <inheritdoc />`，只追加确有必要的差异说明。

MiniExcel 的 4 个导出入口、2 个导入入口共 6 个 inheritdoc 对应 IExcelExporter/IExcelImporter，已读接口和实现，未发现复制上游摘要或无目标问题。

其余 186 个既有 inheritdoc，以及隐式接口、显式接口、抽象实现、override，仍需在第 4 / 7 批按项目和真实上游契约逐项核对：

| 上游契约 | 实现 / 复核范围 |
| --- | --- |
| IExcelExporter / IExcelImporter | Npoi 和 MiniExcel 的 Exports/Imports 入口 |
| ICsvExporter / ICsvImporter | Core/Csv/CsvEntityExporter.cs、CsvEntityImporter.cs |
| IFileExportCommitter、IAtomicFileSystem | Core/IO/DefaultFileExportCommitter.cs、AtomicFileCommitter.cs |
| IExcelMappingPlan*.cs、IExcelCompiledMappingColumn | Core/Mappings/*.cs |
| IExcelValidationBinding、IExcelValidationRule、INamedExcelValidationRule | Core/Mappings/ExcelValidationBinding.cs、Core/Validations/ExcelValidationRules.cs |
| INpoiAsyncStagingFactory / INpoiAsyncStaging | Npoi/IO/NpoiAsyncStaging.cs |
| Stream / IDisposable、内部文件/图片适配接口 | Core/Csv/CsvPipelineSupport.cs、Npoi/Exports/NpoiNonDisposingStream.cs、Npoi/Imports/NpoiFailureWorkbook*.cs、Npoi/Extensions/SheetExtensions.Picture.cs |

接口实现不复制 summary/param/returns；第三方基类的实现同样优先 inheritdoc，不因本地未安装文档而认定无有效目标。不能用 inheritdoc 覆盖普通辅助方法、字段或构造函数。上游缺失先补上游，再补实现；额外流所有权、取消、异常转换、Provider 限制只写差异 remarks/必要 exception。

## 7. 实施批次

本轮按 7 批推进。一次执行只处理第一个未完成批次，验证并更新记录后停止，不自动推进下一批。目录授权仅限第 1 节的 198 个非排除文件；同一文件按当前批次的成员对象分工处理，不因打开文件而处理其全部内容。

### 第 1 批：接口、抽象类和公共契约

范围：

- Abstractions/Bing/Offices/{Csv,Exports,Imports,IO,Providers,Conversions,Validations}/I*.cs。
- Abstractions/Bing/Offices/Configurations/{IExcelMappingConfigurationLoader,IMappingProfileRegistry,IMappingProfileResolver,MappingProfileContracts}.cs。
- Abstractions/Bing/Offices/Attributes/{FilterAttributeBase,DecoratorAttributeBase}.cs。
- Abstractions 的 ExcelImport.cs、ExcelExport.cs、ExcelSheetExportBuilder.cs、ImportMappingBuilder.cs、ExportMappingBuilder.cs、FluentSetting.cs、ExcelModelAliasRegistry.cs：公共方法/泛型标签及类型公共契约。
- Core/Mappings/IExcelCompiledMappingColumn.cs；Core 的 Configurations/ExcelMappingConfigurationLoader.cs、Exceptions/BingOfficesExceptionDispatcher.cs、Extensions/*.cs：无继承目标的公共方法契约。
- Npoi/IO/NpoiAsyncStaging.cs 内部接口、Npoi/Imports/NpoiFailureWorkbookFileSystem.cs 内部接口、Npoi/Extensions/SheetExtensions.Picture.cs 的 IPictureMutationAdapter；Core/IO/AtomicFileCommitter.cs 的 IAtomicFileSystem。
- Npoi/Extensions/*.cs、ExcelHelper.cs：无可继承契约的公共静态入口（与第 4 批重载审计衔接）。

清单：接口/抽象成员及公共新方法三行摘要；参数/泛型/返回语义；上游属性/字段若属于契约本身，补必要文档；精简契约类型摘要。构造函数留第 3 批、实现留第 4 批、类状态字段留第 6 批。

完成条件：可被实现继承的上游文档有效，提供上游到实现对照，公开入口未出现空泛/错误标签。

### 第 2 批：DTO、Entity、Options 和枚举

范围：

- Abstractions/Bing/Offices/{Configurations,Csv,Exports,Imports,Styles,Conversions}/ 的 DTO/请求/结果/Options/枚举/配置载体。
- Abstractions/Validations/ExcelValidationContext.cs 和 IExcelValidationBinding.cs 中的枚举。
- Core/Metadata/*.cs、Mappings/{ExcelPropertyMap,ExcelTypeMap,ExcelMappingStyleAndLayout}.cs 的载体类型。
- Core/Dates/ExcelDateParser.cs 中 ExcelDateOffsetPolicy 枚举。
- Npoi/ExcelColumnPlan.cs、Imports/{ExcelImportExecutionOptions,ExcelImportRuntime,ExcelImportErrorCollector,ValidationRangeIndex}.cs 中的数据/状态载体。
- Npoi/IO/NpoiAsyncStaging.cs 的内部策略枚举。
- MiniExcel/Imports/MiniExcelExcelImporter.cs 中 ColumnBinding 和嵌套异常的类型职责/可读载体属性。

清单：模型类型及全部 183 个枚举成员文档/三行摘要；默认值、单位、null 和布尔语义；修复 ExcelImportRuntime 类型职责。本次未发现独立 Entity 目录，不虚构 Entity 实施文件。

完成条件：DTO/Options/枚举说明准确；构造函数、全局访问措辞仍归第 3 批，字段归第 6 批。

### 第 3 批：构造函数和属性文案统一

范围：第 1 节四项目的全部构造函数、属性（包含嵌套类型），按访问器复核；排除生成文件，不改声明。

清单：118 个构造函数统一固定句式及三行格式；MiniExcel 4 个缺失构造函数补签名标签，2 个现有错误句式修正；720 个属性复核访问措辞、三行格式，MiniExcel ColumnBinding.Column/Property/Header 和 MiniExcelRowException.Error 共 4 个只读属性补全文档。能继承的属性用 inheritdoc，不重复独立摘要；保持与第 1 批上游一致。

完成条件：句式/访问前缀缺口清零、init/表达式体分类正确、属性摘要简洁。不给字段套属性措辞。

### 第 4 批：实现类、重写成员和重载方法

范围：Core/{Csv,IO,Mappings,Validations,Configurations} 的实现；Npoi/{Imports,NpoiMappingPlanFactoryResolver.cs,Exports,IO,Extensions,ExcelHelper.cs}；MiniExcel/{Imports,Exports}；Abstractions/Providers/UniqueTracker.cs 及 Configurations/Exports/Imports 的重载族。

清单：复核 192 个 inheritdoc、所有接口实现/显式实现/抽象实现/override；修正 ExcelValidationBinding 三个属性；类型核心职责与 Provider/Stream 差异拆分 summary/remarks；第 5.2 节重载族逐项确认，不按字符串机械统一。

完成条件：上游可继承，无重复契约、无无效继承目标；同职责重载核心摘要一致。需要独立摘要的实现类型/方法按三行格式；普通辅助方法归第 5 批。

### 第 5 批：private、internal、static 辅助方法

优先补缺失的确切文件：

- MiniExcel/Exports/MiniExcelExcelExporter.cs 的 CreateWorkbookRows、EnumerateRows、配置/能力校验及模板清理。
- MiniExcel/Imports/MiniExcelExcelImporter.cs 的缓冲导入、泛型反射委托、物化、绑定、关系、错误收集、同步/异步流复制等辅助方法。
- MiniExcel/Internals/{MiniExcelMappingPlanBuilder,MiniExcelRawDateSerialReader,MiniExcelValueAdapter,MiniExcelXlsxPreflight}.cs。
- Npoi/Imports/NpoiXlsxZipPreflight.cs。
- provider-shared/ExcelXlsxZipPreflight.cs。

其余复核范围：Core/{Configurations,Csv,Dates,Exceptions,Extensions,IO,Mappings,Validations,Internals} 和 Npoi/{Imports,Exports,IO,Extensions,Resolvers,Internals} 的 private/internal/protected/static 辅助方法，以及 Abstractions Configurations/Imports/Exports 的非公共方法；不重做已完成的公共入口契约。

清单：清理当前 79 个方法缺失中尚未由前批处理的成员；8 个 MiniExcel 内部缺失类型中未归模型批次的委托/泛型缓存/辅助类型补摘要和泛型标签；按真实行为补 param/returns/remarks，处理 ImportTypedSheetRows 多行正文、bool 两种结果、null 和真实异步语义。辅助类型核心摘要三行；确认过时共享来源文案，IsExternalInit 的归属确认后只补类型文档或记录排除证据。

完成条件：非公开/静态/委托不遗漏、标签与签名一致；不把 CreateKey 坐标编码误写为业务缓存机制。

### 第 6 批：private 字段、其他字段、常量、缓存键和配置键

优先文件：MiniExcel/{Exports/MiniExcelExcelExporter,Imports/MiniExcelExcelImporter,Internals/MiniExcelMappingPlanBuilder,Internals/MiniExcelRawDateSerialReader,Internals/MiniExcelValueAdapter}.cs。

清单：

- 补 16 个字段和 3 个 XML 命名空间常量；覆盖 private readonly、static readonly 和私有嵌套类中的 internal 字段。
- 全量复核四项目及 shared 的 201 个字段、45 个常量；已存在的 172 个字段单行摘要、32 个常量单行摘要统一三行。
- 核对 Abstractions/Configurations/MappingProfileRegistry.cs、Core/Mappings/ExcelMappingPlanCacheKey.cs、Core/Exceptions/BingOfficesExceptionDispatcher.cs、Core/Dates/ExcelDateParser.cs 和 Npoi 的反射/样式/staging 缓存文案。
- MiniExcel InvokerCache/SyncInvokerCache.Values 按工作簿根泛型和行实体 Type 隔离；Invokers 按实体 Type 缓存反射委托；DateRule 是规则实例。只写代码证实的生命周期/共享边界，不保证未实现的淘汰策略。
- 字段不用“获取或设置”；键、常量明确具体用途、单位/边界/默认值，避免重复名称。

完成条件：private 字段无遗漏；全部字段/键/常量摘要格式和语义准确，无业务代码改动。

### 第 7 批：最终审计和构建验证

范围：全量 198 个候选文件（归属确认后的兼容桩决定需记录）。

清单：缺失和非三行摘要复扫；XML/cref、标签、重载、bool/null、类型职责、inheritdoc 全量审计；构建四个生产项目和解决方案；运行 Core/Npoi/MiniExcel 职责与集成测试、双 TFM API 契约和文档示例验证；按当前测试项目维护最终生产符号到测试方法映射。

验收差异必须相对本轮实施起点，只允许注释/必要相邻格式及计划记录。不得把工作区已有业务变化归为本任务，不通过删减现有更改或测试来使差异“纯注释”。

### 单行摘要文件索引（134 个文件）

下列是本轮确定的格式问题清单，按当前批次所属成员处理，不能据此跨批修改整个文件；花括号是文件名展开，不表示额外目录授权。其余缺失文件在第 3.4 节，已符合格式但需语义整改的文件见第 5 / 6 节。

- `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/{BindFilterAttribute}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExcelDynamicColumnMergeMode,ExcelMappingConfiguration,ExcelMappingDocumentFactory,ExcelMappingDynamicColumnConfiguration,ExcelMappingDynamicValidationConfiguration,ExcelMappingLayoutConfiguration,ExcelMappingStyleConfiguration,ExcelModelAliasRegistry,ExportMappingBuilder,FluentSetting,ImportMappingBuilder,MappingConfigurationCloner,MappingConfigurationMerger,MappingDocumentCloner,MappingProfileRegistry,ProfileDescriptor}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Conversions/{ExcelCellValue,ExcelConversionContext}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Csv/{CsvImportError,CsvImportOptions,CsvImportResult,ICsvExporter,ICsvImporter}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Exceptions/{BingOfficesExceptionDerived,BingOfficesExceptions}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{ExcelChartDefinition,ExcelColumnPlacement,ExcelComment,ExcelDynamicColumnCloner,ExcelExport,ExcelExportPolicies,ExcelHeaderCell,ExcelHeaderRow,ExcelSheetExportBuilder,ExcelWorkbookExportRequest,ExcelWorkbookMetadataOptions}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/{ExcelFailurePolicies,ExcelImageData,ExcelImport,ExcelImportError,ExcelImportPolicies,ExcelImportResult,ExcelResourceLimits,ExcelSheetPolicies,ExcelWorkbookImportRequest,ExcelWorkbookImportResult}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/IO/{IFileExportCommitter}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Providers/{IExcelMappingPlan,IExcelMappingWorkbookPlan,UniqueTracker}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Styles/{ExcelBorderStyle,ExcelCellStyle,ExcelCellStyleReset,ExcelColor}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Validations/{ExcelValidationContext,IExcelValidationBinding}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Attributes/{ColumnNameAttribute,DataFormatAttribute,DecimalScaleAttribute,ValueMappingAttribute}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/{ExcelDateAttribute,ExcelMaxLengthAttribute,ExcelMaxValueAttribute,ExcelRangeAttribute,ExcelRegexAttribute}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/{ExcelMappingConfigurationLoader,ExcelMappingDocumentValidator,ExcelMappingTextReader}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Csv/{CsvDynamicTypeResolver,CsvEntityExporter,CsvEntityImporter,CsvHeaderBinder,CsvPipelineSupport,CsvPropertyBinding}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Dates/{ExcelDateParser}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Exceptions/{BingOfficesExceptionDispatcher}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Extensions/{CsvStreamExtensions,ExcelStreamExtensions,MappingProfileServiceCollectionExtensions,ProfileDescriptorFactory}.cs`
- `src/Bing.Offices.Core/Bing/Offices/IO/{AtomicFileCommitter,DefaultFileExportCommitter}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Mappings/{ExcelDynamicMappingColumn,ExcelMappingColumn,ExcelMappingPlan,ExcelMappingPlanCacheKey,ExcelMappingPlanFactory,ExcelMappingPlanFactoryProvider,ExcelMappingStyleAndLayout,ExcelMappingWorkbookPlan,ExcelPropertyMap,ExcelTypeMap,ExcelTypeMapFactory,ExcelValidationBinding,ExcelValueConverterBindingResolver,IExcelCompiledMappingColumn}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Metadata/{MergedRegionInfo,PictureInfo,PictureStyle}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Validations/{ExcelValidationRules}.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Exports/{MiniExcelExcelExporter}.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/{MiniExcelExcelImporter}.cs`
- `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/{MiniExcelXlsxPreflight}.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/{ExcelColumnPlan,ExcelHelper,NpoiMappingPlanFactoryResolver}.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiExcelExporter,NpoiExportColumnPlanner,NpoiExportPlanBuilder,NpoiNonDisposingStream,NpoiStyleCache}.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Extensions/{SheetExtensions.Picture,WorkbookExtensions}.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Imports/{ExcelImportErrorCollector,ExcelImportExecutionOptions,ExcelImportRuntime,NpoiExcelImporter,NpoiFailureWorkbookAnnotationWriter,NpoiFailureWorkbookCopier,NpoiFailureWorkbookDiagnostics,NpoiFailureWorkbookFileSystem,NpoiFailureWorkbookPreflight,NpoiFailureWorkbookSerialization,NpoiFailureWorkbookWriter,NpoiImportPlanBuilder,NpoiImportRowMaterializer,NpoiImportSheetExecutor,NpoiRelationBinder,NpoiWorkbookValidationPipeline,ValidationRangeIndex}.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Internals/{NpoiWorkbookPlanKeyBuilder}.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/IO/{NpoiAsyncStaging}.cs`
- `src/provider-shared/{ExcelXlsxZipPreflight}.cs`

## 8. 风险和待确认项

- IsExternalInit.cs 的来源/归属：已确认该文件为仓库跟踪的项目自有兼容桩，并补充最小职责摘要；后续不再按生成文件或第三方源码排除。
- 共享文件由两个 Provider 编译，物理修改一次，构建两个 Provider；Core 的旧 ExcelXlsxZipPreflight.cs 已删除，不能沿用旧路径或“Core 共享实现”文案。
- 标签结构检查不判断异常稳定传播、true/false、null 全部语义；实施必须读真实代码和测试，不凭名称补说明。
- 当前存在业务和测试架构改动，后续批次开始时建立工作区基线；若范围内源码继续变化，先刷新对应文件统计，不抹掉其他改动。
- 合并符号模型不是逐项目构建结果，继承目标/泛型重载/条件编译在最终按 netstandard2.0、net6.0、net8.0 实际项目确认。
- 缓存可确认键维度和容器类型，不应从 ConcurrentDictionary 推导整个导入器线程安全，或虚构缓存淘汰/失效规则。
- 单行摘要格式问题广泛，但只允许注释局部修改；禁止全文件或全仓库格式化，保留编码和换行习惯。
- 历史记录中的测试文件和 netstandard/net6/net8 组合不再代表当前测试架构；验证命令按当前项目更新。
- 此计划不授权修改业务代码、成员签名、访问级别、命名空间、公开 API、测试、快照或项目配置。发现真实行为/API 问题另行报告，不借注释任务修复。

## 9. 验证命令

以下为后续实施时使用的 PowerShell 命令，本规划阶段不执行构建和测试，不把历史成功记录当成本轮结果。

```powershell
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

dotnet restore Bing.Offices.sln
dotnet build src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj -c Release --no-restore
dotnet build src/Bing.Offices.Core/Bing.Offices.Core.csproj -c Release --no-restore
dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-restore
dotnet build src/Bing.Offices.MiniExcel/Bing.Offices.MiniExcel.csproj -c Release --no-restore
dotnet build Bing.Offices.sln -c Release --no-restore

$testProjects = @(
    'tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj'
    'tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj'
    'tests/Bing.Offices.Npoi.Tests/Bing.Offices.Npoi.Tests.csproj'
    'tests/Bing.Offices.Npoi.Tests.Integration/Bing.Offices.Npoi.Tests.Integration.csproj'
    'tests/Bing.Offices.MiniExcel.Tests/Bing.Offices.MiniExcel.Tests.csproj'
    'tests/Bing.Offices.MiniExcel.Tests.Integration/Bing.Offices.MiniExcel.Tests.Integration.csproj'
)
foreach ($testProject in $testProjects) {
    foreach ($tfm in @('net6.0', 'net8.0')) {
        dotnet test $testProject -c Release -f $tfm --no-build
    }
}
foreach ($tfm in @('net6.0', 'net8.0')) {
    dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f $tfm --no-build --filter 'FullyQualifiedName~PublicApiContractTest'
}
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-build

git diff --check
git diff -- src docs/plans/chinese-comments-plan.md
```

每批先构建受影响项目及其直接依赖；shared 改动需构建 Npoi 和 MiniExcel。最终执行全套命令，检查每个退出码，不以最后一个命令成功掩盖前序失败。网络/依赖失败单独记录，不增加 NoWarn 规避文档警告。

静态筛查（文本检索只是初筛，不能代替 Roslyn 签名及继承核验）：

```powershell
rg -n '<summary>.*</summary>|<remark>' src -g '*.cs' -g '!**/bin/**' -g '!**/obj/**' -g '!*.g.cs' -g '!*.generated.cs' -g '!*.Designer.cs' -g '!**/Migration/**' -g '!**/Migrations/**' -g '!*ModelSnapshot.cs'
rg -n '<inheritdoc|<param|<typeparam|<returns|<remarks' src -g '*.cs' -g '!**/bin/**' -g '!**/obj/**' -g '!*.g.cs' -g '!*.generated.cs' -g '!*.Designer.cs'
Get-ChildItem output/release -Recurse -Filter 'Bing.Offices*.xml' |
    ForEach-Object { [xml](Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8) | Out-Null }
```

首条检索无命中时 rg 返回 1 是预期。另外需用语法树检查所有 summary 的三行结构、标签与参数对应、构造函数 cref、属性访问前缀和成员覆盖；按项目编译符号核对隐式实现。编译后确认四个生产程序集的 XML 都存在、可解析，不能仅解析残留文档文件。若输出被占用，隔离到明确的临时输出路径，不删除当前构建目录。

## 10. 验收标准

- 经排除归属确认后的全部候选文件和 2,404 个声明复扫；缺失摘要清零，未保留需要排除的声明。
- 全部独立 summary 已完成所属批次整改并采用三行格式；最终复扫未发现单行或非三行摘要。
- 所有类型职责准确、一句话核心说明；Provider 策略、来源、作用域和兼容性位于必要 remarks，修复已知复制/过时摘要。
- 118 个构造函数采用固定句式和合法 cref；720 个属性按访问方式描述；所有字段/常量（含 private）无遗漏。
- 接口/抽象契约先有效，再使用 inheritdoc；3 个接口属性重复摘要关闭；195 个已有 inheritdoc 和其他实现/override 目标有效。
- 同职责重载核心摘要一致；方法明确用途，差异由 param/returns/remarks 表达，无机械复述或虚构规则。
- param/typeparam/returns 对应真实签名，无重复、错名、多余、缺失；bool 双结果、nullable 条件和异步最终结果准确。
- XML 可解析、无错误 remark 标签、无无效 cref；不通过禁用警告替代修复。
- 注释任务相对实施起点的新增差异只包含注释和必要格式，不改业务逻辑、签名、访问级别、命名空间、API、测试或项目配置；原有改动完整保留。
- 受影响项目构建、最终解决方案构建、双 TFM Core/Provider/集成/API 契约和 Docs 验证通过；失败必须记录真实原因。
- 每批记录日期、文件清单、注释/inheritdoc 增量、复扫、待确认项和验证结果；最终维护真实生产符号到当前测试方法映射。
- 每次执行一个未完成批次并停止；本轮已按该规则完成第 1-7 批，本次在第 7 批完成后停止。

## 11. 执行记录

### 本轮计划状态（2026-09-18 / Skill 1.2.1）

当前状态：**第 1-7 批已完成**。

| 批次 | 状态 | 本轮实施记录 |
| --- | --- | --- |
| 1. 接口、抽象类和公共契约 | 已完成 | 见本节“第 1 批执行记录（2026-09-18）” |
| 2. DTO、Entity、Options 和枚举 | 已完成 | 见本节“第 2 批执行记录（2026-09-18）” |
| 3. 构造函数和属性文案统一 | 已完成 | 见本节“第 3 批执行记录（2026-09-18）” |
| 4. 实现类、重写成员和重载方法 | 已完成 | 见本节“第 4 批执行记录（2026-09-18）” |
| 5. private、internal、static 辅助方法 | 已完成 | 见本节“第 5 批执行记录（2026-09-18）” |
| 6. private 字段、其他字段、常量、缓存键和配置键 | 已完成 | 见本节“第 6 批执行记录（2026-09-18）” |
| 7. 最终审计和构建验证 | 已完成 | 见本节“第 7 批执行记录（2026-09-18）” |

本轮完成：Roslyn 全量语法/文档结构扫描，MiniExcel/shared 只读语义审查，接口实现候选定位及具体文案核对，并完成第 1-7 批接口、抽象类、公共契约、DTO、Options、元数据、枚举、构造函数、属性、实现类、重载、辅助方法、字段、常量、缓存/配置键注释及最终审计、构建和测试验证。扫描及实施均以当前工作区为基线，临时只读扫描脚本和结果不作为新增仓库文件。

计划验证：每批完成后重新执行范围审计、受影响项目构建、XML 解析和 `git diff --check`；工作区原有改动均保留。

### 第 1 批执行记录（2026-09-18）

- **实际修改文件（21 个）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExportMappingBuilder,IExcelMappingConfigurationLoader,IMappingProfileRegistry,IMappingProfileResolver,MappingProfileContracts}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Csv/{ICsvExporter,ICsvImporter}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{ExcelExport,ExcelSheetExportBuilder}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/IO/IFileExportCommitter.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImport.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Providers/{IExcelMappingPlan,IExcelMappingWorkbookPlan}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Validations/IExcelValidationBinding.cs`
  - `src/Bing.Offices.Core/Bing/Offices/{Exceptions/BingOfficesExceptionDispatcher,IO/AtomicFileCommitter,Mappings/IExcelCompiledMappingColumn}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Extensions/{SheetExtensions.Picture,WorkbookExtensions}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiFailureWorkbookFileSystem.cs`
- **注释增量**：统一 102 个 `<summary>` 块为三行格式；修订 10 个 `<param>` 和 5 个 `<returns>` 契约说明，其中包含 Profile 配置、流所有权、布尔结果及映射计划语义。新增 `inheritdoc` 0 个；本批只完善上游契约，未修改实现成员。
- **范围复扫**：105 个 Abstractions 接口契约声明，以及 Core/NPOI 的编译映射、原子文件、staging、失败工作簿和图片适配接口均无缺失 summary、非三行 summary 或签名标签错误；公共扩展/辅助入口的当前批次摘要已复核。构造函数、DTO/枚举、实现/override、私有辅助方法和字段仍按计划留待后续批次。
- **文案与继承**：接口和抽象契约摘要只保留核心用途；参数、返回值和流保持打开等调用约束写入对应标签。没有复制实现类契约文案，也没有新增 `inheritdoc`；已有实现继承关系保持不变。
- **构建与静态验证**：定义 `$taskDocDir = Join-Path $env:TEMP 'bing-offices-batch1-docs'` 后，以下三项目令均通过，0 警告、0 错误：`dotnet build src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj -c Release --no-restore -v:minimal -p:UseSharedCompilation=false "-p:DocumentationFile=$taskDocDir\\Bing.Offices.Abstractions.xml"`，以及 Core、Npoi 对应项目的相同 Release/no-restore/隔离文档输出参数，并设置 `-p:BuildProjectReferences=false` 串行构建已验证的依赖。三个临时 XML 文档均按 UTF-8 解析通过；`git diff --check` 退出码为 0（仅有既有 LF/CRLF 提示）；本批源码差异非 XML 行数为 0。普通并行/共享编译首次因既有构建进程锁定项目级 XML/DLL 输出而失败，隔离输出并禁用共享编译后复跑通过。
- **待确认项**：无本批必须确认的业务语义；`IsExternalInit.cs` 的归属仍待维护者确认。第 2-7 批未执行，尤其是构造函数/属性统一、实现成员 `inheritdoc`、private 方法和字段/键审计仍未完成。

### 第 2 批执行记录（2026-09-18）

- **实际修改文件（23 个）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/{ExcelFormat}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExcelDynamicColumnMergeMode,ExcelMappingDocument}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Csv/{CsvImportError}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exceptions/{BingOfficesExceptions}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{ExcelChartDefinition,ExcelExportPolicies,ExcelWorkbookMetadataOptions}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Imports/{ExcelFailurePolicies,ExcelImportPolicies,ExcelSheetPolicies}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Styles/{ExcelBorderStyle,ExcelCellStyle}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Validations/{IExcelValidationBinding}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Dates/{ExcelDateParser}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/{ExcelMappingStyleAndLayout}.cs`
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/{MiniExcelExcelImporter}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiStyleCache}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/{NpoiAsyncStaging}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/{ExcelImportErrorCollector,ExcelImportExecutionOptions,ExcelImportRuntime,ValidationRangeIndex}.cs`
- **注释增量**：调整 18 个 DTO、Options、元数据或状态载体类型文档块；统一 99 个枚举成员的三行摘要，并修订 `ExcelFormat` 的 2 个成员说明。新增 5 个合法 `<remarks>`，用于隔离映射方向、请求实例归属、内部执行边界、根/子错误收集器和校验索引裁剪策略。新增 `inheritdoc` 0 个，未改动构造函数、属性访问措辞、方法、字段或常量。
- **文案与语义**：`ExcelImportRuntime` 改为行数及图片资源配额跟踪职责；`ExcelMappingDocument` 和 `ExcelWorkbookMetadataOptions` 将请求/方向边界下沉到 `<remarks>`；NPOI 错误收集器、校验范围索引和执行选项保留核心职责摘要；MiniExcel 补齐 `ColumnBinding`、工作表异常和行异常类型摘要。枚举成员覆盖当前扫描到的 183 个成员，未发现缺失或非三行摘要。
- **范围复扫**：第 2 批选定类型和枚举成员的缺失/非三行数量均为 0；XML 文档结构和 C# 语法检查无错误。全量剩余缺失为 112 个，集中在 `IsExternalInit.cs` 及 MiniExcel/private 辅助成员、Npoi/shared 预检方法，留待后续批次；构造函数、全量属性访问措辞、方法、字段和实现类 `inheritdoc` 未提前处理。
- **继承检查**：本批只处理数据载体和枚举，不涉及接口实现、显式实现、抽象成员实现或 override；现有 `inheritdoc` 保持不变，新增数量为 0。
- **构建与静态验证**：按依赖顺序以隔离 XML 输出和禁用共享编译构建 Abstractions、Core、Npoi、MiniExcel，四个项目均 0 警告、0 错误；4 份临时 XML 文档按 UTF-8 解析通过；`git diff --check` 退出码为 0；本批文件差异中的非注释代码行数为 0。
- **待确认项**：本批无必须确认的业务语义。`IsExternalInit.cs` 的归属仍待维护者确认；构造函数/属性统一、实现成员 `inheritdoc`、private/internal/static 方法、字段/常量/缓存键以及最终全量审计仍未执行。

### 第 3 批执行记录（2026-09-18）

- **实际修改文件（88 个 C# 文件）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/BindFilterAttribute.cs`；`Configurations/{ExcelMappingConfiguration,ExcelMappingDynamicColumnConfiguration,ExcelMappingDynamicValidationConfiguration,ExcelMappingLayoutConfiguration,ExcelMappingStyleConfiguration,ExcelModelAliasRegistry,ExportMappingBuilder,ImportMappingBuilder,MappingProfileRegistry,ProfileDescriptor}.cs`；`Conversions/{ExcelCellValue,ExcelConversionContext}.cs`；`Csv/{CsvImportError,CsvImportOptions,CsvImportResult}.cs`；`Exceptions/{BingOfficesExceptionDerived,BingOfficesExceptions}.cs`；`Exports/{ExcelColumnPlacement,ExcelComment,ExcelExportPolicies,ExcelHeaderCell,ExcelHeaderRow,ExcelSheetExportBuilder,ExcelWorkbookExportRequest,ExcelWorkbookMetadataOptions}.cs`；`Imports/{ExcelFailurePolicies,ExcelImageData,ExcelImport,ExcelImportError,ExcelImportPolicies,ExcelImportResult,ExcelResourceLimits,ExcelSheetPolicies,ExcelWorkbookImportRequest,ExcelWorkbookImportResult}.cs`；`Providers/UniqueTracker.cs`；`Styles/{ExcelCellStyleReset,ExcelColor}.cs`；`Validations/ExcelValidationContext.cs`。
  - `src/Bing.Offices.Core/Bing/Offices/Attributes/{ColumnNameAttribute,DataFormatAttribute,DecimalScaleAttribute,ValueMappingAttribute}.cs`；`Attributes/Filters/{ExcelDateAttribute,ExcelMaxLengthAttribute,ExcelMaxValueAttribute,ExcelRangeAttribute,ExcelRegexAttribute}.cs`；`Configurations/ExcelMappingConfigurationLoader.cs`；`Csv/{CsvEntityExporter,CsvEntityImporter,CsvHeaderBinder,CsvPipelineSupport,CsvPropertyBinding}.cs`；`Exceptions/BingOfficesExceptionDispatcher.cs`；`Extensions/MappingProfileServiceCollectionExtensions.cs`；`Mappings/{ExcelDynamicMappingColumn,ExcelMappingColumn,ExcelMappingPlan,ExcelMappingPlanFactory,ExcelMappingStyleAndLayout,ExcelMappingWorkbookPlan,ExcelPropertyMap,ExcelTypeMap,ExcelValidationBinding,IExcelCompiledMappingColumn}.cs`；`Metadata/{MergedRegionInfo,PictureInfo,PictureStyle}.cs`。
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Exports/MiniExcelExcelExporter.cs`；`Imports/MiniExcelExcelImporter.cs`；`Internals/MiniExcelMappingPlanBuilder.cs`。
  - `src/Bing.Offices.Npoi/Bing/Offices/ExcelColumnPlan.cs`；`Exports/{NpoiExcelExporter,NpoiExportPlanBuilder,NpoiNonDisposingStream}.cs`；`Imports/{ExcelImportErrorCollector,ExcelImportExecutionOptions,ExcelImportRuntime,NpoiExcelImporter,NpoiFailureWorkbookPreflight,NpoiFailureWorkbookSerialization,NpoiImportPlanBuilder,NpoiImportRowMaterializer,NpoiImportSheetExecutor,ValidationRangeIndex}.cs`；`IO/NpoiAsyncStaging.cs`。
- **注释增量**：统一四个生产项目 118 个构造函数摘要为 `初始化一个 <see cref="当前类型" /> 类型的实例。`，其中补齐 4 个缺失摘要、修正 MiniExcel 2 个错误句式；复核并统一 720 个属性摘要的三行结构和“获取/设置/获取或设置/获取或初始化”访问前缀，补齐 MiniExcel `ColumnBinding.Column`、`Property`、`Header` 与 `MiniExcelRowException.Error` 4 个缺失属性摘要。新增摘要文档 8 个，新增 `<inheritdoc />` 0 个；生产源码既有 `inheritdoc` 仍为 192 个。
- **文案与继承处理**：构造函数 `<see cref>` 按当前类型生成，泛型类型引用使用合法 XML 文档语法；属性只保留核心用途，未把字段或访问器写成“获取或设置”。能继承的属性保持既有 `inheritdoc`，没有复制接口或基类契约文案；本批未处理实现方法、重载方法、private/internal/static 方法、字段、常量或缓存键。
- **范围复扫**：四项目 118/118 个构造函数的缺失、错误句式和非三行数量均为 0；720/720 个属性的缺失、访问前缀错误和非三行数量均为 0；C# 语法扫描无错误。构造函数固定句式检查命中 118 行，MiniExcel 4 个新增只读属性均使用“获取”前缀。
- **构建与静态验证**：按依赖顺序隔离 XML 输出并禁用共享编译，`dotnet build` 通过：Abstractions/Core `netstandard2.0`，Npoi 与 MiniExcel 的 `net6.0`、`net8.0` 均 0 警告、0 错误；4 份临时 XML 文档按 UTF-8 解析通过；`git diff --check` 退出码为 0（仅有既有 LF/CRLF 提示）。本批新增差异仅为 XML 注释及必要的三行排版，工作区既有非注释改动完整保留，未修改业务逻辑、成员签名、访问级别、命名空间、公开 API、测试或项目配置。
- **待确认项**：本批无必须确认的业务语义；`src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs` 的兼容桩归属仍待维护者确认。第 5-7 批（辅助方法、字段/常量/键及最终审计）尚未执行。

### 第 4 批执行记录（2026-09-18）

- **实际修改文件（20 个 C# 文件）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExcelModelAliasRegistry.cs`；`Csv/{ICsvExporter,ICsvImporter}.cs`；`Exports/IExcelExporter.cs`；`Imports/IExcelImporter.cs`；`Providers/UniqueTracker.cs`。
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`；`IO/{AtomicFileCommitter,DefaultFileExportCommitter}.cs`；`Mappings/ExcelValidationBinding.cs`。
  - `src/Bing.Offices.MiniExcel/Bing/Offices/{Exports/MiniExcelExcelExporter,Imports/MiniExcelExcelImporter}.cs`。
  - `src/Bing.Offices.Npoi/Bing/Offices/{ExcelHelper,NpoiMappingPlanFactoryResolver}.cs`；`Exports/NpoiExcelExporter.cs`；`Imports/NpoiExcelImporter.cs`；`IO/NpoiAsyncStaging.cs`；`Extensions/{CellExtensions,SheetExtensions.MergedRegion,SheetExtensions.Picture}.cs`。
- **注释增量**：修订或统一 60 个 `<summary>` 文档块（包含实现类型、同步/异步重载、筛选重载和三行排版）；新增 28 个 `<remarks>`、2 个 `<exception>`，将 Provider 策略、流所有权、文件提交分工、别名校验和筛选差异下沉到对应标签。`ExcelValidationBinding.Kind`、`IsRaw`、`ErrorMessage` 3 个实现属性改用 `/// <inheritdoc />`；生产源码 `<inheritdoc />` 总数由 192 增至 195。
- **实现与继承复核**：复核接口实现、显式接口实现、抽象成员实现和 override；审计结果 `OverrideNoInherit=0`、`ExplicitNoInherit=0`。实现类未复制上游已有的 summary、param 或 returns；CSV/Excel、同步/异步、文件/流及 NPOI 图片/合并区域重载保持一致核心用途，输入载体、筛选范围和策略差异通过 param、returns 或 remarks 表达。
- **范围复扫**：四个生产项目和 `provider-shared` 共 2,404 个声明重新扫描；当前剩余缺失 104 个，均属于计划第 5 批及之后的辅助成员/兼容桩范围（Abstractions 1、MiniExcel 95、Npoi 1、shared 7），未新增第 4 批缺失。目标文件 XML 结构、参数/泛型/返回标签和语法检查无新增错误；目标文件差异中的非注释代码行数为 0。
- **构建与静态验证**：按依赖顺序以隔离 XML 输出和 `UseSharedCompilation=false` 构建 Abstractions、Core、Npoi（net6.0/net8.0）和 MiniExcel（net6.0/net8.0），均 0 警告、0 错误；临时目录中的 4 份 XML 文档按 UTF-8 解析通过；`git diff --check` 退出码为 0，仅保留既有 LF/CRLF 提示。
- **待确认项**：本批无必须确认的业务语义。`src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs` 的兼容桩归属仍待维护者确认；第 5 批 private/internal/static 辅助方法、第 6 批字段/常量/缓存键和第 7 批最终全量审计尚未执行。

### 第 5 批执行记录（2026-09-18）

- **实际修改文件（按目录归组）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExcelMappingDocumentFactory,ExcelModelAliasRegistry,ImportMappingBuilder,MappingConfigurationCloner,MappingConfigurationMerger,MappingDocumentCloner,MappingProfileRegistry}.cs`；`Exports/{ExcelChartDefinition,ExcelColumnPlacement,ExcelDynamicColumnCloner,ExcelSheetExportBuilder,ExcelWorkbookMetadataOptions}.cs`；`Imports/ExcelWorkbookImportRequest.cs`；`Providers/UniqueTracker.cs`。
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/{ExcelMappingConfigurationLoader,ExcelMappingTextReader}.cs`；`Csv/{CsvDynamicTypeResolver,CsvEntityExporter,CsvEntityImporter,CsvHeaderBinder,CsvPipelineSupport,CsvPropertyBinding}.cs`；`Dates/ExcelDateParser.cs`；`Exceptions/BingOfficesExceptionDispatcher.cs`；`Extensions/{CsvStreamExtensions,ExcelStreamExtensions,MappingProfileServiceCollectionExtensions,ProfileDescriptorFactory}.cs`；`IO/AtomicFileCommitter.cs`；`Mappings/{ExcelDynamicMappingColumn,ExcelMappingColumn,ExcelMappingPlan,ExcelMappingPlanCacheKey,ExcelMappingPlanFactory,ExcelMappingPlanFactoryProvider,ExcelMappingStyleAndLayout,ExcelMappingWorkbookPlan,ExcelPropertyMap,ExcelTypeMapFactory,ExcelValidationBinding,ExcelValueConverterBindingResolver}.cs`；`Validations/ExcelValidationRules.cs`。
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Exports/MiniExcelExcelExporter.cs`；`Imports/MiniExcelExcelImporter.cs`；`Internals/{MiniExcelMappingPlanBuilder,MiniExcelRawDateSerialReader,MiniExcelValueAdapter,MiniExcelXlsxPreflight}.cs`。
  - `src/Bing.Offices.Npoi/Bing/Offices/{ExcelColumnPlan,ExcelHelper,NpoiMappingPlanFactoryResolver}.cs`；`Exports/{NpoiExcelExporter,NpoiExportColumnPlanner,NpoiExportPlanBuilder,NpoiNonDisposingStream,NpoiStyleCache}.cs`；`Imports/{ExcelImportErrorCollector,ExcelImportRuntime,NpoiExcelImporter,NpoiFailureWorkbookAnnotationWriter,NpoiFailureWorkbookCopier,NpoiFailureWorkbookDiagnostics,NpoiFailureWorkbookFileSystem,NpoiFailureWorkbookPreflight,NpoiFailureWorkbookSerialization,NpoiFailureWorkbookWriter,NpoiImportPlanBuilder,NpoiImportSheetExecutor,NpoiRelationBinder,NpoiWorkbookValidationPipeline,NpoiXlsxZipPreflight,ValidationRangeIndex}.cs`；`Extensions/SheetExtensions.Picture.cs`；`Internals/NpoiWorkbookPlanKeyBuilder.cs`；`IO/NpoiAsyncStaging.cs`。
  - `src/provider-shared/ExcelXlsxZipPreflight.cs`。
- **注释增量**：补齐 79 个缺失方法文档和 5 个 MiniExcel 内部委托/泛型缓存/计划委托类型摘要；为泛型委托和辅助方法补齐参数、泛型参数、返回值标签，补充真实同步/异步、bool、null、日期系统、坐标键、错误上限和流位置语义。将 265 个本批范围内既有非公开/静态辅助摘要统一为三行格式；`ImportTypedSheetRows` 的行号解释移入 `<remarks>`。新增 `inheritdoc` 0 个，生产源码 `<inheritdoc />` 总数保持 195 个。
- **文案与继承处理**：MiniExcel 映射计划构建器不再声称由 Core 直接提供，预检适配器改为在类型 `remarks` 说明共享编译链接；NPOI/shared 预检明确 ZIP 分支、资源限制和 `date1904` true/false 语义；`CreateKey` 说明为物理行列索引编码，不描述为业务缓存。辅助类型和委托保留核心职责摘要，未复制接口/基类契约，也未改变实现成员继承关系。
- **范围复扫**：四个生产项目和 `provider-shared` 共 2,404 个声明重新扫描；方法缺失 0、非公开/静态方法单行摘要 0、参数/typeparam/returns 不匹配 0、override/显式接口实现缺少 inheritdoc 0。当前剩余缺失为 Abstractions 的 `IsExternalInit.cs` 兼容桩 1 个，以及 MiniExcel 第 6 批字段/常量 19 个；剩余单行摘要仅为未重做的公共非静态入口/公共类型。C# 语法扫描无错误。
- **构建与静态验证**：按依赖顺序以 `UseSharedCompilation=false` 和隔离文档输出构建 Abstractions、Core、Npoi（net6.0/net8.0）和 MiniExcel（net6.0/net8.0），均 0 警告、0 错误；`$env:TEMP\bing-offices-batch5-docs` 下 4 份 XML 文档按 UTF-8 解析通过；`git diff --check` 退出码为 0（仅有既有 LF/CRLF 提示）。本批源码编辑仅为 XML 注释及必要的三行排版，工作区原有业务和测试改动完整保留。
- **待确认项**：`src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs` 是编译器兼容桩，未确认归属前保留缺口并记录，不添加无意义摘要；MiniExcel 私有字段、缓存/配置键和 XML 命名空间常量留待第 6 批，第 7 批负责最终全量审计和解决方案级验证。

### 第 6 批执行记录（2026-09-18）

- **实际修改文件（按目录归组）**：
  - `src/Bing.Offices.MiniExcel/Bing/Offices/{Exports/MiniExcelExcelExporter,Imports/MiniExcelExcelImporter,Internals/MiniExcelMappingPlanBuilder,Internals/MiniExcelRawDateSerialReader,Internals/MiniExcelValueAdapter}.cs`。
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExportMappingBuilder,ImportMappingBuilder}.cs` 及其余包含字段/常量文档的范围文件；全量字段/常量候选统一三行格式。
  - `src/Bing.Offices.Core/Bing/Offices/{InternalConst,Dates/ExcelDateParser,Exceptions/BingOfficesExceptionDispatcher,Mappings/ExcelMappingPlanCacheKey,Validations/ExcelValidationRules}.cs`；`Configurations/{ExcelMappingConfigurationLoader,ExcelMappingDocumentValidator,ExcelMappingTextReader}.cs`；以及其余包含字段/常量文档的范围文件。
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiExcelExporter,NpoiExportPlanBuilder,NpoiNonDisposingStream,NpoiStyleCache}.cs`；`Imports/{ExcelImportErrorCollector,ExcelImportRuntime,NpoiExcelImporter,NpoiFailureWorkbookSerialization,NpoiImportPlanBuilder,NpoiImportSheetExecutor,NpoiRelationBinder,NpoiWorkbookValidationPipeline,ValidationRangeIndex}.cs`；`Extensions/CellExtensions.cs`；`IO/NpoiAsyncStaging.cs`；`Internals/InternalConst.cs`。
  - `src/provider-shared/ExcelXlsxZipPreflight.cs` 未发现字段或常量目标，本批未修改。
- **注释增量**：补齐 MiniExcel 16 个字段和 3 个 XML 命名空间常量摘要；统一四项目 172 个既有字段摘要和 32 个既有常量摘要为三行格式；修订缓存选项、异常 Data 键、日期/文本限制、工作簿边界和 NPOI staging/反射/样式缓存的实际用途与作用域。字段/常量摘要直接描述用途，不使用属性访问措辞；新增 `inheritdoc` 0 个，生产源码总数保持 195 个。
- **文案与语义处理**：`InvokerCache<TWorkbook>.Values` 和 `SyncInvokerCache<TWorkbook>.Values` 明确按工作簿根泛型与行实体运行时类型隔离；MiniExcel `Invokers` 明确只缓存反射委托；`DateRule` 明确为共享规则实例；映射计划缓存序列化选项的隔离说明下沉到 `<remarks>`；常量补充实际数值、单位、默认值或边界；字段不承诺未实现的缓存淘汰或生命周期策略。
- **范围复扫**：四个生产项目和 `provider-shared` 的 2,404 个声明中，字段 201/201、常量 45/45 均有三行摘要，字段/常量缺失、单行和非三行数量均为 0；参数、typeparam、returns 不匹配均为 0；`OverrideNoInherit=0`、`ExplicitNoInherit=0`；C# 语法扫描无错误。全量仍保留 1 个待确认的 `IsExternalInit` 兼容桩类型缺口。
- **构建与静态验证**：按依赖顺序使用 `UseSharedCompilation=false`、隔离 XML 输出和 `--no-restore` 构建 Abstractions（netstandard2.0）、Core（netstandard2.0）、Npoi（net6.0/net8.0）和 MiniExcel（net6.0/net8.0），均 0 警告、0 错误；4 份临时 XML 文档按 UTF-8 解析通过；`git diff --check` 通过（仅有既有 LF/CRLF 提示）。本批仅修改 XML 注释及必要三行排版，未改变业务逻辑、成员签名、访问级别、命名空间、公开 API、测试或项目配置。
- **待确认项**：`src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs` 的兼容桩归属仍待维护者确认，因此未添加无意义摘要；第 7 批负责最终全量审计、解决方案构建、测试和生产符号到测试方法追溯。本轮未执行第 7 批。

### 第 7 批执行记录（2026-09-18）

- **实际修改文件（12 个 C# 文件）**：`src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs`；`src/Bing.Offices.Abstractions/Bing/Offices/Configurations/MappingConfigurationMerger.cs`；`src/Bing.Offices.Abstractions/Bing/Offices/Csv/CsvImportOptions.cs`；`src/Bing.Offices.Abstractions/Bing/Offices/Exceptions/BingOfficesExceptions.cs`；`src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelChartDefinition.cs`；`src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelExportPolicies.cs`；`src/Bing.Offices.Abstractions/Bing/Offices/Imports/{ExcelImportPolicies,ExcelResourceLimits,ExcelSheetPolicies}.cs`；`src/Bing.Offices.Core/Bing/Offices/Csv/CsvPipelineSupport.cs`；`src/Bing.Offices.Core/Bing/Offices/Extensions/MappingProfileServiceCollectionExtensions.cs`；`src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs`。
- **注释增量**：为项目自有的 `IsExternalInit` 兼容类型补充 1 个三行 `<summary>`；将 27 个此前单行 `<summary>`（9 个类型、18 个方法）重排为三行；修订 `TryParseValue`、`TryParseWorkbookDate`、异步 `WriteField` 和 `CsvImportOptions.Validate` 4 个成员的核心用途、布尔返回和额外行为说明。生产源码 `<inheritdoc />` 总数保持 195 个，本批未复制上游契约文案。
- **最终范围复扫**：198 个候选文件、2,404 个声明全部通过；类型 311、构造函数 118、方法 826、属性 720、字段 201、常量 45、枚举成员 183 均无缺失或非三行摘要。`Missing=0`、`OneLine=0`、`NonThreeLine=0`、`ParamBad=0`、`TypeParamBad=0`、`ReturnBad=0`、`ConstructorBad=0`、`PropertyBad=0`、`LongSummary=0`、`XmlError=0`、`OverrideNoInherit=0`、`ExplicitNoInherit=0`、摘要长度超过 120 字为 0；`IsExternalInit.cs` 已完成归属确认后的项目内注释，不再保留待处理缺口。
- **文案与继承复核**：构造函数固定句式、属性访问前缀、字段用途表达、重载核心 summary 一致性和应下沉的 remarks 已复核；接口实现、显式接口实现、抽象成员实现和 override 均保持正确 `inheritdoc`。未发现复制上游契约、错误/过时/无意义注释或需要新增 inheritdoc 的成员。
- **构建、文档与测试验证**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal -p:UseSharedCompilation=false` 通过，0 警告、0 错误；`output/release` 下 11 份生产 XML 文档全部按 UTF-8 解析通过；`git diff --check` 退出码为 0（仅有既有 LF/CRLF 提示）。`dotnet test Bing.Offices.sln -c Release --no-build --no-restore -v:minimal --nologo` 全部通过：Core 单元 196/196（net6、net8）、Core 集成 29/29（net6、net8）、MiniExcel 单元 39/39（net6、net8）、MiniExcel 集成 9/9（net6、net8）、NPOI 单元 539/539（net6、net8）、NPOI 集成 30/30（net6、net8）、Docs 10/10（net8），合计 1,694 个测试通过、0 失败、0 跳过。
- **生产符号到测试方法追溯**：`IExcelImporter.ImportAsync`、`NpoiExcelImporter.ImportAsync` 由 `Bing.Offices.Npoi.Tests.AsyncPipelineTest.ExcelAsync_ShouldUseAsyncOuterStream_AndMatchSyncImport` 覆盖；Excel/CSV 文件异步导出与提交由 `AsyncFileExport_PreCanceled_ShouldNotCreateTarget`、`CsvAsync_ShouldUseAsyncReadWrite_AndMatchSyncResult` 以及 `DefaultFileExportCommitterTest.CommitAsync_NewTarget_ShouldWriteAndMove`、`CommitAsync_ExistingTarget_ShouldReplaceAfterSuccessfulWrite`、`CommitAsync_WriteFailure_ShouldPreserveExceptionAndTarget`、`CommitAsync_PreCanceled_ShouldNotCreateFiles`、`CommitAsync_CanceledAfterWrite_ShouldCleanTemporaryFile` 覆盖；MiniExcel 对应异步边界由 `MiniExcelProviderTest.AsyncImport_MidReadCancellation_ShouldKeepSourceOpen`、`MiniExcelProviderTest.AsyncFileExport_PreCanceled_ShouldPreserveExistingTargetAndCleanTempFile`、`MiniExcelProviderTest.AsyncFileExport_MidFlightCancellation_ShouldPreserveExistingTargetAndCleanTempFile` 及 `MiniExcelProviderIntegrationTest.ExportToFileAsync_ThenImportFromFileAsync_ShouldRoundTripRealXlsx` 覆盖；Core/NPOI 公开扩展由 `PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature` 覆盖；公共 API 由 `PublicApiContractTest.PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline`、`PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` 覆盖；Markdown 示例由 `DocsConsumerTest.DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually` 覆盖。
- **差异和待确认项**：本批新增差异仅为 XML 注释、三行排版及本计划记录；工作区原有业务、测试架构、删除和未跟踪改动均保留，未修改业务逻辑、成员签名、访问级别、命名空间、公开 API、测试或项目配置。本批无待确认项。

### 历史记录（2026-09-15，旧范围 / 旧版规范）

以下记录原样保留用于追溯；仅属于上一轮完成状态，旧基线统计、测试路径及验证结果不用于判定本轮完成。

历史状态：**旧轮次第 1-7 批已完成**。

| 批次 | 状态 | 实际涉及文件 | 复扫/验证结果 | 备注 |
| --- | --- | --- | --- | --- |
| 1. 接口、抽象类和公共契约 | 已完成 | 20 个文件已修改；21 个范围文件中 `FilterAttributeBase.cs` 无需修改 | 范围复扫未发现缺失契约标签；XML 文档可解析；Release 构建通过；`git diff --check` 通过 | 仅补充/整理契约注释，未处理实现类、枚举成员、构造函数、字段或内部 `Build` 方法 |
| 2. DTO、Entity、Options 和枚举 | 已完成 | 28 个 C# 文件：Abstractions 15 个、Core 9 个、Npoi 4 个；其中 `ExcelSheetExportBuilder.cs`、`NpoiAsyncStaging.cs` 与第 1 批重叠 | 目标类型/属性/typeparam 缺口复扫为 0；27 个枚举及 97 个枚举成员均有摘要；XML 文档可解析；目标 Release 构建通过；`git diff --check` 通过 | 补齐 DTO/Options/元数据/策略枚举文档，明确默认值、单位、范围和 `null` 语义；未处理构造函数、全局属性措辞、方法、字段和实现类 inheritdoc |
| 3. 构造函数和属性文案统一 | 已完成 | 85 个 C# 文件：Abstractions 41 个、Core 30 个、Npoi 14 个 | 112 个构造函数摘要精确匹配；缺失属性文档和访问器前缀候选清零；XML 解析、Release 构建和 `git diff --check` 通过 | 仅修改 XML 注释；实现方法、private 方法和字段留待后续批次 |
| 4. 实现类、重写成员和重载方法 | 已完成 | 4 个 C# 文件：Core 3 个、Npoi 1 个 | 实现/override 复扫通过；XML 解析、Release 构建和 `git diff --check` 通过 | 仅修改 XML 注释，未触及第 5 批 |
| 5. private、internal、static 辅助方法 | 已完成 | 31 个 C# 文件：Core 15 个、Npoi 16 个 | 方法目标复扫通过；XML 解析、Release 构建和 `git diff --check` 通过 | 仅补方法及内部适配器契约注释，未触及第 6 批 |
| 6. private 字段、其他字段、常量、缓存键和配置键 | 已完成 | 19 个 C# 文件；Abstractions 8 个、Core 7 个、Npoi 7 个 | 字段/常量复扫未发现缺失；Release 构建、XML 解析和 `git diff --check` 通过；非 XML C# 差异为 0 | 新增 113 个字段/常量摘要，修订 2 个缓存摘要；未执行第 7 批 |
| 7. 最终审计和构建验证 | 已完成 | 43 个文件：Abstractions 10 个、Core 13 个、Npoi 20 个 | 188 个候选文件、2,271 个目标声明复扫通过；Release 构建、核心/集成测试、API snapshot、XML 解析和 `git diff --check` 通过 | 确定缺失、标签、文案和 inheritdoc 候选均已关闭 |

基线扫描日期：2026-09-15。基线工作区在扫描开始时无未提交改动；计划生成阶段除本计划文档外未修改任何文件。

### 第 1 批执行记录（2026-09-15）

- **实际修改文件（20 个）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/DecoratorAttributeBase.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExportMappingBuilder,FluentSetting,IExcelMappingConfigurationLoader,ImportMappingBuilder}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Conversions/IExcelValueConverter.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Csv/{ICsvExporter,ICsvImporter}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{ExcelExport,ExcelSheetExportBuilder,IExcelExporter}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Imports/{ExcelImport,IExcelImporter}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/IO/IFileExportCommitter.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Providers/IExcelMappingPlan.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Validations/{IExcelValidationBinding,IExcelValidationRule,INamedExcelValidationRule}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/IExcelCompiledMappingColumn.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`（仅内部 staging 接口）
- **注释增量**：113 个类型/成员文档项完成补充或规范化；新增 XML 注释 245 行，移除不适用的裸 `Task` `<returns>` 或重排 13 行，净增 232 行。其中新增 `<summary>` 9 行、`<remarks>` 7 行、`<typeparam>` 22 行、`<param>` 109 行、`<returns>` 92 行；新增 `inheritdoc` 0 个。
- **范围复扫**：第 1 批接口、抽象类、公共入口和内部 staging 接口的缺失 `summary`/`typeparam`/`param`/`returns` 检查结果为 0；`Task`/`ValueTask` 返回规则、`bool` 语义、重载标签和 XML 结构已复核。实现类、override、枚举成员、构造函数、属性访问措辞、private 方法和字段按计划留在后续批次。
- **构建与静态验证**：`dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release`（联动 Abstractions/Core，net6.0 与 net8.0）通过，0 警告、0 错误；三个生产程序集的 Release XML 文档均成功生成并可解析；`git diff --check` 通过；C# 差异新增和删除的非 XML 行均为 0。首次普通权限还原受 NuGet SSL/凭据限制失败，受控网络权限重试后通过。
- **inheritdoc 说明**：本批只建立上游契约，未修改实现类，因此没有新增 `inheritdoc`；实现/override 对照与补充安排在第 4 批。
- **待确认项**：无代码语义待确认项。构建依赖外部 NuGet 源，若后续环境无网络需使用已缓存的还原资产；`IsExternalInit.cs` 的排除归属仍按原计划保留待维护者确认。

### 第 2 批执行记录（2026-09-15）

- **实际修改文件（28 个）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExcelColumnConfiguration,ExcelMappingConfiguration,MappingConfigurationCloner,MappingDocumentCloner,MappingProfileRegistry,ProfileDescriptor}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{ExcelDynamicColumnCloner,ExcelDynamicColumnDefinition,ExcelExportPolicies,ExcelSheetExportBuilder,ExcelWorkbookMetadataOptions}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Imports/{ExcelResourceLimits,ExcelSheetPolicies,ExcelWorkbookImportRequest,ExcelWorkbookImportResult}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/{ExcelDynamicMappingColumn,ExcelMappingColumn,ExcelMappingPlan,ExcelMappingPlanFactory,ExcelMappingStyleAndLayout,ExcelMappingWorkbookPlan,ExcelValidationBinding}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Metadata/{MergedRegionInfo,PictureInfo}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/{ExcelImportExecutionOptions,ExcelImportRuntime,ValidationRangeIndex}.cs`
- **注释增量**：补齐 20 个缺失类型 `<summary>`、7 个缺失属性 `<summary>`、3 个缺失 `<typeparam>` 和 `NpoiAsyncStagingStrategy` 的 3 个枚举成员摘要；另修订动态列、导出策略、资源限制、工作表选择器、工作簿元数据及图像/合并区域元数据中的默认值、单位、零基边界和 `null` 语义。新增 `inheritdoc` 0 个，现有 `inheritdoc` 计数保持 126 个。
- **范围复扫**：第 2 批范围 93 个文件内，目标类型、属性和 `<typeparam>` 缺失项均为 0；27 个枚举、97 个枚举成员均有摘要；连续 XML 文档块可解析。构造函数固定句式、属性访问器措辞、方法和字段仍按计划留在后续批次。
- **构建与静态验证**：`dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-restore` 通过，联动 Abstractions/Core 的 net6.0 与 net8.0 目标均成功；8 个 Release XML 文档文件成功生成并解析；`git diff --check` 通过；C# 差异中除 `MergedRegionInfo.cs`、`PictureInfo.cs` 文件末尾换行外无非 XML 行变更。
- **inheritdoc 说明**：本批处理 DTO、Options、元数据和枚举，不新增接口实现或 override 文档；实现类继承关系与重复契约文案安排在第 4 批。
- **待确认项**：本批无需业务语义确认。`IsExternalInit.cs` 的排除归属、构造函数/属性全量措辞统一、private 方法和字段补全以及 inheritdoc 对照仍按原计划待后续批次处理。

### 第 3 批执行记录（2026-09-15）

- **实际修改文件（85 个 C# 文件）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/{BindFilterAttribute}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExcelColumnConfiguration,ExcelModelAliasRegistry,ExportMappingBuilder,ImportMappingBuilder,MappingProfileRegistry,ProfileDescriptor}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Conversions/{ExcelCellValue,ExcelConversionContext}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Csv/{CsvImportError,CsvImportOptions,CsvImportResult}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exceptions/{BingOfficesExceptionDerived,BingOfficesExceptions}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{ExcelChartDefinition,ExcelColumnPlacement,ExcelComment,ExcelDynamicColumnDefinition,ExcelExportPolicies,ExcelHeaderCell,ExcelHeaderRow,ExcelSheetExportBuilder,ExcelWorkbookExportRequest,ExcelWorkbookMetadataOptions}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Imports/{ExcelFailurePolicies,ExcelImageData,ExcelImport,ExcelImportError,ExcelImportPolicies,ExcelImportResult,ExcelResourceLimits,ExcelSheetPolicies,ExcelWorkbookImportRequest,ExcelWorkbookImportResult}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Providers/{IExcelMappingPlan,UniqueTracker}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Styles/{ExcelBorderStyle,ExcelCellStyle,ExcelCellStyleReset,ExcelColor}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Validations/ExcelValidationContext.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Attributes/{ColumnNameAttribute,DataFormatAttribute,DecimalScaleAttribute,ValueMappingAttribute}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Attributes/Decorators/HeaderAttribute.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/{ExcelDateAttribute,ExcelMaxLengthAttribute,ExcelMaxValueAttribute,ExcelRangeAttribute,ExcelRegexAttribute}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/{CsvEntityExporter,CsvEntityImporter,CsvHeaderBinder,CsvPipelineSupport,CsvPropertyBinding}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Exceptions/BingOfficesExceptionDispatcher.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Extensions/MappingProfileServiceCollectionExtensions.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/{ExcelDynamicMappingColumn,ExcelMappingColumn,ExcelMappingPlan,ExcelMappingPlanFactory,ExcelMappingStyleAndLayout,ExcelMappingWorkbookPlan,ExcelPropertyMap,ExcelTypeMap,ExcelValidationBinding,IExcelCompiledMappingColumn}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Metadata/{MergedRegionInfo,PictureInfo}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/ExcelColumnPlan.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiExcelExporter,NpoiExportPlanBuilder,NpoiNonDisposingStream}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/{ExcelImportErrorCollector,ExcelImportRuntime,NpoiExcelImporter,NpoiFailureWorkbookPreflight,NpoiFailureWorkbookSerialization,NpoiImportPlanBuilder,NpoiImportRowMaterializer,NpoiImportSheetExecutor,ValidationRangeIndex}.cs`
- **注释增量**：112 个构造函数摘要统一为 `初始化一个 <see cref="当前类型" /> 类型的实例。`，其中补齐 26 个缺失摘要并修订 86 个既有摘要；148 个属性摘要按 `获取`、`设置`、`获取或设置`、`获取或初始化` 统一；补齐 21 个缺失属性文档，其中 16 个使用 `inheritdoc`、5 个使用独立摘要；将 `DataFormatAttribute` 属性参考链接下沉到 `<remarks>`。本批新增 `inheritdoc` 16 个。
- **范围复扫**：三个生产项目共识别 112 个构造函数，精确句式为 112/112，缺失为 0；本批识别的 21 个缺失属性文档均已补齐，712 个已识别属性的访问器前缀无残留不一致；已有 `inheritdoc` 未被替换。摘要长度和 XML 结构复核通过，未新增方法、字段或成员标签。
- **文案与继承处理**：构造函数保留原有参数标签和额外说明；属性摘要仅保留核心用途，外部参考说明置于 `<remarks>`；Stream override、staging 接口实现属性和其他可继承属性使用 `inheritdoc`。接口实现方法、重载方法和其余实现类继承文档按第 4 批处理。
- **构建与静态验证**：`dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-restore` 通过，NPOI 的 net6.0/net8.0 与联动项目目标均成功且无警告错误；8 个 Release XML 文档文件按 UTF-8 解析通过；`git diff --check` 通过（仅有 Git 的 LF/CRLF 提示）；C# 差异非 XML 代码行数为 0。
- **待确认项**：未发现需要业务语义确认的构造函数或属性。`IsExternalInit.cs` 的排除归属、复杂继承关系的方法 `inheritdoc`、private/internal 方法和字段注释仍按原计划留待后续批次；本批未执行第 4 批及之后工作。

### 第 4 批执行记录（2026-09-15）

- **实际修改文件（4 个）**：
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvPipelineSupport.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingColumn.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelValidationBinding.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
- **注释增量**：补齐 `CsvPipelineSupport` 中 16 个流包装器 override；补齐 `NpoiAsyncStaging` 中 13 个 staging 接口实现/混合流 override；为 `ExcelValidationBinding.Validate` 新增 1 个 `inheritdoc`；将 `ExcelMappingColumn` 的 4 个编译映射契约属性改用 `inheritdoc`。本批新增 `inheritdoc` 共 34 个，并对字节上限、绑定取消令牌、底层流所有权和 staging 阈值迁移等差异行为补充必要 `<remarks>`。
- **范围复扫**：已审计的 63 个 override 成员均有文档（缺失 0）；loader 与 NPOI Extensions 的重载族保留语义一致且可区分的 summary；未发现显式接口实现遗漏；`NpoiNonDisposingStream.Dispose` 因存在额外释放语义保留独立说明。private 字段未在本批处理，仍由第 6 批负责。
- **继承处理**：接口实现、嵌套 `Stream` override 和 staging 实现使用 `/// <inheritdoc />`；实现差异仅追加 `<remarks>`，未复制上游 summary、param 或 returns。内部 `DefaultPictureMutationAdapter` 及其无完整上游契约的辅助实现留待后续范围复核。
- **构建与静态验证**：`dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-restore -v:minimal` 通过，联动项目的 netstandard2.0、net6.0 与 net8.0 目标均成功，0 警告、0 错误；8 个 Release XML 文档文件成功解析；`git diff --check` 通过（仅有 Git 的 LF/CRLF 提示）；忽略既有 LF/CRLF 与文件末尾换行差异后，C# 差异中的非 XML 代码行数为 0。
- **待确认项**：无本批必须确认的业务语义。`IsExternalInit.cs` 的排除归属仍待维护者确认；`DefaultPictureMutationAdapter` 的内部接口契约、private/internal/static 辅助方法和字段/键注释按原计划留待第 5-6 批。本批未执行第 5 批及之后工作。

### 第 5 批执行记录（2026-09-15）

- **实际修改文件（31 个）**：
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/{ExcelMappingConfigurationLoader,ExcelMappingDocumentValidator}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/{CsvEntityExporter,CsvEntityImporter,CsvPipelineSupport}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Dates/ExcelDateParser.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Extensions/{CsvStreamExtensions,ExcelStreamExtensions}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/IO/AtomicFileCommitter.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/{ExcelMappingPlanCacheKey,ExcelMappingPlanFactoryProvider,ExcelTypeMapFactory,ExcelValidationBinding,ExcelValidationBindingFactory,ExcelValueConverterBindingResolver}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiExcelExporter,NpoiExportPlanBuilder,NpoiExportSheetWriter}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/{NpoiExcelImporter,NpoiFailureWorkbookAnnotationWriter,NpoiFailureWorkbookDiagnostics,NpoiFailureWorkbookPreflight,NpoiFailureWorkbookWriter,NpoiImportPlanBuilder,NpoiImportRowMaterializer,NpoiImportSheetExecutor,NpoiRelationBinder,NpoiXlsxZipPreflight}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Extensions/{SheetExtensions.Picture,WorkbookExtensions}.cs`
- **注释增量**：补齐或修订 126 个方法/内部契约 XML 文档项：Core 47 个方法；Npoi 37 个缺失方法文档、32 个既有方法签名标签，以及 `IPictureMutationAdapter` 与 `DefaultPictureMutationAdapter` 的 10 个契约/实现文档项。新增 `inheritdoc` 5 个，用于内部图片适配器实现；异步方法补充取消与结果语义，必要的资源所有权、流位置和预检行为说明置于 `<remarks>`。
- **范围复扫**：第 5 批目标方法 123 个声明均有 `<summary>` 或 `<inheritdoc />`；全目录静态命中剩余 11 项均为计划排除的局部函数或委托类型。已复核 private/internal/static 辅助方法、重载 summary 一致性、`<param>`/`<typeparam>`/`<returns>` 标签与实际签名；未修改 private 字段、其他字段、常量或缓存/配置键，留待第 6 批。
- **文案与继承处理**：方法 summary 仅保留核心用途，参数、返回值、取消行为和资源边界按需使用标签或 `<remarks>`；同步/异步及泛型反射委托重载保持同族表述。图片适配器先补充内部接口契约，默认实现改用 `/// <inheritdoc />`，未复制上游契约文案。
- **构建与静态验证**：`dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-restore -v:minimal` 通过，联动 Abstractions/Core 的 netstandard2.0、net6.0 与 net8.0 目标均成功，0 警告、0 错误；8 个 Release XML 文档文件按 UTF-8 解析通过；`git diff --check` 通过（仅有 Git 的 LF/CRLF 提示）；忽略既有换行差异后，C# 差异中的非 XML 代码行数为 0。
- **待确认项**：无本批必须确认的业务语义。局部函数和委托类型继续按规则不机械添加 XML；`IsExternalInit.cs` 的排除归属仍待维护者确认；第 6 批字段、常量、缓存键和配置键以及第 7 批最终审计尚未执行。

### 第 6 批执行记录（2026-09-15）

- **实际修改文件（19 个）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/{Imports/ExcelImport,Exports/ExcelExport,Exports/ExcelSheetExportBuilder,Providers/UniqueTracker}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExportMappingBuilder,ImportMappingBuilder,FluentSetting,MappingProfileRegistry}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/{Csv/CsvPipelineSupport,Dates/ExcelDateParser,Exceptions/BingOfficesExceptionDispatcher}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/{ExcelMappingPlanFactory,ExcelMappingPlanCacheKey,ExcelValidationBinding}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/{NpoiExcelImporter,NpoiImportPlanBuilder,NpoiImportSheetExecutor,NpoiRelationBinder}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiExcelExporter,NpoiExportPlanBuilder}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
- **注释增量**：Abstractions/Core 范围新增 91 个字段/常量摘要并修订 2 个缓存摘要；Npoi 新增 17 个反射委托缓存/导入导出依赖/staging 状态字段摘要；另补充 Core 的 4 个校验/异常字段和 1 个默认日期格式常量摘要。合计新增 113 个字段/常量 `<summary>`，修订 2 个缓存键摘要；新增 `inheritdoc` 0 个。
- **范围复扫**：计划范围内的字段、`const`、`readonly`、`static readonly` 以及缓存键/配置键均已核对，未发现未注释候选（0）；private 字段无遗漏。缓存说明覆盖模型/方向/配置或工作簿/父子/键类型的隔离维度；阈值、容量、日期格式和异常 Data 键说明单位、默认值或作用阶段。字段摘要均为直接用途，未使用“获取或设置”；最长摘要未超过 120 个纯文本字符。
- **文案与重载检查**：本批只修改字段和常量文档，未改变方法签名或方法文档；既有重载 summary 一致性复核无新增问题。未发现应下沉而仍留在字段/常量 `<summary>` 中的冗长行为说明。
- **继承处理**：本批字段和常量没有可继承成员，不新增 `inheritdoc`；既有实现/override 的 `inheritdoc` 保持不变，当前生产源码扫描计数为 181 个。
- **范围与差异验证**：C# 差异按忽略行尾空白规则检查，非 XML 代码行数为 0；未修改业务逻辑、成员签名、访问级别、命名空间、公开 API、测试或项目配置。
- **构建与静态验证**：`dotnet build src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj -c Release --no-restore -v:minimal` 通过，联动 Abstractions/Core 的 netstandard2.0、net6.0 与 net8.0 目标均成功，0 警告、0 错误；`output/release` 下 8 个 XML 文档文件按 UTF-8 解析通过；`git diff --check` 退出码为 0（仅有 Git 的 LF/CRLF 提示）。
- **待确认项**：本批未发现需要业务语义确认的字段、常量或缓存键；`IsExternalInit.cs` 的排除归属仍待维护者确认。第 7 批最终全量审计、测试和生产符号到测试方法追溯尚未执行。

### 第 7 批执行记录（2026-09-15）

- **实际修改文件（43 个）**：
  - `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{ExcelMappingDocumentFactory,ExportMappingBuilder,ImportMappingBuilder}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exceptions/{BingOfficesExceptionDerived,BingOfficesExceptions}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{ExcelComment,ExcelSheetExportBuilder,ExcelWorkbookExportRequest}.cs`
  - `src/Bing.Offices.Abstractions/Bing/Offices/Imports/{ExcelImport,ExcelWorkbookImportResult}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Csv/{CsvEntityExporter,CsvEntityImporter,CsvHeaderBinder,CsvPipelineSupport}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/{Exceptions/BingOfficesExceptionDispatcher,Extensions/ProfileDescriptorFactory,Extensions/PropertyInfoExtensions,IO/AtomicFileCommitter}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/{ExcelPropertyMap,ExcelValidationBinding,ExcelValueConverterBindingResolver}.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/ExcelColumnPlan.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/ExcelHelper.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiExcelExporter,NpoiExportColumnPlanner,NpoiExportPlanBuilder,NpoiNonDisposingStream,NpoiStyleCache}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Extensions/{CellExtensions,SheetExtensions,SheetExtensions.MergedRegion,SheetExtensions.Picture}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/{NpoiExcelImporter,NpoiFailureWorkbookDiagnostics,NpoiFailureWorkbookFileSystem,NpoiFailureWorkbookPreflight,NpoiImportPlanBuilder,NpoiImportSheetExecutor,NpoiRelationBinder}.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Resolvers/ColorResolver.cs`
- **最终注释统计**：2,271 个目标声明全部覆盖，其中 `<summary>` 2,085 个、`<inheritdoc />` 186 个；现有 `<param>` 1,917 个、`<typeparam>` 123 个、`<returns>` 459 个。第 7 批补齐审计发现的嵌套类型、委托、构造函数、辅助方法和字段标签，并校正签名标签顺序及 override 文档；未新增行为测试。
- **全量复扫**：排除 `bin`、`obj`、生成文件、迁移/快照和 `IsExternalInit.cs` 后，188 个候选文件、2,271 个目标声明通过 Roslyn 审计；缺失摘要 0、`<param>` 不匹配 0、`<typeparam>` 不匹配 0、`<returns>` 缺失 0。112/112 个构造函数符合规定句式；属性访问前缀不匹配 0；summary 超过 120 个纯文本字符 0；override 缺少 `inheritdoc` 0；显式接口实现缺少 `inheritdoc` 0。
- **继承和文案复核**：实现类未复制上游契约的 summary、param 或 returns；`Stream` override 的额外资源所有权/编码行为均使用 `inheritdoc` 加 remarks 表达；重载核心 summary 保持一致，复杂约束下沉到参数或 remarks。编译器未报告无效 XML 或 `cref` 警告。
- **构建和测试**：`dotnet build Bing.Offices.sln -c Release --no-restore -v:minimal` 通过，0 警告、0 错误；核心测试 net6.0 与 net8.0 均 729/729 通过；集成测试 net6.0 与 net8.0 均 38/38 通过；`PublicApiContractTest` 双 TFM 均 9/9 通过；`Bing.Offices.Docs.Tests` net8.0 为 10/10 通过。`output/release` 下 9 个 XML 文件按 UTF-8 解析通过，其中生产程序集文档文件 8 个。
- **差异检查**：`git diff --check` 退出码为 0，仅保留 Git 的 LF/CRLF 提示；C# 差异中的非 XML 代码行数为 0，3 个历史无文件尾换行的文件仅补充了末尾换行。未修改业务逻辑、成员签名、访问级别、命名空间、公开 API、测试或项目配置。
- **最终生产符号 -> 测试方法追溯**：
  - `IExcelImporter.ImportAsync` / `NpoiExcelImporter.ImportAsync` -> `Bing.Offices.Tests.AsyncPipelineTest.ExcelAsync_ShouldUseAsyncOuterStream_AndMatchSyncImport`。
  - `IExcelExporter.ExportToFileAsync`、`ICsvExporter.ExportToFileAsync` -> `Bing.Offices.Tests.AsyncPipelineTest.AsyncFileExport_PreCanceled_ShouldNotCreateTarget`。
  - `ICsvImporter.ImportAsync` / `ICsvExporter.ExportAsync` -> `Bing.Offices.Tests.AsyncPipelineTest.CsvAsync_ShouldUseAsyncReadWrite_AndMatchSyncResult` 及取消场景测试。
  - `IFileExportCommitter.CommitAsync` / `DefaultFileExportCommitter` -> `Bing.Offices.Tests.DefaultFileExportCommitterTest` 的新目标、替换、失败、取消和临时文件清理测试。
  - Core/NPOI 公开扩展完整签名 -> `Bing.Offices.Tests.PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature`。
  - 公共 API 分类和双 TFM snapshot -> `Bing.Offices.Tests.PublicApiContractTest`；Markdown 文档示例 -> `Bing.Offices.Docs.Tests.DocsConsumerTest.DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually`。
- **待确认项**：无本批必须确认的业务语义；`src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs` 仍按兼容桩排除，是否纳入项目自有文档范围待维护者确认。
