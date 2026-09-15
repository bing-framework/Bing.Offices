# 中文 XML 注释补全计划

## 1. 目标和范围

本计划仅覆盖 `src/*` 下的生产 C# 源码。基线阶段只扫描、统计和制定实施顺序；执行阶段按下述批次逐批处理，不调整成员签名、业务逻辑、项目配置或公共 API。第 1-7 批均已完成。

纳入的实际项目如下：

- `src/Bing.Offices.Abstractions`：85 个候选源文件，重点是接口、抽象基类、公共请求/结果模型、Options、枚举和 Fluent Builder。
- `src/Bing.Offices.Core`：58 个候选源文件，重点是 CSV、映射编译、配置加载、校验、文件提交和反射辅助逻辑。
- `src/Bing.Offices.Npoi`：45 个候选源文件，重点是 NPOI 导入导出、工作簿校验、失败工作簿、异步 staging、扩展方法和缓存。

扫描对象包括类型、构造函数、具名方法、属性、索引器、事件、字段、常量及枚举成员，不因 `private`、`internal`、`protected`、`static`、`readonly` 或异步/泛型形式而省略。局部函数不强制 XML 注释；只有复杂或非直观逻辑才考虑普通中文注释。

扫描方法以 Roslyn 语法树为主，并辅以文本检索和代表性文件人工抽查。缺失统计按声明计数；同一字段声明包含多个变量时按变量数计数。文案、标签和 `inheritdoc` 统计属于候选项，实施时必须结合继承关系和真实语义逐项确认。

## 2. 排除范围

- 构建产物：`bin/**`、`obj/**`。
- 生成文件：`*.g.cs`、`*.generated.cs`、`*.Designer.cs`；当前实际排除 `src/Bing.Offices.Core/Bing/Offices/Resources.Designer.cs`。
- EF Core `Migration`/`Migrations`、`ModelSnapshot`。
- 自动生成客户端、代理代码和第三方源码。
- `src/Bing.Offices.Abstractions/System/Runtime/CompilerServices/IsExternalInit.cs`：按平台兼容桩处理，不纳入业务注释补全；若维护者确认该文件属于项目自有公共文档范围，再单独纳入。
- `tests/**`、`benchmarks/**`、`build/**` 及 `docs/**` 不属于注释补全范围，仅可用于理解行为和执行验证。

物理扫描发现 `src` 下共有 190 个 `.cs` 文件；排除上述 `Resources.Designer.cs` 和 `IsExternalInit.cs` 后，计划候选文件为 188 个。三个 `AssemblyInfo.cs` 没有本计划目标成员，保留在文件范围中但不会产生实施项。

## 3. 扫描统计

### 3.1 按项目统计

| 项目 | 候选文件 | 目标声明 | 缺少 XML 注释 | 缺口占比 |
| --- | ---: | ---: | ---: | ---: |
| `Bing.Offices.Abstractions` | 85 | 995 | 130 | 13.1% |
| `Bing.Offices.Core` | 58 | 650 | 102 | 15.7% |
| `Bing.Offices.Npoi` | 45 | 626 | 123 | 19.6% |
| **合计** | **188** | **2,271** | **355** | **15.6%** |

缺失项分布在 57 个文件中；1,916 个目标声明已存在某种 XML 文档。这里的“已存在”只表示声明前有 XML 文档，不代表内容准确、完整或符合本计划规范。

### 3.2 按成员种类统计

| 成员种类 | 声明总数 | 缺少注释 | 主要缺口 |
| --- | ---: | ---: | --- |
| 类型 | 294 | 39 | internal/private 辅助类型、嵌套类型 |
| 构造函数 | 112 | 26 | internal 17、private 6、public 3 |
| 方法 | 741 | 140 | private 67、public 38、internal 24、隐式访问 8、protected 3 |
| 属性 | 716 | 33 | public 19、internal 5、隐式访问 5、private 4 |
| 索引器 | 0 | 0 | 当前未发现 |
| 事件 | 0 | 0 | 当前未发现 |
| 字段（不含常量） | 184 | 113 | private 112、internal 1 |
| 常量 | 41 | 1 | private 配置/格式常量 |
| 枚举成员 | 183 | 3 | `NpoiAsyncStagingStrategy` 等内部枚举 |

### 3.3 文案和标签候选统计

| 候选问题 | 数量 | 说明 |
| --- | ---: | --- |
| 已注释构造函数未使用规定句式 | 73 / 86 | 必须按当前类型名逐项统一，不能只替换“创建”“使用……初始化”等词 |
| 属性摘要与访问方式用语不匹配 | 314 / 683 | 按 `get`、`set`、`get/set`、`get/init` 复核；表达式体属性按只读处理 |
| override 或显式接口成员未使用 `inheritdoc` | 44 / 72 | 语法级候选；先确认上游契约是否已有有效注释 |
| 已使用 `<inheritdoc />` | 126 | 需反查上游成员，排除无有效继承目标或额外行为未说明的情况 |
| `<param>` 名称/数量与签名不一致 | 191 | 主要是已有 summary 但缺少参数标签；也需排除真实错名和多余标签 |
| `<typeparam>` 名称/数量与签名不一致 | 29 | Fluent Builder 和泛型导入导出入口较集中 |
| `<returns>` 与返回类型规则不一致 | 165 | 非 `void`、非裸 `Task`/`ValueTask` 方法应有返回说明；继承成员除外 |
| summary 超过 120 个纯文本字符 | 1 | 只是冗长说明的自动筛选下限，短于阈值仍需人工判断是否应移入 `<remarks>` |

上述文案和标签数量可能互相重叠，不能相加为问题总数。隐式接口实现无法仅凭局部语法可靠判断，必须在第 4 批通过编译符号或接口对照补充审计。

独立文本扫描还将 `src/Bing.Offices.Core/Bing/Offices/Attributes/DataFormatAttribute.cs` 中 `BuiltinFormat`、`CustomFormat` 的两处 summary 标记为人工复核项：其中包含外部链接和格式背景说明，即使字符数未全部超过自动阈值，也应保留核心用途并把扩展资料下沉到 `<remarks>`。

## 4. 问题分类

1. **确定缺失**：355 个声明没有 XML 文档，最大来源是字段、非公开辅助方法和内部嵌套类型。
2. **公共契约不完整**：`ExcelImport.cs`、`ExcelSheetExportBuilder.cs`、`ImportMappingBuilder.cs`、`ExportMappingBuilder.cs` 等公开 Fluent API 多数只有 summary，缺少与签名一致的 `<param>`、`<typeparam>`、`<returns>`。
3. **DTO/Options/枚举不完整**：请求、结果、资源限制、失败策略、配置模型和内部枚举仍有类型、属性或枚举成员缺口，默认值、单位、可空性和策略含义需要补清。
4. **文案不统一**：大量构造函数使用“创建……”“使用……初始化……”而非规定句式；部分属性使用“是否……”“表示……”等名词性摘要，未体现访问方式。
5. **继承文档不规范**：`Stream` 包装器、异常类型、接口实现和校验规则中存在 override/实现成员无 `inheritdoc`、复制上游说明或未说明差异行为的候选。
6. **辅助成员缺失**：`CsvPipelineSupport.cs`、`NpoiAsyncStaging.cs`、导入导出执行器和配置校验器中的 private/internal/static 方法及字段最集中。
7. **标签语义不完整**：除标签名称/数量外，还要检查 `bool` 的 true/false 语义、可空返回条件、异步结果措辞、范围/单位以及实际传播异常。
8. **机械化或过时注释**：旧式“Excel格式”“错误消息”“获取或设置”空泛说明，以及重复成员名但未说明业务用途的注释，应在补全时一起修订。

## 5. 文案规范问题

- 构造函数 summary 统一为：`初始化一个 <see cref="当前类型" /> 类型的实例。`。泛型类型使用可解析的 `cref` 写法；参数职责写入 `<param>`，不要塞入 summary。
- 只读、只写、可读写和 init-only 属性分别以“获取……”“设置……”“获取或设置……”“获取或初始化……”开头。摘要应补充业务含义，不只复述属性名。
- 方法 summary 只说明核心用途；缓存策略、线程安全、流所有权、取消后状态、文件提交原子性、资源上限和兼容性原因移入 `<remarks>`。
- 语义一致的同步/异步、泛型/非泛型及不同输入载体重载保持相同核心措辞，只在输入形式和返回类型上表达差异。
- `Task`/`ValueTask` 说明异步操作但不添加 `<returns>`；`Task<T>`/`ValueTask<T>` 说明最终结果。`bool` 返回值必须解释 true/false，可空返回值必须说明返回 `null` 的条件。
- `<param>` 说明含义、格式、范围、单位和空值规则；`<typeparam>` 说明类型职责与约束。只记录代码明确抛出或稳定传播、且调用方需要处理的 `<exception>`。
- 优先使用合法的 `<see cref="..." />`、`<paramref name="..." />`、`<typeparamref name="..." />`；文本中的 XML 特殊字符必须转义。
- 字段、常量、缓存键和配置键要说明作用域、生命周期、隔离维度、默认值或用途，禁止“缓存键”“用户 ID”一类机械复述。

重点人工复核文件：

- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/ExcelImport.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/ExcelSheetExportBuilder.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ImportMappingBuilder.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/ExportMappingBuilder.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/ExcelHelper.cs`
- `src/Bing.Offices.Core/Bing/Offices/Csv/CsvPipelineSupport.cs`

## 6. inheritdoc 使用问题

- 接口实现、显式接口实现、抽象成员实现和 `override`，在上游已有准确文档时优先只写 `/// <inheritdoc />`，不得复制 `<summary>`、`<param>` 和 `<returns>`。
- 上游契约缺少或内容失真时，先在第 1 批完善接口/抽象成员，再处理实现；不得用 `inheritdoc` 掩盖空白上游契约。
- 实现有缓存、资源所有权、额外验证、异常转换、取消边界或同步/异步差异时，使用 `<inheritdoc />` 后追加 `<remarks>`，只描述差异。
- 重点复核 44 个 override/显式实现候选，以及无法由语法扫描识别的隐式接口实现。高风险位置包括：
  - `src/Bing.Offices.Core/Bing/Offices/Csv/CsvPipelineSupport.cs`
  - `src/Bing.Offices.Core/Bing/Offices/IO/AtomicFileCommitter.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Validations/ExcelValidationRules.cs`
  - `src/Bing.Offices.Core/Bing/Offices/Mappings/ExcelMappingPlan*.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiNonDisposingStream.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiFailureWorkbookSerialization.cs`
  - `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs`
- 对现有 126 个 `<inheritdoc />` 逐项确认可解析的上游目标；特别检查显式接口成员、嵌套 `Stream` 实现和同名重载，避免继承到错误或空文档。

## 7. 实施批次

### 第 1 批：接口、抽象类和公共契约

目标：先建立可被实现继承的完整上游文档，补齐公开 Fluent API 的标签与返回语义。

涉及目录或文件：

- `src/Bing.Offices.Abstractions/Bing/Offices/{Csv,Exports,Imports,IO,Providers,Conversions,Validations}/I*.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/{IExcelMappingConfigurationLoader,IMappingProfileRegistry,IMappingProfileResolver,MappingProfileContracts}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Attributes/{FilterAttributeBase,DecoratorAttributeBase}.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/{Imports/ExcelImport,Exports/ExcelExport,Exports/ExcelSheetExportBuilder,Configurations/ImportMappingBuilder,Configurations/ExportMappingBuilder,Configurations/FluentSetting}.cs`
- `src/Bing.Offices.Core/Bing/Offices/Mappings/IExcelCompiledMappingColumn.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/IO/NpoiAsyncStaging.cs` 中的内部 staging 接口。

完成条件：接口、抽象成员和公开调用入口的 summary、泛型参数、参数、返回值和必要异常完整；形成“上游契约 -> 实现类”的 inheritdoc 对照表。

### 第 2 批：DTO、Entity、Options 和枚举

目标：补齐数据载体、配置对象、请求/结果、错误模型、策略枚举及全部枚举成员，明确默认值、单位、范围和可空性。本仓库未发现独立 Entity 层；按 DTO/Options/元数据模型执行。

涉及目录或文件：

- `src/Bing.Offices.Abstractions/Bing/Offices/{Configurations,Csv,Exports,Imports,Styles,Conversions}/`
- `src/Bing.Offices.Core/Bing/Offices/{Metadata,Mappings}/`
- `src/Bing.Offices.Npoi/Bing/Offices/{ExcelColumnPlan.cs,Imports/ExcelImportExecutionOptions.cs,Imports/ExcelImportRuntime.cs,Imports/ExcelImportErrorCollector.cs,Imports/ValidationRangeIndex.cs}`
- 枚举重点：`NpoiAsyncStagingStrategy`、导入失败/校验/资源策略、CSV 公式注入策略、映射合并模式和 Excel 样式枚举。

完成条件：所有 DTO/Options/枚举类型及成员均有准确说明；布尔值、默认值、字节/行列上限和 `null` 语义可由文档独立理解。

### 第 3 批：构造函数和属性文案统一

目标：处理 26 个缺失构造函数、73 个构造函数句式候选、33 个缺失属性及 314 个属性访问措辞候选。

涉及目录或文件：

- 全量复扫三个 `src` 项目的构造函数和属性。
- 优先处理 `Bing.Offices.Abstractions` 的 Exceptions、Imports、Exports、Styles。
- 优先处理 `Bing.Offices.Core` 的 Attributes、Csv、Mappings、Metadata。
- 优先处理 `Bing.Offices.Npoi` 的 Imports、Exports、IO。

完成条件：构造函数使用规定句式；属性按访问器使用规定动词；参数说明不挤入 summary；表达式体和 init-only 属性分类正确。

### 第 4 批：实现类、重写成员和重载方法

目标：基于第 1 批上游契约处理接口实现、显式接口实现、抽象实现、override 和语义一致的重载，消除重复契约文案。

涉及目录或文件：

- `src/Bing.Offices.Core/Bing/Offices/{Csv,IO,Mappings,Validations}/`
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMappingConfigurationLoader.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/{Imports/NpoiExcelImporter.cs,Exports/NpoiExcelExporter.cs}`
- `src/Bing.Offices.Npoi/Bing/Offices/{Exports/NpoiNonDisposingStream.cs,Imports/NpoiFailureWorkbookSerialization.cs,IO/NpoiAsyncStaging.cs}`
- `src/Bing.Offices.Npoi/Bing/Offices/Extensions/*.cs` 中语义一致的重载族。

完成条件：可继承成员正确使用 `<inheritdoc />`；差异行为位于 `<remarks>`；同步/异步和不同格式重载的核心 summary 高度一致。

### 第 5 批：private、internal、static 辅助方法

目标：补齐 140 个缺失方法，优先处理 67 个 private、24 个 internal 和 3 个 protected 方法；对 public 缺口中不属于前四批的静态扩展入口一并处理。

涉及目录或文件：

- `src/Bing.Offices.Core/Bing/Offices/Csv/{CsvPipelineSupport,CsvEntityImporter,CsvEntityExporter}.cs`
- `src/Bing.Offices.Core/Bing/Offices/{Configurations,Dates,Extensions,IO,Mappings}/`
- `src/Bing.Offices.Npoi/Bing/Offices/{Imports,Exports,IO,Extensions}/`
- 热点文件：`NpoiAsyncStaging.cs`、`NpoiImportSheetExecutor.cs`、`NpoiXlsxZipPreflight.cs`、`SheetExtensions.Picture.cs`。

完成条件：所有具名辅助方法有核心用途说明和签名一致的标签；本地函数不机械添加 XML 注释；异步方法准确描述结果和取消行为。

### 第 6 批：private 字段、其他字段、常量、缓存键和配置键

目标：补齐 113 个字段和 1 个常量缺口，并复核已有字段/键注释是否说明生命周期、隔离和边界。

涉及目录或文件：

- `src/Bing.Offices.Abstractions/Bing/Offices/{Imports/ExcelImport.cs,Exports/ExcelExport.cs,Exports/ExcelSheetExportBuilder.cs,Providers/UniqueTracker.cs,Configurations/*.cs}`
- `src/Bing.Offices.Core/Bing/Offices/{Csv/CsvPipelineSupport.cs,Mappings/ExcelMappingPlanFactory.cs,Mappings/ExcelMappingPlanCacheKey.cs,Validations/ExcelValidationRules.cs,Configurations/*.cs,IO/AtomicFileCommitter.cs}`
- `src/Bing.Offices.Npoi/Bing/Offices/{Imports,Exports,IO}/`，重点是导入导出依赖字段、staging 状态、样式缓存和反射委托缓存。
- 常量/键重点：文档大小与深度、列/别名/校验数量、NPOI 运算符映射、日期格式、异常 Data 键、缓存容量和工作簿格式上限。

完成条件：所有字段、`const`、`readonly`、`static readonly` 均有准确 XML 注释；缓存键说明隔离维度，配置键/上限说明单位和作用阶段，依赖字段说明职责而非只写类型名。

### 第 7 批：最终审计和构建验证

目标：全量复扫 188 个候选文件，清零确定缺失，并人工关闭所有文案、标签和继承候选。

实施清单：

- 重新生成按“最终生产符号 -> 测试方法”的成员级追溯映射；注释任务不新增行为测试，但必须证明现有职责测试覆盖未因误改而变化。
- 检查最终 diff，只允许 `.cs` 内 XML/必要普通注释以及本计划/执行记录变化，不允许业务代码、签名、命名空间、可见性或项目配置变化。
- 编译三个生产项目和解决方案，检查 XML 文档语法、无效 `cref`、缺失标签及双 TFM 结果。
- 运行核心与集成测试，并执行公共 API snapshot 测试，确认注释变更未改变二进制契约。

## 8. 风险和待确认项

- **扫描精度**：缺失项来自语法树，可信度高；属性用语、标签和 inheritdoc 是静态候选，受条件编译、隐式接口实现、重载解析和上游依赖文档影响，实施前必须语义复核。
- **兼容桩归属**：`IsExternalInit.cs` 当前按平台兼容/第三方样板排除；需维护者确认是否希望项目自有化并补注释。
- **公共契约优先级**：Fluent Builder 同时属于公共契约和构建 DTO。为避免重复修改，方法在第 1 批完成，状态字段在第 6 批完成。
- **历史文案语义**：部分 NPOI 扩展、失败工作簿和流包装器的资源所有权/异常语义不能只从签名判断，应结合调用方和测试，不编造行为。
- **inheritdoc 可解析性**：第三方基类（尤其 `Stream`）的文档是否随目标框架可用可能不同；必要时保留独立说明，但必须记录不使用 inheritdoc 的理由。
- **生成文档警告**：`common.props` 已在 Debug/Release 生成 XML 文档。注释补全可能暴露无效 XML、错误 `cref` 或缺失标签警告，不能通过新增 `NoWarn` 规避。
- **文件规模与冲突**：`ExcelImport.cs`、`CsvPipelineSupport.cs`、`NpoiAsyncStaging.cs` 等热点文件修改量大，建议一批一提交或一批一审查，避免和业务开发交叉。
- **范围边界**：本计划不授权修改测试、快照、迁移记录、项目配置或业务代码；若实施中发现真实 API/行为问题，应另建任务。

## 9. 验证命令

以下命令在实施完成后于仓库根目录执行。PowerShell 会话先固定 UTF-8：

```powershell
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

dotnet restore Bing.Offices.sln
dotnet build Bing.Offices.sln -c Release --no-restore

dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-build
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net6.0 --no-build
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-build

dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net6.0 --no-build --filter "FullyQualifiedName~PublicApiContractTest"
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -c Release -f net8.0 --no-build --filter "FullyQualifiedName~PublicApiContractTest"
```

附加静态核验：

```powershell
git diff --check
git diff -- src
rg -n "<inheritdoc\s*/>|<summary>|<param|<typeparam|<returns|<remarks" src -g "*.cs" -g "!**/bin/**" -g "!**/obj/**" -g "!*.g.cs" -g "!*.generated.cs" -g "!*.Designer.cs" -g "!**/Migrations/**" -g "!*ModelSnapshot.cs"
```

若默认输出被其他进程占用，使用唯一的 `-p:OutputPath=<临时目录>` 隔离构建，不清理或覆盖现有输出目录。构建后确认 `Bing.Offices.Abstractions.xml`、`Bing.Offices.Core.xml`、`Bing.Offices.Npoi.xml` 均生成且 XML 可解析。

## 10. 验收标准

- 188 个候选文件完成复扫，355 个确定缺失项清零，或每个保留项都有文件、符号和理由明确的审计记录。
- 类型、构造函数、方法、属性、字段、常量和枚举成员完整；当前不存在索引器和事件，后续新增时仍按同一规则检查。
- 公开契约先完整，再处理实现；接口实现、显式接口实现、抽象实现和 override 正确使用 `<inheritdoc />`，没有重复复制上游文档。
- 所有构造函数 summary 符合规定句式；属性按访问方式使用“获取”“设置”“获取或设置”“获取或初始化”。
- 方法 summary 只说明核心用途；调用方需要关注的复杂约束位于 `<remarks>`；语义一致的重载文案一致或高度一致。
- `<param>`、`<typeparam>`、`<returns>` 与实际签名和返回语义一致；无多余、错名、遗漏标签，无无效 XML 或无法解析的 `cref`。
- 字段、常量、缓存键和配置键说明用途、范围、生命周期、单位或隔离维度，不存在机械化、错误、过时、冗长或无意义的注释。
- 最终 diff 不包含业务逻辑、公共 API、签名、命名空间、可见性或项目配置变更。
- Release 构建、核心测试、双 TFM 集成测试和公共 API snapshot 测试通过；三个生产程序集的 XML 文档文件生成且可解析。
- 最终执行记录包含“最终生产符号 -> 测试方法”的追溯映射，以及每个批次的文件清单、复扫结果和验证结果。

## 11. 执行记录

当前状态：**第 1-7 批已完成**。

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
