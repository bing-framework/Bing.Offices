# Bing.Offices 系统架构与设计文档

> 本文档描述 `Bing.Offices` 当前主干（2.0.0 包身份）的系统架构、核心模块与数据流。
> 阅读对象：类库维护者、二次开发 Provider 的工程师、以及需要在生产环境部署 Excel/CSV 导入导出的应用架构师。
> 配套图示见 [architecture.html](architecture.html)（可交互分层图、导入/导出时序图、映射编译流水线）。

---

## 1. 系统定位与边界

### 1.1 定位

`Bing.Offices` 是 Bing 应用框架中的 **Excel / CSV 导入导出引擎**。它以「Workbook Request」为中心 API，把「实体模型 ↔ 表格文档」的映射、校验、转换、布局、样式、错误收集与文件提交，封装成一套与具体表格库解耦的管线。

### 1.2 能力范围

**当前交付：**

- Excel 导入 / 导出（XLS、XLSX）
- CSV 导入 / 导出
- 显式注册的 XLSX 低内存服务：SXSSF 导出、ZIP/XML 逐行导入
- Workbook Request：异构多 Sheet、动态列、模板、命名区域、样式、结构化错误
- XLSX 柱状图、折线图、饼图
- Sync / Async 双 API，共享 Mapping / Validation / Conversion / 错误结果 / 取消合同
- 文件原子提交（临时文件 → flush → 取消复查 → 替换目标）

**明确不在范围内：**

- Word、PDF（不提供相关程序集）
- 零 GC 或进程内存固定上限承诺
- Excel Async 的 DOM 阶段完全异步

### 1.3 关键设计约束（Non-Goals → Design Rules）

这些约束是本项目大量设计决策的根源，理解它们才能理解后面的模块划分：

| # | 约束 | 由此产生的设计 |
|---|---|---|
| C1 | NPOI `WorkbookFactory.Create` / DOM 操作 / `workbook.Write` **没有异步 API** | Async API 只异步化外围文件/Stream IO；生产代码禁止 `Task.Run` / `.Result` / `.Wait` 伪装异步 |
| C2 | 既有完整 Workbook API 使用 DOM；新增低内存 API 独立注册 | 普通 XLSX 使用 SXSSF 窗口／ZIP XML 逐行处理；模板等高级功能保留 DOM；均使用有限资源配置 |
| C3 | 模板流在 `WorkbookFactory.Create` 阶段必须**同步可读** | async-only 模板流不属于支持的输入合同 |
| C4 | Provider 可替换（NPOI 只是其中之一） | 契约与引擎下沉到 `netstandard2.0` 的 Abstractions / Core，NPOI 细节全部关在 Npoi 程序集内 |
| C5 | 映射配置来源多、优先级复杂 | 四级合并 + 单一「Plan 编译器」作为唯一生效点 |

---

## 2. 分层架构

系统分三层，依赖方向严格单向：**Npoi → Core → Abstractions**。

```
┌───────────────────────────────────────────────────────────────┐
│  调用方（ASP.NET Core / 控制台 / 后台任务）                     │
│  services.AddBingOfficesNpoi()                                │
└───────────────────────────┬───────────────────────────────────┘
                            │ IExcelImporter / IExcelExporter
                            │ ICsvImporter  / ICsvExporter
┌───────────────────────────▼───────────────────────────────────┐
│  L2  Bing.Offices.Npoi            net6.0; net8.0              │
│      Provider 适配层（唯一依赖 NPOI 的程序集）                   │
│      ├─ Imports/   导入执行、行物化、关系绑定、失败工作簿          │
│      ├─ Exports/   导出执行、列规划、Sheet 写入、样式缓存          │
│      ├─ Imports/NpoiXlsxZipPreflight  ZIP/XML 资源预检          │
│      ├─ Extensions/ 七个用户扩展容器 + DI 注册                   │
│      └─ Resolvers/  ColorResolver 等 NPOI 具体化解析            │
└───────────────────────────┬───────────────────────────────────┘
                            │ IExcelMappingPlanFactory
                            │ IExcelMappingPlan / IExcelMappingWorkbookPlan
┌───────────────────────────▼───────────────────────────────────┐
│  L1  Bing.Offices.Core            netstandard2.0              │
│      Provider-neutral 引擎层                                   │
│      ├─ Mappings/       Plan 编译器 + 计划缓存 + TypeMap 反射缓存 │
│      ├─ Configurations/ JSON/XML 文档加载与校验                  │
│      ├─ Csv/            CSV 管线（CsvHelper）                  │
│      ├─ IO/             AtomicFileCommitter 原子文件提交        │
│      ├─ Validations/    默认 7 条校验规则                       │
│      ├─ Datas/ Dates/   日期解析、类型转换辅助                   │
│      └─ Extensions/     ExcelMappingPlanFactoryProvider（组合根）│
└───────────────────────────┬───────────────────────────────────┘
                            │ 契约、DTO、策略、异常
┌───────────────────────────▼───────────────────────────────────┐
│  L0  Bing.Offices.Abstractions    netstandard2.0              │
│      契约层（无第三方实现依赖）                                  │
│      ├─ Imports/  Exports/  Csv/      请求、结果、策略、错误码    │
│      ├─ Configurations/  Fluent Builder、Document、Merger、     │
│      │                   Profile Registry                      │
│      ├─ Providers/       IExcelMappingPlan、UniqueTracker      │
│      ├─ Conversions/     IExcelValueConverter、ExcelCellValue   │
│      ├─ Validations/     IExcelValidationRule、Binding         │
│      ├─ IO/              IFileExportCommitter                  │
│      ├─ Exceptions/      异常基类与错误码三元组                  │
│      └─ Styles/          provider-neutral 样式模型             │
└───────────────────────────────────────────────────────────────┘
```

### 2.1 分层职责与准入规则

| 层 | 允许依赖 | 禁止事项 |
|---|---|---|
| Abstractions | 仅 BCL | 引用 NPOI / CsvHelper 等具体实现；包含业务执行逻辑 |
| Core | Abstractions + Bing.Utils + CsvHelper + DI.Abstractions | 引用任何 NPOI 类型 |
| Npoi | Abstractions + Core + NPOI + DI.Abstractions | 在 abstrations/core 里泄漏 `IWorkbook` / `ISheet` 等 NPOI 类型 |

**关键设计点：** NPOI 的类型（`IWorkbook`、`ISheet`、`ICell`）**没有出现在任何公共契约中**。Provider 只能看到 `IExcelMappingPlan` / `IExcelMappingColumn` / `ExcelCellValue` 这类 provider-neutral 视图。这是「可换 Provider」这一目标的全部实现手段。

### 2.2 构建矩阵

| 程序集 | TFM | 关键依赖 |
|---|---|---|
| `Bing.Offices.Abstractions` | `netstandard2.0` | 无第三方依赖 |
| `Bing.Offices.Core` | `netstandard2.0` | `Bing.Utils`、`CsvHelper 33.1.0`、`Microsoft.Extensions.DependencyInjection.Abstractions 2.2.0` |
| `Bing.Offices.Npoi` | `net6.0`、`net8.0` | `NPOI`、`Microsoft.Extensions.DependencyInjection.Abstractions`（6.0.0 / 8.0.0，按 TFM 条件引用） |

Abstractions / Core 保持 `netstandard2.0` 是为了让 .NET Framework 与 .NET Core 各版本都能消费契约与引擎；只有 NPOI 适配层需要现代化 TFM。

---

## 3. 核心模块

### 3.1 模块总览

| 模块 | 主要载体 | 一句话职责 |
|---|---|---|
| M1 配置（Mapping Configuration） | `Configurations/`（Abstractions + Core） | 把 4 类来源解析为统一的 `ExcelMappingConfiguration` / `ExcelMappingDocument` |
| M2 计划编译（Plan Compiler） | `Core/Mappings/` | 合并配置 + 反射模型 → 不可变 `IExcelMappingPlan`，并缓存 |
| M3 校验（Validation） | `Abstractions/Validations`、`Core/Validations`、`Npoi/Imports` | 7 条内置规则 + Workbook 原生 DataValidation + 唯一性追踪 |
| M4 转换（Conversion） | `Abstractions/Conversions`、`Core/Mappings/ExcelValueConverterBindingResolver` | 单元格值 ↔ CLR 值的双向转换与绑定 |
| M5 导入执行（Import） | `Npoi/Imports/` | 输入防护 → DOM → 逐 Sheet 执行 → 错误收集 → 关系绑定 → 失败工作簿 |
| M6 导出执行（Export） | `Npoi/Exports/` | 计划 → 列规划 → 逐 Sheet 写入 → 样式 → 图表 → 序列化 |
| M7 CSV | `Core/Csv/`、`Abstractions/Csv/` | 独立于 Excel 的文本表格管线，复用 M1–M4 |
| M8 IO 与提交 | `Core/IO/`、`Npoi/IO/` | 原子文件提交、异步暂存、流复制与限额 |
| M9 异常与诊断 | `Abstractions/Exceptions`、`Core/Exceptions` | 错误码三元组 + Observer 分发 |
| M10 组合根 | `ExcelMappingPlanFactoryProvider`、`ExcelNpoiServiceCollectionExtensions` | 装配全部默认实现 |

---

### 3.2 M1 配置模块

**五类来源与优先级**（低 → 高）：

```
Convention/Default  <  Attribute  <  Profile  <  JSON/XML Document  <  Request Fluent
   (约定回退)        (特性标注)    (强类型Profile)   (外部配置文档)        (调用方代码)
```

| 来源 | 解析入口 | 产出 |
|---|---|---|
| Attribute | `ExcelTypeMapFactory.Create<T>()` 反射 `ColumnNameAttribute` / `DataFormatAttribute` / `ExcelIgnoreAttribute` / `ValueMappingAttribute` / `DynamicColumnAttribute` | `ExcelTypeMap<T>` |
| Profile | `ProfileDescriptorFactory.Create()` → `ImportMappingBuilder` / `ExportMappingBuilder` / `FluentSetting.Build(MappingSourceKind.Profile)` | `ProfileDescriptor` |
| JSON / XML | `ExcelMappingConfigurationLoader.FromJsonDocument` / `FromXmlDocument` | `ExcelMappingDocument` |
| Request Fluent | 调用方通过 `ImportMappingBuilder` / `ExportMappingBuilder` 构建 | `ExcelMappingConfiguration` |

**关键类型：**

- `ExcelMappingDocument`：规范化文档，含 `TenantId`、`ConfigurationVersion`、模型别名，以及按方向（`MappingDirection`）分离的配置。
- `MappingConfigurationMerger.Merge(lower, higher, sourceKind)`：**唯一的合并算法实现**。高优先级覆盖低优先级，并处理 `Clear` / `Reset` 标记、`MergeColumn`、动态列 `Append` / `Replace`、`MergeStyle` / `MergeLayout`。
- `MappingProfileRegistry`：键为 `(ProfileName, Direction, ModelType)` 的 `Dictionary` + `lock`。重复注册**直接抛异常**，不由 DI 注册顺序覆盖。同时实现 `IMappingProfileRegistry` 与只读的 `IMappingProfileResolver`。
- `ExcelModelAliasRegistry`：`ConcurrentDictionary<alias, (modelType, profileName)>` 白名单，被 Plan 编译器与文档校验器共同使用。
- `MappingConfigurationCloner` / `MappingDocumentCloner`：构建后的配置是**独立快照**，修改输入集合不会影响已构建映射。

**DI 注册：** `MappingProfileServiceCollectionExtensions.AddMappingProfile<T>()` / `AddMappingProfiles(assembly)` —— 校验契约 → 扫描去重 → 以 Singleton 注册。

**JSON/XML 加载与安全：**

`ExcelMappingConfigurationLoader`（门面）→ 流重载先经 `ExcelMappingTextReader.ReadLimitedText`（限制 1 MB）→ 分派：

- JSON：`JsonDocument.Parse`（`MaxDepth = 32`）+ `Validator.ValidateJsonElement` + `JsonSerializer`
- XML：`Validator.ValidateXmlShape` + `XmlSerializer`，且 `DtdProcessing = Prohibit`、`XmlResolver = null`，未知节点/属性直接抛 `XmlMappingValidationException`

两者最终都进入 `ExcelMappingDocumentValidator.ValidateDocument`：schema 版本必须为 2、列/别名/校验数量上限、字符串 ≤ 4096、业务别名禁用 CLR 名、`placementKey` 格式、`modelAlias` 白名单。

`DefaultExcelMappingConfigurationLoader` 负责把 `BingOfficesException` 转发给 `IBingOfficesExceptionObserver`。

---

### 3.3 M2 计划编译模块

这是整个系统的心脏：**唯一把「配置 + 模型反射」固化为「可执行不可变计划」的地方**。

```
ExcelMappingPlanFactoryProvider.CreateDefault(...)        ← public 静态门面 / 组合根
        │  (RegisterDefault 用 TryAddSingleton 注册为进程级 Singleton)
        ▼
ExcelMappingPlanFactory : IExcelMappingPlanFactory        ← internal，含缓存
        │
        ├── ResolveDocument<T>()
        │     document 方向配置
        │       └─ MappingConfigurationMerger.Merge(profileDescriptor.Configuration, ...)
        │            └─ MappingConfigurationMerger.Merge(configuration, requestConfiguration, Request)
        │
        ├── CreatePlan<T>()
        │     ExcelTypeMapFactory.Get<T>(configuration)     ← 静态 ConcurrentDictionary<Type, object>
        │       └─ ApplyConfiguration：按列索引/标题/别名/值映射/校验名覆盖，校验重复标题
        │     └─ 遍历 Properties → CreateColumn / CreateDynamicColumn
        │           ├─ ExcelMappingColumn          （固定列）
        │           └─ ExcelDynamicMappingColumn   （动态列）
        │                 ├─ ExcelValueConverterBindingResolver.Resolve(...) → ValueConverters
        │                 └─ BindValidationRules(...) → ValidationBindings
        │
        └── CreateWorkbook<T>() → ExcelMappingWorkbookPlan
              └─ ExcelMappingSheetPlan(name, mapping)  × N
```

**类型职责表：**

| 类型 | 职责 |
|---|---|
| `IExcelMappingPlanFactory` | Provider-neutral 契约，4 个 `Create` / `CreateWorkbook` 重载（含/不含 request 配置） |
| `ExcelMappingPlanFactory` | 合并 → 编译 → 缓存 |
| `ExcelTypeMapFactory` | 静态反射缓存，把 CLR 类型编译为 `ExcelTypeMap<T>` |
| `ExcelTypeMap<T>` | 只读 `IReadOnlyList<ExcelPropertyMap>` |
| `ExcelPropertyMap` | 单属性的不可变映射：Title / Aliases / Ignored / ValueMap / ConverterName / ValidationRuleNames / 编译后的 `Getter` / `Setter` / `ImageMultiplicity` |
| `ExcelMappingColumn` | 运行时固定列，同时实现 `IExcelMappingColumn` + `IExcelCompiledMappingColumn` |
| `IExcelCompiledMappingColumn` | 暴露编译后的 `Property` / `Getter` / `Setter` / `Attributes`，供 Provider 反射取值落值 |
| `ExcelMappingWorkbookPlan` | 按 Sheet 顺序组织的 `ExcelMappingSheetPlan` 集合 |
| `IExcelMappingPlan` | Provider 看到的最终视图：`Columns` / `DynamicColumns` / `Style` / `Layout` / `ProfileName` / `ModelAlias` |

**计划缓存（重要）：**

- 结构：`ConcurrentDictionary<string, Lazy<IExcelMappingPlan>>` + `ConcurrentQueue<string> _cacheOrder`
- 并发安全：`Lazy` 使用 `ExecutionAndPublication`，保证同键并发只编译一次
- 容量：默认 **256**，`TrimCache()` 按近似 FIFO 淘汰
- 生命周期：工厂注册为 **Singleton**（进程级）
- **无 TTL、无主动失效** —— 靠缓存键变化自然分代
- 缓存键 `ExcelMappingPlanCacheKey.Create<T>`：`TenantId` + `ModelType.AssemblyQualifiedName` + `Direction` + `ConfigurationVersion` + 合并后 `Configuration`，经 `JsonSerializer(DefaultIgnoreCondition = Never)` 序列化后 **SHA-256 → Base64**

> 注意：Npoi 侧的 `NpoiWorkbookPlanKeyBuilder` **不是缓存**，它只用于 `NpoiImportPlanBuilder` / `NpoiExportPlanBuilder` 的 `GroupBy` 分组（键 = `itemType.AssemblyQualifiedName` + 文档 JSON + 单方向 request 配置 JSON）。两者容易混淆。

---

### 3.4 M3 校验模块

校验分为 6 个层次，跨越配置期到执行期：

| 层 | 时机 | 载荷 | 失败表现 |
|---|---|---|---|
| L0 配置构造期 | `ExcelImport.Build()` | `ResourceLimits.Validate` / `FailureOptions.Validate`、枚举与选择器唯一性 | `ArgumentOutOfRangeException` 等 |
| L1 计划期 | `NpoiImportPlanBuilder` / `CreateColumns` | 映射合法性、动态列定义、只读属性 | `BingOfficesConfigurationException`(Plan) |
| L2 结构/表头 | `NpoiImportSheetExecutor.CreateColumns` | 重复列、缺列、未知动态列、超 `MaxReadColumns` | `NpoiSheetStructureException` → 转为 `InvalidHeader` 错误 |
| L3 Workbook 原生 | `NpoiWorkbookValidationPipeline.Validate` | Sheet 内 DataValidation：LIST / TEXT_LENGTH / INTEGER / DECIMAL / DATE / TIME | `WorkbookValidation` 错误码；不支持的规则按 `UnsupportedFeaturePolicy` 处理 |
| L4 原始值规则 | `NpoiImportRowMaterializer.ValidateRawValues` | `ExcelValidationBinding.IsRaw` 的绑定（Required / Regex / MaxLength / MaxValue），**在类型转换之前**用原始文本校验 | `Validation` / `MaxLength` / `MaxValue` |
| L5 转换 + 规则 + 唯一 | `TryCreateItem` / `ValidateColumnValue` | 类型转换、其余规则、`UniqueTracker.TryReserve` | `ValueConversion` / `Validation`（唯一冲突带 `FirstRowNumber`） |
| L6 关系 | `NpoiRelationBinder.BindCore` | 父子导航绑定 | `Relationship` 错误码 |

**内置规则：** `ExcelValidationRules.CreateDefault()` 返回 7 条 —— Required、Regex、Range、MaxValue、MaxLength、DateTime、Duplication。契约是 `IExcelValidationRule.CanValidate(attr)` + `Validate(attr, ExcelValidationContext)`；命名规则走 `INamedExcelValidationRule`。

**绑定：** `ExcelMappingPlanFactory.BindValidationRules` 遍历属性上的 `FilterAttributeBase`，`ResolveValidationRule`（优先 `BindFilterAttribute.RuleType`，否则 `CanValidate` 匹配）生成 `ExcelValidationBinding.Attribute(...)` 或 `.Named(...)`，落入 `ExcelMappingColumn.ValidationBindings`。

**错误收集：** `ExcelImportErrorCollector` 持有根错误列表并受 `MaxErrors` 限流（达到上限调用 `MarkTruncated`）；每个 Sheet 通过 `CreateChild()` 拿到 Sheet 视图。

---

### 3.5 M4 转换模块

- `IExcelValueConverter`：`CanConvert` / `TryConvertFrom` / `TryConvertTo`
- `ExcelConversionContext`：`Value` / `PropertyName` / `PropertyType` / `SheetName` / `Row` / `Column` / `Culture` / `Cell`
- `ExcelCellValue`：provider-neutral 单元格快照 —— `Value` / `Text` / `ExcelCellKind` / `CachedKind` / `Formula` / `ErrorCode` / `FormatIndex` / `IsDate1904`
- `ExcelCellKind`：`Empty | Text | Number | Boolean | DateTime | Formula | Error`
- 绑定解析：`ExcelValueConverterBindingResolver.Resolve(converters, converterName, propertyType)` —— 支持按名称精确匹配与按能力匹配（能力判定结果用 `ConditionalWeakTable` 缓存），结果写入 `ExcelMappingColumn.ValueConverters`

**设计意图：** Provider 把「NPOI 单元格」翻译成 `ExcelCellValue` 这一中立结构，转换器因此完全不需要知道 NPOI 的存在。

---

### 3.6 M5 导入执行模块

**入口契约：**

```csharp
ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
    ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken ct = default);
Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
    ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken ct = default);
```

返回 `ExcelWorkbookImportResult<TWorkbook>`：`Workbook`、`Sheets`(`IReadOnlyList<ExcelSheetImportResult>`)、`Errors`、`ErrorsTruncated`、`MaxErrors`、`IsSuccess`。

**严格时序（`NpoiExcelImporter`）：**

| Step | 动作 | 位置 |
|---|---|---|
| 1 | 参数守卫：source / request 非空、`CanRead`、`ThrowIfCancellationRequested` | `Import` / `ImportAsync` |
| 2 | **输入缓冲**：`NpoiStreamCopier.Copy` / `CopyAsync` → `MemoryStream`，逐块累加并施加上限 | `ImportCore` / `ImportAsync` |
| 3 | **ZIP 预检**：`NpoiXlsxZipPreflight.Validate`（仅可 seek 且命中 ZIP 签名时） | `ImportBufferedCore` |
| 4 | 建立 DOM：`WorkbookFactory.Create(bufferedSource)`、`IsDate1904()` | 同上 |
| 5 | 建容器：`new TWorkbook()`、Sheet 结果列表、`ExcelImportErrorCollector(MaxErrors)`、`Dictionary<object, SourceLocation>`、`ExcelImportRuntime(limits)` | 同上 |
| 6 | 解析 Sheet：`ResolveSheet` / `ResolveSheetIndex`（重复物理 Sheet 抛 `ArgumentException`） | 同上 |
| 7 | **建计划**：`NpoiImportPlanBuilder.Create` → `CreateWorkbook<T>` | 同上 |
| 8 | 逐 Sheet 循环：跳过缺失/隐藏（记 `InvalidHeader`）→ `ImportTypedSheet` → `ImportTypedSheetCore` → **`NpoiImportSheetExecutor.Execute`** | 同上 |
| 9 | 关系绑定：`NpoiRelationBinder.Bind` | 同上 |
| 10 | 失败工作簿回写：`NpoiFailureWorkbookWriter.Write` | 同上 |
| 11 | 组装 `ExcelWorkbookImportResult<TWorkbook>` | 同上 |

**`NpoiImportSheetExecutor.Execute` 内部：**

```
取表头行 → CreateColumns<T>()（结构校验）
        → 图片索引 → sheet.GetDataValidations() → ValidationRangeIndex.Create
        → UniqueTracker
        → 逐行循环：
            空行判定
            → TryConsumeRow（超 MaxRows 处理）
            → L3 Workbook 原生校验
            → L4 原始值校验
            → TryCreateItem（转换 → L5 规则 → 唯一性 → 落值）
```

**输入防护（双保险）：**

1. **`MaxInputBytes`（默认 128 MiB）** —— 在 `NpoiStreamCopier.Copy/CopyAsync` 的每块累加处生效，超限抛 `BingOfficesResourceLimitException`(Stage=Open)。同步与异步路径均覆盖。
2. **`NpoiXlsxZipPreflight`** —— 针对 XLSX 的 ZIP 炸弹与 XXE 防护：条目数、单条/总解压字节、压缩比、`sharedStrings`/`styles`/`worksheet` 字节、路径穿越（`/`、`\`、`..`）、重复 entry、缺失 `xl/workbook.xml`，并对每个 XML 条目用 `XmlReader`（禁 DTD）检查字符数与嵌套深度。

> **重要边界：** `MaxInputBytes` 保护的是**原始输入字节**，**不等于**解压后的 DOM 峰值内存保护。

**失败工作簿回写链：**

触发条件：`ExcelImportFailureOptions.Mode != None` 且 `errors.Count > 0`。

```
NpoiFailureWorkbookPreflight.Validate      预算校验（候选错误行/图片/单元格/目标对象）
        ▼  按模式分派
NpoiFailureWorkbookCopier.CreateErrorRowsWorkbook     (ErrorRowsOnly)
NpoiFailureWorkbookAnnotationWriter.AnnotateErrors    (AnnotatedOriginal，按 CommentConflictPolicy)
        ▼
WriteSummary → 生成 _ImportErrors 表
        ▼
NpoiFailureWorkbookSerialization.LimitedWriteStream   受 MaxSerializedBytes 约束
        ▼
IFailureWorkbookFileSystem 清理临时文件（清理失败走 DiagnosticSink 或包成 Cleanup 异常）
```

异步路径先写 `INpoiAsyncStaging.WriteStream`（前缀 `bing-offices-failure-async-`），结束后 `FlushAsync` + `CopyToAsync` 到 `Destination`。

---

### 3.7 M6 导出执行模块

**入口契约：**

```csharp
void Export(ExcelWorkbookExportRequest request, Stream destination, CancellationToken ct = default);
Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination, CancellationToken ct = default);
void ExportToFile(ExcelWorkbookExportRequest request, string path, CancellationToken ct = default);
Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path, CancellationToken ct = default);
```

**Fluent 构建入口：**

```csharp
var request = ExcelExport.Workbook(workbook => workbook
    .Format(ExcelFormat.Xlsx)
    .Metadata(new ExcelWorkbookMetadataOptions { Author = "业务系统" })
    .UseTemplate(templateStream, leaveOpen: false)
    .AddSheet("订单", orders, sheet => sheet
        .Mapping(profileOrDocument)
        .DynamicColumns(columns)
        .Chart(chart => chart.Column(...))));
```

`ExcelSheetExportBuilder<T>` 暴露：`HeaderRowIndex` / `DataRowStartIndex`、`DynamicColumns`、`UnknownDynamicValues`、`Sheet`/`Header`/`BodyStyle`、`HeaderRows`、`Mapping(config | document)`、`Culture`、`ColumnWidth`、`CommentConflicts`、`TemplateCellOverwrite`、`UseTemplateRegion`、`Hidden`、`Chart`。

**严格时序（`NpoiExcelExporter`）：**

| Step | 动作 |
|---|---|
| S1 | 校验目标可写；逐 Sheet `ValidateSheetName`（≤31 字符、非法字符、重名） |
| S2 | `CreateWorkbook`：无模板 → `ExcelHelper.PrepareWorkbook(format, metadata)`；有模板 → `WorkbookFactory.Create(NpoiNonDisposingStream)`，仅 `MetadataSpecified` 时 `ApplyWorkbookMetadata` |
| S3 | `NpoiExportPlanBuilder.Create`：按 `NpoiWorkbookPlanKeyBuilder` 键分组 → 反射分派 → `CreateWorkbook<T>`，回填 `Dictionary<Sheet, IExcelMappingPlan>`；异常转 `BingOfficesConfigurationException`(Plan) |
| S4 | 每 Sheet `WriteSheet`（委托经 `ConcurrentDictionary` 缓存）→ `WriteTypedSheet<T>`：取/建 Sheet、`Hidden`、索引校验、`ResolveTemplateOrigin`（命名区域 → row/col）、`CreateDynamicDefinition`、`CreateColumns<T>`、`ValidateDynamicColumns` |
| S5 | `NpoiExportSheetWriter.Write<T>`：表头 + `WriteCustomHeaders` → 逐行逐列 `WriteCell`（`DynamicGetter` → `ConvertTo` → `WriteValue`）→ `ApplyRequestStyles` → `ApplyHeaderStyle<T>` → `ApplyWrapText<T>` → `ApplyColumnWidths` → `MergeColumns<T>` → `CreateCharts` |
| S6 | `workbook.Write(new NpoiNonDisposingStream(dest, ct), false)` |
| S7 | `finally` 中按 `LeaveTemplateOpen` 释放模板流 |

`ExportToFile` / `ExportToFileAsync` 在 S1–S7 之外包裹 `_fileExportCommitter.Commit(..., "Excel")`。

**列规划：** `NpoiExportColumnPlanner.CreateColumns<T>` 先收集非 `Ignored` 固定列（跳过动态容器属性），再取唯一的 `IsDynamicColumn` 属性，动态列按 `Order` → `Key` 排序落位：

1. `PhysicalColumnIndex`（越界抛异常）
2. `Placement.Before` / `After`（按 Key `FindIndex` 插入）
3. 否则尾部 `Add`

`ExcelColumnPlan` 是不可变列计划：持 Getter/Setter、`IsDynamic`/`IsMerged`/`IsUnique`、`Formatter`/`DecimalScale`/`ValueMap`、转换器与 Header/BodyStyle，提供 `ConvertTo` / `ConvertFrom` / `WriteValue`。物理定位：**列 = `originColumn + columnIndex`，行 = `originRow + HeaderRowIndex / DataRowStartIndex`**。

**样式缓存 `NpoiStyleCache`：**

- 存储：`ConditionalWeakTable<IWorkbook, Cache>` —— **随工作簿 GC 回收**，按工作簿隔离
- 键：`CreateKey(style)`（40+ 属性 + Reset 的确定性串）；`Compose` 用 `baseStyle.Index + ":" + CreateKey`；`ApplyHeaderAttribute` 用 `FontIndex|Name|Size|Bold|Color` 与 `baseStyle.Index|fontKey`
- 线程安全：4 个字典各自 `lock`（`_styles` / `_composedStyles` / `_fonts` / `_headerStyles`）

**模板语义：**

| 策略 | 行为 |
|---|---|
| `PreserveTemplate`（默认） | 保留目标单元格的模板样式与批注，只写值（公式会被值覆盖） |
| `ReplaceTemplate` | 写前重置为 `GetCellStyleAt(0)` 并 `RemoveCellComment` |

未显式调用 `Metadata(...)` 时模板 metadata 保留；显式调用时 6 个字段覆盖模板值，XLS 与 XLSX 策略一致。

**图表：** 仅 `NpoiExportSheetWriter.CreateCharts` 生成，要求 `workbook is XSSFWorkbook` 且能建 `XSSFDrawing`，否则抛 `NotSupportedException` —— **只支持 XLSX，HSSF/.xls 不支持**。经 `ChartDataFactory` 支持 Column / Line / Pie（Pie 强制单 Series）。

---

### 3.8 M7 CSV 模块

CSV 是**独立的文本管线**，但完整复用 M1–M4 的映射、校验、转换与错误模型。

| 维度 | CSV | Excel |
|---|---|---|
| 底层 | CsvHelper 33.1.0 | NPOI |
| 同步 IO | `StreamWriter` + `CsvRecordWriter` | NPOI DOM 序列化 |
| 异步 IO | `CsvAsyncRecordWriter`（StringBuilder + `CsvWriter`，`Encoder.Convert` 分块 `WriteAsync`，**真异步**） | `NpoiAsyncStaging`（Memory / TempFile / Hybrid，默认 TempFile，8 MB 迁移阈值） |
| 公式注入防护 | `CsvFormulaInjectionPolicy.Escape`（默认）：给以 `= @ + -` 开头的字段加 `'`，但 `+N` / `-N` 带符号数豁免；`None` 则完全不处理 | 不适用 |
| 默认编码 / 分隔符 | UTF-8（无 BOM）、`,`、`"`、`\r\n`（均带校验） | — |
| 样式 / 图表 / 合并 / 列宽 | 不支持 | 支持 |
| 导入读取 | `CsvLimitedReadStream` / `CsvCancellationStream` / `CsvRecordReader`（RFC 4180） | `NpoiStreamCopier`（同步 / 异步 + maxBytes） |

CSV 列来源：`CsvPropertyBinding` + `CsvExportColumn`；动态列取字典值（`CsvDynamicTypeResolver`）。两者共用 `IFileExportCommitter`。

---

### 3.9 M8 IO 与原子提交

**`AtomicFileCommitter` 流程：**

```
temp = path.Guid.tmp
  → 检查取消
  → FileStream(CreateNew, None, Async)
  → write(...)
  → Flush(FileStream.Flush(flushToDisk: true))
  → 再次检查取消        ← 最终提交点的取消复查
  → Exists ? File.Replace : File.Move
  → temp = null（标记成功）
失败/取消 → Cleanup 删除 temp
```

- 内容期异常**原样抛出**；其余包成 `BingOfficesFileCommitException`
- 清理失败只写 `exception.Data`，**不覆盖主异常**
- **保证：** 导出失败或取消不会截断已有目标文件
- **已知缺口：**
  - `Exists → Replace/Move` 之间存在 TOCTOU 竞态
  - `File.Replace` 对跨卷 / 锁定 / 不支持的 FS 会失败
  - 无目录 fsync，进程崩溃可能残留 `.tmp`
  - 取消只在边界检查，写入中途不可中断
- **直接写入调用方 Stream 时无回滚保证**，失败后可能已产生部分写入

---

### 3.10 M9 异常与诊断模型

**三维分类法：**

```csharp
enum BingOfficesErrorCode  { ConfigurationInvalid, ImportFailed, ExportFailed,
                             ResourceLimitExceeded, FileCommitFailed,
                             UnsupportedFeature, UserExtensionFailed }

enum BingOfficesOperation  { Configuration, Import, Export, FileCommit }

enum BingOfficesStage      { Open, Preflight, Plan, Read, Validate,
                             Write, Serialize, Commit, Cleanup }
```

`BingOfficesException : InvalidOperationException` 携带 `Code` / `Operation` / `Provider` / `Stage` / `SheetName` / `RowIndex` / `ColumnIndex` / `PropertyName`（行列均为**一基**）。

**分发：** `BingOfficesExceptionDispatcher.Observe(exception)` 把已分类异常送交所有 `IBingOfficesExceptionObserver`，并用 `ObservedKey` 去重后**重抛**。这是唯一的异常观测入口 —— 应用层接一个 Observer 即可集中上报。

**导入错误（数据级，非异常）：** `ExcelImportError(Code, Message, SheetName, RowIndex, ColumnIndex, PropertyName, ColumnKey, Header, RawValue, FirstRowNumber)`，`Code` 取值 `InvalidInput` / `InvalidHeader` / `ValueConversion` / `Validation` / `MaxLength` / `MaxValue` / `Relationship` / `WorkbookValidation` / `ResourceLimit`。

`ImageResourceLimitException`（`ExcelImportRuntime` 内部类）在出口被译为 `BingOfficesResourceLimitException`；非 OOM / StackOverflow 的未知异常被译为 `BingOfficesImportException`(NPOI, Read)。

---

### 3.11 M10 NPOI 用户扩展容器

`Bing.Offices.Npoi.Extensions` 提供 7 个容器（同一容器内区分 Try/Throw 两种行为）：

| 容器 | 覆盖内容 |
|---|---|
| `CellExtensions` | 单元格读写、`CellExtensions.ConditionalFormatting`、`CellExtensions.MergeRegion` |
| `SheetExtensions` | `SheetExtensions.MergedRegion`、`SheetExtensions.Picture` |
| `WorkbookExtensions` | Workbook 级操作 |
| `RowExtensions` | 行级操作 |
| `CellStyleExtensions` | 样式构造与复制 |
| `FontExtensions` | 字体构造 |
| `InternalExtensions` | 内部辅助 |

`Extensions.Service.cs` 中的 `Extensions.Service.cs` 同类文件亦提供 provider-neutral 文件/字节扩展（`Bing.Offices.Extensions`）。

**组合根：**

```csharp
services.AddBingOfficesNpoi();
```

注册内容：

| 注册项 | 生命周期 |
|---|---|
| 7 条默认 `IExcelValidationRule` | Singleton（`TryAddEnumerable`） |
| `IExcelMappingPlanFactory`（缓存容量 256） | Singleton |
| `IExcelMappingConfigurationLoader` | Singleton |
| `IFileExportCommitter` | Singleton |
| `ICsvImporter` / `ICsvExporter` | Transient |
| `IExcelImporter` / `IExcelExporter` | Transient |

全部使用 `TryAdd*`，**已预注册的实现保持优先** —— 这是替换默认实现的官方途径。

---

## 4. 数据流

### 4.1 配置解析与合并流

见 [3.2](#32-m1-配置模块) 与 [3.3](#33-m2-计划编译模块)。核心不变式：

> **只有 Plan 编译器执行最终合并。** 任何上游模块都不得自行合并配置。

### 4.2 Excel 导入数据流

```
Stream ──► NpoiStreamCopier ──► MemoryStream ──► NpoiXlsxZipPreflight
                                                      │
                                    WorkbookFactory.Create ▼
                                                      │
                          ┌───────────────────────────┘
                          ▼
              NpoiImportPlanBuilder ──► IExcelMappingPlanFactory.CreateWorkbook<T>
                          │                        （Plan 缓存命中则直接复用）
                          ▼
        ┌─────── 每个 Sheet ───────┐
        │ NpoiImportSheetExecutor  │
        │   CreateColumns          │
        │   → 逐行：原生校验 → 原始值校验 → 转换 → 规则 → 唯一 → 落值
        │   → ExcelImportErrorCollector
        └──────────┬───────────────┘
                   ▼
        NpoiRelationBinder.Bind（父子关系）
                   ▼
        NpoiFailureWorkbookWriter.Write（可选）
                   ▼
        ExcelWorkbookImportResult<TWorkbook>
        { Workbook, Sheets, Errors, ErrorsTruncated, IsSuccess }
```

### 4.3 Excel 导出数据流

```
ExcelWorkbookExportRequest
   │
   ├─ 无模板 → ExcelHelper.PrepareWorkbook(format, metadata)
   └─ 有模板 → WorkbookFactory.Create(NpoiNonDisposingStream)
   │
   ▼
NpoiExportPlanBuilder ──► IExcelMappingPlanFactory.CreateWorkbook<T>
   │
   ▼
NpoiExportColumnPlanner.CreateColumns<T>  （固定列 + 动态列物理落位）
   │
   ▼
NpoiExportSheetWriter.Write<T>
   表头 → 逐行逐列(Getter→Converter→WriteValue) → 样式 → 列宽 → 合并 → 图表
   │
   ▼
workbook.Write(NpoiNonDisposingStream(dest, ct))
   │
   ▼
[ExportToFile 路径] IFileExportCommitter.Commit → AtomicFileCommitter
```

### 4.4 CSV 数据流

```
ICsvImporter：Stream → CsvCancellationStream → CsvLimitedReadStream → CsvRecordReader
             → CsvHeaderBinder → CsvPropertyBinding → 转换/校验 → 实体
             → CsvImportResult { Items, Errors }

ICsvExporter：IEnumerable<T> → CsvPropertyBinding（含动态列）
             → CsvFormulaInjectionPolicy.Escape
             → CsvAsyncRecordWriter / CsvRecordWriter（UTF-8 无 BOM）
             → Stream 或 IFileExportCommitter
```

---

## 5. 并发、缓存与线程安全

| 缓存 | 位置 | 键 | 并发机制 | 淘汰 |
|---|---|---|---|---|
| Mapping Plan 缓存 | `ExcelMappingPlanFactory` 实例字段 | SHA-256(Base64) 复合键 | `ConcurrentDictionary` + `Lazy(ExecutionAndPublication)` | FIFO，容量 256 |
| TypeMap 反射缓存 | `ExcelTypeMapFactory` 静态 | `Type` | `ConcurrentDictionary` | **无淘汰** |
| 转换器能力缓存 | `ExcelValueConverterBindingResolver` | — | `ConditionalWeakTable` | GC 驱动 |
| NPOI 样式缓存 | `NpoiStyleCache` | 属性串 / `baseStyle.Index` 组合 | `ConditionalWeakTable` + 4 个 `lock` | 随 Workbook GC |
| 导出委托缓存 | `NpoiExcelExporter` | Sheet 类型 | `ConcurrentDictionary` | 无 |
| Profile Registry | `MappingProfileRegistry` | `(Name, Direction, ModelType)` | `Dictionary` + `lock` | 无（重复注册抛异常） |

**并发结论：** 引擎自身的共享状态基本都是并发安全的（`ConcurrentDictionary` / `Lazy` / `lock`），但 **`ExcelMappingPlanFactory` 注册为 Singleton**，其缓存是进程级共享的。

---

## 6. 资源限制与安全防护矩阵

| 防护项 | 默认值 | 生效位置 |
|---|---|---|
| `MaxInputBytes` | 128 MiB | `NpoiStreamCopier` |
| `MaxRows` | 无限 | `NpoiImportSheetExecutor.TryConsumeRow` |
| `MaxErrors` | 无限 | `ExcelImportErrorCollector` |
| `MaxPictures` / `MaxPictureBytes` / `MaxTotalPictureBytes` | 无限 | `BuildImageIndex` / `ExcelImportRuntime` |
| `MaxTrackedUniqueValues` | 无限 | `UniqueTracker` |
| `MaxZipEntries` | 10 000 | `NpoiXlsxZipPreflight` |
| `MaxZipEntryUncompressedBytes` | 64 MiB | 同上 |
| `MaxZipTotalUncompressedBytes` | 512 MiB | 同上 |
| `MaxZipCompressionRatio` | 1000 | 同上 |
| `MaxSharedStringsBytes` | 64 MiB | 同上 |
| `MaxStylesBytes` | 16 MiB | 同上 |
| `MaxWorksheetBytes` / `MaxTotalWorksheetBytes` | 64 MiB / 256 MiB | 同上 |
| `MaxXmlCharacters` / `MaxXmlDepth` | 64 MiB / 256 | 同上 |
| `MaxSerializedBytes`（失败工作簿） | — | `NpoiFailureWorkbookSerialization` |
| CSV 公式注入 | `Escape` | `CsvFormulaInjectionPolicy` |
| XML 映射文档外部实体 | 禁止 | `DtdProcessing.Prohibit` + `XmlResolver = null` |
| 映射文档文本长度 | 1 MB | `ExcelMappingTextReader.ReadLimitedText` |

**部署建议（官方文档口径）：** 处理不受信任文件时，除上述限额，还应设置**进程内存 / CPU 配额**。已知缺口：`MaxInputBytes` 与 `MaxSerializedBytes` 都不覆盖解压后 DOM 的峰值；**未被映射的图片不会进入图片限制器的扫描范围**。

---

## 7. 取消合同

`CancellationToken` 检查点（导入侧）：

- 入口守卫
- `NpoiStreamCopier` 每块复制前后
- ZIP 预检：开头、每条目、每次 `XmlReader.Read`
- 每 Sheet、每关系
- `Execute` 每行
- `BuildImageIndex` 每张图
- `NpoiRelationBinder.BindCore` 三个阶段每次迭代
- 失败工作簿的复制 / 序列化 / 清理

取消统一抛 `OperationCanceledException`；`NpoiExcelImporter` 使用 `when (ct.IsCancellationRequested && 非基类)` 把它归一化为 `new OperationCanceledException(ct)`。

导出侧：`ExportToFile` 在 `AtomicFileCommitter` 的**写前与替换前**各检查一次；写入中途不可中断。

**语义承诺：** 取消不会损坏已有目标文件（`ExportToFile` 路径）。

---

## 8. 扩展点

| 想做什么 | 扩展方式 |
|---|---|
| 自定义值转换 | 实现 `IExcelValueConverter`（或 `INamedExcelValueConverter`）并注册到 DI；在映射上写 `ConverterName` 或隐式按能力匹配 |
| 自定义校验规则 | 实现 `IExcelValidationRule` / `INamedExcelValidationRule`，注册到 DI；用 `BindFilterAttribute` 显式绑定规则类型 |
| 复用映射配置 | 实现 `IImportMappingProfile<T>` / `IExportMappingProfile<T>` / `IMappingProfile<TImport, TExport>`，用 `AddMappingProfile<T>("stable-name")` 注册 |
| 外部化映射 | 提供 JSON / XML v2 文档，经 `ExcelMappingDocument` 注入 |
| 替换默认实现 | 在 `AddBingOfficesNpoi()` **之前**注册自己的实现（全部 `TryAdd*`，预注册者优先） |
| 集中异常上报 | 实现 `IBingOfficesExceptionObserver` 并注册 |
| 自定义文件提交 | 实现 `IFileExportCommitter` |
| 扩展 NPOI 便捷方法 | 使用 `Bing.Offices.Npoi.Extensions` 的 7 个容器 |
| 接入新 Provider | 仅依赖 `IExcelMappingPlanFactory` + `IExcelMappingPlan`，不触碰 NPOI |

---

## 9. 测试架构

| 测试项目 | 定位 |
|---|---|
| `Bing.Offices.Tests` | 单元测试主体（37 个测试文件）。含公共 API 契约快照（`PublicApiContractTest`）、扩展覆盖率（`PublicExtensionCoverageTest`）、映射/校验/CSV/IO/预检等职责级测试；`Models/Bugs/Issue1..9` 固化历史缺陷回归 |
| `Bing.Offices.Tests.Integration` | 真实文件 IO 集成测试：`ExcelAsyncFileIntegrationTest`、`CsvAsyncFileIntegrationTest`、`ExcelImporterIntegrationTest` |
| `Bing.Offices.Docs.Tests` | 把 `docs/excel/*.md` 中的示例代码**编译并执行**，防止文档腐化 |
| `Bing.Offices.ProfileFixtures` | 供程序集扫描注册测试使用的外部 Profile |
| `Bing.Offices.ResourceProbe` | 暂存/资源矩阵探针：`StagingEntrypointMatrix`、`StagingResourceMatrix` |
| `Bing.Offices.Consumer.Net6` / `Consumer.Net8` | 双 TFM 消费方验证，用于 nupkg 可消费性检查（由发布任务中的独立 PackageConsumer 执行） |

**测试门槛（摘自 `AGENTS.md`）：**

- 修改 Excel/CSV 公共接口、默认实现、Provider 分支、映射、缓存键、文件提交或异步 IO 时，**必须同步增加职责级直接测试**
- 每个受影响的默认实现必须有独立测试；**禁止只依赖综合测试间接覆盖**
- CSV/Excel 输出测试必须断言**完整结果、文件内容或结构化错误**，不得只断言片段
- 缓存相关的修改必须覆盖**隔离、命中、未命中、重复渲染、失败后的对象状态**
- NPOI Provider 的 XLS/HSSF 与 XLSX/XSSF 分支都必须有成功与失败场景
- 异步 API 必须使用**真实异步 IO**；生产代码禁止 `Task.Run` / `.Result` / `.Wait`
- API 变化必须更新双 TFM snapshot、成员级追溯与迁移记录；未经成员级审批不得删除公共成员

---

## 10. 已知边界与风险

| # | 边界 / 风险 | 影响 | 缓解 |
|---|---|---|---|
| R1 | 既有 NPOI DOM 全量内存 | 大文件 OOM | 普通 XLSX 可显式选择低内存服务；DOM 继续使用 `ExcelResourceLimits` + ZIP 预检 + 进程级配额 |
| R2 | `MaxInputBytes` ≠ 解压后 DOM 峰值保护 | 高压缩比文件仍可打爆内存 | ZIP 压缩比与解压字节限制 |
| R3 | 未被映射的图片不进入图片限制器扫描 | 图片型攻击面残留 | 依赖进程级内存配额 |
| R4 | 计划缓存无 TTL / 主动失效，容量 256 | 动态生成大量映射时缓存抖动 | 稳定 `TenantId` + `ConfigurationVersion`；必要时显式指定缓存容量 |
| R5 | `ExcelTypeMapFactory` 静态缓存无淘汰 | 长期运行进程持有 Type 引用 | 类型数量通常有限，风险可控 |
| R6 | `AtomicFileCommitter` 的 `Exists → Replace/Move` TOCTOU | 极端并发下的替换竞态 | 同一路径避免并发写 |
| R7 | 无目录 fsync；取消不可中断写入中途 | 崩溃残留 `.tmp`；大文件取消延迟 | 定期清理 `.tmp`；任务级超时 |
| R8 | 直接写调用方 Stream 无回滚 | 失败后产生部分写入 | 需要强一致时使用 `ExportToFile*` |
| R9 | 图表仅 XLSX | `.xls` 请求图表直接 `NotSupportedException` | 请求侧提前判定 |
| R10 | Async 不覆盖 DOM 阶段 | 高并发下线程池仍被 DOM 构建占用 | 按容量规划部署；避免在线程池饥饿敏感路径上同步导出大文件 |
| R11 | 模板流必须同步可读 | async-only 流不可作为模板 | 用 `MemoryStream` 或可 seek 文件流 |
| R12 | 库不承诺 API 向后兼容（见 README 免责声明） | 升级需回归 | 依赖 `nuget-migration.md` 与成员级追溯 |

---

## 11. 附录：关键类型索引

### 11.1 契约（Abstractions）

| 命名空间 | 关键类型 |
|---|---|
| `Bing.Offices.Imports` | `IExcelImporter`、`ExcelWorkbookImportRequest<T>`、`ExcelWorkbookImportResult<T>`、`ExcelSheetImportResult`、`ExcelImportError`、`ExcelImportErrorCode`、`ExcelImportPolicies`、`ExcelSheetPolicies`、`ExcelFailurePolicies`、`ExcelResourceLimits`、`ValidateMode` |
| `Bing.Offices.Exports` | `IExcelExporter`、`ExcelWorkbookExportRequest`、`ExcelExport`、`ExcelSheetExportBuilder<T>`、`ExcelExportPolicies`、`ExcelDynamicColumnDefinition`、`ExcelChartDefinition`、`ExcelWorkbookMetadataOptions`、`ExcelComment` |
| `Bing.Offices.Csv` | `ICsvImporter`、`ICsvExporter`、`CsvImportOptions`、`CsvImportResult`、`CsvImportError`、`CsvExportOptions<T>`、`CsvFormulaInjectionPolicy` |
| `Bing.Offices.Configurations` | `ExcelMappingConfiguration`、`ExcelMappingDocument`、`ExcelMappingDocumentFactory`、`MappingConfigurationMerger`、`MappingProfileRegistry`、`IMappingProfileResolver`、`FluentSetting<TImport, TExport>`、`ImportMappingBuilder`、`ExportMappingBuilder`、`ProfileDescriptor`、`ExcelModelAliasRegistry` |
| `Bing.Offices.Providers` | `IExcelMappingPlanFactory`、`IExcelMappingPlan`、`IExcelMappingColumn`、`IExcelDynamicMappingColumn`、`IExcelMappingStyle`、`IExcelMappingLayout`、`IExcelMappingWorkbookPlan`、`UniqueTracker` |
| `Bing.Offices.Conversions` | `IExcelValueConverter`、`INamedExcelValueConverter`、`ExcelConversionContext`、`ExcelCellValue`、`ExcelCellKind` |
| `Bing.Offices.Validations` | `IExcelValidationRule`、`INamedExcelValidationRule`、`IExcelValidationBinding`、`ExcelValidationContext` |
| `Bing.Offices.Exceptions` | `BingOfficesException`、`BingOfficesErrorCode`、`BingOfficesOperation`、`BingOfficesStage`、`IBingOfficesExceptionObserver` |
| `Bing.Offices.IO` | `IFileExportCommitter` |
| `Bing.Offices.Styles` | `ExcelCellStyle`、`ExcelColor`、`ExcelBorderStyle`、`ExcelCellStyleReset` |

### 11.2 引擎（Core）

| 命名空间 | 关键类型 |
|---|---|
| `Bing.Offices.Mappings` | `ExcelMappingPlanFactory`、`ExcelMappingPlanFactoryProvider`、`ExcelMappingPlanCacheKey`、`ExcelTypeMapFactory`、`ExcelTypeMap<T>`、`ExcelPropertyMap`、`ExcelMappingColumn`、`ExcelDynamicMappingColumn`、`ExcelMappingWorkbookPlan`、`ExcelValueConverterBindingResolver`、`IExcelCompiledMappingColumn` |
| `Bing.Offices.Configurations` | `ExcelMappingConfigurationLoader`、`DefaultExcelMappingConfigurationLoader`、`ExcelMappingTextReader`、`ExcelMappingDocumentValidator` |
| `Bing.Offices.Csv` | `CsvEntityImporter`、`CsvEntityExporter`、`CsvHeaderBinder`、`CsvPropertyBinding`、`CsvPipelineSupport`、`CsvDynamicTypeResolver` |
| `Bing.Offices.IO` | `AtomicFileCommitter`、`DefaultFileExportCommitter` |
| `Bing.Offices.Validations` | `ExcelValidationRules` |
| `Bing.Offices.Exceptions` | `BingOfficesExceptionDispatcher` |
| `Bing.Offices.Dates` | `ExcelDateParser` |
| `Bing.Offices.Extensions` | `MappingProfileServiceCollectionExtensions`、`ProfileDescriptorFactory` |

### 11.3 Provider（Npoi）

| 目录 | 关键类型 |
|---|---|
| `Imports/` | `NpoiExcelImporter`、`NpoiImportPlanBuilder`、`NpoiImportSheetExecutor`、`NpoiImportRowMaterializer`、`NpoiRelationBinder`、`ValidationRangeIndex`、`ExcelImportErrorCollector`、`ExcelImportRuntime`、`NpoiWorkbookValidationPipeline`、`NpoiXlsxZipPreflight`、`NpoiFailureWorkbookWriter`、`NpoiFailureWorkbookPreflight`、`NpoiFailureWorkbookCopier`、`NpoiFailureWorkbookAnnotationWriter`、`NpoiFailureWorkbookSerialization`、`NpoiFailureWorkbookDiagnostics`、`NpoiFailureWorkbookFileSystem` |
| `Exports/` | `NpoiExcelExporter`、`NpoiExportPlanBuilder`、`NpoiExportColumnPlanner`、`NpoiExportSheetWriter`、`NpoiStyleCache`、`NpoiNonDisposingStream` |
| `IO/` | `NpoiAsyncStaging`、`NpoiStreamCopier` |
| `Extensions/` | `ExcelNpoiServiceCollectionExtensions`、7 个扩展容器 |
| 根 | `ExcelColumnPlan`、`ExcelHelper`、`NpoiMappingPlanFactoryResolver`、`Resolvers/ColorResolver` |

---

## 12. 相关文档

- [Excel 高级导入导出](../excel/README.md)
- [Mapping Profile](../excel/mapping-profile.md)
- [JSON/XML 映射文档](../excel/mapping-json-xml.md)
- [导入校验](../excel/import-validation.md)
- [动态列](../excel/dynamic-columns.md)
- [异常与 Observer](../excel/exceptions-and-observers.md)
- [Sync/Async IO](../excel/async-io.md)
- [NPOI 扩展](../excel/npoi-extensions.md)
- [NuGet 迁移](../excel/nuget-migration.md)

## 13. 显式低内存服务

`AddBingOfficesNpoiStreaming` 注册 `IExcelStreamingExporter` 和 `IExcelStreamingImporter`，保持原服务默认行为。Abstractions 仍为 netstandard2.0，NPOI 仍为 2.7.4；没有 Open XML SDK 依赖。

导出链路：Workbook Request → 全配置预检与共享映射计划 → SXSSF（inline strings、有限行/文本窗口、创建单元格时应用最终样式）→ 压缩临时 XML → 文件打包 → 同步/真实异步复制或原子提交 → Dispose SXSSF → Close 基础 XSSF。NPOI 的同步打包阶段不能承诺立即响应取消，写出边界会检查令牌。

导入链路：当前输入位置 → 有额度文件暂存 → ZIP 元数据与真实关系解析 → 有限样式表/共享字符串内存预算及磁盘定长索引 → Sheet XML 当前行 → 共享列绑定/标量转换/校验 → 等待消费者 → 有界摘要。只为表头绑定创建一行有限 XSSF，正文不进入 DOM。未消费部件仅验证 ZIP 元数据，不声称扫描其全部 XML 内容。

每次调用独立拥有临时文件和工作簿。同一个服务实例可以接收独立调用，不对同一 Workbook 并发操作。逐行消费者负责业务事务，已完成的外部写入不回滚。共享字符串查询缓存与内存表共用 8 MiB 计量预算，落盘索引不保留全量偏移数组。

库自有文件每调用默认 512 MiB；NPOI 的系统 poifiles 目录单独监测，不属于该硬额度。异常清理尝试全部资源，不覆盖主异常；清理失败通过异常或异常 Data 可见。

详见 [低内存 API、默认额度与迁移](../excel/streaming.md)。容量验收为 1 万／10 万行 × 10 列、重复／高唯一字符串，共 8 场景；不设固定 MB 门槛，不承诺首版不支持的模板、图片、图表、合并或公式计算。
