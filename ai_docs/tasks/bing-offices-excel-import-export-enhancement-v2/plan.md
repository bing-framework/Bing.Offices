# Bing.Offices Excel 导入导出增强 V2 实施计划

> Task ID：`bing-offices-excel-import-export-enhancement-v2`  
> Workflow：Universal Agent Workflow V4  
> 计划日期：2026-08-20  
> 状态：READY_FOR_EXECUTION  
> 唯一生产入口：Workbook Request  
> 计划性质：未发布 API 的破坏性收敛，不保留无价值兼容层

## 1. 任务摘要

本计划在现有 Workbook Request、多 Sheet、typed 动态列、模板、样式、关系和基础图表能力之上，完成 12 项 Excel 导入导出需求，并治理当前“公开模型已出现但生产接线不完整”的问题。实施顺序先消除会破坏现有数据格式和模板坐标的 P0 缺陷，再建立固定列/动态列共享的 Column Plan，随后实现布局、批注、选择器、读取范围、规范化、唯一性、失败工作簿、Workbook Data Validation 和图片导入，最后收敛公共 API、拆分 NPOI 大文件并完成测试、文档、互操作和 Benchmark。

计划只定义实施契约。本轮未修改业务代码、测试、项目配置或数据库，未创建 `execution.md`/`review.md`，未执行 Git 写操作、发布或实施任务。

## 2. 目标与非目标

### 2.1 目标

1. 保持 `ExcelExport.Workbook(...)`、`ExcelImport.Workbook<TWorkbook>(...)` 为唯一生产请求入口，核心 I/O 为 Stream-first。
2. 固定列和动态列共享同一不可变 Column Plan、Reader/Writer、Style、Layout、Normalization、Validation、Unique、Ignore 和 Image 配置。
3. 实现 `EXP-001`～`EXP-004`、`IMP-001`～`IMP-011`，并以 `IMP-012` 治理调用体验和 Web 场景。
4. 公共 API 使用 provider-neutral 类型；NPOI 类型只存在于适配器 internal 实现。
5. 所有结构错误、单元格错误和关系错误进入统一结果；失败工作簿受 Stream 所有权和资源上限控制。
6. 通过直接 internal 单元测试、组合集成测试、真实 Excel/LibreOffice 互操作、Public API approval 和 Benchmark 建立证据。
7. 允许未发布 API 的 Breaking Change，删除测试兼容路径和无真实消费者证据的旧 API。

### 2.2 非目标

1. 不支持 Microsoft 365 Threaded Comments；本轮批注仅指 NPOI `IComment` 对应的传统 Note/批注。
2. 不在首版承诺复合唯一键；首版唯一性范围为单 Sheet、单列。
3. 不在未验证前承诺 Named Range 列表、跨 Sheet/相对/自定义公式 Data Validation 完整支持。
4. 不新增左侧“行头”专用管线；该需求按自定义列头/多级表头解释，左侧行头用普通固定列表达。
5. 不把路径或 `byte[]` 便利方法提升为第二条执行管线，不在核心路径依赖真实数据库、网络或 Office 自动化。
6. 不扩展与本任务无关的 Word、PDF、CSV 功能。

## 3. 需求清单与 Traceability Matrix

| 原始需求 | 需求 ID | 当前状态（源码核验） | 目标 Task | 直接测试/组合测试 |
| --- | --- | --- | --- | --- |
| 1. 换行、对齐、居中 | EXP-001 | 部分完成；样式字段已存在，Body/Dynamic-only 有提前返回，整样式替换可丢 NumberFormat | P0-002、P2-001 | UT-STYLE-01～04、IT-04 |
| 1. 固定/AutoFit/Adaptive 列宽 | EXP-002 | 未实现；生产主链无列宽 Writer，旧 `SheetSetting.AutoColumnWidthEnabled` 未接 Workbook Request | P1-001、P2-001 | UT-WIDTH-01～05、IT-04、BM-01 |
| 1. 动态列统一能力 | EXP-003 | 部分完成；typed 动态值已读写，但 Converter/Validator、Unique、Whitespace、Image 未接线 | P0-002、P1-002、P4-001、P7-001 | UT-COL-01～04、IT-03 |
| 1. 自定义表头和批注 | EXP-004 | 多级表头部分完成；模板偏移错误；批注未实现 | P0-002、P2-002 | UT-HEADER-01～04、UT-COMMENT-01～04、IT-04 |
| 2. 按 Sheet 名称/下标 | IMP-001 | 名称已接主链且精确查找；下标只在测试兼容 Adapter；比较策略不可配 | P1-001、P3-001 | UT-SELECT-01～06、IT-01 |
| 3. 指定读取列范围 | IMP-002 | 未实现；`MaxColumnLength` 只是异常超宽阈值 | P1-001、P3-001 | UT-RANGE-01～05、IT-01 |
| 4. 错误消息/标注原文件 | IMP-003 | 结构化 Cell Error 部分完成；Header/Layout 会抛；无失败工作簿 | P5-001、P5-002 | UT-ERROR-01～04、UT-FAIL-01～05、IT-05 |
| 5. 仅错误数据文件 | IMP-004 | 未实现 | P5-001、P5-003 | UT-FAIL-06～10、IT-06 |
| 6. Header 大小写策略 | IMP-005 | 硬编码 `OrdinalIgnoreCase` | P1-001、P3-001 | UT-HEADER-05～07、IT-03 |
| 7. 首尾 Trim | IMP-006 | 硬编码在 NPOI `GetStringValue()`，Header/正文不可独立配置 | P1-001、P3-002 | UT-NORM-01～04、IT-03 |
| 8. 删除全部 Unicode 空白 | IMP-007 | 未实现 | P1-001、P3-002 | UT-NORM-05～08、BM-02 |
| 9. 指定列是否允许重复 | IMP-008 | `[Duplication]` 仅固定属性；无动态列/比较/首行坐标 | P4-001 | UT-UNIQUE-01～07、IT-03 |
| 10. 指定列/属性 Ignore | IMP-009 | Attribute/Profile/Configuration 已有；无 Request 级表达式；优先级需统一 | P1-002、P4-001 | UT-IGNORE-01～05、IT-01 |
| 11. Excel 原生 Data Validation | IMP-010 | 未实现；当前 Validation 只处理库内规则 | P1-001、P6-001 | UT-WDV-01～09、IT-05、IT-07 |
| 12. 图片单元格导入 | IMP-011 | 仅公开低层 Drawing/Picture 扩展；Importer 不扫描、不绑定 | P1-001、P7-001 | UT-IMAGE-01～09、IT-05～07 |
| 调用体验/Web 治理 | IMP-012 | Stream 主入口、取消、调用方流不关闭已存在；失败输出、资源限制和 Web 示例缺失 | P5-001、P8-001、P9-001 | UT-STREAM-01～06、IT-08、IT-10 |

## 4. 仓库、Git、构建和测试基线

### 4.1 已确认仓库事实

- 解决方案：`Bing.Offices.sln`，包含 Abstractions、Core、Npoi、Unit、Integration、Benchmark 和 BuildScript。
- 目标框架：Abstractions/Core 为 `netstandard2.0`；Npoi 为 `net8.0;net7.0;net6.0;netcoreapp3.1`；Unit 为 `net8.0;net7.0;net6.0;net5.0;netcoreapp3.1`；Integration 为 `net6.0;net8.0`；Benchmark 为 `net8.0`。
- 包锁：各项目使用 `packages.lock.json` 和 `RestorePackagesWithLockFile=true`；NPOI 由 `version.dev.props` 固定为 `[2.7.4]`；BenchmarkDotNet 为 `0.14.0`。
- 测试：xUnit 2.4.2；Unit 和 Integration 均直接引用 Npoi 项目。
- Public API：当前 `PublicApiContractTest` 只 approval 公开顶层类型，不 approval 成员签名；仍列出 `ExcelHelper`、NPOI Extensions、TypeMap、旧异常、全局 Setting 等大量公开类型。
- 资源：Unit 项目有 18 个历史 `.xlsx` 资源，主要用于旧导入回归；没有覆盖本轮模板偏移、批注、Data Validation、图片、失败工作簿和图表组合的明确互操作夹具。
- 编辑器静态诊断：分析时 `src`、`tests`、`benchmarks` 未报告诊断错误；这不等价于 `dotnet build/test` 通过。

### 4.2 指定证据的可用性

| 指定证据 | 结果 | 处理 |
| --- | --- | --- |
| `Bing.Offices-implementation-review-20260820.md` | `NOT_VERIFIABLE`：仓库内未找到 | 使用用户提示中的报告摘要作为次级输入，以源码/测试为当前态；Executor 开始前再次搜索，如补入仓库则记录差异 |
| `ai_docs/excel-import-export-enhancement-implementation-plan.md` | `NOT_VERIFIABLE`：仓库内未找到 | 以现有 `ai_docs/excel/00-overview.md`～`08-validation.md` 和 `implementation-progress.md` 替代旧计划证据 |
| `docs/excel-import-export-enhancement-api.md` | `NOT_VERIFIABLE`：仓库内未找到 | 以 `docs/excel/README.md` 和真实公共 API 替代草案证据 |
| 当前目标 `plan.md` | 原先不存在 | 本计划为首次创建，不存在旧计划合并问题 |
| Git 状态 | `NOT_VERIFIABLE`：当前 Planner 工具集不能执行 `git status` | `implementation-progress.md` 记载 2026-08-18 已有大量未提交改动；Executor 第一条命令必须记录 `git status --short --branch`，不得回退陌生改动 |

### 4.3 历史命令证据与本轮限制

`ai_docs/excel/implementation-progress.md` 记录上一轮：locked restore 因锁文件不一致失败，`--force-evaluate` 后恢复；Release build 通过；Unit net8 97、net6 99 通过；Integration net8/net6 各 10 通过；pack 通过。这是 2026-08-18 的历史证据，不代表 2026-08-20 当前工作树。

本轮 Planner 未运行 restore/build/test/benchmark，只进行了只读源码、配置、测试和编辑器诊断分析。因此当前 Build/Test/Benchmark 状态统一为 `NOT_VERIFIABLE`，Phase 0 必须重新建立基线。

## 5. 当前实现与真实调用链

### 5.1 导出调用链

`IExcelExporter.Export(request, destination, token)` → `NpoiExcelExporter.Export` → 创建/加载 Workbook → 对每个 `ExcelSheetExportRequest` 反射调用 `WriteTypedSheet<T>` → `ExcelTypeMapFactory.Get(profile, requestConfig)` → `CreateColumns` 合并固定列和 typed 动态列 → 写属性表头/数据 → `ApplyRequestStyles` → Attribute Header/Wrap/Merge → `CreateCharts` → 内存缓冲序列化后复制到调用方流。

真实进入生产链：异构多 Sheet、typed 动态 Key/Alias/DataType/NumberFormat、Before/After/物理位置、模板命名区域起点、请求样式缓存、XLSX Column/Line/Pie 图表、取消检查、调用方目标流保持打开。

未达到预期：

- `ApplyRequestStyles` 只检查 Header/Sheet style，故仅 BodyStyle 或仅 Dynamic style 时整体跳过。
- NPOI Style Cache 从空样式创建并整对象赋给 Cell，不能逐属性叠加模板/写值阶段的 DataFormat。
- `WriteCustomHeaders` 使用绝对 Row/Column，未叠加模板原点。
- 图表范围列索引是 Column Plan 相对索引，未叠加模板首列；显式 Anchor 是否相对模板也未形成统一约定。
- 动态 `ConverterName`、`ValidatorName` 没有被 `WriteRequestCell` 使用。
- 无列宽 Writer、Comment Writer 和资源限制。
- 导出必定完整序列化到 `MemoryStream` 后再复制，避免了部分输出但峰值内存至少包含 Workbook + 完整输出缓冲。

### 5.2 导入调用链

`IExcelImporter.Import(source, request, token)` → 整流复制到内存 → `WorkbookFactory.Create` → 对每个请求按 `GetSheetIndex(name)` 选 Sheet → `ImportTypedSheetCore<TWorkbook,TItem>` → 组装 `ExcelImportExecutionOptions<T>` → `ImportSheet` → `CreateColumns` → raw validation → conversion → configured validation → Setter → 成功实体加入根集合 → `BindTypedRelation`。

真实进入生产链：按精确名称选 Sheet、隐藏 Sheet 结构化错误、每 Sheet Header/Data 起点、固定和 typed 动态 Header/Alias/DataType、Mapping/Profile/Attribute Overlay、配置/特性校验、单 Sheet重复校验、结构化 Cell Error、关系绑定、不可寻址输入缓冲和取消。

未达到预期：

- Sheet 下标只存在于 `tests/.../ExcelCompatibilityTestAdapter.cs`，不是生产 API。
- Header 和 Alias 比较全部硬编码 `OrdinalIgnoreCase`。
- `GetStringValue()` 无条件 Trim 文本，Cell Reader 已丢失原始空白，Header/正文无法独立策略化。
- `MaxColumnLength` 只对 `header.LastCellNum` 抛异常，不是读取范围；缺失 Header、超宽、重复 Header 多处直接抛出。
- `DynamicTargetGetter` 被构建但未传入执行选项；最终总是调用映射属性 Setter，表达式目标没有真实作用。
- 动态 `ConverterName/ValidatorName` 未解析，动态列也跳过 Attribute/Configured Validation 和 Duplication。
- 重复状态只保存字符串 HashSet，不返回首次出现行；比较/空值策略不可配。
- 关系错误统一写入伪 Sheet `关系` 且无源 Sheet/Row/Key。
- 无 Workbook Data Validation 读取、Drawing Index、失败工作簿输出。
- `ExcelSheetImportResult.Items` 是公开 `object` 集合，根 Workbook 已有类型化集合，形成重复且弱类型结果面。

### 5.3 映射、DI 和低层能力

- `ExcelTypeMapFactory` 静态缓存 Attribute 映射，请求 Overlay 返回新 TypeMap；优先级代码为 Request Configuration > Profile > Attribute/Default，基本方向正确。
- `HeaderMappings` 独立于 Mapping/Profile，造成第二种标题配置方式。
- `ExcelIgnoreAttribute` 和 `ExcelColumnConfiguration.Ignored` 已生效，但没有 Sheet Request 的 `Ignore(x => ...)`。
- DI 将 Importer/Exporter 注册为 transient，规则/映射 Loader 为 singleton；请求状态主要为局部变量，可复用。
- `SheetExtensions.Picture.cs` 能枚举/添加/移动 HSSF/XSSF 图片，但公开 NPOI 类型且不在 Importer 主链。
- `SheetSetting.AutoColumnWidthEnabled`、旧异常、`ExcelHelper` 和多个公开 NPOI 扩展属于旧 API 面；Workbook Request 主链没有消费其中大部分能力。

## 6. 完成状态与评分

### 6.1 能力分类

**已完成并进入主链**：Workbook Request、异构多 Sheet、按名称导入、Header/Data 起点、固定列基础映射、typed 动态值读写、模板 Workbook/命名区域基础起点、请求样式基础缓存、父子关系基础绑定、结构化 Cell Error 基础字段、Stream 输入输出/取消、XLSX 三类基础图表。

**部分完成/需要重构**：Wrap/Alignment、动态列、Custom Header、模板偏移、结构化错误、Duplicate、Ignore、Mapping 优先级、Stream 内存策略、Public API approval、关系错误坐标。

**仅有底层类型或测试兼容能力**：Sheet Index（测试 Adapter）、Picture Drawing 扩展、旧 AutoColumnWidth Setting、动态 Converter/Validator 字段、DynamicTargetGetter、旧单 Sheet Result。

**未实现**：统一 Width Layout、Comment Writer、Read Range、可配置名称比较/空白、AnnotatedOriginal、ErrorRowsOnly、Workbook Data Validation、图片绑定、失败工作簿资源上限、Docs Consumer、本轮 Benchmark 矩阵。

### 6.2 当前完成度

对本轮 16 个功能 ID（EXP 4 + IMP 12）按“完整=1、部分=0.5、未实现=0”计分：完整 1 项（IMP-012 的基础 Stream 契约不算完整），部分 8 项，未实现 7 项，功能满足度约 `31%`。工程基础（Workbook 主链、测试框架、DI、锁文件、基础 Benchmark）约 `55%`。按功能 70%、工程基础 30% 加权，本轮范围当前完成度约 `38%`。

该评分低于旧报告整体 62%，原因是本轮加入了失败工作簿、Workbook Data Validation、图片、列宽、规范化和资源治理等尚未实现的较大需求；它不是对上一轮 Workbook Request 工作的否定。Executor 完成每个 Phase 后应按相同口径更新 `execution.md`，不得修改本计划伪造进度。

## 7. 关键缺陷与根因

| 缺陷 | 根因 | 影响 | 优先级 |
| --- | --- | --- | --- |
| Body/Dynamic 样式不生效 | `ApplyRequestStyles` guard 未包含 Body/Dynamic style | EXP-001/003 | P0 |
| NumberFormat/模板样式丢失 | Style 使用整对象替换而非 base + overlay 合成 | 数据类型和显示格式回归 | P0 |
| 模板非 A1 布局错位 | Header、Chart、Merge 等没有共享 Layout Origin | EXP-004、图表 | P0 |
| 动态扩展点无效 | 固定/动态列只有局部合并，转换/校验仍分支实现 | EXP-003、IMP-008/010/011 | P0 |
| 结构错误直接抛出 | Header Binding 与 Result 聚合耦合，缺少结构错误通道 | IMP-001～005 | P0 |
| 原始空白不可恢复 | Provider Reader 提前 Trim | IMP-006/007 | P1 |
| API 平行且重复 | 历史 Setting/Helper/Extensions、HeaderMappings、重复位置字段未收敛 | 开发体验、维护性 | P1 |
| 大文件职责混杂 | Importer/Exporter 内嵌 Planning/Reading/Writing/Validation/Layout | 单测困难，后续功能互相影响 | P1 |
| 内存峰值不可控 | 输入/输出全量缓冲且无文件/失败文件上限 | 大文件和恶意输入风险 | P0/P1 |

## 8. API 目标设计

### 8.1 公共策略类型

采用枚举/值对象，禁止新增冲突 Boolean：

```text
ExcelColumnWidthMode: None / Fixed / AutoFit / Adaptive
ExcelWhitespacePolicy: Preserve / Trim / RemoveAll
ExcelNameComparison: Ordinal / OrdinalIgnoreCase
ExcelSheetSelector: ByName(name) / ByIndex(zeroBasedIndex)
ExcelImportFailureWorkbookMode: None / AnnotatedOriginal / ErrorRowsOnly
ExcelImportValidationMode: Disabled / ConfiguredRules / WorkbookRules / ConfiguredAndWorkbook
ExcelImageMultiplicityPolicy: First / All / Fail
ExcelCommentConflictPolicy: Preserve / Append / Replace / Fail
ExcelUnsupportedFeaturePolicy: Fail / Report
```

### 8.2 统一 Column 配置

公共 Builder 仍用强类型表达式配置固定属性，用稳定 Key 配置动态列，但两者编译成同一个 internal `ExcelColumnPlan`。Plan 至少包含：Key、Title/Aliases、物理列、目标类型、Getter/Setter、Converter、Configured Validator、Workbook Validator、Header/Body Style Overlay、Width、Comment、Whitespace、Unique、Ignore、Image。Writer/Reader 只消费 Plan，不再按 fixed/dynamic 维护平行业务分支。

### 8.3 强制收敛结论

| 现有 API | 结论 | 迁移方向 |
| --- | --- | --- |
| `MaxColumnLength` | 直接 Breaking Rename 为 `MaxColumnCount` | 新增独立 `ReadColumns(startIndex,count)`；前者只做安全上限 |
| `HeaderMappings` | 删除 | 合并到 `ExcelMapping`/Profile 的 `HasTitle`/Alias；Request Overlay 最高优先级 |
| 动态 `ConverterName/ValidatorName` | 不保留当前字符串单值形态 | 合并到统一 Column 配置；Converter 预绑定为唯一命名转换器，Validation 使用规则名称集合；无引用字段删除 |
| `PhysicalColumnIndex` + `Placement.At` | 删除前者，保留单一 `Placement.At(index)` | fixed/dynamic 均编译为同一布局约束 |
| `HeaderRows` | 改为独立 Header Layout | 明确 Title Rows、Property Header Row、模板相对坐标；避免“自定义 Header 不能覆盖属性 Header”的隐式职责 |
| `RowIndex/ColumnIndex` | Breaking Rename 为 `RowNumber/ColumnNumber` | 公共错误始终一开始；internal 坐标明确 `rowIndex/columnIndex` 为零开始 |
| `HasMany` string key | 升级为 `HasMany<TKey>` 并接收 `IEqualityComparer<TKey>` | 默认 `EqualityComparer<TKey>.Default`；错误带父/子来源坐标和 Key |
| `ExcelSheetImportResult.Items` | 删除公开 object 集合 | 根 Workbook 保留强类型集合；Sheet Result 只暴露统计、来源行和错误，必要时提供受控泛型查询而非 object 列表 |
| `ExcelImportResult<T>` | 删除生产公开类型 | 只保留 Workbook Result；测试 Adapter 迁移后删除 |
| `ExcelSetting`/`SheetSetting` | 无主链消费者则删除 | 所有配置进入 Workbook/Sheet/Column Request |
| 旧 Converter/旧异常 | `ICellValueConverter` 和未被主链契约使用的 Office 异常删除 | 统一 provider-neutral `IExcelValueConverter` 和结构化 Result；配置错误使用标准异常 |
| Importer/Exporter 具体类 | 调整为 internal | 消费者通过 `IExcelImporter`/`IExcelExporter` + `AddNpoi()`；测试通过 DI 或 InternalsVisibleTo |
| TypeMap/PropertyMap/低层 Resolver/Writer/Extension | 调整为 internal | 公共只保留 Mapping Builder/Profile 和 provider-neutral DTO |
| `ExcelHelper`/公开 NPOI Extensions | 删除或 internal | 无真实外部消费者证据时不保留；必要能力由 internal adapter 使用 |

### 8.4 Stream 和失败输出契约

- 输入流、导出目标流、失败工作簿目标流均由调用方拥有，核心方法不关闭它们；模板流继续由 `UseTemplate(stream, leaveOpen)` 显式控制，因为 Workbook 生命周期可能持有模板。
- Import 请求声明 Failure Mode；`None` 不要求失败输出流，其他模式要求调用方提供可写失败目标流。参数组合在读取前验证，失败流只在有错误且模式非 None 时写入。
- `byte[]`/路径只作为薄封装；`byte[]` 文档标明仅适合小文件。
- CancellationToken 在输入复制、Sheet/Row/Column 循环、Validation/图片索引、失败 Workbook 复制和输出复制中检查。

## 9. Breaking Change 与迁移策略

项目当前未正式发布，默认一次性破坏性收敛，不添加 Obsolete 平行 API。实施时先在 `ai_docs` 记录 API 决策和迁移表，再更新生产代码、测试 Adapter、Public API approval 和文档消费者。只有 Phase 0 的 Git/包分析发现真实外部项目引用证据时，才允许增加最薄迁移包装，并必须写明删除版本；NuGet 包历史本身不构成保留无效 API 的证据。

迁移示例必须覆盖：单 Sheet Options → Workbook Request；SheetIndex → `ExcelSheetSelector.ByIndex`；HeaderMappings → Mapping Overlay；Duplication Attribute → `RequireUnique`；旧 Result → Workbook Result；直接 `new NpoiExcelImporter/Exporter` → DI 接口；NPOI PictureInfo → `ExcelImageData`。

## 10. internal 化、删除和审批清单

**确定删除/收敛**：测试 `ExcelCompatibilityTestAdapter.cs` 及其生产 API 假象、`ExcelImportResult<T>`、`HeaderMappings`、`MaxColumnLength`、动态重复物理索引字段、legacy `ICellValueConverter`、弱类型 Sheet Items。

**确定 internal 化**：NPOI Importer/Exporter 实现类、`NpoiStyleCache`、Column Plan/Reader/Writer、Validation Range Index、Drawing Index、Failure Workbook Writer、NPOI Resolver。

**候选删除（需 P0-001 用引用搜索确认）**：`ExcelHelper`、`ExcelSetting`、`SheetSetting`、`ExcelTypeMap<T>`、`ExcelPropertyMap`、`ExcelTypeMapFactory`、`PictureInfo`、`PictureStyle`、`MergedRegionInfo`、`OfficeHeaderException`、`OfficeEmptyLineException`、`OfficeDataConvertException`、公开 Cell/Row/Sheet/Workbook/Font/Style Extensions、旧 Color 类型。

Public API approval 必须从“公开顶层类型名单”升级为类型 + 成员签名基线，并增加断言：Abstractions/Core 公开签名不得引用 `NPOI.*`，Npoi 包对消费者只公开 DI 注册入口和必要 provider 注册元数据。

## 11. 文件、目录和命名空间拆分

### 11.1 已确认需修改的现有文件

- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/*`：Workbook/Sheet Builder、动态列、Header、Chart、接口。
- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/*`：Workbook/Sheet Builder、Request、Result、Error、接口。
- `src/Bing.Offices.Abstractions/Bing/Offices/Styles/ExcelCellStyle.cs`。
- `src/Bing.Offices.Abstractions/Bing/Offices/Configurations/*`。
- `src/Bing.Offices.Core/Bing/Offices/Configurations/ExcelMapping.cs`、`Mappings/*`、`Attributes/*`、`Extensions/ExcelStreamExtensions.cs`。
- `src/Bing.Offices.Npoi/Exports/NpoiExcelExporter.cs`、`NpoiStyleCache.cs`。
- `src/Bing.Offices.Npoi/Imports/NpoiExcelImporter.cs`、`ExcelImportExecutionOptions.cs`。
- `src/Bing.Offices.Npoi/Extensions/CellExtensions.cs`、`SheetExtensions.Picture.cs`、`Extensions.Service.cs`。
- Unit、Integration、Benchmark、Public API 测试及现有 Excel 文档。

### 11.2 候选新增/拆分文件

以下是职责目标，不要求机械一类一文件；Executor 在每个 Task 中按最小可测试边界落地：

```text
src/Bing.Offices.Abstractions/Bing/Offices/
  Exports/Columns/ExcelColumnLayout.cs
  Exports/Comments/ExcelComment.cs
  Imports/ExcelSheetSelector.cs
  Imports/ExcelImportPolicies.cs
  Imports/Failures/ExcelImportFailureOptions.cs
  Imports/Images/ExcelImageData.cs
  Workbooks/ExcelResourceLimits.cs

src/Bing.Offices.Npoi/
  Exports/Planning/NpoiExportPlanCompiler.cs
  Exports/Writing/NpoiSheetWriter.cs
  Exports/Layout/NpoiColumnWidthWriter.cs
  Exports/Styles/NpoiStyleComposer.cs
  Exports/Comments/NpoiCommentWriter.cs
  Imports/Planning/NpoiImportPlanCompiler.cs
  Imports/Reading/NpoiSheetReader.cs
  Imports/Reading/NpoiCellReader.cs
  Imports/Validation/NpoiWorkbookValidationReader.cs
  Imports/Validation/ValidationRangeIndex.cs
  Imports/Images/NpoiDrawingIndex.cs
  Imports/Failures/NpoiAnnotatedWorkbookWriter.cs
  Imports/Failures/NpoiErrorRowsWorkbookWriter.cs

 tests/Bing.Offices.Tests/Excel/
  Exports/*Tests.cs
  Imports/*Tests.cs
  PublicApi/*Tests.cs
 tests/Bing.Offices.Tests.Integration/Excel/*Tests.cs
 tests/Bing.Offices.Docs.Tests/ (候选新测试项目，加入解决方案)
 tests/Bing.Offices.Tests.Integration/Resources/Interop/*
```

命名空间按现有 `Bing.Offices.Exports`、`Imports`、`Styles` 保持消费者稳定；internal NPOI 可细分 `Bing.Offices.Npoi.Exports.Planning/Layout/Comments` 与 `Imports.Reading/Validation/Failures/Images`。拆分验收以职责可独立测试为准，不设任意行数门槛。

## 12. 分阶段实施计划

### Phase 0：仓库基线与阻塞 P0

#### P0-001 基线、消费者证据和决策冻结

- **优先级/依赖**：P0；无依赖；阻塞全部任务。
- **当前证据**：历史 locked restore 曾失败；当前 Git 状态和当前构建未验证；三份指定设计/审查文档缺失；Public API 只验证顶层类型。
- **目标**：保存当前工作树、依赖、测试、API、文件尺寸/职责和消费者引用基线，不修改陌生改动；冻结 Breaking API 决策。
- **已确认文件**：解决方案、所有 `.csproj`/lock/props、`PublicApiContractTest.cs`、`implementation-progress.md`。
- **候选文件**：本计划列出的候选删除类型及任何 Executor 搜索到的仓库内消费者项目；实施证据只写 `execution.md`。
- **步骤**：记录 Git status/diff；执行 locked restore；若仅 lock 漂移，查明 csproj 变更来源后单独更新锁并记录；运行基线 build/unit/integration/public API/benchmark smoke；用符号引用确定候选 API 是否有生产消费者；记录当前 Importer/Exporter 行数和职责；形成 API 决策表。
- **API/数据/配置影响**：本 Task 不改 API；确定后续 Breaking 清单和锁文件处理方式。
- **测试**：执行第 19 节基线命令；失败必须原样记录，不得先改代码掩盖。
- **性能**：运行现有 1K 基准 smoke，记录环境、时间和 Allocated，不作为新功能验收阈值。
- **风险**：大量未提交改动可能与计划重叠；必须逐文件保留用户修改。
- **验收**：`execution.md` 有 Git、restore/build/test/API/benchmark 基线及消费者引用证据；缺失文档保持 `NOT_VERIFIABLE` 或补充真实路径。

#### P0-002 修复样式合成、模板坐标、动态死字段和结构错误基线

- **优先级/依赖**：P0；依赖 P0-001；阻塞 Phase 1～2、5。
- **当前证据**：`ApplyRequestStyles` guard 漏 Body/Dynamic；整样式替换；Custom Header 和 Chart Column 未应用模板列偏移；动态 Converter/Validator 与 DynamicTarget 未接线；Header/MaxColumn 多处抛异常。
- **目标**：在不新增最终策略 API 前建立正确的 style overlay、统一 template origin 和结构错误通道；删除或标记待统一的死字段，防止后续功能建在错误基础上。
- **已确认文件**：`NpoiExcelExporter.cs`、`NpoiStyleCache.cs`、`NpoiExcelImporter.cs`、Export/Import Request/Builder、现有 Workbook/P0 测试。
- **候选文件**：`NpoiStyleComposer.cs`、`NpoiLayoutOrigin.cs`、结构错误内部模型测试。
- **步骤**：从模板/当前 CellStyle clone 后逐属性覆盖；明确 Request > Profile > Attribute > Default 且 column > region > sheet；所有 Header/Merge/Chart 使用同一 origin；仅明确为绝对 Sheet 坐标的 Chart Anchor 不偏移并写入契约；把 Header missing/duplicate/too-wide 转为 Result Error；决定动态字段在 P1 接线前保留或立即改为统一配置；移除未使用 DynamicTarget Getter 或接入统一 Setter 计划。
- **API/数据/配置影响**：不新增平行 Wrap API；允许修正现有行为；错误代码可先 internal 扩充，最终公共命名在 P5-001 完成。
- **测试**：仅 BodyStyle、仅 DynamicStyle、NumberFormat + Wrap、模板原样式叠加、B3 Header/Merge/Chart、自定义 Header 冲突、Header 结构错误不抛；完整断言 CellType/DataFormat/style/坐标。
- **性能**：Style key 必须包含 base style identity + overlay；同一合成结果复用，样式数与唯一组合数成正比。
- **风险**：NPOI HSSF/XSSF clone 差异；图表显式范围相对语义可能是 Breaking。
- **验收**：P0 回归全绿；Body/Dynamic-only 生效；NumberFormat 不丢；非 A1 模板组合坐标正确；结构错误进入 Result。

### Phase 1：公共策略与统一 Column Plan

#### P1-001 定义公共策略、资源限制和 Sheet/Read/Failure 契约

- **优先级/依赖**：P0；依赖 P0-002。
- **当前证据**：当前以 Boolean/硬编码行为表达策略，缺少 Width、Whitespace、Selector、Read Range、Failure/Validation/Image 模型和资源上限。
- **目标**：一次性批准 provider-neutral 策略值对象和默认值，避免后续各 Phase 反复改 API。
- **已确认文件**：Abstractions Export/Import Builder、Request、Style、Error、接口。
- **候选文件**：第 11.2 节 Abstractions 新类型；`ExcelResourceLimits`。
- **步骤**：实现第 8.1 枚举；定义 Width（字符单位、mode/fixed/min/max/sample/template policy）、Selector、Read Range、Failure Output、Validation Mode、Image Data/Multiplicity、Comment、资源限制；在 Build 阶段验证互斥参数和正数上限；所有默认值写入 XML 注释与 approval。
- **API/数据/配置影响**：Breaking rename `MaxColumnLength`；新增失败目标流契约；公共类型不得引用 NPOI。
- **测试**：枚举非法值、Selector 互斥、Read Range 溢出、Width min/max、Failure mode 缺输出、资源上限非法组合。
- **性能**：默认禁用高成本 AutoFit/Workbook Validation/Image/Failure Workbook；限制对象不可变且请求级。
- **风险**：一次性 API 面较大；必须先用 Docs Consumer 草案编译确认可用性。
- **验收**：所有策略有明确默认值/索引/单位/所有权；无冲突 Boolean；Public API 签名无 NPOI。

#### P1-002 编译统一不可变 Column Plan

- **优先级/依赖**：P0；依赖 P1-001。
- **当前证据**：固定/动态只在基础位置和值写入上合并，转换/校验/Ignore/Unique/Image 仍分支；Converter 每 Cell 解析。
- **目标**：固定与动态配置编译为同一种 internal Column Plan，所有反射、表达式、Converter/Validator 在 Sheet 执行前绑定。
- **已确认文件**：Mapping/Profile/TypeMap、Export `CreateColumns`、Import `CreateColumns`、动态定义。
- **候选文件**：Export/Import Plan Compiler、共享 provider-neutral plan contracts（internal）。
- **步骤**：定义优先级 Request > Profile > Attribute > Default；将 Ignore、Title/Alias、Placement、Style、Layout、Normalization、Unique、Comment、Image 编译到列；命名转换器/规则要求唯一匹配；统一动态 Target Setter；检测 Key/Title/Alias/物理位置冲突；删除 `HeaderMappings` 和重复位置字段。
- **API/数据/配置影响**：Mapping Builder 增加 Request expression overlay；动态列改用统一 Column 配置面；TypeMap 逐步 internal。
- **测试**：固定/动态计划等价、优先级、冲突、命名解析缺失/重复、Setter、并发请求隔离、未知动态策略。
- **性能**：每 Sheet 编译一次；逐 Cell 禁止反射、LINQ converter 查找和规则字符串查找；缓存仅含静态类型元数据，请求配置不得进入全局 cache key。
- **风险**：静态 TypeMap 缓存污染、多租户请求串扰。
- **验收**：Reader/Writer 仅消费 Column Plan；固定/动态没有第二套同名业务能力；并发隔离测试通过。

### Phase 2：导出布局与批注

#### P2-001 Wrap/Alignment 与 Fixed/AutoFit/Adaptive Width Writer

- **优先级/依赖**：P1；依赖 P1-002。
- **当前证据**：Style 字段存在；无 Width Writer；旧 AutoColumnWidth Setting 不在主链。
- **目标**：实现 Sheet/region/column style overlay 和四种 Width Mode，公共单位为 Excel 字符宽度。
- **已确认文件**：Style、Export Builder/Request、Exporter、Style Cache。
- **候选文件**：Column Layout、Style Composer、Width Estimator/Writer 测试。
- **步骤**：实现居中快捷方法为 style 配置糖；Fixed 转 NPOI 1/256 并 clamp；AutoFit 在写完后调用实际 Cell 测量且显式开启；Adaptive 按换行逐段估算 ASCII/CJK/全角/Emoji，采样数据行并应用 min/max；模板宽度 Preserve/Override 明确；格式化显示值参与估算。
- **API/数据/配置影响**：新增 Width 配置；删除旧 Setting auto-width。
- **测试**：单位转换、边界、中文/英文/Emoji/多行、格式化数值、采样、模板保留、fixed/dynamic 优先级、NumberFormat 不丢。
- **性能**：AutoFit O(rows×columns)；Adaptive O(sampleRows×columns)，默认采样上限；无宽度模式不扫描正文。
- **风险**：NPOI AutoSize 对字体/平台测量差异；断言使用范围和相对关系，不依赖单一机器像素值。
- **验收**：重开 Workbook 后 ColumnWidth、Style、DataFormat 与策略一致；100K Adaptive 不扫描超过采样上限。

#### P2-002 自定义 Header Layout 与 Comment Writer

- **优先级/依赖**：P1；依赖 P0-002、P2-001。
- **当前证据**：HeaderRows 支持 merge 但职责混合且模板偏移错误；无 Comment Writer。
- **目标**：实现标题/属性表头清晰分层、模板相对布局、Header/Body Comment 和冲突策略。
- **已确认文件**：HeaderRow/Cell、Sheet Export Builder、Exporter。
- **候选文件**：Header Layout、Comment DTO/Writer/Cache。
- **步骤**：预编译相对 origin 的 Header Layout；预检重叠、越界、模板既有 merge/comment 冲突；实现 Preserve/Append/Replace/Fail；支持作者、文本、可见性；Body Comment 可由静态定义或行值委托产生；缓存 CreationHelper/anchor/style，限制 MaxComments。
- **API/数据/配置影响**：`HeaderRows` Breaking 迁移到 Header Layout；Comment provider-neutral。
- **测试**：跨行/列 merge、非 A1 偏移、冲突策略、同 Cell 追加、动态列 comment、MaxComments、XLS/XLSX 支持差异。
- **性能**：无 Comment 配置不创建 Drawing；Comment 数 O(configured cells)，达到上限返回明确错误。
- **风险**：传统 Note 与 Threaded Comment 语义混淆；文档明确只支持传统批注。
- **验收**：重开 `.xls/.xlsx` 检查 Comment author/text/visible、Merged Region 和模板保留；冲突不静默覆盖。

### Phase 3：导入选择、读取与规范化

#### P3-001 Sheet Selector、Read Range、MaxColumnCount 和名称比较

- **优先级/依赖**：P1；依赖 P1-002。
- **当前证据**：生产只按精确名称，Builder 按名称去重时忽略大小写；Read Range 缺失；结构错误部分抛出。
- **目标**：按 `ByName/ByIndex` 选择，Header/Sheet 比较独立配置，读取范围真正限制绑定管线。
- **已确认文件**：Import Builder/Request、Importer、ExecutionOptions。
- **候选文件**：Selector Resolver、ReadRange、Header Binder。
- **步骤**：ByIndex 零开始；名称按独立策略解析并检测歧义；隐藏/missing/out-of-range/duplicate selection 结构化；只枚举 `[start,count]` 内 Header Cells；HeaderMatch 只要求范围内配置列；超出 MaxColumnCount 在绑定前拒绝；稀疏行按物理列索引处理。
- **API/数据/配置影响**：Sheet 方法改接 Selector，保留便利 `Sheet("name",...)` 仅作为创建 Selector 的薄糖；`MaxColumnCount` 与 Read Range 分离。
- **测试**：名称大小写、重复名称歧义、index 0/末尾/越界、隐藏、稀疏 Header、范围外必填列不误报、恶意 LastCellNum。
- **性能**：只为读取范围建立 binding；超宽输入在进入逐行处理前失败。
- **风险**：Excel 自身不允许同 Workbook 完全重复名称，但大小写策略可能产生语义重复。
- **验收**：所有选择/范围错误进入 Result；范围外 Cell 不转换、不验证、不写 Setter。

#### P3-002 Header/Body Whitespace Normalization

- **优先级/依赖**：P1；依赖 P3-001。
- **当前证据**：`GetStringValue()` 对 String/default 无条件 Trim，已破坏原始文本；无 RemoveAll。
- **目标**：Cell Reader 保留原始 typed value/text；Header 与正文分别执行 Preserve/Trim/RemoveAll。
- **已确认文件**：CellExtensions、Importer ReadCellValue、Column Plan、Header Binder。
- **候选文件**：NpoiCellReader、ExcelStringNormalizer。
- **步骤**：移除 provider Reader 的 Trim；只对字符串执行策略；RemoveAll 用 `char.IsWhiteSpace` 单遍扫描，先检测是否需分配；Numeric/Date/Boolean/Formula typed 值不转字符串后规范化；Unique 在规范化后运行；RawValue 保留原始值，另记录 normalized 值仅在需要时。
- **API/数据/配置影响**：Sheet HeaderWhitespace 和默认 BodyWhitespace；Column 可覆盖 Body。
- **测试**：ASCII space/tab/CRLF、全角空格、NBSP、Unicode whitespace、中间空白、纯空白、Numeric/Date/Formula 不变、Header 与正文独立。
- **性能**：O(text length)，无变化返回原字符串；禁止逐 Cell Regex。
- **风险**：改变当前隐式 Trim 默认行为；迁移指南明确默认值并由 API 决策决定 Preserve 或 Trim。
- **验收**：三策略结果精确；RawValue 可追溯；Benchmark 显示 RemoveAll 分配受控。

### Phase 4：唯一性与 Ignore

#### P4-001 统一 Unique、Ignore 和关系 Key

- **优先级/依赖**：P1；依赖 P1-002、P3-002。
- **当前证据**：Duplication 仅 Attribute 固定列，HashSet 无首次行；Ignore 缺 Request expression；HasMany 仅 string key。
- **目标**：固定/动态列同样支持 `RequireUnique()`，Request Ignore 复用 Mapping Overlay，关系升级泛型 Key/Comparer。
- **已确认文件**：Duplication Attribute/Rule、Mapping Builder/TypeMap、Importer、Relation Request。
- **候选文件**：Unique Plan/Tracker、关系索引。
- **步骤**：定义空值和比较策略；用 `Dictionary<normalizedValue, sourceRow>`；失败行回滚本行新增；Error 带首次/重复行；Ignore 在 Header binding 前移除 Column Plan；被忽略列不得读/转/验/unique/set；Attribute 迁移到同一 plan；`HasMany<TKey>` 使用 comparer 并保留来源行映射。
- **API/数据/配置影响**：新增 `Ignore(x=>...)`、`RequireUnique()`；旧 Duplication 可直接删除或仅作为 Attribute 输入编译到同一 Plan，禁止第二套执行器。
- **测试**：默认允许、空值、大小写、规范化后重复、首次行、失败回滚、动态列、Ignore 三层优先级、关系非 string key/comparer/错误来源。
- **性能**：Unique/Relation O(n)，内存 O(unique values)；MaxRows/MaxErrors 限制极端输入。
- **风险**：可变/复杂 TKey comparer；首版文档限制为稳定可比较键。
- **验收**：无嵌套扫描；固定/动态行为一致；Ignored 属性全链路零副作用。

### Phase 5：结构化错误与失败工作簿

#### P5-001 统一 Error Model、统计和失败输出 Stream 契约

- **优先级/依赖**：P0；依赖 P1-001、P3-001、P4-001。
- **当前证据**：Error 缺 ErrorCode string、Validation Source、PropertyPath、首次重复行、关系来源；Sheet Items 弱类型；部分结构错误抛出。
- **目标**：所有可预期输入错误进入统一结果，配置/编程错误仍抛标准异常；批准失败输出流契约。
- **已确认文件**：Import Error/Code/Workbook Result/Sheet Result、Importer、接口/扩展。
- **候选文件**：Failure Options、Error Source/Location DTO、Error Collector。
- **步骤**：Rename Row/Column Number；增加 Header/ColumnKey/PropertyPath/ErrorCode/Message/RawValue/ValidationSource/FirstOccurrence/RelationSource；区分 workbook/sheet/row/cell 范围；MaxErrors 达到后截断并报告；删除 object Items；定义 FailureWritten/FailureSize 等结果元数据；验证 stream ownership/cancellation。
- **API/数据/配置影响**：Breaking Error/Result；Failure destination 参数/执行选项新增。
- **测试**：Header/Layout/Cell/Relation/Unsupported/limit 错误、统计去重、坐标一开始、配置错误仍抛、失败流写异常/预取消/保持打开。
- **性能**：Error 收集受 MaxErrors；RawValue 不保留大二进制图片副本，仅摘要/元数据。
- **风险**：错误代码枚举未来扩展；使用稳定字符串 code 或可扩展值对象并 approval。
- **验收**：需求列出的所有错误字段可追踪；无预期输入错误逃逸为未分类异常。

#### P5-002 AnnotatedOriginal Writer

- **优先级/依赖**：P1；依赖 P5-001、P2-002。
- **当前证据**：不存在失败 Writer；Importer 已有完整输入缓冲，可作为原 Workbook 副本来源。
- **目标**：在原 Workbook 副本标记错误 Cell，并生成 `_ImportErrors` 汇总 Sheet，尽量保留原结构。
- **已确认文件**：Importer 输入缓冲、Style/Comment 低层能力。
- **候选文件**：Annotated Writer、Error Style Cache、Summary Sheet Writer。
- **步骤**：仅有错误时重开/复用副本；按 Cell 分组合并多错误批注；clone 原样式后叠加红边/填充；非 Cell 错误只入汇总；避免名称冲突；写入失败流前检查 MaxFailureWorkbookSize；保留公式、CellType、merge、validation、drawing；取消时不产生后续块。
- **API/数据/配置影响**：Failure Mode `AnnotatedOriginal`；错误样式可 provider-neutral 配置但有安全默认。
- **测试**：多错误批注合并、样式缓存、非 Cell 汇总、公式/样式/merge/validation/picture 保留、已有 comment 冲突、无错误不输出、size limit。
- **性能**：最多一次 Workbook 副本和一次序列化；style/font/comment anchor 缓存；避免每错误创建样式。
- **风险**：NPOI 重写高级 OOXML 部件可能有保真限制；互操作矩阵必须记录。
- **验收**：重开失败文件逐项断言；原值和 CellType 不被 DTO 重建替换；资源上限生效。

#### P5-003 ErrorRowsOnly Writer

- **优先级/依赖**：P1；依赖 P5-001、P7-001 的图片复制接口可后接。
- **当前证据**：不存在；不能从已转换 DTO 恢复原 CellType。
- **目标**：按源 Workbook 原始 Header/失败行复制，每源行一次，多 Sheet 分别输出并汇总。
- **已确认文件**：Importer source row tracking、Error Model。
- **候选文件**：Error Rows Writer、Cell/Style/Merge/Validation/Picture Copy helpers。
- **步骤**：按 Sheet+SourceRow 去重；复制原 Header/Layout 和原失败行；追加四个来源/错误列；同一行多错误聚合；调整 merge/data validation 范围；复制关联图片并重算 anchor；Sheet 名冲突/31 字符处理；无错误不写。
- **API/数据/配置影响**：Failure Mode `ErrorRowsOnly`；不新增返回 byte[] 核心 API。
- **测试**：行去重、原 CellType/Formula/DataFormat、附加列完整、多 Sheet、图片 anchor、无错、写失败、size limit。
- **性能**：按失败行集合复制，不扫描 DTO；样式映射缓存；输出大小预估和硬上限。
- **风险**：跨失败/成功行 merge 无法原样复制；定义裁剪/报告策略，禁止静默错误 merge。
- **验收**：重开文件只含 Header+唯一失败行；来源字段和汇总精确；原始 Cell 类型保留。

### Phase 6：Workbook Data Validation

#### P6-001 Rule Reader、Range Index 和支持矩阵

- **优先级/依赖**：P1；依赖 P1-002、P3-001、P5-001。
- **当前证据**：Importer 只运行 Attribute/Named rules；无 Sheet Data Validation 读取。
- **目标**：实现 Workbook Rules 与 Configured Rules 独立/组合模式，预编译每 Sheet Range Index。
- **已确认文件**：Validation abstractions、Importer、NPOI Sheet API。
- **候选文件**：Workbook Validation Reader、Rule DTO/Compiler、Range Index、Interop samples。
- **步骤**：先用最小 HSSF/XSSF 样本验证 NPOI API；支持显式列表、整数/小数、日期/时间、文本长度、常量操作数和直接 Cell Range；构建按行区间/列区间查询的索引；Unsupported 按 Fail/Report，不得通过；Configured + Workbook 错误标来源；范围外规则不执行。
- **API/数据/配置影响**：Validation Mode 和 Unsupported Policy；公共不暴露 NPOI constraint。
- **测试**：每种规则成功/失败/边界、重叠规则、稀疏范围、Unsupported、HSSF/XSSF、范围外、组合模式。
- **性能**：Sheet 级预编译一次；逐 Cell 查询不得遍历全部规则；Benchmark 1K/10K/100K 多规则。
- **风险**：Named Range、跨 Sheet、自定义/相对公式和 HSSF/XSSF 差异。
- **验收**：首版支持项全部有直接单测和重开集成断言；待验证项明确 Report/Fail，无静默放行。

### Phase 7：图片导入

#### P7-001 Drawing Index、Image Binding 和资源限制

- **优先级/依赖**：P1；依赖 P1-002、P3-001、P5-001；为 P5-003 图片复制提供接口。
- **当前证据**：低层 Picture 扩展可枚举 HSSF/XSSF，但公开 NPOI/legacy DTO，Importer 不调用。
- **目标**：仅在配置图片列时每 Sheet 扫描 Drawing 一次，以 top-left anchor 绑定 provider-neutral 图片数据。
- **已确认文件**：`SheetExtensions.Picture.cs`、`PictureInfo/PictureStyle`、Importer Column Binding。
- **候选文件**：`ExcelImageData`、Drawing Index、Image Binder/Limits、Failure copy adapter。
- **步骤**：索引图片且忽略非图片 Shape；校验 anchor；处理 merged/cross-cell/floating/hidden row/column；First/All/Fail；绑定 `byte[]`、`ExcelImageData`、只读集合；限制单图 bytes、总图数/bytes、可选像素尺寸/格式；错误不包含完整 binary；将复制能力接入两个失败 Writer。
- **API/数据/配置影响**：Column `AsImage(...)`；公共 Image Data 只含 bytes/content type/尺寸/anchor metadata，不含 NPOI。
- **测试**：XSSF/HSSF、top-left、merged、跨格、一格多图、非图片/损坏 anchor、hidden、类型不匹配、所有资源上限。
- **性能**：无图片列零扫描；每 Sheet O(shapes + pictures)，索引 O(pictures)；避免重复 byte[] copy，明确数据所有权。
- **风险**：图片像素解码可能引入高成本/漏洞依赖；首版尺寸检查应使用安全元数据读取且可关闭。
- **验收**：图片列真实写入目标属性；Public API 无 NPOI；失败文件图片 anchor 正确；限制可阻止超大输入。

### Phase 8：API 治理与文件拆分

#### P8-001 删除兼容层、internal 化和完整 Public API approval

- **优先级/依赖**：P1；依赖 P2～P7 功能 API 稳定。
- **当前证据**：公开 TypeMap、NPOI Extensions/Helper、legacy Result/Converter/Setting；测试 Adapter 维持旧调用；approval 不含成员。
- **目标**：公共面只保留消费者需要的 Workbook Request、策略、Mapping、Result、DI；删除未发布兼容层。
- **已确认文件**：Public API 测试、Compatibility Adapter、所有公开类型。
- **候选文件**：成员签名 approval 快照、InternalsVisibleTo、Docs Consumer。
- **步骤**：完成第 8.3/10 节清单；迁移所有测试到生产 Workbook API；删除 Adapter；具体 NPOI 实现 internal；增加 API signature generator/approval；扫描公开签名 NPOI；更新 DI 测试和 XML 注释。
- **API/数据/配置影响**：本任务执行最终 Breaking cut；不保留 Obsolete 平行层。
- **测试**：API exact baseline、无 NPOI signature、DI 解析、所有旧测试迁移后行为等价、文档消费者编译。
- **性能**：无直接影响；确保 internal 化不引入反射服务定位。
- **风险**：仓库外消费者不可见；以未发布约束和 P0 消费证据为依据。
- **验收**：Compatibility Adapter 删除；无重复/死 API；所有公开成员有中文 XML 注释；approval 精确到成员。

#### P8-002 按职责拆分 Importer/Exporter 和 NPOI Extensions

- **优先级/依赖**：P1；依赖 P8-001 API 冻结，可与其后半段并行。
- **当前证据**：Exporter 混合 Workbook/Template/Planning/Writing/Style/Merge/Chart；Importer 混合 buffering/selection/header/conversion/validation/duplicate/relation/result；Extensions 暴露大量低层方法。
- **目标**：入口类只做参数验证、生命周期和 orchestration；算法类可直接 internal 单测。
- **已确认文件**：两个 NPOI 大文件、Style Cache、Cell/Sheet Extensions。
- **候选文件**：第 11.2 节 NPOI 目录。
- **步骤**：按 Planning/Reading/Writing/Layout/Validation/Failures/Images 拆分；共享 Workbook/Sheet context；消除反射 `MakeGenericMethod` 可行时使用已编译非泛型 plan；低层 Extensions 改成窄职责 internal services；保持行为测试先绿再移动。
- **API/数据/配置影响**：仅 internal 结构，不扩大公共 API。
- **测试**：每个 internal 组件直接测试；全量回归；文件移动不改变 public approval。
- **性能**：比较拆分前后 Benchmark；不得因抽象引入逐 Cell 虚调用/反射/重复枚举。
- **风险**：大规模移动导致 diff 难审；每次只拆一个职责并独立验证。
- **验收**：入口文件无具体算法和低层 NPOI shape/style 细节；每个核心组件有直接测试；性能无显著回退。

### Phase 9：完整验证、文档和交付

#### P9-001 Unit/Integration/Docs Consumer 与文档收口

- **优先级/依赖**：P0；依赖 P8-002。
- **当前证据**：现有 Unit 偏历史大文件，Integration 仅 9 个测试且多依赖兼容 Adapter；docs 只有简短入口；无 Docs Consumer。
- **目标**：完成第 13～15 节矩阵，编译所有文档 API 示例，更新 ai_docs/docs/README。
- **已确认文件**：现有 Unit/Integration、`docs/excel/README.md`、`ai_docs/excel/*`、根 README。
- **候选文件**：职责级测试、Docs Consumer 项目、API/迁移/性能/interop 文档。
- **步骤**：迁移兼容测试；增加组合夹具；建立 Docs Consumer；文档覆盖默认值、错误语义、Stream ownership、12 项需求、多 Sheet 订单/明细和 ASP.NET Core 上传/错误文件返回；ai_docs 记录 ADR、执行证据、Breaking、Interop、Benchmark。
- **API/数据/配置影响**：文档只描述 approval 后真实 API。
- **测试**：第 19 节全部自动命令；示例项目编译；ASP.NET 示例使用内存 TestServer/等价消费者测试，不启动外部服务。
- **性能**：文档明确 AutoFit/图片/失败 Workbook 成本和限制。
- **风险**：示例漂移；Docs Consumer 必须引用项目而非复制伪签名。
- **验收**：第 12 节需求均可从文档示例追到测试；Build/Unit/Integration/API/Docs Consumer 全绿。

#### P9-002 Benchmark、真实互操作、最终 Diff Review

- **优先级/依赖**：P0；依赖 P9-001。
- **当前证据**：Benchmark 仅 1K、单 Sheet、基础 Import/Export；真实 Office 互操作未验证。
- **目标**：执行第 16 节矩阵，记录 Excel/LibreOffice 样本和最终资源/兼容结论。
- **已确认文件**：Benchmark 项目、历史 xlsx resources。
- **候选文件**：新增 benchmark classes、Interop resources/manifest、人工验证记录。
- **步骤**：扩展参数矩阵；单独 feature benchmarks 避免笛卡尔爆炸；生成 NPOI 样本并加入 Excel/LibreOffice 外部样本；自动重开断言；人工打开检查；`git diff --check`、格式、pack、完整 diff review；确认只含计划执行相关改动。
- **API/数据/配置影响**：根据实测只收紧默认 limit，不在最终阶段新增未评审 API。
- **测试**：Benchmark + Interop + 全量回归；人工项不可执行时标 `NOT_VERIFIABLE` 并写恢复步骤。
- **性能**：记录时间、Allocated、Gen0/1/2、峰值内存、输出大小；对基线显著回退给出根因和接受/修复结论。
- **风险**：100K×10 Sheet×图片/失败文件可能超出 CI；分 smoke/nightly，资源上限测试使用最小越界夹具。
- **验收**：结果写入 ai_docs；所有可自动项通过；人工互操作有应用版本、步骤、样本哈希和通过标准；无自动 commit/push/tag/PR/publish。

## 13. 单元测试矩阵

测试写入 `tests/Bing.Offices.Tests` 的职责目录；每个 xUnit 方法使用英文 `Method_State_Expected`，中文 XML 测试目的，AAA 结构。internal 类型通过 `InternalsVisibleTo` 直接测试，不 Mock 被测内部算法。

| ID | Given / When / Then | 关键断言 |
| --- | --- | --- |
| UT-STYLE-01～04 | base style + Sheet/Region/Column overlay；仅 Body/动态 style；NumberFormat+Wrap | 每属性优先级、CellType/DataFormat、Style key/cache count |
| UT-WIDTH-01～05 | None/Fixed/AutoFit/Adaptive，CJK/Emoji/换行/格式化/采样 | 完整 NPOI width、min/max、采样次数、模板策略 |
| UT-HEADER-01～07 | 多级/merge/模板偏移/冲突/Ordinal vs IgnoreCase | 完整 merged range 和 Header binding，不仅 Contains |
| UT-COMMENT-01～04 | Preserve/Append/Replace/Fail、作者/可见性、限制 | Comment 完整文本/作者/冲突结果/数量 |
| UT-COL-01～04 | fixed/dynamic 同配置、converter/validator/placement/target | 同 Plan 字段、解析一次、Setter 实际目标、冲突失败 |
| UT-SELECT-01～06 | ByName/ByIndex、隐藏、缺失、越界、比较歧义 | 结构错误 code/sheet，不抛未分类异常 |
| UT-RANGE-01～05 | start/count、稀疏、范围外必填、MaxColumnCount、溢出 | 仅范围列进入 binding/conversion/validation |
| UT-NORM-01～08 | Preserve/Trim/RemoveAll + Unicode/typed formula | 精确字符串、原 RawValue、非字符串类型不变、无 regex |
| UT-UNIQUE-01～07 | 默认允许、空、比较器、规范化、动态、首行、回滚 | 首次/重复 RowNumber、O(n) tracker 行为 |
| UT-IGNORE-01～05 | Attribute/Profile/Request override | Header/convert/validate/unique/set 全不执行 |
| UT-ERROR-01～04 | workbook/sheet/row/cell/relation/config error | 完整字段、统计、MaxErrors、配置错误边界 |
| UT-FAIL-01～10 | Annotated/RowsOnly、合并错误、原始行、汇总、无错、写失败 | CellType/style/comment/row de-dup/source fields/stream state |
| UT-WDV-01～09 | list/numeric/date/time/length/direct range/overlap/unsupported/HSSF-XSSF | 规则 source、range lookup、Fail/Report |
| UT-IMAGE-01～09 | anchor/merged/multi/shape/hidden/damage/limits/types | provider-neutral data、multiplicity、无 NPOI 泄露 |
| UT-STREAM-01～06 | non-seekable、pre/mid cancel、LeaveOpen、write failure、size limit | 流保持打开、无无效失败输出、取消传播 |
| UT-API-01～03 | exported type/member baseline、NPOI signature scan、DI | exact approval、无 NPOI、接口可解析 |

Mock 边界：只替代故障/不可寻址 Stream、取消令牌触发器、可选图片元数据读取器和外部时钟（如汇总文件需要时间）；NPOI Workbook、Style、Drawing、Validation 使用真实内存对象。

## 14. 集成与互操作测试矩阵

集成测试写入 `tests/Bing.Offices.Tests.Integration/Excel`，必须重新打开输出 Workbook 并检查真实对象，不允许只断言文件非空。

| ID | 场景 | 必须断言 |
| --- | --- | --- |
| IT-01 | 1/5/10 Sheet 异构导入导出 + name/index/read range | Sheet、CellType、Header/Row/Column 坐标、枚举次数 |
| IT-02 | 订单/明细 `HasMany<TKey>` | comparer、导航集合、relation source row/key |
| IT-03 | 动态租户列 + normalization + unique + error output | dynamic key/type、首次重复行、failure workbook |
| IT-04 | 模板 B3 + 多级 Header + comment + dynamic + chart + width | MergedRegion、Comment、ColumnWidth、Style/DataFormat、Chart range/anchor |
| IT-05 | AnnotatedOriginal 保留公式/style/merge/validation/picture | Formula/CellType、DataFormat、validation、drawing anchor、错误汇总 |
| IT-06 | ErrorRowsOnly 复制原失败行 | 原 CellType/公式/style、唯一失败行、来源列、图片 anchor |
| IT-07 | `.xls/.xlsx` Validation/Image/Comment 支持差异 | 支持项通过，Unsupported 有结构结果，不静默 |
| IT-08 | non-seekable input/failure output/cancel/write failure | 调用方流不关闭、取消、无错误时不写 |
| IT-09 | 多租户并发 Mapping/Style/Converter/Cache | 无列定义、style、converter、unique 状态串扰 |
| IT-10 | ASP.NET Core 上传 + 返回错误 Excel 的消费者示例 | Stream ownership、content type/name、错误 Workbook 可重开 |

真实互操作样本必须分别由 Excel 和 LibreOffice 创建：非 A1 模板、Data Validation、图片、传统批注、公式、合并、图表。每个样本附来源应用/版本、格式、哈希、预期清单；CI 自动重开，人工检查显示和保真。环境未安装应用时标 `NOT_VERIFIABLE`，不得声称通过。

## 15. 性能与资源计划

### 15.1 复杂度目标

- AutoFit：O(rows×columns)，仅显式启用；Adaptive：O(sampleRows×columns)。
- Unique/Relation：O(n)，内存 O(unique values/keys)。
- Workbook Validation：Sheet 预编译 O(rules×ranges)，Cell 查询使用范围索引，禁止 O(cells×rules)。
- Normalization：O(text length)，无 Regex，未变化不分配。
- Drawing：仅图片列启用时 O(shapes+pictures)，每 Sheet 一次。
- Style/Font/DataFormat/Comment Anchor：Workbook 级缓存，数量与唯一组合数相关。
- AnnotatedOriginal：最多一次完整副本/序列化；ErrorRowsOnly 仅复制失败行及关联 drawing。
- 不可寻址输入允许一次受限缓冲；禁止无上限 `ToArray()` 和同一 Workbook 重复序列化。
- 多 Sheet 数据源只枚举一次；请求缓存不可跨租户保存可变定义。

### 15.2 Benchmark 矩阵

| 维度 | 值 |
| --- | --- |
| Rows | 1K / 10K / 100K |
| Sheets | 1 / 5 / 10 |
| Columns | 固定 / 固定+动态混合 |
| Features | baseline / Adaptive / AutoFit（限 1K/10K）/ Comment / Validation / Image / Annotated / RowsOnly |
| Metrics | Mean/P95（外部脚本可选）、Allocated、Gen0/1/2、峰值工作集、输出大小 |

Benchmark 不做所有维度笛卡尔积：基础规模矩阵、单 feature 10K、极限 100K 三组运行。CI 跑 short smoke，完整矩阵由 nightly/人工命令运行并归档 JSON/Markdown。

### 15.3 资源限制

必须评估并给出安全默认值与零/负值语义：`MaxFileSize`、`MaxRows`、`MaxSheets`、`MaxColumnCount`、`MaxDynamicColumns`、`MaxErrors`、`MaxComments`、`MaxPictures`、`MaxPictureBytes`、`MaxTotalPictureBytes`、`MaxFailureWorkbookSize`。限制在分配/解码/逐行处理前尽早检查；命中限制生成稳定结构错误或参数异常（按输入错误/配置错误边界），不得静默截断，只有 MaxErrors 允许“停止收集并附截断标记”。

## 16. 文档计划

### 16.1 ai_docs

更新 `ai_docs/excel` 或新增本任务实施证据文档，包含：API/Column Plan/Failure Writer/Data Validation/Image ADR；每 Task 实施进度；验证命令和产物；Breaking Changes；NPOI 2.7.4 待验证矩阵；Excel/LibreOffice 互操作记录；Benchmark 结果；最终生产符号 → 测试方法追溯表。

### 16.2 docs 和 README

扩展 `docs/excel/README.md` 并按主题拆页：公共 API/default/error；Stream ownership；普通/动态导出；Header/Comment/Width；name/index/read range；comparison/whitespace/unique/ignore；AnnotatedOriginal；ErrorRowsOnly；Workbook Data Validation；图片；订单/明细；ASP.NET Core 上传与失败文件；Breaking migration；性能/资源限制。根 README 只保留准确入口。

建立 `tests/Bing.Offices.Docs.Tests` 或等价编译验证，所有代码示例引用真实项目 API。文档不得声明 Threaded Comment、未验证 Validation 公式或 XLS 高级能力。

## 17. 风险、阻塞和待验证

| ID | 风险/阻塞 | 处理 |
| --- | --- | --- |
| R-01 | 当前工作树未知且历史记录为大量未提交改动 | P0 先记录 Git，逐文件协作，不 reset/checkout/删除陌生内容 |
| R-02 | 三份指定报告/API 文件缺失 | 保持 NOT_VERIFIABLE；如后续出现，按证据优先级对照源码 |
| R-03 | locked restore 历史失败 | 先 locked restore；只在确认项目依赖变化后更新 lock |
| R-04 | NPOI 2.7.4 Validation/Comment/Drawing HSSF-XSSF 差异 | 最小样本 + Unsupported Policy + Interop 记录 |
| R-05 | NPOI 重写 OOXML 可能损失高级部件 | Annotated 保真集成 + Excel/LibreOffice 人工检查；不做未验证承诺 |
| R-06 | 完整 Workbook 缓冲导致峰值内存 | 文件/输出上限、一次缓冲/序列化、Benchmark 峰值工作集 |
| R-07 | 样式/字体数量爆炸 | base+overlay cache key、组合数断言、资源上限 |
| R-08 | API 大幅 Breaking | 未发布策略、Docs Consumer、迁移表、一次性 cut，不保留平行主链 |
| R-09 | netcoreapp3.1/net5 运行时/包支持警告 | 编译全部目标；测试至少 net6/net8，旧 runtime 不可用则记录 NOT_VERIFIABLE |
| R-10 | 100K/10 Sheet 完整矩阵 CI 超时 | smoke/nightly 分层，保留可复现命令和产物 |

待验证：Named Range 列表、跨 Sheet/相对/自定义公式 Data Validation；HSSF comment/image anchor 细节；Excel/LibreOffice 模板/图表/validation/picture 保真；Chart Anchor 在模板下最终相对/绝对约定；外部真实消费者证据；当前 Git/构建/测试状态。

## 18. 验证命令

以下命令均基于仓库真实解决方案和项目路径，在仓库根目录 PowerShell 执行。先设置 UTF-8 控制台编码；命令本身不写中文文件。

```powershell
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

git status --short --branch
dotnet restore .\Bing.Offices.sln --locked-mode
dotnet build .\Bing.Offices.sln -c Release --no-restore

dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore
dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -f net6.0 -c Release --no-restore
dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore
dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -f net6.0 -c Release --no-restore

dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore --filter FullyQualifiedName~PublicApiContractTest

dotnet format .\Bing.Offices.sln --verify-no-changes --no-restore
dotnet pack .\Bing.Offices.sln -c Release --no-build --no-restore
git diff --check
```

Docs Consumer 当前不存在，基线为 `NOT_VERIFIABLE`。P9-001 创建并加入解决方案后执行真实路径：

```powershell
dotnet test .\tests\Bing.Offices.Docs.Tests\Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore
```

现有 Benchmark smoke 命令：

```powershell
dotnet run --project .\benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj -c Release --no-build -- --filter *StreamPipelineBenchmarks* --job short
```

若 locked restore 因确认过的项目依赖变更失败，只能在记录差异后执行一次：

```powershell
dotnet restore .\Bing.Offices.sln --force-evaluate
```

随后必须重新执行 `--locked-mode` 并通过。Excel/LibreOffice 人工互操作没有可编造的 CLI；未安装时标 `NOT_VERIFIABLE` 并保存样本、应用版本和人工检查清单。

## 19. Definition of Done

1. 12 项原始需求全部映射到稳定 Task ID、直接测试和组合测试；`IMP-012` 明确为支撑治理。
2. 每项能力可从公共配置追踪到 immutable Column/Workbook Plan 和 NPOI 生产实现。
3. 固定列和动态列不存在第二套 Reader/Writer/Layout/Validation/Unique/Image 管线。
4. Style 逐属性合成；Body/Dynamic-only 生效；NumberFormat、模板样式和 CellType 不丢。
5. Width、Comment、Header Layout 和模板非 A1 组合通过重开断言。
6. Sheet Selector、Read Range、MaxColumnCount、Header Comparison、Whitespace 策略行为明确且结构错误不逃逸。
7. Unique 为 O(n)，包含首次行、比较/空值/回滚和动态列；Ignore 全链路无副作用。
8. AnnotatedOriginal、ErrorRowsOnly、Workbook Validation 和 Image 有资源上限、取消和 Stream 所有权测试。
9. 多 Sheet、动态列、模板、关系、失败文件、Validation、图片和 Web 消费者有组合集成测试。
10. 公共 API 不暴露 NPOI，无无效、重复或仅兼容未发布版本的 API；具体 NPOI 实现和低层算法 internal。
11. Public API approval 精确到成员签名，所有公开成员有中文 XML 注释。
12. `docs` 示例与真实 API 一致并通过 Docs Consumer 编译；`ai_docs` 有 ADR、实施、验证、Breaking、Interop 和 Benchmark 证据。
13. Release Build、Unit net6/net8、Integration net6/net8、Public API、Docs Consumer、format、pack 和 `git diff --check` 通过。
14. Benchmark 已执行并记录，或因明确环境缺失标 `NOT_VERIFIABLE` 且给出恢复命令；真实 Office 互操作同理。
15. 最终生产符号 → 测试方法追溯映射完整。
16. 未自动 Commit、Push、Tag、PR 或发布，未回退用户已有改动。

## 20. V4 执行与 Review 交接

- Executor 不修改本计划 Checkbox 或文本伪造进度；逐 Task 实施证据、命令、结果、偏差写入 `execution.md`。
- 独立 Reviewer 只验收并写 `review.md`，不修改代码。
- `NEEDS_FIX` 使用结构化 `FIX-xxx`；Fixer 默认只处理 `MUST_FIX` 和必要依赖，不修改 `review.md`；修复后重新 Review。
- 任一 Task 因 NPOI/环境不可验证时使用 `BLOCKED`/`NOT_VERIFIABLE`，不得以文件非空替代行为验收。
- AI 默认不执行 Commit、Push、Tag、PR 或发布。

推荐后续命令：

```text
/execute-plan bing-offices-excel-import-export-enhancement-v2
/review-plan bing-offices-excel-import-export-enhancement-v2

如 Review=NEEDS_FIX：
/fix-review bing-offices-excel-import-export-enhancement-v2
/review-plan bing-offices-excel-import-export-enhancement-v2
```
