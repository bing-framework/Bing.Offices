# MiniExcel RC API、性能与 Entity 能力收口计划

## 任务信息

- Task ID：`BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001`
- 状态：`PLANNED`
- 计划基线：分支 `feat/miniexcel-provider`，提交 `94bb52e84ffc70634067b04541857433ef7af9df`
- 任务类型：RC 架构收敛、Breaking API、Provider SPI、性能/内存治理、Entity/List/Template 能力、测试与发布收口
- 执行边界：本文件只定义实施顺序和门禁，不实施源码，不提交、推送、打标签、创建 PR 或发布包。

## 1. 规划依据与当前结论

### 1.1 已核对证据

- 规范：`AGENTS.md`、`.editorconfig`、`.gitattributes`。
- 项目：`Bing.Offices.sln`、`common.props`、`framework.props`、`version.props`、`version.dev.props`、各生产/测试/Benchmark/Consumer 项目。
- 历史任务：
  - `BING-OFFICES-MINIEXCEL-PROVIDER-20260915-001`
  - `BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
  - `BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001`
- 当前调用链：Abstractions 请求与 builder、Core mapping plan、NPOI/MiniExcel importer/exporter、关系绑定、RawDate 补偿、文件提交与取消。
- 当前测试与工具：六个 Provider 测试项目、Docs、Consumer、ResourceProbe、BenchmarkDotNet、API snapshot 和 CI。

历史报告仅作为线索。所有 PASS、性能数字和包消费结果必须由最终候选重新生成，不继承历史结论。

### 1.2 不重复实现的能力

- NPOI 与 MiniExcel Provider 已存在，且已接入 DI；本任务不重写 Provider。
- Abstractions/Core 保持 `netstandard2.0`；Provider 保持 `net6.0`/`net8.0`。
- workbook/list 型请求、映射、converter、validation、value map、错误聚合、取消、临时文件和原子提交已有基础设施，应复用。
- NPOI 关系绑定已采用强类型缓存调用和父键索引，复杂度已接近 `O(P+C)`；先冻结合同，再抽取共享引擎。
- 生产程序集之间当前无 `InternalsVisibleTo`；现有 IVT 仅面向测试项目，允许保留。
- 测试已按 Common/NPOI/MiniExcel、Unit/Integration 拆分，CI 和根 `artifacts/` 体系已有基础。

### 1.3 当前缺口分类

| 分类 | 内容 |
|---|---|
| Breaking cleanup | `ValidateMode` 重命名；清理死 overload；删除未使用参数；统一公开友好 API；关系 selector 改为纯函数合同 |
| 行为变化 | Dynamic 最多一个且计划期失败；MiniExcel 集合合同从 `IList` 对齐 `ICollection`；关系 first-parent-wins/重复键诊断；merge anchor/conflict 规则 |
| 性能治理 | Sheet-level binding；RawDate 过滤；去逐行 Dictionary clone、Activator、反射和 DynamicInvoke；MiniExcel importer 拆责 |
| 新功能 | List/Entity/Template/Workbook 同步异步友好 API；Entity 固定 Cell、Merge、List Region、多 Sheet；正式 Provider SPI 与能力声明 |
| 发布治理 | 当前候选 API snapshot、包 Consumer、Benchmark、100K/500K/1M、2CPU/4GiB、同候选证据链 |

### 1.4 版本与 API 基线硬门禁

执行前后都记录 SHA-256；下列文件不得由本任务修改：

| 文件 | 当前 SHA-256 |
|---|---|
| `version.props` | `EC5201C9A9EC7F4981879FFA604C5BA2B35F89F064AB88C790ECFB6644DC5306` |
| `version.dev.props` | `0A2F75F575DCC863D51F256A0D823F400FC33369FB1C7EB823CE6EBA9321872B` |
| `common.props` | `1C830F653129C28F59094F3F5886EB1BA1640ACBFE8984FF34F24FEDAB10ECA4` |
| `framework.props` | `5B64D45A8FC516449251449C619E8F0434481A295D18A8B727AC4978FF5FDE35` |

`build/api-snapshot-baseline.json` 当前哈希为 `951F9944C39B843CFDFE8DF92BA7C4592F4A9CAB2F459540E4FA6B7EF1F4AAF0`，但其 candidate identity 早于当前 HEAD。它只能作为旧基线输入；Breaking API 经成员级审批后才能更新 approved baseline。

## 2. 目标架构与依赖顺序

```text
用户友好 API
  Import/Export List | Entity | Template | Workbook（同步/异步）
                         ↓
单一 Request/Result + OperationKind + Plan 构建
                         ↓
Provider SPI（执行合同 + 能力声明，无运行时路由）
                         ↓
一次 Provider 分派
  ├─ List：Provider-native fast path
  ├─ Entity：layout plan + Provider-native executor
  ├─ Template：template/layout plan + preflight
  └─ Workbook：现有高级组合能力
```

原则：不创建四套 engine/interface；高级布局对象不得进入 List fast path；保持“第一个注册的 Provider 生效”，不增加 provider key、router、factory 或运行时切换。

依赖顺序：

```text
Phase 0
  → Phase 1 API 冻结
  → Phase 2 SPI
  → Phase 3 MiniExcel 拆责
  → Phase 4 Dynamic plan
  → Phase 5 List/RawDate/materialization
  → Phase 6 Relation
  → Phase 7 Entity
  → Phase 8 Template
  → Phase 9 合同/专项测试
  → Phase 10 Benchmark 决策
  → Phase 11 容量与资源门禁
  → Phase 12 文档、包、API、RC 审查
```

Phase 4、5、6 在 Phase 3 后可分支实现，但合并前必须通过共同的 List 合同测试。Phase 7 依赖 API/SPI 冻结；Phase 8 依赖 Entity layout 和 merge preflight。

## 3. 分阶段执行计划

### Phase 0：冻结当前候选与可信基线

#### Task `P0-BASELINE`

- 目标：形成可复现的 before candidate，防止把历史报告、脏工作区或不同二进制混入性能和发布判断。
- 证据：当前 HEAD/分支、clean status、SDK/runtime、CPU/内存、包锁定、版本文件哈希、API baseline identity。
- 当前调用链：不改调用链；只捕获当前公开 API、测试、Benchmark、资源证据。
- 已确认文件：`Bing.Offices.sln`、`version*.props`、`build/api-snapshot-baseline.json`、`.github/workflows/ci.yml`、`benchmarks/`、`tests/`。
- 候选文件：本任务目录下 `baseline.md`；`artifacts/baseline/<candidate-id>/`。
- 公共 API：无变化。
- Breaking Change：无。
- 实施步骤：定义 candidate id 为 `commit + dirty fingerprint + TFM + RID + configuration`；clean HEAD 先跑 restore/build/六项目测试/API capture/现有基准；保存命令、退出码、原始日志和文件哈希；baseline 产物只读归档，不被 after 覆盖。
- 测试：当前六个职责级测试项目分别跑 `net6.0`/`net8.0`；当前失败也原样记录。
- Benchmark：保存现有 A-I 可覆盖项的 before；缺失项标 `NOT_AVAILABLE_BEFORE`，不得捏造提升比例。
- 风险：首次完整 baseline 成本高；历史环境无法复现。
- 回滚条件：工作区不干净、candidate identity 缺失、before/after 环境不同则停止性能比较并重新捕获。
- 验收：`baseline.md` 能从原始日志追到同一候选；版本文件哈希与上表一致。

### Phase 1：API Breaking Cleanup 与公开形状冻结

#### Task `P1-API-CLEANUP`

- 目标：在 RC 前一次完成命名、死代码、统一请求/结果与友好入口的 Breaking 收敛。
- 证据：`Imports/ValidateMode.cs`；`ExcelImport.cs`；import request；两个 Provider 的调用点；现有 `IExcelImporter`/`IExcelExporter`；API snapshot。
- 当前调用链：builder → workbook request → provider importer/exporter；公开 API 目前以 workbook/list 组合为主，没有 Entity/Template 一等入口。
- 已确认文件：`src/Bing.Offices.Abstractions/Bing/Offices/Imports/ValidateMode.cs`、`ExcelImport.cs`、`ExcelWorkbookImportRequest.cs`、`IExcelImporter.cs`、`IExcelExporter.cs`；NPOI/MiniExcel importer；相关测试/docs/API snapshot。
- 候选文件：`ExcelValidationFailureMode.cs`；通用 import/export request/result、operation kind、友好扩展类；迁移记录。
- 公共 API：
  - `ValidateMode` → `ExcelValidationFailureMode`；builder/property 使用 `ValidationFailureMode`；保留独立的 `ExcelImportValidationMode`。
  - 新增 `ImportList/ExportList`、`ImportEntity/ExportEntity`、`ImportForTemplate/ExportForTemplate`、`ImportWorkbook/ExportWorkbook` 及 Async。
  - 目标方向为一个泛型核心请求/结果执行合同，友好方法只组装请求；最终签名在本 Phase 的 API design record 中冻结。
- Breaking Change：不保留 obsolete wrapper；关系 selector 纯函数约束在 Phase 6 落地；所有删除均进入成员级迁移表。
- 实施步骤：先列出全部符号/调用点；验证 `IsErrorLimitReached(ICollection<ExcelImportError>, object)` 的真实 overload resolution 后删除死 overload；删除 `ValidateHeaderCount` 未使用的 `workbookRequest`；设计 `OperationKind` 和 base/derived request，使 List 请求不携带 entity/template layout；公开 `ExcelImportResult<TResult>` 或等价统一结果；更新两个 Provider 编译面；生成 API diff。
- 测试：每个入口同步/异步、stream/file、成功/失败/取消；rename 编译测试；死 overload 调用点为零；完整错误对象断言。
- Benchmark：只做 API shim 零额外层级的 smoke；正式 fast path 在 Phase 10。
- 风险：过度抽象形成四套 engine；泛型结果破坏现有错误合同；误删实际被 overload 绑定的方法。
- 回滚条件：友好 API 需要四次 Provider dispatch、List 依赖 layout、或无法保持现有 workbook 行为时退回 design review，不进入 SPI。
- 验收：API 评审记录批准；所有 Breaking 成员具备 old→new 迁移项；无兼容 wrapper；双 TFM 编译。

### Phase 2：正式第三方 Provider SPI

#### Task `P2-PROVIDER-SPI`

- 目标：只依靠公开契约实现第三方 Provider，消除对内部类型、生产 IVT 和反射 hack 的需求。
- 证据：现有 `IExcelMappingPlan*`、`ExcelMappingPlanFactoryProvider`、异常/唯一性/文件提交服务、DI 注册和 production IVT 搜索结果。
- 当前调用链：Provider 注册默认 mapping factory → importer/exporter 执行；`IExcelImporter/IExcelExporter` 同时承担用户面和执行面，能力声明缺失。
- 已确认文件：Abstractions/Core 的 Plans、Imports、Exports；两个 Provider 注册扩展；各 `.csproj`/`AssemblyInfo`。
- 候选文件：`IExcelProviderCapabilities`、最小执行上下文/descriptor；`tests/Bing.Offices.ThirdPartyProvider.Tests/` 或独立 fixture 项目；CI/solution。
- 公共 API：最小 SPI 包含 importer/exporter 执行合同、稳定 mapping metadata、错误/取消/流所有权/提交合同、只读 capability；compiled accessor 仅在 public-only fixture 证明必需时引入，并标记高级基础设施属性。
- Breaking Change：现有 provider interface 若重塑为统一执行合同属于批准范围；禁止生产 IVT。
- 实施步骤：划分 User API/SPI/Internal；能力维度至少含 List、Workbook、Entity、Template、Merge、Async、XLS/XLSX；能力只用于 preflight/诊断，不做路由；第三方 fixture 只引用已打包公开程序集；检查依赖方向；保持 first-registration-wins。
- 测试：fixture 编译并执行 List sync/async；unsupported 产生结构化异常且无部分输出；扫描 production→production IVT、非公开反射和 keyed/router 类型为零。
- Benchmark：SPI 包装对 List 分派只允许常数级一次分支，纳入 A。
- 风险：为了共享编译 accessor 过早扩张 API；capability 被误用成运行时 provider selector。
- 回滚条件：SPI 需要暴露 Provider 私有补偿、Core internal implementation 或多 Provider router 时缩减契约重新评审。
- 验收：第三方 fixture 不使用 IVT/反射，公开包即可编译运行；NPOI/MiniExcel 注册语义不变。

### Phase 3：MiniExcel Importer 结构拆分

#### Task `P3-MINIEXCEL-DECOMPOSE`

- 目标：在性能改写前拆开读取、header plan、materialization、dynamic、relation、错误/取消职责，保持行为等价。
- 证据：`MiniExcelExcelImporter.cs` 同时包含 query、dictionary clone、Activator、reflection、dynamic、relation、RawDate 协调。
- 当前调用链：Import sheet → RawDate 全表索引 → MiniExcel Query → row dictionary clone → reflection materialize → dynamic/relation。
- 已确认文件：`src/Bing.Offices.MiniExcel/**/MiniExcelExcelImporter.cs`、`MiniExcelRawDateSerialReader.cs`、MiniExcel tests。
- 候选文件：同项目内 `MiniExcelSheetReader`、`MiniExcelSheetPlanBuilder`、`MiniExcelRowMaterializer`、`MiniExcelRelationCoordinator` 等内部类。
- 公共 API：无变化。
- Breaking Change：无。
- 实施步骤：以职责提取而非重写；保留错误顺序、RowIndex、取消点、流所有权；用 characterization tests 锁定；拆分后再允许 Phase 4-6 改算法。
- 测试：原有 MiniExcel 测试零行为漂移；新增每个内部职责的直接测试；异常类型和完整错误集合一致。
- Benchmark：运行 before/after 等价基准，拆分不得造成显著回退。
- 风险：提取时改变延迟枚举、stream 生命周期或错误顺序。
- 回滚条件：任何 characterization 差异且无法说明为已批准行为变化时回滚该提取。
- 验收：主 importer 只负责 orchestration；热点职责可独立测试；双 TFM 通过。

### Phase 4：Dynamic Column Sheet-level Plan

#### Task `P4-DYNAMIC-PLAN`

- 目标：每个实体最多一个 `[DynamicColumn]`，在 plan build 失败；header/physical column/binding 每 Sheet 构建一次。
- 证据：Core factory 当前允许多个 dynamic 属性；NPOI 执行时才拒绝；MiniExcel 每行 FindHeader/FindPhysicalColumnIndex/GetValue/SetValue。
- 当前调用链：type map → fixed bindings；每行再检测未知 header 并反射访问 dynamic dictionary。
- 已确认文件：Core `ExcelTypeMapFactory`/mapping plan；`DynamicColumnAttribute`；NPOI column plan；MiniExcel importer/exporter；动态列测试。
- 候选文件：共享只读 `ExcelDynamicColumnPlan`/sheet binding descriptor；Provider-specific physical binding。
- 公共 API：除非 SPI fixture 需要，不公开运行时字典或 setter；属性规则写入公开合同/XML docs。
- Breaking Change：多个 `[DynamicColumn]` 从执行期/不一致行为改为 plan-build 确定性失败。
- 实施步骤：Core 建图时计数；标准化 header 一次；构建 known/unknown physical column map；无 dynamic 属性直接走零分配分支；有 dynamic 时按 unknown count 预分配字典；alias/whitespace/case/read range 在计划期解析。
- 测试：0/1/2 dynamic 属性；0/5/20/100 extra；alias、空白、大小写、unknown、ReadColumnRange、重复 header；两个 Provider 独立成功/失败。
- Benchmark：D 矩阵；普通实体必须无 dynamic allocation，吞吐不因功能存在回退。
- 风险：header normalizer 与现有 converter context 列号不一致。
- 回滚条件：RowIndex/ColumnIndex、unknown-header 行为或普通实体分配恶化则不合并。
- 验收：每行不做 header 匹配；多属性在 provider IO 前失败；计划缓存隔离/命中/失败状态有测试。

### Phase 5：List Fast Path、Materialization 与 RawDate

#### Task `P5-LIST-FAST-PATH`

- 目标：保留 Provider-native List 最快路径，消除 MiniExcel 每行 clone/Activator/反射，并把 RawDate 索引限定到必要坐标。
- 证据：MiniExcel 每行 `ToDictionary`、`Activator.CreateInstance`、`PropertyInfo.SetValue/GetValue`；RawDate 对每 Sheet 扫描并缓存所有 numeric cell，实际仅日期转换读取。
- 当前调用链：见 Phase 3；NPOI 已有 native cell path，但 column getter/setter仍需测量后优化。
- 已确认文件：MiniExcel importer/exporter、`MiniExcelRawDateSerialReader`；Core compiled mapping；NPOI column plan；List tests/benchmarks。
- 候选文件：Provider-specific compiled constructor/accessor/binding cache；按列/行范围过滤的 RawDate index。
- 公共 API：无新增；List 请求必须不构建 entity/template/merge 对象。
- Breaking Change：MiniExcel sheet target 从内部 `IList` 假设修正为公开承诺的 `ICollection<T>`，属于 bug fix/行为对齐。
- 实施步骤：直接按 source key 读取原 row dictionary，不 clone；一次编译 `new TItem()` 或 constructor delegate；一次编译 getter/setter；bindings 存 physical column/source key；先完成 header/column plan，再判断实际日期列和行范围；无日期列完全跳过 RawDate；只索引日期列×有效行。是否用 compact array/segment 必须由基准决定。避免为获取 header 引入双全量 Query；比较“同一枚举器继续读取”和“轻量 header prepass”后选取。
- 测试：不同 dictionary comparer、缺失/null key、value map、converter/validation context、日期显式 converter、1900/1904、dynamic 日期、read range/max rows、无日期 skip、取消、stream ownership、`ICollection` 非 `IList`。
- Benchmark：A、B、C、E、F；记录 wall time、throughput、allocated bytes、Gen0/1/2、peak working set、RawDate index count/bytes。
- 风险：MiniExcel 返回字典生命周期不稳定；compiled expression 在 AOT/TFM 下差异；日期 header 发现导致重复扫描。
- 回滚条件：正确性变化、List P95/allocated/peak 任一关键指标回退超过批准阈值，或优化只在 microbenchmark 有效而 E2E 无收益。
- 验收：非动态、非日期 List 无额外高级能力分配；每行无 dictionary clone/Activator/PropertyInfo；无日期时 RawDate 调用计数为零。

### Phase 6：Relation 合同与共享 O(P+C) 引擎

#### Task `P6-RELATION-ENGINE`

- 目标：将 selector 明确定义为纯、确定性函数，统一两个 Provider 的 `O(P+C)` 关系绑定。
- 证据：NPOI 已使用强类型 cached invoker + parent dictionary；MiniExcel 使用 `DynamicInvoke` + `FirstOrDefault`，为 `O(P*C)`；现有注释暗含 selector 副作用语义。
- 当前调用链：sheet materialize → relation descriptor → provider-specific binder → parent navigation collection。
- 已确认文件：NPOI `NpoiRelationBinder`；MiniExcel importer relation block；Abstractions relation builder/descriptor；relation tests。
- 候选文件：Core provider-neutral relation plan/engine；强类型 comparer adapter/cache。
- 公共 API：XML docs 明确 selector 必须纯且确定；自定义 comparer、null/missing/duplicate 合同明确。
- Breaking Change：不再保证 selector 每次/每子项被调用的副作用；这是有意 Breaking 行为。
- 实施步骤：先用合同测试冻结 NPOI 语义；父项单次建索引，`TryAdd` 保留 First Parent Wins；重复键同时产生既定结构化错误；子项单次 lookup；复用强类型委托，不在 hot loop DynamicInvoke；共享引擎仅依赖对象集合/强类型 descriptor，不依赖 NPOI/MiniExcel。
- 测试：空集合、null key、missing parent、duplicate parent、custom comparer、非 `IList` ICollection、selector 调用次数、错误上限/顺序、取消。
- Benchmark：G，P/C 比例矩阵和两个 Provider/共享引擎比较。
- 风险：抽取时削弱 NPOI 已有性能或改变重复键错误。
- 回滚条件：NPOI 基准或合同回退；共享抽象需要 Provider DOM。
- 验收：两 Provider 复杂度和结果一致；hot loop 无 `DynamicInvoke`/线性父项搜索。

### Phase 7：Entity 一等能力

#### Task `P7-ENTITY`

- 目标：支持单个业务聚合对象的固定 Cell、Merge、List/Table Region、多 Sheet 导入导出。
- 证据：当前 workbook builder 只把 sheet 绑定到 `ICollection<TItem>`；NPOI 有底层 cell/merge 能力但无 Provider-neutral entity layout；MiniExcel 不支持 offset/template layout。
- 当前调用链：用户只能把对象拆成列表/sheet request；无 entity layout plan/executor。
- 已确认文件：Abstractions Imports/Exports/builders/plans；Core mapping/converter/validation；NPOI workbook/cell extensions；Provider capabilities。
- 候选文件：`ExcelEntityLayoutBuilder<TEntity>`、cell/range value objects、fixed-cell/list-region/merge descriptors、layout validator；NPOI entity executor；Entity tests/fixtures。
- 公共 API：`ImportEntity[Async]`、`ExportEntity[Async]`；typed builder 绑定属性到 `Sheet!A1`、merged range、list region；复用统一 result/error。
- Breaking Change：无旧 API 删除；布局冲突在 provider IO 前失败。
- 实施步骤：定义 immutable layout；统一 A1/range parser；固定值、converter、validation、value map 复用既有列合同；merge 值仅写 anchor，读取 merge 内任意 cell 均解析 anchor；构建 sheet-level merge interval/index；检测 cell-cell、cell-region、region-region、merge overlap；NPOI 先完整实现；MiniExcel 按 capability preflight，要么原生高效完整支持，要么明确 unsupported 且零输出。
- 测试：单/多 sheet、fixed cell、merged cell 非 anchor 读取、list region 起止、空/越界、冲突、转换/验证、错误坐标、取消、已有目标保护、XLS/HSSF 和 XLSX/XSSF 成功/失败；MiniExcel 支持或 fail-fast 两套确定性断言。
- Benchmark：H，Entity 固定字段数×明细行数；merge index 查找成本。
- 风险：把布局引擎做成第二套 mapping/converter；merge 全表线性搜索。
- 回滚条件：无法复用现有错误/转换合同，或 NPOI 每 cell 扫描 merge list。
- 验收：NPOI 功能完整；MiniExcel 无伪支持/部分文件；List fast path 不引用 entity layout。

### Phase 8：Template 一等能力

#### Task `P8-TEMPLATE`

- 目标：把“基于既有模板填充聚合对象”作为独立用户语义，同时复用 Entity layout 与 Provider 执行基础。
- 证据：当前 `UseTemplate`/`UseTemplateRegion` 面向 tabular origin；MiniExcel 明确拒绝 template/styles/offset。
- 当前调用链：workbook export request 携带 template → Provider；无 ImportForTemplate/ExportForTemplate 和固定区域合同。
- 已确认文件：现有 template request/options；NPOI template pipeline；MiniExcel unsupported preflight；文件提交测试。
- 候选文件：template request/layout adapter、template merge validator、Template fixtures/tests。
- 公共 API：`ImportForTemplate[Async]`、`ExportForTemplate[Async]`，命名与 Entity 区分但不复制 engine。
- Breaking Change：existing merge mismatch 从隐式覆盖/不确定行为改为 preflight fail-fast。
- 实施步骤：读取模板结构生成只读 sheet/merge index；完全相同 merge 复用，不重复添加；交叉、包含但不相等、anchor 不一致均报结构化冲突；值写 anchor；保留未映射样式/公式/图片/logo；import 使用同一 fixed/list region layout；NPOI 完整实现；MiniExcel 仅在当前版本原生且高效时实现，否则 capability unsupported。
- 测试：existing exact merge、mismatch/overlap、跨 sheet、公式/样式/图片保留、模板缺失、取消、原子提交/临时文件清理、目标文件不损坏。
- Benchmark：I，模板装载、merge preflight、填充和提交分段指标。
- 风险：对模板做全量复制造成高峰内存；失败后留下部分文件。
- 回滚条件：不能在写出前识别冲突，或未映射内容无法稳定保留。
- 验收：所有冲突在 commit 前确定；模板原件不变；失败零部分输出。

### Phase 9：测试架构与合同矩阵补齐

#### Task `P9-TEST-MATRIX`

- 目标：为每个受影响默认实现建立职责级直接测试，并拆分巨型测试文件。
- 证据：六项目结构已存在，但 MiniExcel importer 测试职责过密；第三方 public-only fixture、Entity/Template、new API 和完整性能合同缺失。
- 当前调用链：现有 unit/integration → provider implementation；部分行为仅综合测试覆盖。
- 已确认文件：`tests/Bing.Offices.Tests*`、`Bing.Offices.Npoi.Tests*`、`Bing.Offices.MiniExcel.Tests*`、consumer/resource projects。
- 候选文件：按 API、binding、dynamic、raw-date、materialization、relation、entity、template、IO/commit 拆分测试类；第三方 fixture。
- 公共 API：测试公开入口，不依赖内部便利路径；必要的内部职责测试继续使用仅面向 tests 的 IVT。
- Breaking Change：测试迁移不改变 API。
- 实施步骤：先建最终生产符号→测试方法 map；巨型文件按责任移动，不复制；每个默认实现单独测；CSV/Excel 输出断言完整文件或结构化内容；NPOI HSSF/XSSF 成功失败；未知实现异常类型固化；Async 使用真实异步 IO，扫描禁止 `Task.Run/.Result/.Wait()`。
- 测试：六项目双 TFM；new API 全组合；cache 隔离/命中/未命中/重复渲染/失败状态；取消、流所有权、资源限制、临时文件、提交。
- Benchmark：无新增算法；benchmark smoke 作为测试入口验证。
- 风险：只拆文件导致覆盖率假提升；测试使用 internal 绕开公开 SPI。
- 回滚条件：symbol-test map 存在无直接测试的受影响生产符号。
- 验收：`symbol-test-map.md` 无空项；provider matrix 完整；public-only fixture 通过。

### Phase 10：Benchmark A-I 与优化保留决策

#### Task `P10-BENCHMARK`

- 目标：用同机、同 runtime、同 workload 的 before/after 数据决定优化是否保留。
- 证据：现有 BenchmarkDotNet 覆盖 provider comparison、部分 RawDate/relation/accessor，但缺少完整矩阵和 E2E 归因。
- 当前调用链：benchmark 项目直接跑 Provider 或反射 private 方法，部分不能代表用户路径。
- 已确认文件：`benchmarks/Bing.Offices.Benchmarks/**`、Benchmark 输出配置、CI probe。
- 候选文件：A-I benchmark classes、dataset generator、candidate manifest、report generator。
- 公共 API：Benchmark 必须优先走公开友好 API；内部 microbenchmark 只作归因，不能替代 E2E。
- Breaking Change：无。
- 实施步骤：A List E2E；B Provider 比较；C RawDate 稀疏/密集和范围；D Dynamic 0/5/20/100；E materialization clone/constructor/accessor；F read/export binding；G relation；H Entity；I Template。固定 seed、列数、行数、TFM/runtime；保存 BDN JSON/CSV/markdown 和环境信息。
- 测试：benchmark smoke 1 iteration；完整 run 独立执行；报告校验 workload/candidate identity 一致。
- Benchmark：A-I 完整矩阵；指标包括 Mean/P95（可测时）、ops/s、MB/s、Allocated、Gen0/1/2、peak working set、GC pause、RawDate index count/bytes、文件大小；以 baseline 中预先批准阈值判定。
- 风险：仅用 microbenchmark 宣称 E2E 改善；before 二进制与 after 源码不一致。
- 回滚条件：Dictionary clone、constructor、compiled accessor、compact RawDate、共享 relation 任一改动未显示 E2E 正收益或关键指标回退，即回滚该优化而非调整阈值。
- 验收：A-I 均有结果或有技术性 `NOT_APPLICABLE`；before/after manifest 可验证；无覆盖旧产物。

### Phase 11：100K/500K/1M 与受限资源验收

#### Task `P11-RESOURCE-GATES`

- 目标：验证大数据、取消、临时文件和 2CPU/4GiB 下的真实资源行为。
- 证据：现有 ResourceProbe/CI 主要完成 100K；500K/1M、生产机器和外部 CI 历史上未验证。
- 当前调用链：ResourceProbe → public API → Provider → stream/temp/commit。
- 已确认文件：ResourceProbe 项目、资源脚本/CI、文件提交实现、`artifacts/`。
- 候选文件：List/Entity/Template 场景矩阵、容器/作业资源限制脚本、resource report。
- 公共 API：只走公开 API。
- Breaking Change：无。
- 实施步骤：先 100K 两 Provider/双 TFM；再 500K、1M；独立运行避免互相污染；2CPU/4GiB 必须由可验证的 OS/container limit 证明，不以本机总规格代替；记录 peak memory、elapsed、GC、输出完整性、取消后目标/临时文件状态。
- 测试：import/export、日期稀疏/密集、dynamic、relation；Entity 大明细；中途取消；磁盘不足/目标存在/提交失败。
- Benchmark：容量测试不是 BDN 回归替代，但复用相同 dataset identity。
- 风险：机器换页导致“完成”但不满足内存门槛；1M XLS 超格式上限。
- 回滚条件：OOM、目标损坏、临时文件泄漏、非线性异常增长；XLS 场景按格式极限明确 `NOT_APPLICABLE`，不得强跑。
- 验收：100K 必须 PASS；500K/1M 和 2CPU/4GiB 必须有真实证据才 PASS，否则明确 `NOT_VERIFIED`，不得标 `COMPLETED`。

### Phase 12：文档、API、Package Consumer 与最终 RC 审查

#### Task `P12-RELEASE-CLOSEOUT`

- 目标：把代码、API、测试、包、Consumer、Benchmark、资源与文档绑定到同一最终候选。
- 证据：`ai_docs/excel/implementation-progress.md` 内容已过时；旧 snapshot identity 过时；Consumer 仅覆盖基础 workbook；历史报告混合候选。
- 当前调用链：build → test → API capture/compare → pack → local-source consumer → benchmark/resource → reports/review。
- 已确认文件：README、`docs/`/`ai_docs/excel/`、API scripts/baseline、consumer projects、CI、package props。
- 候选文件：将旧 `implementation-progress.md` 重命名为 `implementation-history.md` 并加 Historical 标识；当前状态/迁移/Provider SPI/Entity/Template 文档；consumer samples；CI；任务报告。
- 公共 API：双 TFM snapshot；每个新增/删除/重命名成员有审批和迁移记录；修改范围内用 `chinese-comments` 技能治理中文 XML 注释。
- Breaking Change：只有审批清单中的 Breaking 可进入最终 snapshot；不得删除未批准成员。
- 实施步骤：clean checkout 生成 candidate manifest；一次 build 后封存 binaries；测试、pack、consumer、API、benchmark/resource 均引用该 manifest/包哈希；Consumer 从 `.nupkg/.nuspec` 发现版本，不改版本属性；检查根外/nested artifacts；独立 review 对照 plan/execution/diff/raw logs。
- 测试：双 TFM build/test；Docs；PackageReference-only net6/net8 consumers 覆盖四类 API 和第三方 fixture；XML doc/build warnings。
- Benchmark：引用 Phase 10 最终报告；不重用不同 SHA 的结果。
- 风险：pack 后源码又变化；API baseline 未审批即覆盖；历史进度文档误导用户。
- 回滚条件：任一 artifact candidate id/包哈希不一致、版本文件变化、API 未审批或根外产物存在。
- 验收：所有门禁 PASS 才 `COMPLETED`；生产机器/外部 CI/500K/1M 未执行则 `NOT_VERIFIED` 并保持 `PARTIAL`。

## 4. 具体 API/SPI 设计约束

1. 友好 API 是薄层：构建 typed request 后只调用一次 Provider core execution。
2. List/Entity/Template/Workbook 是 operation shape，不是四套 service hierarchy。
3. List fast path 不分配 layout、merge、template、capability matrix 等高级对象。
4. Provider capabilities 是只读 preflight 信息；不选择 Provider，不改变 first-registration-wins。
5. `ExcelImportResult<TResult>`（或评审通过的等价类型）统一承载结果、errors、sheet diagnostics；不得为 Entity 复制错误体系。
6. Entity/Template layout 必须 immutable、可缓存、cache key 包含类型和所有影响渲染/读取的配置。
7. Provider-specific compensation 保留在 Provider：RawDate 只属于 MiniExcel；NPOI DOM/merge index 只属于 NPOI。
8. MiniExcel unsupported 必须在创建部分输出前失败，错误包含 operation/capability/provider。

## 5. 验证命令与产物布局

执行阶段按当前仓库实际项目名补全参数，但不得改变以下语义：

```powershell
dotnet restore Bing.Offices.sln
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test <each-of-six-test-projects> -c Release -f net6.0 --no-build --results-directory artifacts/tests/net6 --logger "trx;LogFileName=<project>.trx"
dotnet test <each-of-six-test-projects> -c Release -f net8.0 --no-build --results-directory artifacts/tests/net8 --logger "trx;LogFileName=<project>.trx"
dotnet test <docs-project> -c Release -f net8.0 --no-build --results-directory artifacts/tests/docs
dotnet pack <each-production-project> -c Release --no-build -o artifacts/packages
dotnet run --project <api-snapshot-tool> -- capture --output artifacts/api/current
dotnet run --project <api-snapshot-tool> -- compare --baseline build/api-snapshot-baseline.json --candidate artifacts/api/current
dotnet run --project benchmarks/Bing.Offices.Benchmarks -c Release -f net8.0 -- --artifacts artifacts/benchmarks/<candidate-id>
git diff --check
git diff --stat
git ls-files --eol
```

Consumer 使用仅指向 `artifacts/packages` 的临时 NuGet source，通过 nuspec 读取实际版本；不得修改 `version.props`/`version.dev.props`。所有日志、trx、coverage、API、packages、consumer、benchmark、resource、temp-audit 只能位于根 `artifacts/`。

任务目录最终报告：

- `baseline.md`
- `execution.md`
- `unit-test-report.md`
- `integration-npoi.md`
- `integration-miniexcel.md`
- `provider-test-matrix.md`
- `symbol-test-map.md`
- `api-diff.md`
- `package-consumer-report.md`
- `benchmark-report.md`
- `resource-report.md`
- `final-report.md`
- `review.md`

## 6. 20 个必须回答的问题

1. **哪些已经完成？** 两个 Provider、基础 workbook/list 流程、mapping/converter/validation/error、取消与文件提交基础、六项目测试分层、NPOI 的索引式 relation、根 artifacts/CI 基础已存在；只做契约补齐和回归，不重写。
2. **纯 API Breaking Cleanup？** `ValidateMode`/builder property 重命名、无 wrapper；删除确认无调用的死 overload；删除 `ValidateHeaderCount` 无用参数；统一四类友好命名和 provider core contract。
3. **行为变化？** dynamic 最多一个且计划期失败；relation selector 纯函数、First Parent Wins/重复键诊断；MiniExcel 支持任意 `ICollection`；merge anchor/冲突 fail-fast；unsupported 零部分输出。
4. **性能优化？** Sheet-level binding、RawDate 筛选、去 row clone/Activator/反射/DynamicInvoke、字典预分配、共享 `O(P+C)` relation、NPOI merge index。
5. **新功能？** List/Entity/Template/Workbook 友好 API、Entity 固定 Cell/Merge/List Region/多 Sheet、Template import/export、Provider capability 与 public-only fixture。
6. **第三方 Provider 最少 Contract？** 统一 import/export core execution、公开 request/result、mapping metadata、错误/取消/流所有权/提交语义、能力声明；compiled accessor 仅经 fixture 证明必要后公开。
7. **允许 Unsupported？** Entity/Template/Merge、XLS、Provider 不原生支持的高级布局可声明 unsupported；基础 List、其宣称支持的格式、取消和错误合同不可静默降级。
8. **List Fast Path 如何不被拖慢？** List request 不含 layout；一次 operation 分支后直接 Provider native reader/writer；无 dynamic/date 时对应对象和 RawDate 零分配/零调用，由 A/D 基准守门。
9. **Dynamic 如何改成 Sheet plan？** Core 限制一个属性；header 读取后一次标准化并建立 physical→fixed/dynamic binding；每行只按数组/已解析 key 写值，未知列字典按数量预分配。
10. **RawDate 如何降内存？** header plan 得到真实日期物理列，再叠加读取行范围，只索引相交坐标；无日期列跳过；compact representation 仅在 C 基准证明收益后使用。
11. **如何消除 Dictionary Clone？** 保留 MiniExcel row dictionary 为只读 source，以计划期解析的 source key 直接 TryGetValue；不更改 source comparer，不为大小写匹配复制整行。
12. **如何消除逐行 Reflection？** 每 type/plan 编译并缓存 constructor/getter/setter/dynamic accessor；Provider hot loop 只调用 typed delegate；NPOI 同样用 E/F 基准决定是否替换现有反射 delegate。
13. **Relation 如何到 O(P+C)？** selector 纯函数；父项扫描一次建带 comparer 的 dictionary，`TryAdd` 保留首项并记录重复；子项扫描一次 lookup；无 DynamicInvoke/FirstOrDefault。
14. **Entity 如何支持 Cell/Merge/List Region？** immutable typed layout 把属性映射到 sheet+A1/range/list origin；merge 内任意地址归一到 anchor；固定字段和明细共享 converter/validation/error；NPOI native executor 执行。
15. **Template Existing Merge/Conflict？** 预读模板 merge index；完全相等复用；交叉/部分包含/anchor 不一致在写入前失败；只向 anchor 写值；失败不提交。
16. **Provider-native 优化？** NPOI 使用 `IWorkbook/ISheet/IRow/ICell`、compiled column plan 和 merge interval index；MiniExcel 使用原生 streaming/query、source-key binding、筛选 RawDate 补偿；不强迫共享 DOM/补偿。
17. **哪些优化必须 Benchmark 才保留？** 所有热点改写：clone 消除、compiled constructor/accessor、RawDate compact/index策略、dynamic preallocation、relation shared engine、header prepass、staging 策略；都需 micro+E2E before/after。
18. **500K/1M 验收？** 支持格式场景完成且内容正确、无 OOM/目标损坏/temp 泄漏，peak memory 和耗时在 Phase 0 预先批准门槛内；2CPU/4GiB 有可验证限制证据。未跑即 `NOT_VERIFIED`。
19. **如何保证同一 Candidate？** immutable candidate manifest 记录 commit/dirty hash、二进制/包哈希、TFM/RID/SDK；一次 build 后封存；每份报告引用 manifest，发现 hash 不同即整组作废重跑。
20. **正式发布还剩哪些 Gate？** Breaking API 审批、双 TFM build/test、public-only SPI fixture、NPOI/MiniExcel 合同、API snapshot 审批、package consumer、A-I benchmark、100K/500K/1M、2CPU/4GiB、外部 CI/生产机、资源/原子提交、文档/XML 注释、版本未变、独立 review。任一必需项未完成不得 `COMPLETED`。

## 7. 总体风险、回滚与完成定义

### 总体风险

- RC 前一次性 Breaking 面较大，API、SPI、Entity layout 若未先冻结会造成反复改写。
- MiniExcel 能力边界可能无法高效实现 Template/merge；不得用临时文件二次解析或部分支持掩盖。
- 性能优化可能提高吞吐却增加峰值内存；必须多指标判定。
- 历史证据与当前候选混用是最高发布风险。

### 总体回滚策略

- 每 Phase 独立可回滚；先保持 characterization，再改变算法。
- 失败时回滚具体优化/能力，不恢复旧命名 wrapper，不修改版本来规避门禁。
- Entity/Template 若 MiniExcel 不能原生完整支持，回滚其实现并保留明确 capability unsupported；NPOI 能力不受影响。
- API/SPI 未批准时停止后续实现，不先写 Provider 私有替代入口。

### 完成定义

仅当以下全部满足才把任务标记为 `COMPLETED`：

- Phase 0-12 所有必需验收通过，MUST/SHOULD review finding 清零。
- 双 TFM build/test、Provider/第三方 fixture、API snapshot、包 Consumer 全通过。
- Benchmark A-I 有可信 before/after；100K、500K、1M 和 2CPU/4GiB 达到已批准门槛。
- 取消、资源限制、流所有权、临时文件、原子提交均有真实证据。
- 同一 candidate manifest 覆盖所有报告和产物，根外/nested artifacts 为零。
- 版本文件及版本属性未变化；生产程序集间 IVT、runtime router、反射 hack 为零。
- 文档、迁移记录、双 TFM API snapshot、中文 XML 注释和 symbol→test map 完整。
- 生产机器或外部 CI 属于发布必需环境时必须实际 PASS；未执行只能标 `NOT_VERIFIED`/`PARTIAL`。
