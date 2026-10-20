# 统一 TODO：待实现与待确认事项

更新日期：2026-10-20。原 29 组未收口任务合并为下列 11 项，旧任务过程资料和生成产物已清理。已实现功能见[功能分类](implemented-features.md)。历史 Finding 先与当前候选复核，再决定关闭或实施；不按旧状态自动重做修复。

状态约定：

- **待实现**：最后一份记录明确存在本地可执行的实现/性能修复项。
- **待确认**：旧 Review 仍开放，但后续执行声称已修复，或代码、文档、候选版本已变化；先对照当前实现、测试和文档，再决定关闭或重新登记修复。
- **待外部验收/审批**：需要维护者批准、目标 Office/部署环境、生产资源或最终远端 CI；不得写成代码已失败或门禁已通过。
- **可选改进**：原记录明确标为非阻断或 OPTIONAL；需明确取舍后安排。

## 合并待办

### TODO-001：Entity Layout 批次直接验收

状态：**待确认**。`BING-OFFICES-ENTITY-LAYOUT-BATCH-20260930-001` 的 Review 记录 `FIX-001`～`FIX-003` 为 OPEN_ACTIONABLE；后续 execution 记录称三项修复已完成，Review 未随之更新。

下一步先核对当前 `EntityLayoutBatchImportTest`、Provider Consumer 和文档，确认动态列/关系/空区域、失败行坐标及计数、取消/回调异常、同步/异步和模板入口已有直接证据，且“全量通过”声明与实际执行范围一致。仅在当前证据确实缺失时补对应测试/文档；随后做独立 Review。

关闭条件：三项历史 Finding 均由当前代码和职责测试复核为已解决，完成范围准确的独立审查；若有缺口，拆成明确可执行的小项再关闭本条。

### TODO-002：确认实体批次完整资源预算契约

状态：**待实现/待确认（先确定设计与审批范围）**。F-006 的全链路对象/内存预算和超限结果语义尚未形成获批的续阶段契约。此项不表示现有临时磁盘额度、共享字符串磁盘索引、NPOI XLS 投影预算或现有结构限制缺失；`Buffered + Complete`、模板和失败工作簿与部分预算组合目前明确在读取前拒绝。

下一步先为 DOM、实体结果集合、唯一值/关系索引、错误集合及失败工作簿暂存定义预算单位、边界和组合规则；明确 `Complete` 与批次回调在超限后的结果/事务语义、跨 Provider 一致性及取消/清理要求。设计和 API 影响经维护者确认后再实现，覆盖 NPOI/ClosedXML、双 TFM、模板/失败输出和组合拒绝边界。

关闭条件：获批的设计明确资源上限和超限终态；实现有预算边界、目标保全、取消和临时资源清理的直接证据；对应 API 审批、Consumer 与正式双 TFM snapshot 完成。不得以缩小 BatchSize 或仅限制临时磁盘容量宣称端到端有界。

### TODO-003：复核字典模板修复后的独立结论

状态：**待确认**。`BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001` 的 Round 2 execution 记录 FIX-010/011/012 已完成，但保留的 Review 仍是 NEEDS_FIX。

下一步由独立审查对当前实现复核 HSSF 显式列表导入、名称扫描取消传递、异步/失败资源边界及对应 XLS/XLSX 直接测试；确认包消费者和文档示例对应当前候选。不要按旧 Review 把已记为完成的三项重新实现。

关闭条件：当前候选复核确认 FIX-010/011/012 已关闭，或把实际仍缺项另列为新 Finding；独立 Review 结果与最新 execution 一致。

### TODO-004：确认 Mapping Profile 增强性缺口

状态：**待确认，可选**。`BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001` 的最新 Review 为 PASS_WITH_ISSUES；FIX-006 是非阻断 SHOULD_FIX，涉及更多四形状 CSV/XLSX Integration、默认 DI Profile 命名文档、版本占位符及 descriptor/merge/cache benchmark。后续功能文档已更新 Profile 命名说明，不能照抄旧文档 Finding。

下一步对照现行 mapping 文档、迁移指南、测试和 Benchmark 参数，标出已解决部分；由维护者决定是否保留剩余代表性集成/基准补充。若没有当前用户价值，明确延期，不自动开启修复循环。

关闭条件：当前文档/代码证据复核完成；保留项有明确范围和验收标准，或由维护者记录延期/不再要求。

### TODO-005：验收高级透视文本 Count 客户端结果

状态：**待确认（外部验收）**。高级透视计划的 T-005/T-006 与 NPOI NPIVOT-002/004 记录仍称 Excel 文本 Count 首开、刷新、保存重开结果为空；后续 XML 结构修复没有新的客户端成功证据。NPOI 文档中旧的 Provider 能力矛盾已在后续执行记录中修正，不再列为代码 TODO。

下一步在可用且可控的 Excel 会话中，以当前候选生成工作簿，逐项执行首开、刷新、保存、关闭和重开；记录版本、源文件/结果哈希、Count/Min/Max 与列轴结果。先确认问题根因和列轴合同；在结果明确前不推进依赖该验收的最终阶段。

关闭条件：当前候选在计划要求的完整客户端矩阵满足数值合同并由独立审查核验；或维护者明确接受兼容边界并记录风险和替代行为。

### TODO-006：补齐需真实应用/部署环境的互操作验收

状态：**待确认（外部验收）**。不同任务的应用和环境合同需按其格式边界分别留证，不得用 OOXML 重开、静态结构测试或其他 Provider 结果替代。

下一步在可控环境完成各项实际适用的矩阵：Entity Failure Workbook 的批注、隐藏摘要 Sheet 和样式显示；透视模板更新及已有透视与业务 Sheet 组合的新夹具首开/修改来源/刷新/保存重开；同一代际工作簿的 Excel/LibreOffice/WPS 打开、保存、重开及结构复核；AsposeCells 在有效许可证、固定字体及目标函数集下执行公式重算、PDF/页面渲染。ClosedXML 的 Linux/字体及容量环境也按获批发布范围确认是否必需。

关闭条件：每个被要求的环境均有版本、输入/输出身份和实际结果；不可用的客户端/商业环境由维护者明确标为不适用或接受限制，不能伪造 PASS。

### TODO-007：逐成员归属并收口正式 API snapshot

状态：**待确认/待审批**。多项旧任务共享 Abstractions 候选，不能把混合差异视为单一任务的批准。2026-10-11 的历史候选记录为 27 项新增、0 项删除，其中 F4 7 项已批准、F5 1 项已批准、F6 17 项与 F7 2 项待审批/缺少审批；2026-10-15 记录又新增 Entity 临时预算 2 个候选成员。以上数字只描述当时候选，需重新捕获当前候选。

下一步对当前源码和所有受支持 TFM 重新生成候选，逐成员记录符号、所属任务、兼容性影响及批准/拒绝结论；重点复核 Entity Batch Projection/Failure/Forward、ClosedXML、ExcelDataReader、临时预算和透视模板更新的各自审批记录。将已批准范围与其他未决成员隔离，未经批准不更新正式 baseline。

关闭条件：当前候选每个新增/删除成员都有可追溯决定；双 TFM snapshot/API contract 通过，包消费者使用与该候选相同的包身份通过，独立 Review 确认身份一致。

### TODO-008：复核已完成的五项 100K 性能修复

状态：**待确认（修复后的独立复审）**。RC-API-PERF 的旧 Review 仍保留 Round 43 `FIX-002C=OPEN_ACTIONABLE`，但 execution 的 Round 53～54 已记录读取分派优化、双 TFM 直接回归和同候选五次确认，明确本地已无可执行性能代码 TODO。不能把旧审查分类重新写成当前待修复回退。

历史 Round 54 相对已批准 before 中位数的结果为：NPOI async -4.88%、MiniExcel sync +22.27%、excel-file-sync -43.99%、excel-file-async -65.18%、excel-throttled-async -49.62%；五项均在 25% 回退预算内。这是旧候选的测量记录，不代表当前候选重新测量。

下一步独立复核 Round 53～54 对应的当前代码、职责测试与可用原始证据，确认是否仍是同一候选；仅在证据已失效或不匹配时安排五目标的精准复测。FIX-002A 为 VERIFIED_BOUNDARY，FIX-002B 已关闭，不无条件重跑 Relation 100K、c16/c64 或完整 A-I。

关闭条件：独立审查确认 FIX-002C 已解决并统一状态；若当前候选有新回退或证据无法复用，再登记具体修复/验证项。生产资源与远端 CI 门禁分别按 TODO-009/010 收口。

### TODO-009：批准并完成正式性能、容量与资源门禁

状态：**待确认/待审批/待外部环境**。多个 RC 任务都有真实测量，但仍缺获批阈值或目标环境证据；探针成功不等于生产容量批准。相关旧 Review 中的性能/资源 OPEN 项先按当前候选复核，不直接当成现行回归。

下一步由负责人确定 workload、阈值、容差、生产环境、批准人和日期；按计划对当前候选补齐所需 Benchmark/Resource Matrix、取消延迟/峰值与保留量证据。仅按获批范围运行 500K/1M、Failure Workbook 双 DOM、2 CPU/4 GiB 等任务要求的矩阵；不同任务的 100K 指标不互相替代。

关闭条件：资源/性能批准或具名 waiver 完整；最终报告准确区分 PASS、FAIL、WAIVED、NOT_VERIFIED，证据绑定当前候选。未经批准不得把 UNAPPROVED、NOT_VERIFIED 或 BLOCKED 改成 PASS。

### TODO-010：形成可复现的最终交付候选和远端门禁

状态：**待确认（最终交付与外部门禁）**。当前多份任务的历史产物来自不同候选，不能合并成一份最终候选通过证据。若旧 Finding 涉及未跟踪报告、锁文件或 API 工具，先对照当前 Git/CI 规则和实际待交付提交确认，不将旧工作树状态自动套用到当前。

下一步在获授权的交付提交中纳入当前候选所需的 lockfiles、API 工具/baseline、审批记录、快照和获准保留的原始证据；从 clean checkout 执行 locked restore、Release build、双 TFM 测试、Docs、pack、package-only Consumer、API identity/compare。对 `BO-RC-20260908-002` 的 FIX-003 另需生产 2C/4G 单槽位证据和最终候选 GitHub Actions 全绿；探针 `SemaphoreSlim(1,1)` 不能替代生产入口配置或采样。

关闭条件：clean checkout 能复现当前候选的全部必需门禁；包、DLL、source 与 API candidate identity 相符；远端 run 的 head SHA 与交付候选一致，且所有计划要求的步骤通过。需要 Git staging、远端运行或生产部署的动作由获授权交付流程完成。

### TODO-011：确认剩余 TFM、复审与可选维护工作

状态：**待确认/可选**。旧任务在不同时间点记录了缺失旧 runtime、未做完整同候选测试、复审待回合及内部职责拆分未完成；不能仅凭历史 NEEDS_FIX 或 PARTIAL 推断当前构建仍失败。

下一步复核历史 NPOI 回归中 `ExcelWorkbookRequestTest.ExportToFile_Success_ShouldCreateAndReplaceTarget` 的 Windows `File.Replace` / `UnauthorizedAccessException`，区分环境限制与当前文件提交缺陷；再按当前项目支持的 TFM 和 CI 矩阵核实 `net5`、`net7`、`netcoreapp3.1` 等旧 runtime 是否仍属发布承诺；逐项确认后续 execution 声称修复的 Finding 是否已有当前独立复审。再由维护者选择是否推进大类职责拆分、Mapping Profile 增强矩阵及低优先级中文注释/生产符号追溯（Post-MiniExcel FIX-010/FIX-014、RC Closure FIX-006、RC Hardening Phase 3、Stabilize P1-REMAINDER-01/05）。

清理前仍需逐项复核的历史细项如下；编号与描述只保留追溯，不直接认定为当前缺陷。

| 原任务 / Finding | 复核对象 | 合并处理 |
| --- | --- | --- |
| RELEASE-HARDENING-20260827-001 / FIX-003 | 请求级 metadata 的 XLS/HSSF 默认保留、顺序隔离、64 并发导出重开与 Docs 用法 | 核对当前职责覆盖，缺项再补 |
| RELEASE-HARDENING-20260827-001 / FIX-004 | NPOI selector 是否重复解析，计划与物化是否复用同一物理 Sheet | 对照当前解析链与边界测试决定是否需修复 |
| RELEASE-HARDENING-20260827-001 / FIX-006 | package-only Consumer 的空缓存、第三方依赖及可信源前提 | 并入 TODO-010 的可复现交付 |
| RELEASE-HARDENING-20260827-001 / FIX-007 | 职责拆分、全 TFM、Benchmark、办公客户端与最终报告 | 按 TODO-006/009/010 和本条逐项决定适用范围 |
| RELEASE-HARDENING-20260831-001 / FIX-006 | XLS/XLSX selector 越界索引、VeryHidden 与客户端政策；Round 10 记录跳过 | 核对当前测试与客户端证据，缺项再登记 |
| RC-HARDENING-20260904-001 / FIX-004 | `SheetExtensions.TryAddPicture` 两个重载的空值、空字节、坐标、图片类型及异常合同 | 对照当前实现与直接测试，不以旧异常描述代替复核 |
| RC-HARDENING-20260904-001 / FIX-005 | 日期多输入格式及 request/document/profile 独立配置 | 现行日期/Profile 文档已有相关能力，确认历史项是否已被后续实现覆盖 |

关闭条件：受支持 TFM 均有当前候选证据，或已批准退出支持并更新发布矩阵；上表细项及修复轮次后的 Review 已复核；可选重构/文档项有明确接受、延期或取消决定。

## 不重新登记的旧路线图项

`BING-OFFICES-EXCEL-CAPABILITY-GAP-20260929-001` 是较早的差距快照。其“分组与分页组合、逐页命名锚点、公共 PivotTable”条目与当前 `docs/excel/nuget-migration.md`、`docs/excel/pivot-tables.md` 已记录的能力重叠；批次唯一性、Workbook 校验、动态列、关系、Failure Workbook 也有后续实现和专题文档。迁移时不把这份旧快照里的清单整体复制为未实现功能。只有 TODO-001/002 的当前验收或资源边界复核确认仍有差距时，才登记新修复项。商业公式重算/PDF/页面渲染仍属 TODO-006 的部署验收。

## 29 个旧任务覆盖映射

| 旧任务 ID | 合并去向 |
| --- | --- |
| `BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001` | TODO-007/009/010 |
| `BING-OFFICES-ADVANCED-PIVOT-TABLES-20261005-001` | TODO-005/007 |
| `BING-OFFICES-CLOSEDXML-PROVIDER-20260922-001` | TODO-006/007/009 |
| `BING-OFFICES-ENTITY-BATCH-PROJECTION-20261006-001` | TODO-002/007 |
| `BING-OFFICES-ENTITY-FAILURE-WORKBOOK-20261006-002` | TODO-006/007 |
| `BING-OFFICES-ENTITY-FORWARD-BATCH-20261006-001` | TODO-002/007 |
| `BING-OFFICES-ENTITY-LAYOUT-BATCH-20260930-001` | TODO-001 |
| `BING-OFFICES-ENTITY-TEMPORARY-BUDGET-20261015-001` | TODO-002/007 |
| `BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001` | TODO-003 |
| `BING-OFFICES-EXCEL-CAPABILITY-GAP-20260929-001` | TODO-001/006；旧功能路线图按“不重新登记的旧路线图项”复核 |
| `bing-offices-excel-import-export-enhancement-v2` | TODO-006/009 |
| `BING-OFFICES-EXCELDATAREADER-FAILURE-WORKBOOK-20261006-001` | TODO-006/007 |
| `BING-OFFICES-MAPPING-PROFILE-VNEXT-20260823-001` | TODO-004/011 |
| `BING-OFFICES-NPOI-PIVOT-CACHE-20261005-001` | TODO-005/007/010 |
| `BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001` | TODO-007 |
| `BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001` | TODO-006/009/011 |
| `BING-OFFICES-PRE-RC-CLEANUP-20260901-001` | TODO-007/009/010/011 |
| `BING-OFFICES-PROVIDER-CONTRACT-HARDENING-20260925-001` | TODO-009/010 |
| `BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001` | TODO-008/009/010/011 |
| `BING-OFFICES-RC-CLOSURE-20260905-001` | TODO-007/009/011 |
| `BING-OFFICES-RC-HARDENING-20260904-001` | TODO-007/009/011 |
| `BING-OFFICES-RC-HARDENING-20260906-001` | TODO-007/009/010/011 |
| `BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001` | TODO-009/010/011 |
| `BING-OFFICES-RELEASE-HARDENING-20260825-01` | TODO-006/009/011 |
| `BING-OFFICES-RELEASE-HARDENING-20260827-001` | TODO-009/010/011 |
| `BING-OFFICES-RELEASE-HARDENING-20260831-001` | TODO-007/009/010 |
| `BO-RC-20260908-002` | TODO-009/010 |
| `TASK-BING-OFFICES-20260821-MAPVAL-V2` | TODO-004/010 |
| `TASK-BING-OFFICES-STABILIZE-20260825-001` | TODO-006/010/011 |

`ai_docs/tasks` 仅保留 3 份现有 CI/API 工具引用的审批附件，见[索引](README.md)。审批范围不等于任务全部验收通过，旧计划、执行、审查和 TODO 不再保留。
