<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001
AI_EXECUTION_FINISHED_AT: 2026-09-28T13:10:54.000Z

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 23

已完成[旧版 Excel API 迁移诊断](../BING-OFFICES-LEGACY-MIGRATION-DIAGNOSTICS-20260928-001/execution.md)：可选分析器新增 `BOM001`、`BOM002`，仅提示已绑定的旧导出服务/选项和动态列标记；双 TFM 各 20/20，真实 NuGet 包消费者输出预期三条警告。业务自定义坐标特性保留人工迁移。以下 Continuation 22 及更早章节保留为历史。

## 当前阶段：Continuation 22

已完成[第三方 Excel Provider 最小项目模板](../BING-OFFICES-PROVIDER-STARTER-20260928-001/execution.md)：独立 net6.0/net8.0 项目从本地 NuGet 包实现基础 XLSX 工作簿导出，双 TFM 各 11/11；CI 新增包消费者验证。模板只声明已实现能力，Entity 合同在实现对应接口后启用。以下 Continuation 21 及更早章节保留为历史。

## 当前阶段：Continuation 21

已完成[报表级工作簿导出策略](../BING-OFFICES-WORKBOOK-EXPORT-STRATEGY-20260928-001/execution.md)：调用方显式选择完整工作簿或前向流式模式；同步、异步、流和文件路径沿用现有 Provider 契约。NPOI XLS/XLSX、ClosedXML XLSX 与 SpreadCheetah XLSX 的职责级测试在两个 TFM 各 18/18，通过本地 NuGet 包的第三方消费者亦在两个 TFM 通过；双 TFM API 快照无删除成员。以下 Continuation 20 及更早章节保留为历史。

## 当前阶段：Continuation 20

已完成[Entity Layout 静态区域冲突诊断](../BING-OFFICES-ANALYZER-AREA-CONFLICT-20260928-001/execution.md)：新增 `BOE005`，在内联常量布局中检查合并、固定单元格和显式有界列表的冲突；双 TFM 分析器职责测试各 17/17，本地分析器包消费者正常编译、`BOE001` 与 `BOE005` 负例均按预期失败。修正最终 Footer 命名锚点的过期文案。以下 Continuation 19 及更早章节保留为历史。

## 当前阶段：Continuation 19

最终 Footer 新增 `FormulaSumDetailRowsAbove`，在分页、分组小计后仍按实际明细行段写入原生 SUM 公式，排除间隔、小计和签字行。NPOI HSSF/XSSF、ClosedXML XLSX 的职责级测试及双 TFM 共享 Provider 合同已通过；阶段计划与最终验证记录见[跨小计明细总计公式](../BING-OFFICES-ENTITY-DETAIL-SUM-20260928-001/execution.md)。以下 Continuation 18 及更早章节保留为历史。

## 当前阶段：Continuation 18

本阶段为独立 Entity Provider 合同包补入第七项 `VerifyFooterFormulas`，覆盖原生 Footer 公式、连续明细求和公式、导入边界及布局阶段跨分页小计冲突拒绝；合同文档和主索引已同步。范围见[阶段计划](../BING-OFFICES-ENTITY-FORMULA-CONTRACT-20260928-001/plan.md)，验证进度见[阶段执行记录](../BING-OFFICES-ENTITY-FORMULA-CONTRACT-20260928-001/execution.md)。以下 Continuation 17 及更早章节保留为历史。

## 当前阶段：Continuation 17

已完成 [Footer 连续明细求和公式辅助方法](../BING-OFFICES-ENTITY-FOOTER-SUM-20260928-001/execution.md)，范围见[阶段计划](../BING-OFFICES-ENTITY-FOOTER-SUM-20260928-001/plan.md)。方法依据连续明细快照、冻结的 `GapRows` 和公式单元格相对行生成原生公式，空明细写入 `=0`；分页/分组中间小计可用，最终 Footer 跨越中间小计时拒绝。NPOI 双 TFM 各 13/13、ClosedXML 双 TFM 各 8/8 定向通过，详情见阶段执行记录。以下 Continuation 16 及更早章节保留为历史。

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 16

已完成 [Entity Layout Footer 显式公式](continuation-16-execution.md)。NPOI XLS/XLSX 与 ClosedXML XLSX 将 `Formula` 内容写为原生公式单元格，保留普通 `Cell` 语义及 Footer marker 导入边界。定向测试、Provider 常规测试、API 审批和第三方包消费者结果见阶段执行记录。以下 Continuation 15 及更早章节保留为历史。

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 15

已完成[独立 Entity Provider 合同测试包](continuation-15-execution.md)，范围见[阶段计划](continuation-15-plan.md)。新包包含 net6.0/net8.0 合同程序集，仅依赖公开 Abstractions；第三方消费者已取消源码链接，从本地 NuGet 产物在双 TFM 执行全部六项合同。CI 增加独立打包及专用源，生产包数量门禁不变。以下 Continuation 14 及更早章节保留为历史。

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 14

已完成[分析器准确性、包消费与命名布局合同](continuation-14-execution.md)，范围见[阶段计划](continuation-14-plan.md)。分析器双 TFM 各 8/8；命名锚点合同及冲突负例、Provider Contract 双 TFM 各 295/295；NPOI HSSF/XSSF 的新增职责测试双 TFM 各 4/4。NPOI、ClosedXML 常规测试和第三方包消费者均通过。本阶段无公共 API 变更、版本号修改或发布。以下 Continuation 13 及更早章节保留为历史。

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 13

已完成[迁移示例与 Entity Layout 编译期诊断](continuation-13-execution.md)，范围见[阶段计划](continuation-13-plan.md)。完整采购单示例通过 Docs 原文执行；独立分析器四项诊断在双 TFM 各 4/4 通过，本地分析器包生成，Release 解决方案构建成功。当前阶段不涉及既有生产 API 变更或包发布。以下 Continuation 12 及更早章节保留为历史。

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 12

已完成 [可复用 Entity Layout Provider 合同源码模板](continuation-12-execution.md)，范围见 [阶段计划](continuation-12-plan.md)。官方 NPOI/ClosedXML 的双 TFM 合同测试各 9/9，Provider Contract 各 291/291，独立 NuGet 包消费者双 TFM 均通过；解决方案 Release 构建及 Docs 测试通过。本轮没有生产代码或公共 API 变更，未覆盖 API baseline。

后续顺序：评估 Roslyn 地址/动态 Key 分析器；独立 NuGet 测试包的版本与发布流程另立任务。以下 Continuation 11 及更早章节保留为历史。

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 11

已完成最终 Footer 命名锚点。NPOI XLS/XLSX 和 ClosedXML XLSX 支持在导出时创建或更新指向最终 marker 的名称，模板冲突返回 Plan 配置错误；旧 Footer 和导入 marker 边界保持兼容。详见 [阶段执行记录](continuation-11-execution.md) 和 [阶段范围](continuation-11-plan.md)。双 TFM API 只新增两个批准成员，最终 compare 无差异；第三方 NuGet 包消费者、Release solution 构建和 `Category!=Large` 常规测试均通过。本阶段不涉及发布。

后续顺序：可复用 Entity Layout Provider 合同模板 → Roslyn 地址/动态 Key 分析器。以下 Continuation 10 及更早章节保留为历史记录。

AI_EXECUTION_STARTED_AT: 2026-09-27T23:07:17.187Z

AI_EXECUTION_RESULT: PASS

## 当前阶段：Continuation 10

已完成跨页签字区组合能力验证及两处边界修复。详见 [阶段执行记录](continuation-10-execution.md) 和 [阶段范围](continuation-10-plan.md)。本轮公共 API 未变，未覆盖 API baseline；NPOI 675/675、ClosedXML 196/196、PublicApi 11/11 均在两个 TFM 通过，Provider Contract 双 TFM 各 282/282（net6.0 并行初次失败及独立通过均已记录），Docs 10/10。构建通过。本轮严格编码检查只覆盖本轮修改文件，API baseline 与 ProfileFixtures.xml 的既存偏差已单列。

后续顺序：尾部命名锚点 → 可复用 Provider 合同模板 → Roslyn 分析器。以下 Continuation 9 及更早章节保留为历史，不代表当前阶段重新运行了包消费者或完成全量独立 Review。

## TODO Continuation 9

本轮按分页相关 TODO 顺序完成分页小计能力并补齐 Provider 合同。实体列表新增 `PageSubtotal`，必须与 `PageBreak` 同时配置，按每个中间页的只读明细快照执行聚合；最后一页继续使用普通 `Footer`，导入按 marker、间隔和 Footer 高度跳过小计区域。分页小计与 `GroupSubtotal` 互斥，marker 缺失、重复或位置不符合页边界时在 Plan 阶段返回结构化配置错误。

本轮交付：

- NPOI XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 均支持分页小计导出、水平分页符、最终 Footer、模板/合并预检和导入边界；移除了两个未使用的内部分页行计算辅助方法。
- 第三方包消费者增加公开 `PageSubtotal` 导出/导入往返验证；`net6.0`、`net8.0` 均输出 `third-party-public-only-provider-ok`。
- Provider Contract 新增 `EntityPageSubtotals` 独立场景档案；NPOI/ClosedXML 作为 Supported，MiniExcel 作为 NotApplicable，预期不从生产能力声明生成。
- 文档、API approval、生产符号到测试追溯已同步；`build/api-snapshot-baseline.json` 记录本轮批准的两个新增成员。

本轮验证：

- NPOI 分页小计直接职责：`2/2`（net6.0、net8.0）通过，覆盖正常往返、分页符、导入跳过、marker 缺失和位置错误；Provider `Category!=Large`：`663/663`（net6.0、net8.0）通过。
- ClosedXML 分页小计直接职责：`2/2`（net6.0、net8.0）通过，覆盖正常往返、分页符、导入跳过、marker 缺失和位置错误；Provider `Category!=Large`：`190/190`（net6.0、net8.0）通过。
- Provider Contract `Category!=Large`：`282/282`（net6.0、net8.0）通过；Docs：`10/10`（net8.0）通过；Public API Integration：`11/11`（net8.0）通过。
- Release solution build：成功，0 错误；保留既有 4 个 `Microsoft.Bcl.Memory` net6.0 警告。
- API capture/compare：最终双 TFM compare 通过，`api-page-subtotal-final/compare-final/api-diff.json` 为 `net6.0={}`、`net8.0={}`；更新前的候选差异证据只包含 `PageSubtotal` 方法和隐藏布局属性两个新增成员，`removed=[]`。
- NuGet package consumer：还原、构建和运行均通过，`net6.0`、`net8.0` 输出 `third-party-public-only-provider-ok`。
- `git diff --check`：无新增空白错误；仅保留 `ProfileFixtures.xml` 的存量 CRLF 规范化提示。

TODO 状态（本轮后）：

### Completed

- 分页小计、marker 导入边界、失败位置校验和 Provider 合同覆盖已完成。
- 双 TFM API snapshot、基线、包消费者、文档和生产符号追溯已同步。

### Open Actionable

- 无。

### Deferred

- 跨页签字区、尾部命名锚点、Footer 公式 DSL、可复用契约测试包发布化、Roslyn 地址冲突分析器、Source Generator、旧版迁移分析器和显式 DOM/Streaming 策略层继续延期。

### Next recommended order

1. 在分页小计稳定后定义跨页签字区与尾部命名锚点的占用模型、模板合并和导入边界。
2. 将现有 Provider Contract 场景整理为可引用测试模板/包，继续保持独立预期档案。
3. 复用布局预检规则实现 Roslyn 地址冲突分析器，覆盖固定地址、锚点、动态 Key、计算列和 Footer marker。
4. 以反射启动成本和迁移规模实测决定 Source Generator 与旧版迁移分析器投入。
5. 待资源探针证明需要后设计显式 DOM/Streaming 策略层，保持 Provider 选择可预测。

### Current Gate

- `IMPLEMENTATION_STATUS`: `PASS`
- `TEST_STATUS`: `PASS`
- `REVIEW_STATUS`: `PASS`
- `RELEASE_STATUS`: `PASS`
- `OPEN_ACTIONABLE`: `0`
- 未自动 commit、push 或发布。

# 实施执行报告

## 执行结论

计划内的实体布局能力、NPOI/ClosedXML Provider 实现、业务组合测试、NuGet 消费者、文档、API 快照和追溯记录均已完成。生产程序集没有新增友元依赖，Entity Layout Provider 通过公开描述契约工作。

状态矩阵：

| 状态 | 结果 |
|---|---|
| `IMPLEMENTATION_STATUS` | `PASS` |
| `TEST_STATUS` | `PASS` |
| `PERFORMANCE_EVIDENCE_STATUS` | `NOT_RUN`（计划明确排除无关 500K/1M 矩阵） |
| `RESOURCE_EVIDENCE_STATUS` | `PASS`（本次新增取消、失败清理和目标完整性职责测试） |
| `EXTERNAL_GATE_STATUS` | `NOT_APPLICABLE`（计划未要求外部 CI 或生产环境门禁） |
| `RELEASE_STATUS` | `PASS` |
| `GOAL_STATUS` | `COMPLETED` |

`OPEN_ACTIONABLE=0`，没有待修复的本地问题。未自动提交、推送、创建 PR、发布包或修改版本号。

## 任务信息

- Task ID：`BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001`
- 基线提交：`6048f5ec70751eb2fcd9b149f178e9c54896a4b1`
- 目标框架：`net6.0`、`net8.0`
- 变更范围：仅 `Bing.Offices`；`D:\MyWorking\Utopa\Utopa.Erp.IE` 未读取后写回、未复制、未提交。

## 变更影响分析

```text
ChangedFiles: Abstractions Entity Layout/属性与 Footer 描述、NPOI/ClosedXML Provider、测试、文档、API 基线与任务证据
ChangedProjects: Bing.Offices.Abstractions, Bing.Offices.Npoi, Bing.Offices.ClosedXml, 受影响测试与 Consumer
ChangedPublicContracts: 属性式固定单元格、动态列分组、未知值策略、Footer、公开隐藏 Provider SPI
ChangedRuntimePaths: Entity Layout 导出/导入、NPOI HSSF/XSSF、ClosedXML XLSX、模板与失败工作簿
ChangedProviders: NPOI, ClosedXML
ChangedTFMs: net6.0, net8.0
ChangedBuildPackaging: Release 构建、NuGet pack、双 TFM API capture/compare
ChangedBenchmarkHarness: No
ChangedDocsOnly: No
AffectedDependents: Provider Contract、Public API、Docs、NuGet Consumer、Entity Layout 业务测试
RiskLevel: HIGH
```

## 计划执行情况

### P1：公共实体布局抽象

- 新增 `ExcelEntityCellAttribute`，支持属性式固定单元格、继承和可选转换器名称。
- 新增 `LayoutFromAttributes` 与 `CellsFromAttributes`，按公开实例属性名稳定排序并保留只读导出/导入错误语义。
- 新增多个动态列分组、未知值策略及其冲突、唯一性和字典写回校验。
- 新增相对明细 Footer：marker、间隔行、固定/聚合单元格、样式、相对合并和预检。
- 新增 `EditorBrowsable(Never)` 的公开 Provider SPI 描述类型，消除生产程序集对内部类型和友元访问的需要。

### P2：Provider 实现

- NPOI 移除 Entity Layout 动态列拒绝分支，支持 HSSF/XSSF 单字典、多分组、属性单元格、Footer 和 marker 导入边界。
- ClosedXML 接入分组描述并实现与 NPOI 一致的 Footer、模板合并、样式、预检、只读属性和错误分类。
- 同步/异步路径沿用现有取消、文件提交、流所有权、DOM 准入和临时文件清理契约。
- `rg InternalsVisibleTo src -g '*.cs'` 复核结果只包含明确测试程序集；未发现生产 Provider 名称或未知程序集友元。

### P3：真实业务测试

- 覆盖商品档案三组动态字典、采购单属性头与合计 Footer、赠品单空明细、盘点单多 Sheet/多区域和导入校验。
- 覆盖重复 Key、物理列冲突、Footer 越界/重叠、marker、聚合异常、取消、提交失败、目标文件完整性和资源清理。
- 覆盖 NPOI XLS/HSSF、NPOI XLSX/XSSF、ClosedXML XLSX，并验证同步/异步结构一致性。
- NuGet Consumer 使用生成包和公开 API 在 `net6.0`、`net8.0` 完成属性布局、动态分组、Footer 和公开提交器运行验证。

### P4：API、文档与交付

- 补齐新增成员中文 XML 注释、Provider 矩阵、动态列/Footer 文档和 NPOI 单据迁移示例。
- 更新 API approval、双 TFM snapshot、生产符号到测试方法追溯表和本执行报告。
- API baseline 已按明确的用户计划审批同步新增成员；双 TFM compare 无删除成员。

## 修改文件

生产代码：

- `src/Bing.Offices.Abstractions/Bing/Offices/Entities/ExcelEntityCellAttribute.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Entities/ExcelEntityListExtensions.cs`
- `src/Bing.Offices.Abstractions/Bing/Offices/Entities/ExcelEntityLayout.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Entities/NpoiEntityLayoutExecutors.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Exports/NpoiExportColumnPlanner.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Imports/ExcelImportExecutionOptions.cs`
- `src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiImportSheetExecutor.cs`
- `src/Bing.Offices.ClosedXml/Bing/Offices/Entities/ClosedXmlEntityLayoutExecutor.cs`

测试、Consumer、文档和交付记录同时更新；完整清单以 `git status --short` 和本任务目录为准。

## API / 数据 / 配置变化

- 新增属性式固定单元格、动态列分组、Footer Builder 和公开隐藏 SPI；旧 Fluent Layout、单字典动态列和 Provider 入口保持源码/二进制兼容。
- `api-final/compare/api-diff.json` 为：

```json
{
  "net6.0": {},
  "net8.0": {}
}
```

- 未变更数据模型、包版本或运行时配置；未公开 `AtomicFileCommitter` 或其他内部文件系统适配器。

## 测试结果

| 验证项 | 结果 |
|---|---|
| NPOI Entity Layout Provider 直接职责筛选（net6.0/net8.0） | `34/34`、`34/34` 通过 |
| ClosedXML Entity Layout/Gap 模板直接职责筛选（net6.0/net8.0） | `15/15`、`15/15` 通过 |
| NPOI 完整测试（Category!=Large） | `645/645`、`645/645` 通过 |
| ClosedXML 完整测试（Category!=Large） | `168/168`、`168/168` 通过 |
| Provider Contract（net6.0/net8.0） | `276/276`、`276/276` 通过 |
| Public API Integration（net6.0/net8.0） | `32/32`、`32/32` 通过 |
| Common Tests（net6.0/net8.0） | `204/204`、`204/204` 通过 |
| Docs Tests（net8.0） | `10/10` 通过 |
| 全解决方案 `Category!=Large` | `2982` 通过，`0` 失败，`0` 跳过 |
| NuGet Consumer | `net6.0`、`net8.0` 均输出 `third-party-public-only-provider-ok` |
| API capture/compare | `net6.0`、`net8.0` 均通过，`removed` 为空 |

NuGet Consumer 的 `net8.0` 运行使用 `UseAppHost=false` 绕过本机旧 `apphost.exe` 文件锁；托管程序集构建和运行均通过。`net6.0` 仅产生 SDK 生命周期提示，不影响结果。

## Build / Typecheck / Lint / Format

- `dotnet build Bing.Offices.sln -c Release --no-restore -m:1 --nologo`：成功，0 错误；4 个 `Microsoft.Bcl.Memory` net6.0 警告为既有依赖警告。
- `git diff --check`：无空白错误；仅报告未改动的 `tests/Bing.Offices.ProfileFixtures/Bing.Offices.ProfileFixtures.xml` 存量 CRLF 规范化提示。
- 新增 C# 文件检查为 UTF-8 BOM、LF；新增/修改 Markdown 为 UTF-8、LF；未引入 Mixed Line Endings。
- 未运行无关的 500K/1M 性能矩阵，符合计划范围。

## 计划偏差

- 计划要求的公开包消费者已完成；由于隔离包缓存第一次只包含本地生成包，恢复时缺少第三方依赖，改用仓库任务缓存中的已存在依赖并重新运行，最终两个 TFM 均通过。这是恢复环境差异，不是源码或 API 问题。
- 本地没有外部 CI/生产环境门禁，因此外部验证记为 `NOT_APPLICABLE`，没有用本地结果冒充外部证据。

## 基线问题

- API baseline 的 candidate identity、双 TFM 成员快照和审批已同步更新；compare 结果为空差异。
- 存量 ProfileFixtures XML 的 CRLF 提示未由本任务引入，也未做全仓库行尾转换。

## 已知问题、限制与后续建议

- MiniExcel、SpreadCheetah、ExcelDataReader、AsposeCells 不声明 Entity Layout 能力，继续按能力矩阵拒绝或不注册，不做伪实现。
- 未加入 Footer 公式 DSL、模板命名锚点、分页小计、Roslyn 分析器、Source Generator 或大型报表自动 Provider 切换；这些保留为计划中的后续任务。
- 性能/容量重矩阵属于 `DEFERRED`，原因是本次没有修改共享热路径且计划明确排除；不影响本次功能验收。

## Reviewer 注意事项

- 独立复核发现的 ClosedXML 只读属性、Footer 越界、模板合并、Footer 样式和测试证据问题均已修复并由定向测试覆盖。
- 重点检查 `ExcelEntityLayout` 的冲突预检、NPOI/ClosedXML 相对 Footer 语义、HSSF/XSSF 分支和公开 SPI 的 `EditorBrowsable(Never)` 边界。
- 生产友元门禁应继续保持精确测试程序集白名单；本次未向任何生产程序集授予 `InternalsVisibleTo`。

## TODO 状态分类

### Completed

P1、P2、P3、P4 及其所需本地验证全部完成。

### Open Actionable

无。

### Blocked Approval

无；新增 API 已获得当前任务计划中的明确批准。

### Blocked External

无；本计划未设置外部 CI/生产环境门禁。

### Not Applicable

外部 CI、生产环境资源探针及无关大规模性能矩阵不属于本任务验收范围。

### Accepted Limitations

非 Entity Layout Provider 不声明该能力；`net6.0` 由当前 SDK 输出生命周期提示；存量 XML 行尾提示保持不变。

### Verified Boundaries

生产友元仅指向测试程序集；Footer、动态列、取消、失败清理和调用方流所有权边界均由职责测试验证。

### Deferred

公式 DSL、命名锚点、分页/分组汇总、契约测试包、分析器、Source Generator 和大报表策略层。

## Git 状态

- 工作区保留本次源码、测试、文档、API 和任务证据改动。
- 未执行 `git add`、`git commit`、`git push`、PR 创建或发布。

## Review 修复记录

### Round 1

- Review 状态：`NEEDS_FIX`
- Fix Scope：`recommended`
- Review 文件：`ai_docs/tasks/BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001/review.md`

#### FIX-001

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/Bing.Offices.ClosedXml/Bing/Offices/Entities/ClosedXmlEntityLayoutExecutor.cs`
  - `tests/Bing.Offices.ClosedXml.Tests/ClosedXmlProviderTest.cs`
- 根因：ClosedXML 导入找到 Footer marker 后将 marker 前一行直接作为明细终点，未扣除 `GapRows`；模板间隔行有内容时会被误读为明细。
- 修复：按 `markerRow - 1 - GapRows` 计算明细终点，并拒绝早于合法 GapRows 位置的 marker；增加模板间隔行保留内容的真实导入回归测试。
- 验证：
  - `EntityTemplate_FooterGapWithTemplateContent_ShouldNotImportGapAsDetail`（net6.0/net8.0）：PASS
  - ClosedXML Entity Layout/Gap 模板直接职责筛选（net6.0/net8.0）：PASS

#### FIX-002

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/Bing.Offices.ClosedXml/Bing/Offices/Entities/ClosedXmlEntityLayoutExecutor.cs`
  - `tests/Bing.Offices.Npoi.Tests/EntityLayoutProviderTest.cs`
  - `tests/Bing.Offices.ClosedXml.Tests/ClosedXmlProviderTest.cs`
- 根因：ClosedXML 对动态字典直接调用 `Activator.CreateInstance` 和 Setter，类型不兼容、无无参构造或 Setter 失败时会泄漏底层异常。
- 修复：集中检查 `IDictionary<string, object>` 兼容性，捕获实例化与属性写入异常并统一包装为 `BingOfficesConfigurationException(BingOfficesStage.Plan)`；保留接口字典的默认创建。
- 验证：
  - 两个 Provider 的动态字典创建职责测试各覆盖可创建、只读、无无参构造和不兼容接口：PASS
  - NPOI/ClosedXML Entity Layout 直接职责筛选（net6.0/net8.0）：PASS

#### FIX-003

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
  - `src/Bing.Offices.Npoi/Bing/Offices/Entities/NpoiEntityLayoutExecutors.cs`
  - `src/Bing.Offices.ClosedXml/Bing/Offices/Entities/ClosedXmlEntityLayoutExecutor.cs`
  - `tests/Bing.Offices.Npoi.Tests/EntityLayoutProviderTest.cs`
  - `tests/Bing.Offices.ClosedXml.Tests/ClosedXmlProviderTest.cs`
- 根因：Footer 未声明 `End` 时只做相对地址检查，未把起始位置、间隔行和 Footer 单元格/合并折算为工作簿绝对行列上限；HSSF 上限低于 XLSX。
- 修复：NPOI 使用 `IWorkbook.SpreadsheetVersion` 做 HSSF/XSSF 物理边界预检，ClosedXML 按 XLSX 上限预检；预检在任何 Footer DOM 写入前执行，并统一返回 `Plan` 配置错误。
- 验证：
  - HSSF/XSSF/ClosedXML 物理行列边界职责测试：PASS
  - 越界场景断言目标流保持空：PASS
  - NPOI/ClosedXML 完整 Provider 测试（net6.0/net8.0）：PASS

#### FIX-004

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 执行状态：`COMPLETED`
- 修改文件：
  - `tests/Bing.Offices.Npoi.Tests/EntityLayoutProviderTest.cs`
  - `tests/Bing.Offices.ClosedXml.Tests/ClosedXmlProviderTest.cs`
  - `ai_docs/tasks/BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001/production-symbol-test-map.md`
  - `ai_docs/tasks/BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001/execution.md`
- 根因：P3 计划场景没有逐方法证据，执行报告把未覆盖的业务和失败矩阵写成已完成。
- 修复：补充三字典商品档案、无表头空赠品单、多 Sheet 盘点单、marker 错误、动态字典创建、Footer 物理边界和模板 GapRows 内容场景，并在追溯表中登记生产符号到测试方法的映射；同步修正测试计数和报告措辞。
- 验证：
  - NPOI Entity Layout Provider 直接职责筛选：`34/34`（net6.0/net8.0）PASS
  - ClosedXML Entity Layout/Gap 模板直接职责筛选：`15/15`（net6.0/net8.0）PASS
  - NPOI Provider：`645/645`（net6.0/net8.0）PASS
  - ClosedXML Provider：`168/168`（net6.0/net8.0）PASS
  - Provider Contract：`276/276`（net6.0/net8.0）PASS
  - `git diff --check`：PASS；仅保留 ProfileFixtures 存量 CRLF 提示

### Round 1 汇总

- MUST_FIX：无
- 已完成：`FIX-001`、`FIX-002`、`FIX-003`、`FIX-004`
- PARTIAL：无
- BLOCKED：无
- FAILED：无
- 回归验证：受影响 Provider、Entity Layout 直接职责和 Provider Contract 均通过；修复前全解决方案证据按未受影响范围复用。
- 下一步：重新进行独立 Review

## Independent Review Follow-up

- 第二轮独立复核期间发现动态字典自定义构造函数或 Setter 抛出 `OperationCanceledException` 时应保持取消语义；ClosedXML 统一创建器和 NPOI 动态组分发器现已透传该异常，不再包装为配置错误。
- Follow-up 定向测试：NPOI/ClosedXML net6.0、net8.0 均通过；受影响 Provider 全集仍为 NPOI `645/645`、ClosedXML `168/168`。

## Round 2 TODO Continuation

本候选版本的 L1 校验矩阵发现 ClosedXML Entity Layout 列表导入没有接入跨行唯一性跟踪：`[ExcelUnique]` 绑定本身不保存状态，重复值因此未产生结构化错误。已在 `ClosedXmlEntityLayoutExecutor.ReadRegion` 接入现有公开隐藏 `UniqueTracker`，并将固定列、动态列的唯一值按行成功提交、失败回滚，同时沿用 `MaxTrackedUniqueValues` 和 `UniqueComparison` 配置。未修改公共成员、文件提交、取消、DOM 准入或序列化格式。

当前验证：

- ClosedXML Entity Layout 直接职责筛选：`20/20`（net6.0、net8.0）通过。
- ClosedXML Provider `Category!=Large`：`176/176`（net6.0、net8.0）通过。
- NPOI Entity Layout 校验矩阵：`1/1`（net8.0）通过；NPOI Provider 源码范围未因本轮修复变化，前轮 `645/645` 双 TFM 证据可复用。
- `Bing.Offices.sln` Release 构建：成功，0 错误；4 个既有 `Microsoft.Bcl.Memory` net6.0 警告。
- `git diff --check`：无新增空白错误；保留 `ProfileFixtures.xml` 存量 CRLF 提示。

本节证据属于新的工作树候选版本；顶部 Round 1 和 Independent Review Follow-up 的历史计数保持不变。本轮修改后的候选版本已完成独立复核，唯一性测试/追溯缺口已补齐，`OPEN_ACTIONABLE=0`。

TODO 状态（当前候选版本）：

### Completed

- ClosedXML Entity Layout 固定列唯一性、逐行提交/回滚和唯一跟踪资源上限已实现并通过双 TFM Provider 验证。
- 新增无效行回滚复用和动态列大小写重复的实体职责测试，已登记到生产符号追溯表。

### Open Actionable

- 无。

### Blocked Approval

- 无。

### Blocked External

- 无。

### Not Applicable

- 本轮未改变共享热路径、公共 API、包图或性能算法；无关的 500K/1M 性能矩阵不重跑。

### Accepted Limitations

- 非 Entity Layout Provider 的能力边界和计划中的后续功能保持原记录。

### Verified Boundaries

- ClosedXML 列表导入重复值、唯一跟踪资源上限、失败行回滚和完整错误坐标已由职责测试覆盖。

### Deferred

- 公式 DSL、命名锚点、分页/分组汇总、契约测试包、分析器、Source Generator 和大报表策略层继续按计划延期。

### Current Gate

- `IMPLEMENTATION_STATUS`: `PASS`
- `TEST_STATUS`: `PASS`
- `REVIEW_STATUS`: `PASS`
- `RELEASE_STATUS`: `PASS`
- `OPEN_ACTIONABLE`: `0`

## TODO Continuation 3

本轮继续实现 TODO 列表中的实体导入校验扩展：显式动态列组现在可以携带内置 `ValidationRules`，实体导入新增 `ValidationFailureMode` 选择。旧的单参数 `ExcelEntityImportOptions` 构造函数和默认首错停止行为保持不变；`Continue` 只扩大当前行的错误收集范围，不改变跨行资源限制、取消、异常观察和流所有权契约。

交付记录：

- `ExcelDynamicColumnDefinition.ValidationRules` 已沿请求克隆、布局克隆、Mapping Plan 和 NPOI/ClosedXML 执行链传递；显式动态组的 `range` 校验会返回动态列坐标和结构化错误。
- `ExcelEntityImportOptions(ExcelResourceLimits, ExcelValidationFailureMode)` 与 `ValidationFailureMode` 已登记到 API approval；双 TFM API snapshot 只新增计划中的 3 项成员，compare `removed` 为空。
- `api-todo-final/capture`、`api-todo-final/compare` 保存本轮双 TFM 快照和比较证据；候选程序集与 8 个生产包的身份哈希已同步到 `build/api-snapshot-baseline.json`。

本轮验证：

- NPOI 新增实体职责测试：`2/2`（net8.0）通过；ClosedXML 新增实体职责测试：`2/2`（net8.0）通过。
- NPOI Provider `Category!=Large`：`650/650`（net6.0、net8.0）通过；ClosedXML Provider：`178/178`（net6.0、net8.0）通过；Core/Common：`204/204`（net8.0）通过。
- `dotnet pack Bing.Offices.sln -c Release --no-restore`：成功；NuGet 包身份重新采集并通过 API compare。
- `git diff --check` 仅保留 `ProfileFixtures.xml` 的既有 CRLF 规范化提示，未新增空白错误。

TODO 状态（本轮后）：

### Completed

- 显式动态列组内置校验规则与导入失败模式已实现、测试并登记追溯。
- 双 TFM 公共 API 审批、快照、候选身份和包身份已同步。

### Open Actionable

- 无。

### Deferred

- Footer 公式 DSL、命名锚点、分页/分组汇总、契约测试包、Roslyn 分析器、Source Generator 和大型报表策略层继续按原计划延期。

## TODO Continuation 4

本轮继续实现模板实体布局的命名锚点能力。`CellNamed` 和 `ListRegionNamed` 使用工作簿命名范围解析运行时坐标，支持工作簿级名称与 Sheet 局部名称；局部名称优先。命名明细区域的行边界仍由实际明细和 Footer 决定，因此拒绝绝对 `End` 配置。Provider 通过 `EditorBrowsable(Never)` 的公开布局副本方法消费解析结果，不恢复生产程序集友元访问。

本轮验证：

- NPOI 命名锚点与校验矩阵：`6/6`（net6.0、net8.0）通过。
- ClosedXML 命名锚点与校验矩阵：`6/6`（net6.0、net8.0）通过。
- NPOI Provider `Category!=Large`：`655/655`（net6.0、net8.0）通过；ClosedXML Provider：`183/183`（net6.0、net8.0）通过。
- Abstractions、NPOI、ClosedXML Release 编译：0 警告、0 错误。
- 公共 API 合同：`11/11`（net6.0、net8.0）通过；API snapshot compare：net6.0/net8.0 通过；新增成员为 7 项，删除成员为 0。
- 文档代码围栏消费测试：`10/10` 通过，命名锚点示例可独立编译执行。
- 生产友元白名单门禁：`1/1`（net6.0、net8.0）通过；仅保留测试程序集友元。
- `git diff --check`：无新增空白错误；仅保留 `ProfileFixtures.xml` 存量 CRLF 提示。

TODO 状态（本轮后）：

### Completed

- 模板命名锚点固定单元格和列表区域已实现，NPOI/ClosedXML 行为和非法模板错误分类已覆盖。

### Deferred

- Footer 公式 DSL、分页/分组汇总、可复用 Provider 契约测试包、Roslyn 分析器、Source Generator、大型报表策略层和旧版自动迁移分析器继续延期。

### Next recommended order

1. 行号上下文和计算列 DSL：先明确导入错误坐标、导出计算值和缓存边界，再扩展 NPOI/ClosedXML 公共契约。
2. 分页小计、分组汇总和签字区：建立与 Footer、模板合并和动态明细边界兼容的布局模型。
3. 可复用 Entity Provider 契约测试包：把固定 Cell、命名锚点、动态组、Footer、取消和资源行为抽成第三方 Provider 可执行合同。
4. Roslyn 地址冲突分析器：在编译期发现重复锚点、固定地址重叠、不可写导入属性和动态 Key 冲突。
5. Source Generator 与旧版迁移分析器：只有启动反射成本或迁移规模有实测证据时再投入。
6. 大型报表 DOM/Streaming 策略层：保持显式选择，待资源探针证明需要后单独设计。

## TODO Continuation 5

本轮继续实现实体布局的“行号上下文和计算列 DSL”。`CalculatedColumn<TValue>` 为列表区域增加导出专用计算列，计算委托接收当前项目、本次导出的只读明细快照、零基明细索引、零基行列索引及对应的一基行列号。计算列支持默认排序、Before/After/物理列位置、表头/正文样式和数字格式；NPOI 与 ClosedXML 使用同一物理列规划语义。

行为边界：计算委托在每个导出单元格执行一次，异常由 Provider 包装为带 Sheet、行、列和 Key 的 `BingOfficesExportException`，`OperationCanceledException` 原样透传。导入保留计算列的物理位置但跳过实体属性绑定，计算委托不会在导入过程中执行；明细快照只覆盖本次导出调用，不跨请求缓存。

本轮验证：

- NPOI 计算列职责测试：`2/2`（net6.0、net8.0）通过；ClosedXML：`2/2`（net6.0、net8.0）通过。
- NPOI Provider `Category!=Large`：`657/657`（net6.0、net8.0）通过；ClosedXML Provider：`185/185`（net6.0、net8.0）通过。
- 文档代码围栏消费者：`10/10` 通过，计算列示例可独立编译执行。
- Abstractions、NPOI、ClosedXML Release 编译：0 错误；编译使用 `UseSharedCompilation=false` 避免并行构建共享输出锁定。

TODO 状态（本轮后）：

### Completed

- 实体列表导出计算列和行上下文已接入 NPOI、ClosedXML，覆盖物理位置、样式、数字格式、导入跳过和失败坐标。
- 新公共 API 已登记 API approval，双 TFM 快照和差异检查待本轮候选程序集重新采集后完成。

### Open Actionable

- 无。

### Blocked Approval

- 无；计算列公共 API 已由本轮用户继续实现指令授权。

### Blocked External

- 无。

### Not Applicable

- 本轮未改变公式计算引擎、共享热路径、包版本或资源算法；不运行无关的 500K/1M 性能矩阵。

### Accepted Limitations

- 计算列只计算并写出值，不提供公式 DSL；导入不反向计算，也不将计算结果绑定回实体。
- MiniExcel、SpreadCheetah、ExcelDataReader 不声明 Entity Layout 计算列能力。

### Verified Boundaries

- 计算委托的快照、调用次数、行列坐标和异常属性已由双 Provider 真实工作簿测试验证。

### Deferred

- Footer 公式 DSL、分页/分组汇总、可复用 Provider 契约测试包、Roslyn 分析器、Source Generator、大型报表策略层和旧版自动迁移分析器继续按计划延期。

### Next recommended order

1. 分页小计、分组汇总和签字区：复用 Footer、计算列和模板合并边界，先明确跨页 marker 与导入语义。
2. 可复用 Entity Provider 契约测试包：抽出固定 Cell、命名锚点、动态组、计算列、Footer、取消和资源行为合同。
3. Roslyn 地址冲突分析器：在编译期发现重复锚点、固定地址重叠、不可写导入属性、动态 Key 和计算列冲突。
4. Source Generator 与旧版迁移分析器：有启动反射成本或迁移规模实测证据后再投入。
5. 大型报表 DOM/Streaming 策略层：保持显式选择，待资源探针证明需要后单独设计。

## TODO Continuation 6

本轮继续实现列表布局的连续分组小计能力。`GroupSubtotal` 按输入明细中相邻键的变化切分分组，复用现有 Footer 描述写出每组小计，并继续写出列表级最终 Footer；不自动排序，也不合并非连续的同名分组。导入时识别并跳过分组小计的间隔和 Footer 区域，避免把聚合行转换为明细。分组 Footer 与最终 Footer 的 marker 必须不同，运行时会在写入目标流前完成行列占用和冲突预检。

本轮交付：

- Abstractions 增加 `GroupSubtotal<TKey>`，Provider 通过 `EditorBrowsable(Never)` 的公开 `IExcelEntityGroupSubtotal` 描述消费，无生产程序集友元依赖。
- NPOI 的 XLS/HSSF、XLSX/XSSF 和 ClosedXML XLSX 均支持连续分组小计、组内聚合、最终 Footer、模板布局和导入跳过；同步/异步共用同一布局语义。
- 文档、API approval、生产符号追溯已登记分组小计的相邻顺序约束、marker 规则和后续迁移边界。

本轮验证：

- NPOI Provider `Category!=Large`：`658/658`（net6.0、net8.0）通过；ClosedXML Provider：`186/186`（net6.0、net8.0）通过。
- NPOI 与 ClosedXML 分组小计真实工作簿往返职责测试各 `1/1`（net8.0）通过，验证小计行、最终合计和导入明细隔离。
- 文档代码围栏筛选：`2/2`（net8.0）通过；新增分组小计示例可独立消费。
- `dotnet pack Bing.Offices.sln -c Release --no-restore`：成功；仅保留既有 `Microsoft.Bcl.Memory` net6.0 警告。
- API snapshot capture/compare：net6.0、net8.0 均通过；仅新增本轮审批的 `GroupSubtotal<TKey>` 及隐藏 Provider SPI 成员，`removed` 为空。
- 公共 API/友元门禁：`11/11`（net6.0、net8.0）通过；计算列与分组小计 SPI 类型均有明确分类并标记 `EditorBrowsable(Never)`，生产程序集无生产友元。
- NuGet 包消费者：`Bing.Offices.ThirdPartyProvider.Consumer` net6.0、net8.0 均通过，输出 `third-party-public-only-provider-ok`；使用离线依赖缓存和本轮生成包验证分组小计公开 API。

TODO 状态（本轮后）：

### Completed

- 连续分组小计、分组 Footer 聚合、导入跳过和运行时占用预检已实现并由 NPOI/ClosedXML 双 Provider 验证。
- 本轮新增公共成员已完成 API approval、双 TFM 快照、候选身份和包身份同步。

### Open Actionable

- 无。

### Deferred

- 跨页分页小计、签字区/尾部命名锚点、可复用 Provider 契约测试包、Roslyn 地址冲突分析器、Source Generator、旧版迁移分析器和显式 DOM/Streaming 策略层继续延期。

### Next recommended order

1. 分页小计与签字区：在分组 Footer 语义稳定后定义跨页 marker、分页边界和模板合并规则。
2. 可复用 Entity Provider 契约测试包：固化固定 Cell、命名锚点、动态组、计算列、Footer、分组小计、取消和资源行为合同。
3. Roslyn 地址冲突分析器：编译期检查锚点、固定地址、动态 Key、计算列和 Footer marker 冲突。
4. Source Generator 与旧版迁移分析器：先用启动反射和迁移规模数据决定投入范围。
5. 大型报表 DOM/Streaming 策略层：保持显式选择，待资源探针证明需要后单独设计。

### Current Gate

- `IMPLEMENTATION_STATUS`: `PASS`
- `TEST_STATUS`: `PASS`
- `REVIEW_STATUS`: `PASS`
- `RELEASE_STATUS`: `PASS`
- `OPEN_ACTIONABLE`: `0`
- 不自动 commit、push 或发布。

## TODO Continuation 7

本轮按推荐顺序补齐 Provider 合同层的连续分组小计覆盖。新增独立的 `EntityGroupSubtotals` 场景档案和跨 Provider 合同测试，使用公开 `ExcelEntity.Layout`、`GroupSubtotal` 与 `Footer` 完成真实工作簿导出/导入；NPOI、ClosedXML 验证小计和最终合计行不会被导入为明细，MiniExcel 明确标记为不适用。该轮不新增生产 API，也不改变已有 Provider 行为。

本轮验证：

- `EntityGroupSubtotal` 定向合同：`3/3`（net6.0、net8.0）通过。
- Provider Contract `Category!=Large`：`279/279`（net6.0、net8.0）通过。
- 档案完整性测试覆盖新增场景，profile 门禁定向测试 `5/5`（net6.0、net8.0）通过，三项 Provider 预期均有独立配置。

TODO 状态（本轮后）：

### Completed

- Provider 合同层已覆盖固定实体、动态列、计算列、Footer 和连续分组小计；第三方 Provider 可据此复用公开行为合同。

### Open Actionable

- 无。

### Deferred

- 分页小计、分页边界和签字区仍需先定义跨页 marker、模板合并、导入截断和页设置的公共模型；本轮不引入未经批准的分页 API。
- Roslyn 地址冲突分析器、Source Generator、旧版迁移分析器和显式 DOM/Streaming 策略层继续延期。

### Next recommended order

1. 分页小计与签字区：先完成公共布局模型和 NPOI/ClosedXML 行为矩阵，再实现 API 与真实文件测试。
2. 将现有 Provider 合同抽成可引用的测试包或测试模板，保留独立预期档案，避免 Provider 直接读取生产能力生成结果。
3. Roslyn 地址冲突分析器：复用布局校验规则，检查锚点、固定地址、动态 Key、计算列和 Footer marker 冲突。
4. Source Generator 与旧版迁移分析器：以反射启动成本和迁移规模实测作为投入门槛。
5. 大型报表 DOM/Streaming 策略层：保持显式选择，待资源探针证明需要后单独设计。

### Current Gate

- `IMPLEMENTATION_STATUS`: `PASS`
- `TEST_STATUS`: `PASS`
- `REVIEW_STATUS`: `PASS`
- `RELEASE_STATUS`: `PASS`
- `OPEN_ACTIONABLE`: `0`
- 未自动 commit、push 或发布。

## TODO Continuation 8

本轮按分页相关 TODO 的最小可交付顺序补充实体列表的水平分页符能力。新增 `PageBreak(int rowsPerPage)`，每页只统计明细行，Provider 在导出工作表中写入分页元数据；不插入空白行，不改变 Footer、marker 或导入明细边界。为避免分页符与连续分组小计的行占用语义冲突，两个配置在布局构建阶段互斥。`ExcelEntityListRegion<TEntity>.PageBreakRows` 作为 `EditorBrowsable(Never)` 的公开布局描述属性，供 Provider 在无生产友元依赖的条件下读取。

本轮交付：

- NPOI XLS/HSSF、XLSX/XSSF 和 ClosedXML XLSX 均写入水平分页符，表头不计入 `rowsPerPage`；默认 Footer、模板样式和流所有权语义保持不变。
- 包消费者增加分页布局的公开 API 导出验证；文档新增分页示例，Provider 能力矩阵和生产符号追溯已同步。
- API approval、双 TFM snapshot、候选身份和 `build/api-snapshot-baseline.json` 已同步；新增成员仅为 `PageBreak` 和隐藏 `PageBreakRows`，删除成员为零。

本轮验证：

- NPOI PageBreak 直接职责：`3/3`（net6.0、net8.0）通过，覆盖 XSSF、HSSF 模板和非法组合。
- ClosedXML PageBreak 直接职责：`2/2`（net6.0、net8.0）通过，覆盖 XLSX 分页元数据和非法组合。
- NPOI Provider `Category!=Large`：`661/661`（net6.0、net8.0）通过；ClosedXML Provider：`188/188`（net6.0、net8.0）通过。
- 公共 API/友元门禁：`11/11`（net6.0、net8.0）通过；API snapshot compare 双 TFM 通过，diff 的 added 仅为本轮两项成员，removed 为空。
- 文档代码围栏：`10/10`（net8.0）通过；包消费者 net6.0、net8.0 编译并运行，均输出 `third-party-public-only-provider-ok`。
- `dotnet pack Bing.Offices.sln -c Release --no-restore --no-build` 成功生成本轮包；离线消费者还原复用既有依赖缓存，外部 NuGet 源不可用仅产生漏洞数据告警。

TODO 状态（本轮后）：

### Completed

- 实体列表水平分页符已实现并由 NPOI/ClosedXML 双 Provider、双 TFM 和包消费者验证。
- 本轮新增成员的中文 XML 注释、API approval、双 TFM 快照、基线、文档与符号追溯已完成。

### Open Actionable

- 无。

### Deferred

- 分页小计、跨页签字区和尾部命名锚点仍未实现。它们需要单独定义跨页 marker、模板合并占用、导入截断和分页设置交互，不能由当前仅写分页元数据的 API 隐式推导。
- 可复用 Entity Provider 契约测试包、Roslyn 地址冲突分析器、Source Generator、旧版迁移分析器和显式 DOM/Streaming 策略层继续延期。

### Next recommended order

1. 定义分页小计与签字区的跨页布局模型和导入边界，再实现 NPOI/ClosedXML 行为矩阵。
2. 将现有固定 Cell、命名锚点、动态组、计算列、Footer、分组小计、分页和资源行为抽成可引用 Provider 合同包。
3. 复用布局预检规则实现 Roslyn 地址冲突分析器，覆盖锚点、固定地址、动态 Key、计算列和 Footer marker。
4. 以反射启动成本和迁移规模实测决定 Source Generator 与旧版迁移分析器的投入。
5. 待资源探针证明需要后设计显式 DOM/Streaming 策略层，保持 Provider 选择可预测。

### Current Gate

- `IMPLEMENTATION_STATUS`: `PASS`
- `TEST_STATUS`: `PASS`
- `REVIEW_STATUS`: `PASS`
- `RELEASE_STATUS`: `PASS`
- `OPEN_ACTIONABLE`: `0`
- 未自动 commit、push 或发布。
