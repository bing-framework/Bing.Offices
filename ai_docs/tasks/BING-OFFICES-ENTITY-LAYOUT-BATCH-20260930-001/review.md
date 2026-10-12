<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-ENTITY-LAYOUT-BATCH-20260930-001
AI_REVIEWED_AT: 2026-10-03T21:32:09.7881841+08:00

# 当前计划差距与后续 TODO

## 范围和结论

本次核对 ERP 字典模板计划、20260929 能力差距报告、后续六个分批任务及当前源码、测试、工作区/暂存区差异。未修改生产代码、测试、计划或执行记录，未重跑全量测试。测试数量引用历史 execution，不作为本次新鲜验证结果。

按上一份差距报告的九个源码能力项计，前六项已接入实现，剩余三个路线图能力：分组分页组合、重复页命名锚点、公共 PivotTable。另有一项商业部署环境验收。六项已有实现不等于六项已经独立完整验收；最新 Entity 批次仍有下列三组收口 TODO。计数按能力/问题组，不代表工作量百分比。

## 计划对照

| 项目 | 当前结论 | 证据与边界 |
| --- | --- | --- |
| ERP 自动字典、名称列表、Entity 原生校验 | 核心已实现 | LookupList writer、两个 WorkbookValidationPipeline、Entity executor；原 FIX-010～012 已有后续修复记录，NPOI 显式列表和取消传递代码已更新 |
| 填写区域保护、两级级联 | 已实现 | 双 Provider writer、构建入口和职责测试；级联仅同一行两级，导入不计算级联公式 |
| 分批跨行唯一性 | 已实现 | ExcelDataReader MaterializeTypedRow 使用 UniqueTracker，失败行回滚 |
| 分批 Workbook 原生校验 | 已实现，有格式边界 | ExcelXlsxWorkbookValidation 接入批次物化；仅 XLSX，XLS/XLSB 保持 Unsupported |
| 分批动态列 | 已实现 | BindHeaders/MaterializeTypedRow 动态分支、动态字典设置和职责测试 |
| 分批跨 Sheet 关系 | 已实现，有内存边界 | ExcelDataReaderBatchRelationCoordinator 保存成功实体，在所有 Sheet 读取后绑定；回调时关系可能尚未完成 |
| 分批 Failure Workbook | 已实现，有 Provider 边界 | NPOI/ClosedXML 完整 DOM 导入与失败产物提交后投影批次；ExcelDataReader 仍拒绝 |
| Entity Layout 分批 | PARTIAL（实现已接入，验收待补） | 两适配器、DI 和 ExcelEntityBatchProjection 已接入普通/模板、同步/异步；参见 FIX-001～003 |
| 分组与分页组合 | DEFERRED，未实现 | ExcelEntityListRegionBuilder.cs:282、310、331 明确互斥；布局 Build 也校验拒绝 |
| 重复页命名锚点 | DEFERRED，未实现 | 仅 Cell/ListRegion AnchorName 和最终 FooterAnchorName；没有每页小计/签字区命名入口 |
| 公共 PivotTable | DEFERRED，未实现 | Abstractions 无模型；ClosedXmlTemplatePreflight.cs:43 拒绝模板 PivotTable 部件 |
| 商业公式重算与渲染 | BLOCKED_EXTERNAL | 需要许可证、字体和实际部署函数集；现有本地测试不能替代部署验收 |

## FIX-001：补齐 Entity 批次组合与失败边界直测

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：计划“实现与测试”第 3 项要求动态列、关系和空列表，但两个 EntityLayoutBatchImportTest 仅各有四个测试方法（NPOI Theory 展开六例）；未覆盖这些组合，也未直接覆盖区域失败行、回调中取消/异常、普通异步和同步模板入口。
- 证据：tests/Bing.Offices.Npoi.Tests/EntityLayoutBatchImportTest.cs；tests/Bing.Offices.ClosedXml.Tests/EntityLayoutBatchImportTest.cs；src/provider-shared/ExcelEntityBatchProjection.cs:122。
- 影响：完整导入器通过测试不能证明新投影器的区域匹配、成功项/SourceRows 配对、错误归属及摘要计数正确。
- 修复目标和要求：用真实工作簿从新入口验证动态字典和关系结果、空列表/无区域、失败行错误与计数；验证 Footer/分页后原坐标和回调取消/异常停止交付；断言完整数据而非仅条数。
- 验证方式：NPOI XLS/XLSX、ClosedXML XLSX，net6.0/net8.0 职责测试；只在相关修改后扩大验证。

## FIX-002：补齐 Entity 批次包消费者

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：双 TFM Consumer 中存在 Workbook 分批调用，但没有 IExcelEntityBatchImporter 调用；最新 execution 记录 API 合同，没有新增 Entity SPI 的打包运行证据。
- 证据：tests/Bing.Offices.Consumer.Net6/Program.cs；tests/Bing.Offices.Consumer.Net8/Program.cs；本任务 execution.md。
- 影响：DI 注册、共享源码打包和新增入口的外部消费尚未闭环。仓库公共 API 变化规则要求对应 Consumer 验证。
- 修复目标和要求：从当前生成的 NuGet 包解析两个 Provider 的 Entity 批次入口，运行实体布局导入并核对批次内容与诊断摘要；记录包身份，不沿用新增入口前的包结果。
- 验证方式：net6.0/net8.0 Consumer 均成功；双 TFM API 检查无未审批删除或改签名。

## FIX-003：修正完成声明和验证证据边界

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：docs/excel/09-providers.md:73 声称空列表已由职责测试覆盖，但当前测试没有空列表；execution 声称全量收口，实际两次解决方案运行分别出现资源探针和临时文件集合失败，独立重跑通过不等于同一次全量通过。
- 证据：本任务 execution.md“测试证据”；两个 EntityLayoutBatchImportTest；AsyncPipelineTest.cs:190、204 对共享临时目录进行前后集合比较。
- 影响：完成度被高估；“前次残留导致失败”没有足够证据，测试期间其他进程创建/删除文件也可产生同样现象。
- 修复目标和要求：补齐测试后更新文档与符号映射；准确记录单项通过与全量未一次通过的区别。若需完整门禁，先隔离共享临时路径或采用覆盖项目与 TFM 的串行策略，再执行一次必要确认，不以连续全量重试代替定位。
- 验证方式：文档声明可追溯到具体方法；最终候选具有对应门禁结果，未通过项明确说明。

## 后续能力 TODO（DEFERRED，不属于当前回归缺陷）

1. 分组分页组合：先统一物理行规划；定义组跨页、页小计/组小计顺序、最终 Footer 公式区段、导入跳过规则和行数上限。
2. 重复页锚点：基于上述规划提供每页小计/签字区稳定名称，定义名称冲突、局部名称、空明细与模板兼容。
3. 公共 PivotTable：独立设计源数据范围、行列字段、聚合和过滤；逐 Provider 定义创建、模板保留、缓存刷新和 Unsupported。

建议顺序：FIX-001 → FIX-002 → FIX-003 → 分组分页组合 → 重复页锚点 → PivotTable。每阶段先明确计划和验收再实施。

## 保留限制与可选扩展

- ACCEPTED_LIMITATION：NPOI/ClosedXML 批次是完整 DOM 结果投影；不能宣称低内存前向读取。
- ACCEPTED_LIMITATION：ExcelDataReader 关系会保留实体以完成最终绑定；启用关系后的内存特征不同于无关系逐批消费。
- ACCEPTED_LIMITATION：Entity Failure Workbook 本轮计划明确排除；可后续单独设计，不计入前三项路线图欠账。
- ACCEPTED_LIMITATION：两级级联导入不计算公式；递归级联、跨行依赖可独立扩展。
- BLOCKED_EXTERNAL：AsposeCells 许可证/字体/部署环境验收一项。
- BLOCKED_APPROVAL：本次没有确认的新审批阻塞；未来 API 设计按成员审批流程执行。
- NOT_APPLICABLE：本次只读差距检查不重跑大规模性能矩阵。
- VERIFIED_BOUNDARY：本次没有新增容量探针结论。

## 最终状态

IMPLEMENTATION_STATUS: 核心入口已实现，尚未完成组合验收
TEST_STATUS: PARTIAL
RELEASE_STATUS: NOT_VERIFIED
OPEN_ACTIONABLE: 3 组验收收口问题
DEFERRED: 3 项路线图源码能力
BLOCKED_EXTERNAL: 1 项部署验收
Next Action: 按上述顺序补齐 Entity 批次直接测试；本次仅审查，不自动修复。
