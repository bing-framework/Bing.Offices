<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: BING-OFFICES-PIVOT-TEMPLATE-UPDATE-20261005-001
AI_REVIEWED_AT: 2026-10-11T00:20:00+08:00

# 独立代码审查

## 结论

Round 1 的 `FIX-001`、`FIX-002` 已关闭，当前没有 `OPEN_ACTIONABLE`。实现、职责测试和最终候选包证据已覆盖计划内的服务端模板透视保留与更新路径。Excel、WPS、LibreOffice 的首次打开、刷新、保存重开不属于用户当前代码验收范围，本审查继续将其记为 `ACCEPTED_LIMITATION`，不作为缺陷或发布阻塞。

```text
IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
PERFORMANCE_EVIDENCE_STATUS: NOT_APPLICABLE
RESOURCE_EVIDENCE_STATUS: PASS
EXTERNAL_GATE_STATUS: PASS
RELEASE_STATUS: BLOCKED_APPROVAL
GOAL_STATUS: STOPPED_BLOCKED
```

当前 Finding 分类：

- `OPEN_ACTIONABLE`：0。
- `BLOCKED_APPROVAL`：1，正式双 TFM API 门禁包含 19 个不属于本任务已批准七成员范围的新增成员。
- `BLOCKED_EXTERNAL`：0；本次已完成独立只读 Review，旧记录中的 Review 路由不可用不再构成当前阻塞。
- `ACCEPTED_LIMITATION`：3，客户端刷新/呈现不在本轮范围；关联透视图、切片器、时间线更新明确提前拒绝；NPOI XLS/HSSF 及其它 Provider 不声明模板透视能力。

## 计划验收矩阵

| 阶段 | 状态 | 审查结论 |
| --- | --- | --- |
| PTPL-001 清单与依赖图 | PASS | `ExcelPivotTemplateParts`、`ExcelPivotTemplateManifest` 和关系清单进入两个 Provider 的模板预检；非标准路径、损坏/外部关系、资源预算、取消和流所有权有直接测试。 |
| PTPL-002 公共快照与执行预检 | PASS | 七个已批准公共成员、深快照、精确更新定位及其它 Provider 提前拒绝已接入；显式保留加新建且无更新请求现在进入相同 Snapshot、Plan 和 Bridge 预检。 |
| PTPL-003 保留/更新输出 | PASS | 原样保留、全部/部分更新、共享缓存拆分、来源扩缩、布局保留、目标冲突、组合更新与新建均进入真实主链；无更新的新建组合会避让原部件路径和缓存标识后恢复模板关系。 |
| PTPL-004 业务合同 | PASS（用户范围内） | 采购/仓储、五聚合、筛选、数字格式、请求复用、失败重试和未知依赖边界均有 NPOI/ClosedXML、同步/异步、双 TFM 文件级合同；客户端行为按用户要求排除。 |
| PTPL-005 交付 | PARTIAL（审批阻塞） | 文档、追溯、Release 构建、职责测试、最终候选身份和普通/第三方双 TFM 包消费者已有证据；仅正式 API 门禁受 19 个其它任务成员审批阻塞。 |

## Findings

### FIX-001

- 状态：`CLOSED`。
- 原严重程度：MEDIUM。
- 复核结论：`ExcelPivotTemplatePreserver.IsPreserveTemplate` 现在只要求显式保留、XLSX 模板，不再要求存在更新请求；`ExcelPivotTemplateBridge.IsRequired/Create` 因此会为 `PreserveTemplatePivotTables() + PivotTable(...)` 建立 Snapshot 和 Plan。
- 实际执行路径：两个 Provider 均先以移除原透视关系的模板生成业务工作簿和新透视，再由 `ExcelPivotTemplateBridge.Merge` 的无更新分支调用 `RenameGeneratedParts`、`CreateCacheIdMap`、`RemapGeneratedCacheIds`，最后复制原部件并合并 Content Types、工作簿缓存关系和工作表透视关系。NPOI、ClosedXML 的同步和异步入口都调用同一桥接链。
- 测试证据：`TemplatePivotPreserveExportTest.PreserveAndCreate_ShouldKeepExistingPivotAndCreateNewPivot` 覆盖两个 Provider 的同步/异步路径，逐字节核对原透视部件，断言两个缓存定义及唯一缓存标识、原透视和新透视、临时 Sheet 清理、模板字节、位置和所有权。`PreserveAndCreate_ShouldRejectNameOrTargetConflictBeforeWriting` 覆盖名称与目标冲突的两个 Provider 同步/异步预检。
- 回归边界：MiniExcel、SpreadCheetah 和 XLS/HSSF 的拒绝边界未放宽。

### FIX-002

- 状态：`CLOSED`。
- 原严重程度：MEDIUM。
- 复核结论：最终候选已在 FIX-001 后重新生成并绑定当前源码、程序集和包；旧候选不再作为本轮完成证据。
- 候选证据：`artifacts/api-candidate-pivottpl-reviewfix-20261010/api-candidate.json` 使用 `logical-v3`，源码清单 SHA-256 为 `739E38686F2CAD0023D2FD36546FEE5F3A9E524423EFCE6E338A7355381795B9`，记录 14 个程序集和八个 `2.0.0-pivottpl-20261010` nupkg 的身份哈希。`api-diff.json` 没有候选内部快照差异。
- 包证据：执行记录核对包内 14 个程序集与候选程序集逐字节一致、消费者隔离缓存中的八个 nupkg 与候选 feed 逐字节一致。普通消费者在 .NET 6、.NET 8 均输出 `package-consumer-ok`；第三方公开 API 消费者在 net6.0、net8.0 均输出 `third-party-public-only-provider-ok`。消费者合同通过公开 Workbook API 对 NPOI、ClosedXML 执行同步/异步模板更新，并检查模板流所有权、透视名称、缓存字段、完整缓存记录和 `refreshOnLoad`。
- API 边界：正式 baseline 未被候选覆盖；双 TFM 仍为 `+27/-0`，本任务七个成员已有 `api-approval.md`，其它成员继续独立审批。

## 非代码阻塞

### API 成员审批

- 状态：`BLOCKED_APPROVAL`。
- 当前双 TFM 候选相对正式基线均为 `+27/-0`。其中本任务七个成员已有 `api-approval.md`，其余 19 个成员未获批准或缺少批准记录，另有一个成员属于其它已批准任务。
- 该阻塞不能转换为本任务代码缺陷，也不得由 Reviewer 更新正式 baseline。维护者完成其它成员审批或在隔离候选中只保留已批准差异后，才能通过正式 API snapshot 门禁。

## 已接受边界

- `ACCEPTED_LIMITATION`：Provider 只重建缓存并设置 `refreshOnLoad=true`，不执行 Excel/WPS/LibreOffice 刷新、公式计算或呈现；用户已明确排除客户端验收，文档没有把结构证据外推为客户端成功。
- `ACCEPTED_LIMITATION`：更新关联透视图、透视切片器或时间线时在业务数据枚举前结构化拒绝；显式保留继续复制原部件和关系。
- `ACCEPTED_LIMITATION`：模板透视保留/更新仅声明 NPOI XLSX 与 ClosedXML XLSX；NPOI XLS/HSSF、MiniExcel、SpreadCheetah 保持提前拒绝。

## 验证证据

Change Impact Analysis：

```text
ChangedProjects: Abstractions, NPOI, ClosedXML, ProviderContract, Consumer, Docs
ChangedPublicContracts: 7 个本任务已批准新增成员
ChangedRuntimePaths: Workbook 模板快照、OPC 清单、透视保留/更新桥接
ChangedProviders: NPOI XLSX, ClosedXML XLSX
ChangedTFMs: net6.0, net8.0
ChangedBuildPackaging: 8 个候选 NuGet 包
RiskLevel: HIGH
```

- Round 1 执行证据：`PreserveAndCreate_` 双 TFM 各 `12/12`；完整 `FullyQualifiedName~TemplatePivot` 双 TFM 各 `168/168`；NPOI、ClosedXML 与 ProviderContract 的 Release 双 TFM 编译通过，0 错误。
- Reviewer 已逐项读取 FIX-001 的条件、Plan、Bridge 合并算法、两个 Provider 调用点及新增成功/冲突断言；没有发现 Stub、旁路或只测模型而未进入 Provider 的情况。
- Reviewer 的附加并行 `--no-build` 复验使用了较早的 156 项二进制，两个 TFM 同时扫描同一系统临时目录时，资源清理快照用例互相观察到对方的临时文件；内容用例通过。随后的重编译因共享工作区中另一进程占用 `Abstractions/obj` 返回 `MSB3491`。这组并行结果不替代修复轮已记录的串行新二进制 `168/168`，也未形成生产代码 Finding。
- 候选身份与包消费者证据均晚于 FIX-001 源码修改；候选保留 dirty-worktree 的显式源码清单，未只绑定 Git HEAD。
- `git diff --check` 对本轮审查范围无空白错误；Reviewer 只更新本任务 `review.md`，未修改生产代码、测试、计划、执行记录、API baseline、CI 或包。
- 当前 Git HEAD：`1d13019ce61985db86fb7994e99b9870cad1767f`；工作区包含大量其它任务的未提交与未跟踪变更，因此正式候选必须使用显式源码和程序集身份，不能只绑定 HEAD。

## 下一步

本任务没有可进入自动修复的 Finding。发布前仅需处理正式 API 门禁中 19 个非本任务成员的审批或隔离候选；客户端验收继续遵循用户已确认的排除范围。
