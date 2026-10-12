<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ENTITY-INDEPENDENT-PAGE-ANCHORS-20261005-001
AI_EXECUTION_FINISHED_AT: 2026-10-05T02:26:57.363346+00:00

# 实施执行报告

## 执行结论

COMPLETED。F-001 计划内源码、职责及常规门禁已经完成；最新修复候选的隔离包消费者和 API 身份比较已通过，仓库基线已精确同步该候选且公共成员无额外差异。review-round1-final.md 独立复核为 PASS，FIX-001～014 均已关闭。以下“修复中”“待执行”记录属于相应历史时点；当前结论以阶段矩阵和最终候选收口节为准。本结论不代表完整 Excel Goal 完成。

## 计划执行情况

| 阶段 | 当前情况 |
| --- | --- |
| P1 模型、布局、快照 | 已实施，模型 net6/net8 各 9/9，包括描述传递、失败状态、名称/marker 和连续求和边界 |
| P2 行规划与导入 | 中间页、末页、空页、空白高度及缺失/重复/错位 marker 已直测；末页错误行号修正已由 ClosedXML 原失败用例通过验证 |
| P3 Provider 与模板 | 三格式名称、两种策略、非法引用、作用域、固定字段、多区域及格式边界均已验证；完整常规回归通过 |
| P4 业务与资源 | 主体职责 NPOI 双 TFM 各 84/84、ClosedXML 各 42/42；新增 Composition 为 NPOI 各 10/10、ClosedXML 各 5/5。最终隔离包消费者四次运行全部成功，资源及同步/异步完整往返通过 |
| P5 API 与交付 | 最新 Release 构建退出 0，0 错误、4 个存量警告；27 组常规门禁 4,805/4,805。API 原始审批各 TFM 12 新增、0 删除，最终内部修复无额外成员差异；已同步通过比较的最终候选身份。独立最终复核 PASS |

## 已完成事项

- 公共 Builder 新描述传入 ListRegion/ListRegionNamed，区域快照及 WithResolvedStart 保留新属性；新增策略枚举已修 BOM。
- 每页 Footer 行段在组小计、分页小计之后、分页符和下一页页眉之前写入，末页和空明细均写入。
- 导入按 marker、GapRows、高度和页容量跳过每页 Footer；无最终 Footer 时补齐末页尾部的空白占用行；保留原始坐标及既有校验路径。
- 两个 Provider 对所有独立系列预检作用域与单格引用后，才清理允许移除的旧名称；默认 Reject，RemoveObsolete 不处理无关名称，旧分页小计约定保留。
- 签字区纳入格式边界、相对合并、运行时区域及固定字段占用检查；明细不能占用签字 marker 文本。
- NPOI 真实 XLS/XLSX 组页组合证明最终 DetailSum 仅含 C2:C3、C9:C10、C16:C16，排除签字区 1001/1002 数值。
- ClosedXML 同步/异步组页组合已断言全部单元格及最终 DetailSum；签字区接收当前页 2/2/1 条明细。
- 导出在目标序列化前读取实际单元格并复用导入扫描校验，不重新执行工厂；固定文本/工厂产生重复 marker 时拒绝，非扫描列允许同文本。

## Review 修复记录 / Round 1

- FIX-001：模型传递和布局校验已实施，职责仍在补全，不登记最终完成。
- FIX-002：Provider、共享规划与导入已接入；当前 NPOI 28 项职责通过，尚待双 TFM、ClosedXML 和资源矩阵。
- FIX-003：新增职责及真实产物正在验证；API、包消费者及最终门禁仍待执行。
- FIX-004：枚举已按 UTF-8 BOM/LF 修正，最终改动文件严格编码检查待收口执行。

### 当前进度审查 FIX-005～007

- 依据 review-progress-20261005.md，保持历史报告不变；当前 runtime 仍为原 review-fix Round 1，不重复注册或伪造新 Round。
- FIX-005：已修复并职责验证。共享 ValidateWrittenMarkers 读取渲染结果，检查首页源区域与全部辅助行段。三格式覆盖固定文本/工厂、规范化、原始坐标、旧输出/模板保全、观察一次及非扫描列完整往返。
- FIX-006：部分完成。模型已双 TFM 9/9；NPOI 84、ClosedXML 42 项当前职责已执行。修改后的 NPOI XLS 损坏夹具已实际用于 net6，net8 将由后续完整回归覆盖；其他组合与共享回归门禁待完成。
- FIX-007：进行中。三个隔离包消费者已增加逐页内容、名称坐标、SPI 描述及异步完整往返；尚未执行新包。API 差异/快照、文档、追溯、最终门禁与独立复核待收口。

## 测试与命令事实

- model-net8 首轮：6/7，失败为测试夹具在保留 A1 再配置 Cell。首次命令中的 PowerShell logger 分号未引用，TRX 写到项目 TestResults；后续命令已正确引用并指定任务目录。保留失败日志，不把命令错误记成生产缺陷。
- npoi-net8-01：构建失败，新测试缺少 Exports/Imports using；已修正。
- npoi-net8-02：20/20。
- npoi-net8-03：24/28；四个打印标题组合用例的名称数量未计入 Excel 自动生成名称，已修正断言。
- npoi-net8-04：28/28，退出 0。上述命令和对应 TRX/日志位于 artifacts/duties；不是前序 SOURCE 候选测试。
- 正在新增测试，因此这些数字只代表对应时点的已执行范围；未运行完整 Release、Category!=Large 或商业客户端验收。
- model-net8-02、model-net6-01：各 9/9，退出 0。
- npoi-net8-05：38/38；npoi-net8-06 新 marker 夹具用错公共属性名编译失败，修正后 -07 为 56/56。
- closedxml-net8-01：9/10，末页缺失 marker 行号偏一；共享扫描已修正。-02/-03 Boundary 夹具括号错误，-04 新异步用例缺 Task using；均已修正并保留日志。
- closedxml-net8-05：40/42，两项失败来自夹具错误地将每页 2 条明细写为期望 3。修正期望并增加全部单元格断言；-06 为 42/42，net6-01 为 42/42。
- npoi-net8-08：82/84，创建全局/局部同名夹具时先设名称导致 NPOI 拒绝；改为先设作用域。-09 为 84/84。
- npoi-net6-02：84/84，已更正损坏 marker 夹具真正用 XLS/HSSF 和 XLSX/XSSF，最大行列用单条数据和扩展 Footer 触发。没有用两条相同 XLSX 用例冒充双格式。

## 最新核查修复 FIX-008～010

- 依据 review-todo-latest-20261005.md 继续原 review-fix Round 1，不改写历史审查或重复注册。
- FIX-008：迁移示例源范围改为 A1:H5，与 A5 表头和第 6 行明细相邻；最终金额用明细数量乘单价聚合，避免把单价当金额列。示例实际输出 HSSF/XLS 与 XSSF/XLSX 模板、读回完整明细，并检查签字区和最终金额/名称；docs-fix-05 退出 0，两个围栏职责通过，全部 23 个 Markdown C# 围栏逐个编译并执行。
- 围栏收口首轮把旧 21 个数量误改为 22，而真实数量为 23；随后新片段顶层 using 不适用于围栏测试的语句上下文，已去掉重复 using。扩展 DTO 支持第一次条件过宽影响 dynamic-columns 片段，已收窄为有真实单据实例的独立页锚点示例。失败日志 docs-fix-01/02/04 保留，没有删除或弱化围栏验证。
- FIX-009：千页名称、长度边界、外部引用与动态/计算列组合职责正在补齐，尚未登记完成。
- FIX-010：8 个生产包和 EntityProviderContracts 包已成功生成；两个普通消费者与第三方消费者的 net6/net8 四次运行全部退出 0。双 TFM 实际成员差异仅 Abstractions 新增 12 条、删除 0，其它七程序集零差异，更新前基线和差异已保留；审批内临时候选身份比较正在执行，尚未更新仓库基线。常规回归和独立复核仍待执行。Release 首次构建暴露 BuildScript 默认编入独立 PivotClientProbe 源码及生成文件；沿既有工具排除规则新增 Compile Remove，一行配置调整后 release-build-02 退出 0。
- 消费者第一次失败来自将同单格的范围表达式与单格公式作字符串比较，已改为解析来源 Sheet、首末行列，仍严格验证实际单格。第三方示例还缺两个动态字典组，已补公共 DynamicColumnGroup、完整四列表头/数据及动态字典读回。两个失败实例发出未处理异常后仍未退出，已仅终止本次创建的 PID 70644/86508，保留失败日志/命令和 consumer-results-failed-01/02.json；其它实例未触碰。验证脚本 stdout/stderr 显式 UTF-8，修正父进程捕获中文时按本机代码页输出导致的解码错误。最终 consumer-results.json 记录 11 次 restore/build/run 命令全部退出 0。

## 第二次核查修复 FIX-011～012

- 依据 review-todo-20261005-round2.md 继续原 review-fix Round 1，runtime 的 active/taskId/mode/reviewRound 已读取确认，不重复注册。
- FIX-011：新增 NPOI Composition net8 首次 10 项中 4 项通过、6 项失败。千页名称与 255 字符边界两种格式通过；外部名称夹具未建立真实外部关联，动态/计算列夹具把 B3 起点的 Footer marker 当作 A 列读取，且计数断言未配置相应值工厂。Luna 仅负责两个 Composition 测试文件的夹具与 XML 注释修复，生产代码由主代理根据修正后的实际结果判断。
- npoi-composition-net8-02：5/10。修正相对坐标与实际聚合后，HSSF 外部名称拒绝场景通过；XSSF 的公开 LinkExternalWorkbook 实际抛 NPOI.Util.RuntimeException（Not Implemented - see bug #57184），不因 XML 文档声明方法就认为可用。其余四项遇读取器关闭输入流后无法 SetLength，夹具改为独立读取流和可扩展输出流。失败命令、TRX 和日志保留，继续修正真实 XSSF 外部关联夹具。
- FIX-011 已完成最小验证。NPOI net8-03 候选误用抽象 POIXMLFactory，编译失败；-04 发现新工作簿的 ExternalLinksTable 尚未初始化，并首次跑到生产空引用：原生校验取动态/计算列的 Property.Name。生产修复跳过不回写实体的计算列，为动态列使用 Key。XSSF 夹具经公开 XSSFFactory/ExternalLinksTable/CT 工作簿关系生成、序列化再读回，明确验证关联后再设置外部名称，不使用反射访问内部成员；-05 的内部 GetCTName 编译失败也保留，最终夹具未使用该方法。
- 两个 Provider 既有合同均为原生校验失败即跳过整行，不继续物化该行。新组合分别构造原生失败与配置 Unique 失败的副本，完整断言两类原始坐标及成功实体，没有删掉配置规则验收。ClosedXML -01 的非法外部名称文法与颜色期望错误修正后 -02 为 5/5；NPOI -06 为 10/10。net6 各为 10/10 和 5/5，artifacts/composition-results.json 四次命令全部退出 0。
- 因 NPOI 内部生产修复，最终交付候选需要重新构建/打包/消费者/API 身份验收；修复前批准快照保存在 artifacts/api-baseline-before-native-fix.json，新的隔离产物归入 artifacts/final-candidate，原包和消费者证据保留。
- FIX-012：隔离消费者 11 个 restore/build/run 命令、9 个 pack 命令均退出 0；api-compare-approved 退出 0，核对成员差异及候选身份后才更新 build/api-snapshot-baseline.json，更新前基线保存在 artifacts/api-baseline-before.json。常规门禁脚本已就绪，尚未执行，不把脚本当作验证结果。
- 本次核查 142 个改动 C# 文件，严格 UTF-8、BOM、LF 和单末尾换行全部通过；历史 encoding-csharp.json 的 NO_BOM 已通过仅增加共享名称辅助文件 BOM 修正，最终字节检查仍将在交付时登记。
- 前面的 FIX 修复说明与测试失败均为历史时点事实；最新状态以本节、阶段矩阵和后续真实命令终态为准。

## 最终候选收口 / FIX-012～014

- FIX-001～011 已完成并验证；FIX-012 的组合、常规门禁、文档、包和 API 证据已经闭环，最后独立审查仍单独记录，不用 Executor 结论替代 Reviewer。
- FIX-013：final-candidate/api-compare.command.json 于 2026-10-05T02:05:25Z 结束、退出 0。内部修复候选与已批准基线的 assemblies 完全一致；仅在该比较通过后更新 approvedAt/candidateIdentity，build/api-snapshot-baseline.json 现与已验证 api-approved-candidate.json 完全一致。原始新增成员差异和修复前基线均保留，不覆盖审批外变更。
- 当前 regular-results-final.json 对应 27 条退出 0 的命令和 27 份最终 TRX，4,805/4,805，失败和跳过均为 0。旧 Docs 的 18/19 失败 TRX 保留在目录，但不混入最终 27 组计数；修正是将计划批准的 Anchors 目录纳入 SourceLayoutTests 能力目录名单。
- 最终 Release 构建退出 0；9 个 pack 命令全部成功，包含八个生产包和 EntityProviderContracts。全新 final-candidate 本地 feed/cache 的消费者 11 条 restore/build/run 命令全部退出 0，普通和第三方 net6/net8 四次实际运行全部成功。
- 新增 Composition 四命令合计 30/30；主体职责与组合数字分别对应职责范围，不与常规门禁相加冒充不同测试数。原生规则失败跳过整行、动态 Unique 状态隔离与配置 Unique 失败的完整坐标分别验证；追溯表未将原生失败路径冒充配置 RollbackRow 的直接证据。
- 最终编码清单核查 275 个文本文件，其中 143 个 C# 严格 UTF-8、BOM、LF、单末尾换行全部通过。三个保留的非 C# 存量/生成例外为 BATCH-RELATIONS 的 plan.md、symbol-test-map.md EOF 和生成 EntityProviderContracts.xml 行尾；未实施全仓库转换。git diff --check 退出 0，保留 ProfileFixtures.xml 存量 CRLF 提示。
- verification-summary.json 汇总最终候选入口；旧包、失败日志、原始 TRX、旧审批比较及历史 NEEDS_FIX 报告保留。
- FIX-014 已完成：执行记录同步，review-round1-final.md 独立 PASS 已落盘，历史报告保留。F-002～007 与环境验收仍属于后续 Goal，不因 F-001 收口而宣布整个目标完成；本地 no-notify Finalizer 单独验证。

## 修改范围与兼容性

Abstractions 的模型、Builder、策略枚举；共享 RowPlan/SubtotalImportScanner；NPOI Entity 导出、导入和布局支持；ClosedXML Entity 执行器；新增职责测试及本任务验证产物。没有增加生产友元、改变旧公开签名、修改提交算法或引入 Task.Run/.Result/.Wait。C# 维持 BOM/LF，文档与验证脚本无 BOM/LF；最终编码清单和保留例外见上节。

## Reviewer 注意事项与后续

最终独立 Review 已核对 marker 缺失/重复定位、空明细与模板缩短、局部名称及非法旧引用、最终公式区间、固定字段/多区域冲突和资源合同，并确认当前 OPEN_ACTIONABLE=0。review.md 及进度报告保留为历史审计证据，最终结论见 review-round1-final.md。

使用 task-finish --no-notify 仅维护本地任务状态；不修改版本号，未自动 commit、push、创建 PR、发布或发送外部通知。完整 Goal 后续按 F-002～007 推进。
