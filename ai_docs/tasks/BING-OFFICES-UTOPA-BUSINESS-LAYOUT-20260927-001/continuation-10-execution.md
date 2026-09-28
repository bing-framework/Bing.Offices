# Continuation 10：跨页签字区执行记录

## 结果

本阶段 COMPLETED。PageBreak + 多行 PageSubtotal + Footer 可以表达中间页小计后签字、末页总计后签字；已用真实工作簿验证，无需重复增加公开 API。尾部命名锚点仍是下一独立阶段，未实现。

修复两个关联问题：

1. 未配置最终 Footer 时，分页小计和签字区把末页明细推移到声明/物理边界外，旧实现分别静默写出或返回运行期导出异常。现在按推移后的实际末行、列宽在写目标流前返回 BingOfficesConfigurationException / Plan。
2. 导入允许只有分页小计却没有后续明细的非法终末页面；现在返回 Plan 配置错误。

## 验证证据

证据目录：`artifacts/continuation-10/`。所有测试使用 Release，常规测试筛选 Category!=Large。

| 层级 | 验证 | 结果 |
|---|---|---|
| L1 | PageSignatures + PageSubtotalMarkerErrors，NPOI | net6.0、net8.0 各 13/13 |
| L1 | 同上，ClosedXML | net6.0、net8.0 各 7/7 |
| L2 | NPOI 常规测试 | net6.0、net8.0 各 675/675 |
| L2 | ClosedXML 常规测试 | net6.0、net8.0 各 196/196 |
| L3 | Provider Contract | net8.0 282/282；net6.0 单独运行 282/282 |
| L3 | PublicApi（包含成员快照） | net6.0、net8.0 各 11/11 |
| L3 | Docs 可执行示例 | net8.0 10/10，17 个代码块 |
| L0 | Solution Release build | 0 错误；4 个既有 Microsoft.Bcl.Memory net6.0 警告 |
| L0 | diff、严格 UTF-8/BOM/EOL | 本轮修改文件通过 |

修复前新增回归先失败：NPOI 5 个失败（4 个越界和 1 个非法末页标记），ClosedXML 3 个失败（2 个越界和 1 个非法末页标记）；原有签字组合实现的成功路径全部通过。修复后以上定向用例全部通过。

签字区往返测试每个配置包含 0、2、5 条明细；NPOI 使用 HSSF/XSSF 模板，ClosedXML 使用 XLSX 模板；覆盖有无表头、同步/异步、GapRows、完整两列内容、每个合并范围、字体加粗、行高/列宽、所有水平分页符、完整明细读回、调用方流仍打开。越界回归断言结构化错误及目标流原始字节完整。

Provider Contract 首次双 TFM 并行执行时，net6.0 的 `SpreadImages_FailedEnumeration_ShouldCleanStagingAndPreserveFile(async:false)` 失败（281/282）。该测试对进程共享临时目录做前后快照，可能观察另一框架仍在使用的文件。等待全部运行结束后只做一次 net6.0 独立确认，282/282。未修改不相关的 SpreadCheetah 代码或该测试，未隐藏首次失败 TRX。

首次 solution 构建被既有输出目录的访问权限阻止；经工具审批提升权限重试成功。未据此修改代码或仓库权限。

## API、文档、注释及范围

- 本轮仅修改两个 Provider 的私有/内部实现、对应职责测试及文档；公共成员快照与批准基线匹配。未覆盖 API baseline，未创建新审批。
- 按 chinese-comments 为本轮触及 helper 补齐中文 XML 注释。没有进行全仓库注释清理。
- 更新 dynamic-columns 的签字区可执行示例，修正分页小计仍列 TODO 的矛盾，说明末页/空明细与打印页限制。
- 子代理独立维护该文档，并只读核查本轮窄范围实现和测试。其最初边界计算漏计默认表头，复核后撤回该 finding；无剩余窄范围问题。这不代表全部历史 Git Diff 已完成独立审查。
- 存量格式偏差：build/api-snapshot-baseline.json 为 CRLF、无末尾换行；ProfileFixtures.xml 为 CRLF。本轮未修改两者，未将其计为本轮编码通过。
- 未运行全 solution 测试、重新打包 Consumer 或无关性能矩阵；公共 API 和打包链未改变。本轮交付依据是当前 Provider/合同/API/Docs 测试，不把上轮 NuGet 包结果冒充当前构建验证。
- 外部 Utopa 仓库未改动；未 commit、push、PR、发布或更改版本号。

## TODO 分类

- Completed：签字区占用及导入边界、两 Provider 回归、文档和追溯（4/4）。本阶段要求验证已完成。
- Open Actionable：无。
- Blocked Approval / Blocked External：无。
- Not Applicable：新增公共 API 审批、新成员快照更新、无关性能 before/after。
- Accepted Limitations：按明细条数的手工分页不保证打印机实际物理页；最后一页签字使用 Footer；GroupSubtotal 与分页小计仍互斥。
- Verified Boundaries：声明区域末行及 HSSF/XSSF 物理末行越界预检。
- Deferred：尾部命名锚点（重复页命名、模板作用域与动态地址更新），可复用 Provider 契约模板/包，Roslyn 检查器，后续有证据再评估 Source Generator。
- No-Progress Check：CHANGED。
- Next Action：STOP（本阶段结束）；下一阶段先设计尾部命名锚点，再推进契约模板。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
REVIEW_STATUS: NARROW_SCOPE_CHECKED
RELEASE_STATUS: NOT_EVALUATED（本轮不是发版任务）
OPEN_ACTIONABLE: 0
