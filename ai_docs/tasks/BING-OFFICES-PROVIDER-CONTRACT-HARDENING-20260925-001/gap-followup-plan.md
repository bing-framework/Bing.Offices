# 后续验收缺口补充计划

Task ID: BING-OFFICES-PROVIDER-CONTRACT-HARDENING-20260925-001

## 分析结论与边界

本轮承接用户“先分析，再实现”的授权。已有独立审查 PASS 是上轮范围的结论；本轮按原验收项核查发现下列尚缺直接证据的路径。当前没有证据要求重新实现 Golden、Entity、校验引擎或原子提交算法，不新增公共 API。

| 编号 | 可执行缺口 | 已有证据的实际边界 | 本轮方案及验收 |
| --- | --- | --- | --- |
| G1 | Failure Workbook 文件目标失败分支缺少公共入口证据 | 首读取消发生在输出产生前，不能证明提交失败清理 | 通过真实目标路径制造提交失败，覆盖 NPOI/ClosedXML 同步异步及两种输出模式；检查目标内容、输入所有权和临时清理。通过序列化字节预算失败验证旧目标保全。复用 Core 写入/取消直接测试并建立映射。 |
| G2 | Workbook 行数上限缺少跨 Sheet 精确边界成功测试 | 现有共享用例只断言合计超限 | 三 Provider、同步异步测试两 Sheet 合计等于 MaxRows；使用不同集合承接两 Sheet，断言完整实体、无错误；强化单 Sheet 超限无部分实体断言。 |
| G3 | Stream best-effort 缺少共同的中途写失败证据 | Provider 局部异步成功测试不能证明失败后所有权 | 注入已写入部分内容后抛出的流，覆盖 NPOI/ClosedXML，断言输入输出未关闭、实际写入发生、失败传播；不声明可回滚。 |
| G4 | 最终报告和最终全解证据不够明确 | 最终报告混有旧计数及已过期 PARTIAL；上轮全解调用只显示前段输出，进程退出不等于退出码 0 | 重新记录一次最终全解测试的可审计日志、退出码和 TRX；当前结果放在报告首部，历史证据保留为历史；核对包 consumer 是否匹配当前候选。 |

## 实施顺序

1. 补 G1/G3 公共文件/流合同测试；先定向运行。只有测试证明生产缺陷才修改对应实现。
2. 补 G2 精确边界测试和全结果断言，定向验证。
3. 跑 ProviderContract 完整双 TFM；若生产代码无变化，复用未变 Provider 单元测试和 API 证据。
4. 最终执行 `dotnet build Bing.Offices.sln --no-restore -c Release -v:minimal -m:1` 与 `dotnet test Bing.Offices.sln --no-restore -c Release -v:minimal -m:1`，保留 TRX 和退出码。仅此证据缺口触发一次新鲜全解，不触发重跑资源矩阵。
5. 更新执行报告、当前最终报告及生产符号映射，独立复核本轮 Diff；只有 OPEN_ACTIONABLE > 0 才修复。

## 影响分析

- 预期改动：ProviderContract 测试、任务文档；生产 API/依赖/TFM/Benchmark 不变。
- 风险：故障测试必须确定性触发，不使用延时竞态；临时目录只删除测试自建路径。
- 选择：优先公共 API + 真文件证据，复用 Core 注入式提交测试；不为测试增加新的生产全局钩子。
- 完成标准：G1–G4 具有真实测试或明确复用证据；记录最终退出码；UTF-8/BOM/EOL 和 git diff --check 通过。
- 外部 CI、其他 OS、字体与生产容量仍为 BLOCKED_EXTERNAL，本轮不将其改写为本地通过。
