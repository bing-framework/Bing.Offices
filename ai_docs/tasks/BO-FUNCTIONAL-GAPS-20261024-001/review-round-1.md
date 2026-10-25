<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BO-FUNCTIONAL-GAPS-20261024-001
AI_REVIEWED_AT: 2026-10-24T17:17:36+08:00

# 功能遗漏修复复核

## 结论与范围

主要修复有效，但发现 1 项由 NPOI 重复绑定保护引入的边界回归，不能沿用 execution.md 中“本地全部关闭”的结论。该执行记录保留为上一轮证据，不在复核阶段改写。

本次对照用户确认的 5 项修复要求、execution.md、当前工作区与暂存区差异、源码、测试及实际产物。任务没有 plan.md；这里是修复结果复核，不是正式的基于计划验收，也未补写或虚构计划。业务代码及现有测试均未修改；只增加本复核记录与被 Git 忽略目录下的诊断程序。一个 luna_worker 独立核对了测试与工作流证据。

```text
IMPLEMENTATION_STATUS: NEEDS_FIX
TEST_STATUS: EXISTING_SUITE_PASS / NEW_BOUNDARY_REPRODUCED
PLAN_ACCEPTANCE_STATUS: NOT_APPLICABLE
PERFORMANCE_EVIDENCE_STATUS: NOT_APPLICABLE（本轮未测量）
RESOURCE_EVIDENCE_STATUS: BLOCKED_EXTERNAL（生产环境证明）
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL（当前候选远端 CI）
RELEASE_STATUS: BLOCKED
OPEN_ACTIONABLE: 1
BLOCKED_EXTERNAL: 2
```

## 逐项核对

| 修复项 | 结论 | 依据 |
| --- | --- | --- |
| GAP-001：Excel/CSV 多个表头绑定同一目标 | PARTIAL | 标题/别名重复绑定保护已接入；CSV 固定属性使用 Ordinal 身份比较。NPOI 固定属性身份却受 HeaderComparison 控制，存在 FIX-001 误拒绝。 |
| GAP-002：无表头 CSV 动态列 | PASS | BindByPosition 按 map.Columns 展开，动态列使用 Order、Key 排序，与导出一致；同步、异步均在消费正文前绑定；缺少定义和展开后的列数上限有直接测试。 |
| GAP-003：所有 CSV 字段长度限制 | PASS | Read/ReadAsync 在绑定和实体物化前检查所有解码字段，并通过原始长度上限提前停止超长字段；已覆盖表头、未知列、引用转义及异步读取。 |
| GAP-004：覆盖率真实采集及缺失门禁 | PASS | collector 已安装；CI/Docs 检查非空 Cobertura 文件；实际 XML 可解析；正、负向返回码证据齐全。 |
| GAP-005：文档示例一致性门禁 | PASS | 不再对 ignored 目录运行无效 git diff；重新生成 Markdown 后通过 docs-check 验证源码示例。6 个示例成功、临时篡改失败的日志与 CLI 返回值一致。 |

## FIX-001：固定属性身份不应沿用表头大小写比较规则

- 严重程度：MEDIUM（P2）
- 处理要求：SHOULD_FIX
- 状态：OPEN_ACTIONABLE
- 问题：`NpoiImportSheetExecutor.CreateColumns` 第 184～185 行以 `options.HeaderComparison` 为 `boundFixedProperties` 选择比较器；默认 OrdinalIgnoreCase 会把不同 CLR 属性 `Name` 和 `name` 当成同一目标。第 216～217 行因此返回 InvalidHeader，拒绝本来具有两个不同列标题的合法导入。
- 证据：实体分别以 `[ColumnName("Upper")] public string Name` 和 `[ColumnName("Lower")] public string name` 映射两个字段。工作簿表头为 Upper/Lower，正文为 first/second。默认比较规则下 XLS/XLSX × 同步/异步 4 个路径均返回 `InvalidHeader: 导入的表格中多个表头映射到同一属性: name`，成功行数为 0。仅切换表头比较为 Ordinal 的 4 个对照路径均成功导入 1 行，且两个属性分别精确等于 first、second。
- 影响：使用大小写不同属性名、但具有不同显式标题的模型，会被本次新保护错误拒绝。表头匹配规则和 CLR 属性身份是两件事；改变前者不应合并后者。
- 修复目标：固定属性去重以实际属性身份或大小写敏感的属性名为准；表头仍按调用方 HeaderComparison 查找。
- 修复要求：保留真正的同一属性被标题/别名重复绑定时的 InvalidHeader；不要通过强制所有调用方改用 Ordinal 表头匹配规避。可与 CSV 的 `StringComparer.Ordinal` 固定属性集合保持一致。
- 验证方式：补充 Name/name 配合不同 ColumnName 的 XLS/XLSX × 同步/异步直接回归，断言完整成功值及空错误；运行新用例与已有 NpoiImportHeaderBindingTest，确保同一属性重复绑定仍失败。此问题不需要修改公共 API 或重建 API baseline。

复现程序与日志：

```text
TestResults/audit-fixes-review/probe/Probe.csproj
TestResults/audit-fixes-review/probe/Program.cs
TestResults/audit-fixes-review/header-identity-probe.log
```

```powershell
dotnet run --project TestResults/audit-fixes-review/probe/Probe.csproj -c Release --no-restore
```

诊断程序退出码为 0 表示上述“4 个失败场景 + 4 个成功对照”均符合断言，不表示该缺陷已修复。

## 验证证据复核与限制

- 重新读取上一轮 5 个 TRX：787 + 787 + 38 + 38 + 10 = 1660，全部 failed=0、notExecuted=0。新增的 58 项用例已包含于该回归；没有覆盖 FIX-001 的模型，因此通过数不能证明这一边界正确。
- 本轮新增一个隔离诊断程序，完成 XLS/XLSX 与同步/异步受控复现和 Ordinal 对照；未重复运行无相关变化的整套测试或 Benchmark。
- 定向 Cobertura XML 存在、可解析，大小 2,540,890 字节；其覆盖范围是原来的 58 项定向用例，不能当作全仓覆盖率。
- coverage-gates.json 中 CI/Docs 的 missing/empty/actual_report 返回码均为 1/1/0；工作流中的实际 Bash 检查与这些结果相符。
- docs-gates.log / docs-gates.json / verify-docs.py 相互一致：5 份 Markdown 生成成功、6 个示例匹配、临时副本中的 ExportWorkbook 不匹配时返回 1；实际 CLI 会将这一返回值传为进程退出码。
- `git diff --check` 通过；暂存区没有变更；本轮没有暂存、提交或修改生产实现。
- 目标 2C4G/semaphore=1 的部署与资源证明、当前候选的远端 CI 全绿证据仍为 BLOCKED_EXTERNAL。它们与本轮发现的代码缺陷分别记录。

复核到此结束；本轮没有进入自动修复循环。
