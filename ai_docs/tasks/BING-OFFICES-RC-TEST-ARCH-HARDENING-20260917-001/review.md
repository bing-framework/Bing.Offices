<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001
AI_REVIEWED_AT: 2026-09-18T09:58:51+08:00

# Round 8 独立审查

## 结论

当前 Review 修复范围验收 PASS：FIX-006 已解决，FIX-001 至 FIX-009 无剩余 OPEN 的 MUST_FIX/SHOULD_FIX，本轮未发现新的代码缺陷。此 PASS 是对前轮修复项和相关回归的结论，不是全计划 COMPLETED，也不是正式 Release PASS。

全计划完成状态仍 PARTIAL，正式发布仍 PARTIAL。API compare 候选 identity、最终同候选四包/Consumer、迁移前展开 Case 证据及正式资源/外部 CI 尚未闭环。本轮只更新 review.md，不修改生产、测试、plan.md 或 execution.md，不提交、推送或发布。

## 阶段验收

| Phase | 状态 | 依据 |
| --- | --- | --- |
| RC-000 | PARTIAL | 基线存在；迁移前展开 Case/TRX 未完整验证，不补造历史结果。 |
| RC-010 | PASS | 方法身份 569→592，Added=23、Removed=0；原 NPOI 三个分支已恢复明确目的地及等价断言。历史展开 Case 缺口归 RC-000。 |
| RC-020 | PASS | Common 隔离与直接 CSV 实现沿用此前验收。 |
| RC-030 | PASS | NPOI 专属架构及已补回的直接边界测试沿用 Round7 验收。 |
| RC-040 | PASS | MiniExcel 专属架构成立，本轮独立双 TFM 各 39/39。 |
| RC-050 | PASS | NPOI 真实文件集成职责沿用此前验收。 |
| RC-060 | PASS | 此前本机集成及 500K/1M Large 记录成立，不代表正式环境批准。 |
| RC-070 | PASS | Aggregate 公共合同职责沿用此前验收。 |
| RC-080 | PASS | Solution/CI/精确 IVT 治理沿用此前验收；外部 CI 未验证。 |
| RC-090 | DEVIATED_OK | RawDateReader 恢复 HEAD 全量索引；热点反射已适配实际签名，轻量 worker 成功。 |
| RC-100 | DEVIATED_OK | Relation 缓存/索引撤回，原逐子项扫描合同恢复。 |
| RC-110 | DEVIATED_OK | Importer 列预绑定全部撤回；preflight 移位保留，语义安全网保留。 |
| RC-120 | PASS | 当前摘要已明确优化撤回与历史证据定位，中文注释沿用此前验收；历史轮次记录不作为当前事实。 |
| RC-130 | PARTIAL | 本轮专项验证通过；最终同候选 API/包/Consumer、正式资源及外部 CI 未闭环。 |
| RC-140 | PASS | 本轮独立复审完成，前轮全部 FIX 已关闭；不代替 RC-130 发布门禁。 |

## 本轮证据与限制

- 读取 skill、AGENTS、完整 plan、execution、旧 review，核对 HotspotProbe 真实调用、生产 diff 和 benchmark/final 当前摘要。暂存区 diff 为空；未跟踪的 HotspotProbe/任务报告按实际文件审查，不因 git diff 不显示而跳过。
- HotspotProbe.cs:198-203 已查找 Read(Stream,string,CancellationToken)，并对缺少方法显式抛 MissingMethodException；不再查找已撤回五参数过滤重载。:41 将 before 标为 investigative numeric-cell replay，明确不是 HEAD Reader baseline。
- RawDateReader 相对 HEAD 无差异；Importer 相对 HEAD 仅有 Relation 委托副作用的两行注释，没有固定/动态列预绑定保留。
- benchmark-report.md:10/14、final-report.md:5/22、execution.md 当前摘要及 Round7 修复记录均明确生产优化全部撤回，不宣称旧样本的当前候选收益。历史 JSONL 不重标。
- 独立执行下面的 worker 命令，退出码 0；3 个样本均 indexCount=300000、workbookBytes=1863156、PID=16096。证据 artifacts/benchmarks/independent-review-round8-raw-date.jsonl。
- 独立 MiniExcel Unit Release --no-build --no-restore：net6/net8 各 39/39，failed=0。TRX：artifacts/tests/independent-review-round8/miniexcel-net6/miniexcel-net6.trx 和 miniexcel-net8/miniexcel-net8.trx。
- 委派窄任务用 XML 核对 Fixer Round8 两份 TRX：各 39 total/executed/passed，failed/notExecuted/error 均 0。
- worker 的单独 JSONL 不含 candidate header；本轮只认定为运行 smoke，不认定为正式性能对照。核对实际运行目录的 DLL SHA-256：MiniExcel D2138E2353D5C0BACAB9A451233455F7A756D02B896D9085F63458327BD5CDE3；Benchmarks E4AE36CE6AB82CAEB09DAC312E4817B5508684C5A5C423833980A41FCD1934E7。正式性能采样仍须完整候选 identity、真实 baseline 和独立进程证据。
- git diff --check 通过，仅 LF/CRLF warning。src/tests/benchmarks 下 artifacts、TestResults、BenchmarkDotNet.Artifacts 目录扫描为零。
- 版本哈希未变：version.props 77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1；version.dev.props 282D9A7C70D79FFC5F354DA57035A479B08A8485D1AF7637B25AF6C815DCAF31。
- 本轮未重新构建完整 Solution；Fixer 已记录 Release build 0 warning/0 error。本轮未重跑其他五项目全矩阵、Docs、API、Pack/Consumer、Large、资源或完整热点矩阵，不将历史结果重标为本轮执行。

专项命令（仓库根）：

```powershell
dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --hotspot-worker artifacts/benchmarks/independent-review-round8-raw-date.jsonl after raw-date 3 3
```

## FIX 收口

### FIX-006 — RESOLVED

实际三参数 Reader 调用成功，悬空反射重载问题已消失；调查 replay 标签已纠正，当前报告不再称 Importer 预绑定保留。原始性能 artifact 保持历史资料，不能证明收益的生产修改已撤回，符合计划允许的收口。不要求为已撤回优化补 before/after，也不要求重跑 Relation 100K 原扫描。

### FIX-007 — RESOLVED

保持 Round7 结论：NpoiProviderBoundaryRegressionTest 三个方法恢复日期系统、四组 DataRowStartIndex、ReadColumns 结构化错误的原 NPOI 分支，独立双 TFM 各 6 个展开 Case 已通过，矩阵和符号映射已补去向。

### FIX-008 — RESOLVED

Relation 逐子项 ParentKey.DynamicInvoke 首匹配扫描恢复，重复求值/短路/异常边界回归保留，本轮 MiniExcel 双 TFM 各 39/39。

### FIX-009 — RESOLVED

违规嵌套构建产物已可恢复地移至根 quarantine，本轮再次扫描为零。

FIX-001 至 FIX-005 保持此前已解决结论。当前没有需交 Fixer 的未解决项。

## 全计划与发布交接

审查修复验收 PASS；全计划 PARTIAL；正式发布 PARTIAL。

仍需闭环的验证工作是：最终回退候选的 API compare identity、同候选四包/Consumer DLL identity、完整最终职责矩阵与发布证据。迁移前 TRX 缺失只能保留 NOT_VERIFIED，不能补造。维护者资源批准、真实 2 CPU/4 GiB、候选 commit 外部 CI 仍 NOT_VERIFIED/BLOCKED；1K/36 场景资源 artifact 且 approval BLOCKED 不等于生产批准。本机历史 Large 通过不代替正式发布环境证据。

不自动提交、PR、tag 或发布；下一阶段应补齐已记录的验证门禁，而不是继续恢复未经证明的优化或重复审查已解决 FIX。

## 当前进度分析

本节为基于现有证据的进度盘点，不是新一轮全量测试或发布批准。Round8 修复审查 PASS 不变；全计划和正式发布仍 PARTIAL。

- 阶段完成率：13/15 = 86.7%，约 87%；尚未闭环阶段 2/15 = 13.3%。按 RC-000 至 RC-140 等权统计，PASS 和 DEVIATED_OK 计入收口，PARTIAL 不计完成；不是代码行数或工时百分比。
- 审查修复完成率：9/9 = 100%，FIX-001 至 FIX-009 已关闭，没有剩余 MUST_FIX/SHOULD_FIX。
- 六个职责项目已建立并隔离；方法身份账为 569→592、Added=23、Removed=0。不同轮次测试证据不能替代最终同候选全矩阵。
- RC-090/100/110 计入已接受的回退收口，不代表目标性能优化已交付；RawDate 过滤、Relation 缓存和 Importer 预绑定均已撤回，无已证明的优化收益。

未完成工作集中于：

1. RC-000：迁移前展开 Case/TRX 缺口需明确历史证据状态与验收处理，禁止补造历史运行结果。
2. RC-130：最终候选六项目双 TFM、Docs、Release build 及同候选 100K 探针证据统一复核，不能混用多轮候选结果。
3. RC-130：API compare 的 candidate source identity mismatch 未闭环，需要维护者对正确候选的 identity 走既定审批，不自动改正式 API baseline。
4. RC-130：最终回退候选四包与两个 PackageReference Consumer 的包/DLL identity 未重验；已有 Consumer PASS 是历史证据，不是最终候选闭环。
5. RC-130：资源矩阵现有 artifact 为 1K/36 场景、approval BLOCKED，需补计划规模与真实资源环境证据及维护者批准人/时间。
6. RC-130：正式环境 500K/1M、真实 2 CPU/4 GiB 和候选 commit 外部 CI 尚未验证；本机 Large 历史通过不可替代。

正式发布就绪状态为未就绪，不用任意百分比替代强制门禁。剩余工作以验证/证据/批准为主，不应再进入重复 fix-review 循环；历史缺口与外部审批也不是单纯继续编码就能消除。
