<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BO-PACKAGE-VALIDATION-20261026
AI_EXECUTION_FINISHED_AT: 2026-10-26T23:28:57.2383630+08:00

# 包消费与候选验证实施报告

## 执行结论

PKG-01～PKG-04 全部完成。本轮补齐新增 API 独立 NuGet 消费验证、当前候选身份验证和 CI 接入。真实打包、隔离恢复、四个消费目标、双 TFM API 比较及三项拒绝用例全部通过。生产公共 API、行为、依赖/包版本不变，批准的 API 基线未修改。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
LOCAL_CANDIDATE_STATUS: PASS
OPEN_ACTIONABLE: 0
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL
RELEASE_STATUS: NOT_VERIFIED
PERFORMANCE_EVIDENCE_STATUS: NOT_APPLICABLE
RESOURCE_EVIDENCE_STATUS: NOT_APPLICABLE（本轮无资源算法/额度变化）
GOAL_STATUS: STOPPED_BLOCKED

## 任务与候选信息

- 用户要求“继续实现相关功能”，承接 BO-API-USABILITY-20261025 复查剩余项；先写 [plan.md](plan.md)，再实施。开始时间为 2026-10-26T15:19:02.069Z。
- Git HEAD：84958c1cc5eec5afbae51d76b0d07b4ab133b748；工作区 dirty，包含用户既有改动及本轮实现。
- 当前源码逻辑清单：ACB31616C30B2A4F0A987A7DF1A7F39BD633AD66F6B8FB0D66D9ADEBADB7D024。
- 证据目录：TestResults/package-validation-20261026；实际 nuspec 版本均为 2.0.0。未将旧 candidateIdentity/nupkg 当作本轮产物。

## 计划执行情况与已完成事项

| 任务 | 状态 | 结果 |
| --- | --- | --- |
| PKG-01 独立消费者 | COMPLETED | 保留旧 Net6/Net8 消费，补六个单表入口；新增 Core-only 双 TFM 消费，四个目标从空缓存恢复运行成功 |
| PKG-02 候选与 API 分离 | COMPLETED | --candidate 提供当前身份，--baseline 决定批准 API；capture 记录实际 base commit；正常比较与三个反例通过 |
| PKG-03 统一脚本与 CI | COMPLETED | 打包/捕获/消费/比较/归档/清理使用同一脚本；CI 接入并 always 上传证据 |
| PKG-04 文档与追溯 | COMPLETED | 使用文档、专题链接、symbol-test-map 和实施报告齐备 |

NPOI 消费新增 XLS/XLSX × 六个入口（Stream、Bytes、File 的同步/异步），每 TFM 成功 12 次、失败 12 次，共 48 个单表场景。断言完整实体、SourceRows、完整 InvalidHeader 错误、状态摘要及流所有权。旧 Workbook/CSV/直接 Provider/NPOI 扩展检查保留。

Core-only 只引用 Core 与对应框架 DI 实现。两个框架分别验证同步/异步 CSV 完整 UTF-8 输出、有效实体、校验失败后继续及状态摘要；没有 Excel Provider 注册。隔离恢复图禁止 ProjectReference、Core 无 NPOI；8 次 Office 包恢复来源均匹配本轮本地源，记录在 summary.restoredSources。

脚本在仓库外系统临时目录复制消费者，阻断仓库构建导入，每个消费项目使用独立空缓存。finally 恢复包版本环境变量并清理自建临时目录，先检查绝对父路径和专属前缀；产物与证据保留。输出目录必须不存在，防止混用旧候选。

## 修改文件与 API/配置变化

- build/ApiSnapshot/Program.cs；新增 build/Verify-PackageCandidate.ps1。
- .github/workflows/ci.yml：用统一脚本替换重复 pack/consumer 步骤，保留既有业务测试、覆盖率、资源契约和 benchmark 枚举，证据上传改为 always()。
- tests/Bing.Offices.Consumer.Net6/Net8 的 csproj、Program.cs；新增 Consumer.Shared/ExcelSingleSheetPackageChecks.cs 与 Consumer.Core 项目。
- docs/excel/README.md；新增 docs/excel/package-validation.md；本任务计划/报告/追溯文件。

--candidate 仅选择身份清单；成员、schema/generator 和审批检查仍取批准基线。未传参数时保持旧路径，capture 与 candidate 互斥；缺失/错误清单明确失败。capture 将实际 base commit 同步到文档，compare 保存本轮身份及双 TFM 快照。

身份沿用 logical-v3，不改历史算法，不代表原始文件签名或构建来源证明。当前有效性由本轮真实构建、打包、消费、源/产物逻辑身份及 API 比较共同验证。没有更新批准基线或制造新审批。

## 测试结果与 Build/Lint/Format

实际运行：

```powershell
pwsh -NoProfile -File build/Verify-PackageCandidate.ps1 -OutputDirectory TestResults/package-validation-20261026
```

以下路径均相对于本轮证据目录。

| 检查 | 结果 | 证据 |
| --- | --- | --- |
| 三个生产包及库/XML 结构 | PASS | packages/、logs/pack-*.log；3 nupkg + 3 snupkg |
| ApiSnapshot Release build | PASS，0 warnings / 0 errors | logs/build-api-tool.log |
| 既有身份契约自检 | PASS | logs/identity-contract.log；LF/CRLF、源/包/程序集/审批篡改拒绝 |
| NPOI net6/net8 独立消费 | 2/2 PASS；新增单表场景 48 次 | logs/run-Bing.Offices.Consumer.Net6-net6.0.log、Net8-net8.0.log |
| Core-only net6/net8 独立消费 | 2/2 PASS | logs/run-Bing.Offices.Consumer.Core-net6.0.log、net8.0.log |
| 双 TFM 当前候选比较 | PASS | comparison/api-diff.json 两个 TFM 均为空 |
| 缺失候选清单 | 按预期拒绝 | logs/reject-missing-candidate.log |
| 篡改程序集身份 | 按预期拒绝 | logs/reject-bad-identity.log |
| 合法候选 + 不匹配 API 基线副本 | 按预期拒绝 | logs/reject-api-mismatch.log；未修改批准基线 |
| PowerShell AST / CI YAML 入口和 artifact | PASS | 本地解析；独立子代理亦确认 AST 有效 |
| UTF-8 / git diff --check / 批准基线无差异 | PASS | 无空白错误；既有 CRLF/LF 提示不影响结果 |

最终标记 package-candidate-ok consumers=4 negativeChecks=3，summary.status=PASS。恢复图保存于 consumers/，完整命令输出在 logs/。消费者及其缓存已清理。

net6 消费构建有 SDK 既有 NETSDK1138 支持周期提示，均为 0 errors；未抑制提示或改变 TFM。实际运行时为 .NET 6.0.36 / .NET 8.0.30。

本轮无生产实现/公开成员变化，复用上轮完整 1,999 次和复查 66 次的未变更业务范围证据，不作为本轮新执行数量。本轮重新构建生产包、验证受影响消费者和 ApiSnapshot；未重跑完整业务测试、benchmark 或部署容量矩阵。

## 计划偏差、基线问题、风险与 Reviewer 注意事项

- 无实质计划偏差；本地范围没有未完成项。
- luna_worker 实施窄范围消费源码；主代理集成和实测。该代理另对未参与编写的工具/脚本/CI 做独立只读复核，未发现可证实的可执行问题；不冒充正式 review.md。
- 历史 NuGet.Config 保留供旧消费者手工调试；统一脚本生成匹配当前候选的临时配置。
- 本地 Windows 通过不能代替 GitHub Ubuntu run。远端需提交/推送后关联实际提交和 artifact；本轮未触发远端工作流。
- 生产资源/并发证明仍属外部门禁，无部署环境或新容量承诺；输入限额不是 DOM 峰值保护。
- TestResults 是本地忽略证据目录，CI 负责归档同类结果；发布/复制候选时应保持包、清单及日志对应。
- Executor 未改旧审查报告；上一轮待补项的当前状态以本报告为准。

## TODO 分类与 Git 收口

- Completed：4/4 计划项、四个消费者、双 TFM 比较、三项拒绝用例、CI/文档。
- Open Actionable：无。
- Blocked Approval：无新增。
- Blocked External：实际 GitHub run、生产资源/并发部署证明。
- Not Applicable：性能 before/after、新资源容量矩阵。
- Accepted Limitations：沿用 logical-v3，不承诺原始字节签名；既有 DOM 同步边界不变。
- Verified Boundaries：Core 不依赖 NPOI；单表成功/失败；错误身份和 API 差异被拒绝。
- Deferred：本地范围无。
- No-Progress Check：CHANGED，上一轮本地验证缺口已补齐。
- Implementation：4/4；Local Candidate Gate：PASS；外部发布门禁不计为通过。
- Next Action：STOP。
- Finalizer：task-finish.mjs BO-PACKAGE-VALIDATION-20261026 --no-notify。
- 保留用户已有 README、docgen 文档、未跟踪工作流及旧任务材料；未自动 git commit、未自动 git push、未自动创建 PR、未发布包、未发送通知。
