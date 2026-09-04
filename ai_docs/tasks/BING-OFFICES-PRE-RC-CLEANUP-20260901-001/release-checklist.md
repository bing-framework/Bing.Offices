# RC 发布清单

## 判定

当前：`No-Go` / `BLOCKED`。

本清单只绑定当前工作树真实证据，不把历史未批准 API hash、ShortRun Benchmark、mapping-only ResourceProbe、缺失 runtime 或旧 consumer 二进制当作发布通过。

## 构建与包

| 门禁 | 状态 | 证据/解除条件 |
| --- | --- | --- |
| Restore | `PASS`（当前环境） | `baseline.md`；后续交付仍需在不可变包身份下 strict restore |
| Release solution build | `PASS` | 已有 Release build 结果；提交后需 clean clone 复核 |
| Unit net6 | `PASS` | `384/384`，正式 API baseline 匹配；TRX：`tests/Bing.Offices.Tests/TestResults/api-baseline-net6-final-rerun.trx` |
| Unit net8 | `PASS` | `384/384`，正式 API baseline 匹配；TRX：`tests/Bing.Offices.Tests/TestResults/api-baseline-net8-final-rerun.trx` |
| Unit netcoreapp3.1/net5/net7 | `BLOCKED` | runtime 缺失，安装后按相同命令重跑 |
| Integration net6/net8 | `PASS` | `30/30`，见 `integration-test-report.md` |
| Docs net8 | `PASS` | `11/11`，但为 ProjectReference consumer；不替代 nupkg consumer |
| nupkg pack | `PASS` | Round 5 任务 `artifacts/packages-round5`；包 hash 见 `package-consumer-report.md` |
| Package-only consumer net8 | `PARTIAL` | Round 5 `package-consumer-rerun2` 在 `C:\nupkg-cache-round5` restore/build/run 成功；长路径 cache `MSB3106` |

## API 与兼容治理

| 门禁 | 状态 | 证据/解除条件 |
| --- | --- | --- |
| 成员级 API diff | `PASS` | `api-diff.md` 已绑定当前 Release 快照；差异仅为已批准 Breaking Change 和 API 收敛 |
| 正式 API baseline | `PASS` | 用户已批准当前 Release API；compare 退出码 `0`；正式文件：`artifacts/api-snapshot-formal-baseline-20260903.json` |
| 删除候选闭环 | `PARTIAL` | `deprecated-removal.md`；legacy/API execution detail 仍待逐符号批准 |
| DI 唯一入口 | `PASS` | `AddNpoi` 已删除；新入口返回同一 `IServiceCollection` |

## 安全、资源与失败输出

| 门禁 | 状态 | 证据/解除条件 |
| --- | --- | --- |
| Mapping/v1/relation/CSV P0 | `PASS`（可运行 TFM） | Unit/Integration 回归和 `test-matrix.md` |
| Failure Workbook | `PASS` | 专项 `14/14`；`MaxSerializedBytes` 仅表示序列化输出上限 |
| ResourceProbe | `PARTIAL` | Excel child process `7/7`，另有 mapping/unique `16/16`；仍需补真实 DOM 扩展、压缩输入和 Failure Workbook 双 DOM 样本 |
| 内存硬上限声明 | `BLOCKED` | NPOI DOM/解压/实体/双 DOM 峰值不能由 `MaxInputBytes` 单独保证；文档必须保持限制表述 |
| NPOI capability fallback | `PASS` | HSSF/XSSF 2.7.4 行属性行为已直接核实；仅保留窄范围 catches |

## 性能门禁

| 门禁 | 状态 | 证据/解除条件 |
| --- | --- | --- |
| StreamPipeline BDN | `PARTIAL` | `9/9` ShortRun；不能代替正式预算 |
| 计划 Benchmark 矩阵 | `BLOCKED` | CSV、getter/setter、完整 Failure Workbook、ExportToBytes/Stream/File、取消等未形成统一批次 |
| 尾延迟 | `UNAPPROVED` | 并发 1/4/16/64 已测；需批准预算/容差/批准人/日期或具名 waiver |
| 100k Import 异常 | `REVIEW_REQUIRED` | BDN 记录 `3` 个异常，需复核后才可作发布结论 |

## 文档与交付

- [x] `docs/excel/README.md` 明确 DOM、ownership、`MaxInputBytes` 和 `MaxSerializedBytes` 边界。
- [x] `docs/excel/nuget-migration.md` 使用新 DI 入口和新失败输出属性。
- [x] `package-consumer-report.md` 已排除旧二进制成功输出，并记录短路径成功和长路径 `MSB3106`。
- [x] `review.md` 已形成独立 `BLOCKED` Review。
- [x] `final-report.md` 已形成最终执行摘要。
- [x] API baseline 批准记录和正式快照已写入任务 artifacts；仍需授权人员纳入最终交付提交。
- [ ] 任务报告、原始 JSONL/BDN、快照和 lockfile 进入获授权的交付提交并可由 clean clone 获取。
- [ ] 缺失 TFM runtime 验证完成。

## 解除顺序

1. 已完成：用户确认两项 Breaking Change 和正式 API baseline；
2. 已完成：在批准 baseline 上重跑可运行 Unit/API compare；仍需补齐缺失 runtime TFM；
3. 形成同提交、完整 Benchmark 矩阵及具名性能预算/waiver；
4. 补齐真实 XLS/XLSX、压缩输入、Failure Workbook 双 DOM 的独立 ResourceProbe；
5. 解决或正式记录长路径 SDK/MSBuild 环境限制，重新验证 package-only consumer；
6. 完成 legacy/public execution detail 的逐符号治理；
7. 由独立 Reviewer 重新复核，满足全部 Go 条件后再讨论发布。

## 安全操作边界

本轮未执行 `git add`、`git commit`、`git push`、`git reset`、`git clean`、tag、publish 或 PR。建议提交分组和 message 仅在 `final-report.md` 中说明，不代表已执行。
