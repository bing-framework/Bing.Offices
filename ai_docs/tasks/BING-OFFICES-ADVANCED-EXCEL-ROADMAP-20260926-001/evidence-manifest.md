# 最终候选证据索引

本轮范围 F01–F06；生产代码冻结后运行。以下路径均相对于仓库根，明确列举以免混用历史运行。

| 证据 | 最终路径 | 结论 |
| --- | --- | --- |
| Release build | `artifacts/open-provider-20260926/final-build.log` | 退出 0、0 errors |
| 全量测试 | `artifacts/open-provider-20260926/final-test.log`、`final-tests/` | 25 TRX，2820/2820 |
| Linux 功能 | `artifacts/open-provider-20260926/docker-final/docker-smoke.log` | 352/352，退出 0 |
| Linux 运行时 | `artifacts/open-provider-20260926/docker-final/dotnet-info.txt` | 固定 SDK/.NET 6 Runtime |
| 缺字体对照 | `artifacts/open-provider-20260926/docker-final/fontless-npoi-result.json` | NPOI 结构化失败；SpreadCheetah 成功读回 |
| 容量 | `artifacts/open-provider-20260926/docker-final/streaming-capacity.jsonl`、`streaming-capacity.log` | 9/9，单次实测，不设置阈值 |
| API 最终捕获与比较 | `artifacts/open-provider-20260926/api-final/capture/`、`api-final/compare/` | 双 TFM 空 diff；最终身份与批准形状匹配 |
| 八包本地制品 | `artifacts/open-provider-20260926/packages-final/` | 2.0.0，不发布 |
| 隔离包消费者 | `artifacts/open-provider-20260926/consumer-final/` | 双 TFM restore/build/run 均通过 |
| 独立审查 | 当前任务 `review.md` | 4 findings CLOSED，OPEN_ACTIONABLE=0 |

最终 API 候选身份关联稳定输出程序集与包。全量测试完成后重新打包以消除并行重建时的制品竞争；不改变生产源码、批准 API 形状或历史 before 快照。

旧 `full-test.log/full-tests/`、根级 `f06-summary.json`、旧 `streaming-capacity.*`、`api-capture-advanced-*`、`packages-advanced-*`、`consumer-advanced-*` 和 `api-final` 根级初始捕获仅为阶段/历史证据，不用于最终门禁结论。最终 API 只采用上表的 capture/compare 子目录。

远程 CI 未运行。Aspose 商业功能、重新计算、高级 Custom 校验及透视表不在本轮验收范围。
