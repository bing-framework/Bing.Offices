# Unit 测试报告

| 职责 | net6 | net8 | TRX |
| --- | ---: | ---: | --- |
| Common Unit | PASS 196 | PASS 196 | `artifacts/tests/review-round3/common-net6/common-net6.trx`; `artifacts/tests/review-round3/common-net8/common-net8.trx` |
| NPOI Unit | PASS 539 | PASS 539 | `artifacts/tests/review-fix-round7/npoi-net6/npoi-net6.trx`; `artifacts/tests/review-fix-round7/npoi-net8/npoi-net8.trx` |
| MiniExcel Unit | PASS 39 | PASS 39 | `artifacts/tests/review-fix-round7/miniexcel-net6/miniexcel-net6.trx`; `artifacts/tests/review-fix-round7/miniexcel-net8/miniexcel-net8.trx` |

- 编译：六项目解决方案 Release build 通过，0 warning/0 error（Consumer net6 的 EOL warning 单独记录）。
- Case 数为实际 TRX 展开结果，不是 Fact/Theory 声明数。
- `Category=Large` 仍按计划在普通 Unit 门禁外单独执行；本报告不把本机 Large 结果替代生产资源门禁。
- Round 3 NPOI Unit net6.0 的历史全量 TRX 曾有 1 个失败；Round 7 重新构建并运行 NPOI Unit 双 TFM，全量各 539/539，通过结果见上表。

## Round 4 修复复验

- MiniExcel Unit 新增 Relation 重复 ParentKey 委托异常边界回归；net6.0/net8.0 均 `39/39 PASS`，TRX 分别为 `artifacts/tests/review-fix-round5/miniexcel-net6/miniexcel-net6.trx`、`artifacts/tests/review-fix-round5/miniexcel-net8/miniexcel-net8.trx`。
- Round 3 的 36 Case 表格和 Round 4 的 38 Case TRX 保留为历史职责基线；Round 5 新增 Case 不回写历史 TRX。

## Round 7 修复复验

- NPOI 新增三组原跨 Provider 分支回归，net6.0/net8.0 均 `539/539 PASS`；MiniExcel 在撤回 RawDate 过滤优化后仍为 `39/39 PASS`。
- 该报告的测试结果来自 Round 7 新 TRX；迁移前双 TFM TRX 仍为 `NOT_VERIFIED`。
