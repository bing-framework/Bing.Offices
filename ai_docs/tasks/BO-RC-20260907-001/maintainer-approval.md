# BO-RC-20260907-001 维护者审批记录

- `approvedBy`: `jian玄冰`
- `approvedAt`: `2026-09-08T14:43:03+08:00`
- 服务器预算：`2C4G`
- 逻辑请求并发档位：`1 / 4 / 16 / 64`
- 最大实际并行度：`1`
- 默认 staging 策略：`TempFile`

## API 基线

批准 `artifacts/api-compare-final/api-candidate.json` 作为 net6.0/net8.0 正式公共 API 基线。批准范围包括 Async API、NPOI Provider User API、NPOI 扩展 namespace 迁移，以及 `artifacts/api-compare-final/api-diff.json` 中列出的 `ExcelMappingDiagnostic` 和旧 diagnostics 重载删除。

删除项必须在发布说明中保留 Breaking Change 和迁移说明；本记录不表示兼容性风险消失。

## 资源预算策略

资源验收采用“逻辑请求并发”和“实际 NPOI DOM 并行度”分离的模型。服务器可以接收 4、16、64 个请求，但进入 NPOI DOM 管线的实际并行度固定为 1；超出部分排队，不得创建额外的并行 DOM 工作簿。

资源策略的最终 `APPROVED` 状态以 formal v8 的 100K 活动槽位矩阵和 smoke v7 的排队请求完成证据为依据：36 个 formal 单元全部完成、无 `budget-failed`、无错误且清理残留为 0；smoke v7 的 36 个排队单元全部完成。formal v8 不宣称 64 个 100K DOM 工作簿同时或逐个完成，生产请求入口必须在调用库之前执行实际并行度为 1 的排队。
