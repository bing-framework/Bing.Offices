# Integration Test Report

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 配置：Release；TFM：`net6.0`、`net8.0`。

## 结果

| 范围 | net6.0 | net8.0 | 证据 |
| --- | ---: | ---: | --- |
| 全量 Integration | 39/39，通过 0 失败 | 39/39，通过 0 失败 | `artifacts/tests/review-fix-round2/net6-integration-final/integration-net6-final.trx`；`artifacts/tests/review-fix-round2/net8-integration-final/integration-net8-final.trx` |
| Docs | 不适用 | 10/10，通过 0 失败 | `artifacts/tests/docs/net8/docs-net8-final.trx` |

MiniExcel 集成测试执行真实 `ExportToFileAsync`、`ImportFromFileAsync` 文件路径往返，断言完整实体、错误集合和非空目标文件。全量 Integration 同时保留 NPOI 既有契约，未改变其 XLS/HSSF 与 XLSX/XSSF 分支。
