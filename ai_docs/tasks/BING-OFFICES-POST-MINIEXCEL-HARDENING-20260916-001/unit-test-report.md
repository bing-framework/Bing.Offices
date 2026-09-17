# Unit Test Report

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 配置：Release；TFM：`net6.0`、`net8.0`。

## 结果

| 范围 | net6.0 | net8.0 | 证据 |
| --- | ---: | ---: | --- |
| 全量 Unit | 761/761，通过 0 失败 | 761/761，通过 0 失败 | `artifacts/tests/review-fix-round2/net6-final/unit-net6-final.trx`；`artifacts/tests/review-fix-round2/net8-final/unit-net8-final.trx` |
| MiniExcel 专项 | 25/25，通过 0 失败 | 25/25，通过 0 失败 | `FullyQualifiedName~MiniExcelProviderTest` 双 TFM 直接职责运行；真实日期矩阵和固定列重排均包含在内 |
| NPOI Preflight | 27/27，通过 0 失败 | 27/27，通过 0 失败 | `artifacts/tests/preflight/net6/preflight-net6.trx`；`artifacts/tests/preflight/net8/preflight-net8.trx` |
| Cross-provider contract | 5/5，通过 0 失败 | 5/5，通过 0 失败 | 由上述全量 TRX 中的 `MiniExcelProviderContractTest` 直接职责方法覆盖；既有 cross-provider TRX 保留在 `artifacts/tests/unit/net6-cross-provider/`、`artifacts/tests/unit/net8-cross-provider/` |

## 覆盖重点

MiniExcel 专项直接覆盖多 Sheet/Mapping、异步 API、DI、XLS fail-fast、unsupported 零输出、四种 `ValidationMode` 边界、ValueMap、Validation、Unique、固定/动态列 converter 物理 `RowIndex`/`ColumnIndex`、Core 日期/1900/1904 标志、真实 1900/1904 serial workbook（同一 fixture 同时映射 `DateTime` 与固定 `+08:00` `DateTimeOffset`，并由 NPOI/MiniExcel 使用同一 request 完整对比）、ZIP preflight、关系 comparer、资源限制、失败工作簿边界、预取消、中途导出/导入取消、流所有权、原子文件取消/临时文件清理。

跨 Provider 合同比较共同标量、动态列、列序、Decimal、Nullable、Enum、Bool、Date、ValueMap、Converter、Validation、Relation 和结构化 Error；样式能力明确记录为 NPOI 支持、MiniExcel 结构化拒绝。
