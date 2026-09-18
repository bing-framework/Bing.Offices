# NPOI Integration 报告

- net6.0：PASS 28，TRX：`artifacts/tests/review-round3/npoi-integration-net6/npoi-integration-net6.trx`
- net8.0：PASS 28，TRX：`artifacts/tests/review-round3/npoi-integration-net8/npoi-integration-net8.trx`
- 覆盖 XLS/HSSF 与 XLSX/XSSF 的真实文件、模板、失败提交和取消清理；富工作簿 reopen 断言合并单元格、批注和图表；专属能力不提升为公共合同。
- `LargeFileRoundTrip_ShouldUseRealPath` 另有 500K/1M 两个 Theory Case（`Category=Large`）；本机本轮未执行，workflow dispatch 的 Large Job 单独验证。
