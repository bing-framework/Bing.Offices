# Docs Test Report

## 结果

- Docs test project：`tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj`
- TFM：`net8.0`
- 结果：`11/11 passed`，退出码 `0`
- TRX：`tests/Bing.Offices.Docs.Tests/TestResults/rc-hardening-docs-final.trx`

## 覆盖

- 文档代码块和 consumer 示例的编译/运行。
- XML documentation extractor、公开 API 示例和迁移片段。
- 异常合同、日期配置、Mapping、DI 和 NPOI extension 相关示例的可消费性。

## 限制

- Docs test 通过不替代无 `ProjectReference` 的 package-only consumer；后者见 `package-consumer-report.md`。
- DateOnly 当前未纳入 API；当前 TFM 矩阵保持不支持，需后续文档持续明确。
- ZIP/XLS/OLE 的资源限制说明仍需随最终 API approval 和发布文档一起确认。

## 结论

Docs executable coverage：`PASS 11/11`；最终 RC 文档收口：`PARTIAL`。
