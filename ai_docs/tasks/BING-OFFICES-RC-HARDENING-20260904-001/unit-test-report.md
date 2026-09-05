# Unit Test Report

## 环境

- OS：Windows 10 `10.0.19045.6466`，x64
- SDK：.NET SDK `10.0.400`
- 运行时：.NET Core 3.1.32、.NET 6.0.36、.NET 8.0.30
- 配置：Release，`--no-restore`
- 测试框架：xUnit 2.4.2 / VSTest adapter

## 当前任务全量结果

| TFM | 总数 | 通过 | 失败 | 跳过 | 退出码 | 结果 |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| `net8.0` | 415 | 414 | 1 | 0 | 1 | 唯一失败为 formal API hash |
| `net6.0` | 415 | 414 | 1 | 0 | 1 | 唯一失败为 formal API hash |
| `netcoreapp3.1` | 415 | 414 | 1 | 0 | 1 | 唯一失败为 formal API hash |

全量结果 TRX：

- `tests/Bing.Offices.Tests/TestResults/rc-hardening-unit-net8-final.trx`
- `tests/Bing.Offices.Tests/TestResults/rc-hardening-unit-net6-final.trx`
- `tests/Bing.Offices.Tests/TestResults/rc-hardening-unit-netcoreapp3.1-final.trx`

唯一失败：`PublicApiContractTest.PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot`。实际 candidate hashes 与 formal baseline 不一致；这是待批准的 API breaking diff，未修改 hash 断言。

## Review 修复回归

- net8：`NpoiXlsxZipPreflightTest` `16/16`，TRX `tests/Bing.Offices.Tests/TestResults/rc-hardening-review-fix-zip-net8.trx`。
- net6：`NpoiXlsxZipPreflightTest` `16/16`，TRX `tests/Bing.Offices.Tests/TestResults/rc-hardening-review-fix-zip-net6.trx`。
- netcoreapp3.1：`NpoiXlsxZipPreflightTest` `16/16`，TRX `tests/Bing.Offices.Tests/TestResults/rc-hardening-review-fix-zip-netcoreapp31.trx`。
- FIX-001 覆盖：`MaxZipCompressionRatio = null` 通过；`0`、负数、`NaN`、正/负无穷被拒绝。

## 覆盖摘要

- 统一异常 code/operation/provider/stage、inner exception、observer 单次通知和 observer failure 隔离。
- 日期默认 ISO、culture independence、显式格式、DateTimeOffset offset policy、1900/1904 serial、公式/Workbook DATE/TIME。
- ZIP preflight、资源限制、DTD/entity、重复 entry、异常路径、取消和源流位置恢复。
- Mapping precedence/cache、CSV RFC4180/公式注入/唯一值、Failure Workbook、原子提交和 NPOI public extension contract。

## 未覆盖或不等价项

- formal API approval 不是 Unit 行为测试，不能由修改断言替代。
- 本任务没有把完整 Failure Workbook 双 DOM 资源矩阵或全部性能矩阵伪装成 Unit 覆盖。
- `get_errors` 对生产和测试目录为 `No errors found`；Release build 仍有第三方 TFM/既有 analyzer warnings。

## 结论

行为回归：`PASS`。正式 API 门禁：`BLOCKED`。Unit 总体状态：`PARTIAL`。
