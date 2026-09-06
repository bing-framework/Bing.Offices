# Integration Test 报告

## 结论

状态：`PASSED`（当前 Windows 支持环境）。net6.0 与 net8.0 各 15/15，零失败、零跳过。

## 环境、命令与结果

- Git：`master` / `9d78ab76e28891ac9b7e0e558f4df669721a4ea6` / dirty（执行前 clean）。
- Windows 10.0.19045 win-x64；SDK 10.0.400；Release。

```powershell
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -c Release -f <net6.0|net8.0> --no-build --no-restore --logger "trx;LogFileName=final-integration-<TFM>.trx" --results-directory <task>/artifacts/test-results/integration
```

| TFM | Passed | Failed | Skipped | 耗时 | Artifact |
| --- | ---: | ---: | ---: | ---: | --- |
| net6.0 | 15 | 0 | 0 | 774 ms | `artifacts/test-results/integration/rc-final-integration-net6.0.trx` |
| net8.0 | 15 | 0 | 0 | 530 ms | `artifacts/test-results/integration/rc-final-integration-net8.0.trx` |

覆盖真实 XLS/XLSX/CSV Stream 与文件 IO、关闭重开、DI、映射、Converter/Validation、文件锁、原子提交、临时文件清理、失败工作簿、取消和资源错误。netstandard2.0 由独立包消费者编译验证；netcoreapp3.1 的真实包运行由 PackageConsumer 验证。

Linux/macOS runner 当前不可用，未执行跨 OS Integration；Windows 文件锁语义不能由本机结果外推到 Unix。外部数据库不在本仓库/本任务调用链中，没有连接生产数据库。
