# Docs Test 报告

## 结论

状态：`PASSED`。最终文档测试 net8.0 为 10/10，零失败、零跳过。

## 证据

```powershell
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-restore --logger "trx;LogFileName=final-docs-after-contracts-net8.0.trx" --results-directory <task>/artifacts/test-results/docs
```

Artifact：`artifacts/test-results/docs/rc-final-docs-net8.0.trx`；耗时约 2 s。

测试从 `docs/excel` Markdown 原文提取全部 10 个 C# fence，逐个编译并执行；同时覆盖 DI、Workbook Request、XLS/XLSX metadata 重开、Mapping Profile、JSON/XML v2、typed CSV、动态列和 ASP.NET Core 上传示例。新增异常、日期、NPOI 扩展三章已加入文档枚举并随测试输出加载。

本项目仍使用生产 `ProjectReference`，因此本报告只证明文档/源码一致，不替代 nupkg 证据；独立包消费见 `package-consumer-report.md`。
