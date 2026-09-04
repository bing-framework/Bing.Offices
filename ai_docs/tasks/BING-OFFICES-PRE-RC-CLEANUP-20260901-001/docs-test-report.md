# 文档测试报告

## 已执行结果

| TFM | 项目 | 总数 | 通过 | 失败 | 退出码 | TRX |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| net8.0 | `tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj` | 11 | 11 | 0 | 0 | `artifacts/docs-net8-rerun/docs-net8-rerun.trx` |

命令：

```powershell
 dotnet test .\tests\Bing.Offices.Docs.Tests\Bing.Offices.Docs.Tests.csproj -c Release --no-restore -f net8.0 --logger "trx;LogFileName=docs-net8-rerun.trx"
```

## 覆盖范围

- Markdown C# fence 提取、独立编译和执行；
- package-style 文档路径执行；
- `AddBingOfficesNpoi` provider-neutral DI 示例；
- Mapping v2 方向请求；
- JSON/XML 文档迁移和 Stream ownership；
- XLS/XLSX metadata roundtrip；
- ASP.NET Core 上传成功/失败响应；
- 文档 consumer 不暴露 NPOI 类型。

该 Docs 测试项目仍包含对工作区生产项目的 `ProjectReference`，因此本报告不替代独立 nupkg consumer 证据。独立包消费者见 `package-consumer-report.md`。

## 状态

`VERIFIED`（net8.0）。
