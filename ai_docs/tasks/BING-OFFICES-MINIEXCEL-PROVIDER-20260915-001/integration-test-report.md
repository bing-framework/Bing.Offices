# 集成与回归报告

## 命令与结果

| 命令类别 | net6.0 | net8.0 |
|---|---:|---:|
| Unit `Bing.Offices.Tests` | 739/739 | 739/739 |
| Integration `Bing.Offices.Tests.Integration` | 39/39 | 39/39 |
| MiniExcel filtered Unit | 10/10 | 10/10 |
| NPOI shared XLSX preflight filtered | 27/27 | 27/27 |
| Docs tests | 不适用 | 10/10 |

核心命令均使用 `-c Release --no-build --no-restore`，并在需要反射 API 测试时显式设置 `NUGET_PACKAGES=C:\Users\jianx\.nuget\packages`。

新增 `MiniExcelProviderIntegrationTest.ExportToFileAsync_ThenImportFromFileAsync_ShouldRoundTripRealXlsx` 通过 `AddBingOfficesMiniExcel()`、真实临时 `.xlsx` 路径、异步文件写入和异步文件读取，断言完整业务值、无错误和非空文件。

## 包消费者

- .NET 6.0.36：`package-consumer-ok ... packages=Bing.Offices.Npoi+Bing.Offices.MiniExcel/2.0.0 ... miniExcelBytes=4211`。
- .NET 8.0.30：`package-consumer-ok ... packages=Bing.Offices.Npoi+Bing.Offices.MiniExcel/2.0.0 ... miniExcelBytes=4208`。
- 两个消费者都使用真实 `.nupkg` `PackageReference`，不是 ProjectReference；同时验证 NPOI 与 MiniExcel 并存时接口和扩展均可用。

## 互操作限制

MiniExcel 集成测试验证的是 MiniExcel 真实生成/读取的 XLSX 和 Bing.Offices 业务结果，不宣称与 NPOI 的 ZIP 字节或样式字节一致。复杂样式、模板、图片、批注、Chart、Failure Workbook 已在 Provider capability matrix 中明确拒绝。
