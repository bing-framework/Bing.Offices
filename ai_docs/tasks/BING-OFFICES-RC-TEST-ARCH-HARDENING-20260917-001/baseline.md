# RC-000 基线

- Task-ID：`BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001`
- 捕获日期：2026-09-17（Asia/Shanghai）
- 当前工作区包含本次迁移修改；未把历史报告当作本次执行结果。

## Git、SDK、Solution
```text
ece61d8f9f7c6fcd1efe1dcf68cc8224165c80cc
feat/miniexcel-provider
 M .github/workflows/ci.yml
 M Bing.Offices.sln
 M src/Bing.Offices.Abstractions/AssemblyInfo.cs
 M src/Bing.Offices.Core/AssemblyInfo.cs
 M src/Bing.Offices.MiniExcel/AssemblyInfo.cs
 M src/Bing.Offices.Npoi/AssemblyInfo.cs
 M tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj
 D tests/Bing.Offices.Tests.Integration/CsvAsyncFileIntegrationTest.cs
 D tests/Bing.Offices.Tests.Integration/ExcelAsyncFileIntegrationTest.cs
 D tests/Bing.Offices.Tests.Integration/ExcelImporterIntegrationTest.cs
 D tests/Bing.Offices.Tests.Integration/MiniExcelProviderIntegrationTest.cs
 D tests/Bing.Offices.Tests/AsyncPipelineTest.cs
 M tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj
 D tests/Bing.Offices.Tests/DependencyContainer.cs
 D tests/Bing.Offices.Tests/ExcelMappingConfigurationLoaderTest.cs
 D tests/Bing.Offices.Tests/ExcelP0RegressionTest.cs
 D tests/Bing.Offices.Tests/ExcelWorkbookRequestTest.cs
 M tests/Bing.Offices.Tests/MappingProfileRegistryTest.cs
 D tests/Bing.Offices.Tests/MiniExcelProviderContractTest.cs
 D tests/Bing.Offices.Tests/MiniExcelProviderTest.cs
 D tests/Bing.Offices.Tests/NpoiCellExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiCellStyleExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiFailureWorkbookPreflightTest.cs
 D tests/Bing.Offices.Tests/NpoiFontExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiRowExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiServiceCollectionExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiSheetExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiSheetPictureExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiWorkbookExtensionsTest.cs
 D tests/Bing.Offices.Tests/NpoiXlsxZipPreflightTest.cs
 D tests/Bing.Offices.Tests/PublicApiContractTest.cs
 D tests/Bing.Offices.Tests/PublicExtensionCoverageTest.cs
 D tests/Bing.Offices.Tests/ReviewFixRegressionTest.cs
 D tests/Bing.Offices.Tests/StreamPipelineTest.cs
 D tests/Bing.Offices.Tests/TemplateAsyncBoundaryTest.cs
 D tests/Bing.Offices.Tests/TestBase.cs
?? ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/
?? tests/Bing.Offices.MiniExcel.Tests.Integration/
?? tests/Bing.Offices.MiniExcel.Tests/
?? tests/Bing.Offices.Npoi.Tests.Integration/
?? tests/Bing.Offices.Npoi.Tests/
?? tests/Bing.Offices.Tests.Integration/MiniExcelProviderContractTest.cs
?? tests/Bing.Offices.Tests.Integration/PublicApiContractTest.cs
?? tests/Bing.Offices.Tests/PublicCoreExtensionCoverageTest.cs

6.0.428 [C:\Program Files\dotnet\sdk]
7.0.410 [C:\Program Files\dotnet\sdk]
8.0.419 [C:\Program Files\dotnet\sdk]
9.0.312 [C:\Program Files\dotnet\sdk]
10.0.201 [C:\Program Files\dotnet\sdk]
10.0.302 [C:\Program Files\dotnet\sdk]
10.0.400 [C:\Program Files\dotnet\sdk]
10.0.401 [C:\Program Files\dotnet\sdk]
Microsoft.AspNetCore.App 6.0.36 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 7.0.20 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 8.0.25 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 8.0.30 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 9.0.14 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 9.0.19 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 10.0.5 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 10.0.10 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 10.0.11 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 10.0.12 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.NETCore.App 6.0.36 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 7.0.20 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 8.0.25 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 8.0.30 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 9.0.14 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 9.0.19 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 10.0.5 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 10.0.10 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 10.0.11 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 10.0.12 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.WindowsDesktop.App 6.0.36 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 7.0.20 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 8.0.25 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 8.0.30 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 9.0.14 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 9.0.19 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 10.0.5 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 10.0.10 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 10.0.11 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 10.0.12 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]

项目
--
benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj
build\ApiSnapshot\ApiSnapshot.csproj
build\BuildScript.csproj
src\Bing.Offices.Abstractions\Bing.Offices.Abstractions.csproj
src\Bing.Offices.Core\Bing.Offices.Core.csproj
src\Bing.Offices.MiniExcel\Bing.Offices.MiniExcel.csproj
src\Bing.Offices.Npoi\Bing.Offices.Npoi.csproj
tests\Bing.Offices.Docs.Tests\Bing.Offices.Docs.Tests.csproj
tests\Bing.Offices.MiniExcel.Tests.Integration\Bing.Offices.MiniExcel.Tests.Integration.csproj
tests\Bing.Offices.MiniExcel.Tests\Bing.Offices.MiniExcel.Tests.csproj
tests\Bing.Offices.Npoi.Tests.Integration\Bing.Offices.Npoi.Tests.Integration.csproj
tests\Bing.Offices.Npoi.Tests\Bing.Offices.Npoi.Tests.csproj
tests\Bing.Offices.ProfileFixtures\Bing.Offices.ProfileFixtures.csproj
tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj
tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj

version.props=77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1
version.dev.props=282D9A7C70D79FFC5F354DA57035A479B08A8485D1AF7637B25AF6C815DCAF31

--- src\Bing.Offices.Abstractions\AssemblyInfo.cs
[assembly: InternalsVisibleTo("Bing.Offices.Tests")]
[assembly: InternalsVisibleTo("Bing.Offices.Npoi.Tests")]
--- src\Bing.Offices.Core\AssemblyInfo.cs
[assembly: InternalsVisibleTo("Bing.Offices.Tests")]
[assembly: InternalsVisibleTo("Bing.Offices.Npoi.Tests")]
--- src\Bing.Offices.MiniExcel\AssemblyInfo.cs
[assembly: InternalsVisibleTo("Bing.Offices.MiniExcel.Tests")]
[assembly: InternalsVisibleTo("Bing.Offices.MiniExcel.Tests.Integration")]
--- src\Bing.Offices.Npoi\AssemblyInfo.cs
[assembly: InternalsVisibleTo("Bing.Offices.Npoi.Tests")]
[assembly: InternalsVisibleTo("Bing.Offices.Npoi.Tests.Integration")]


```

## 版本哈希

- `version.props`：`77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1`
- `version.dev.props`：`282D9A7C70D79FFC5F354DA57035A479B08A8485D1AF7637B25AF6C815DCAF31`
- 两个版本文件均未修改；CI 改为 checkout 后捕获并在 gate 结束时比较。

## 测试基线

- 迁移前静态声明：Common `532`、旧 Integration `32`；这是计划提供的静态盘点，未能在迁移前重新执行双 TFM，因此迁移前 TRX/展开 Case 清单为 `NOT_VERIFIED`。
- 迁移后执行：六个职责项目均完成 net6/net8，TRX 路径见 `unit-test-report.md` 与 integration reports。
- 本轮使用 `method-identity-audit.py` 的 UTF-8 锚定解析器关联独立 Fact/Theory 属性与方法声明：HEAD 两个旧测试项目为 `569` 个方法，当前六个职责项目为 `592` 个，`After=Before+Added-Removed` 对账为 Added=`23`、Removed=`0`。当前职责分项为 Common `166`、NPOI Unit `332`、MiniExcel Unit `35`、Aggregate `28`、NPOI Integration `23`、MiniExcel Integration `8`；其中 144 个 unchanged、391 个 moved、34 个 renamed/moved。完整逐身份清单见 `method-identity-audit.md`。这些是方法声明数，不是 Theory 展开 Case 数；Theory 数据仍以矩阵和 TRX 分开记录。

## 风险与证据边界

- 根 `artifacts/` 仅作为本次证据目录；src/tests/benchmarks 下 nested 产物扫描结果写入最终报告。
- 资源探针为本机 36 场景、rowCount=1000，机器参数和生产批准不等价。
- API compare 当前候选 identity 与 baseline source manifest 不一致；见 `api-diff.md`，不能写成 Release PASS。
