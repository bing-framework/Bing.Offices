# 生产符号到测试追溯

| 生产符号 | 关键行为 | 测试项目/方法 |
|---|---|---|
| `MiniExcelExcelExporter.Export` | XLSX、多 Sheet、映射、unsupported preflight | `Bing.Offices.Tests/MiniExcelProviderTest.cs`: `RoundTrip_ShouldSupportMultipleSheetsAndMapping`, `Xls_ShouldBeRejectedAsUnsupported`, `UnsupportedPresentationOptions_ShouldFailBeforeWriting` |
| `MiniExcelExcelExporter.ExportAsync` | 真实异步写入、取消 | `AsyncRoundTrip_ShouldUseMiniExcelAsyncApi`, `PreCanceledAsync_ShouldNotWriteOrRead` |
| `MiniExcelExcelExporter.ExportToFileAsync` | 原子文件提交与真实路径 | `Bing.Offices.Tests.Integration/MiniExcelProviderIntegrationTest.cs`: `ExportToFileAsync_ThenImportFromFileAsync_ShouldRoundTripRealXlsx` |
| `MiniExcelExcelImporter.Import` | ZIP preflight、映射、资源错误、关系 | `ResourceLimit_ShouldFailBeforeMiniExcelParser`, `Relations_ShouldHonorCustomKeyComparer`, `FailureWorkbook_ShouldBeRejectedBeforeParser` |
| `MiniExcelExcelImporter.ImportAsync` | QueryAsync、异步 roundtrip | `AsyncRoundTrip_ShouldUseMiniExcelAsyncApi` |
| `MiniExcelMappingPlanBuilder` | Core plan、动态列合并与缓存 invoker | `DynamicColumns_ShouldRoundTripWithNamedConverter`；全量映射回归 |
| `MiniExcelValueAdapter` | fixed/dynamic converter、ValueMap、类型转换 | MiniExcel dynamic roundtrip；全量 Unit 映射/转换测试 |
| `MiniExcelXlsxPreflight` | MiniExcel provider ZIP/XML 安全边界 | `ResourceLimit_ShouldFailBeforeMiniExcelParser`；Core/NPOI 共享预检 27/27 |
| `ExcelXlsxZipPreflight.Validate` | ZIP entry、路径、DTD、压缩比、取消、预算 | `NpoiXlsxZipPreflightTest` 27/27（net6/net8） |
| `AddBingOfficesMiniExcel` | null/DI 生命周期/默认服务 | `AddBingOfficesMiniExcel_ShouldRegisterProviderServices` |
| `IExcelExporter`/`IExcelImporter` public surface | 无第三方类型泄漏、双 TFM additive API | `PublicApiContractTest`；API snapshot capture/compare |
| MiniExcel package surface | 真实 PackageReference 消费 | `Consumer.Net6`/`Consumer.Net8` 程序，日志 `package-consumer-ok` |

未列出的 Provider 内部 helper 只由上述职责测试和全量回归覆盖，不作为公共 API 承诺。
