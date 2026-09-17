# 需求矩阵

| 需求 | 状态 | 实现/证据 | 直接验证 |
|---|---|---|---|
| 独立 `Bing.Offices.MiniExcel` 项目，双 TFM | 已完成 | `src/Bing.Offices.MiniExcel/Bing.Offices.MiniExcel.csproj`，net6.0/net8.0 | Release solution build；API snapshot |
| 稳定 MiniExcel 依赖 | 已完成 | `MiniExcelPackageVersion=1.46.0` | MiniExcel nuspec 依赖检查 |
| XLSX、异构多 Sheet、Stream | 已完成 | `MiniExcelExcelExporter`/`MiniExcelExcelImporter` 使用真实 `SaveAs`/`Query` | `RoundTrip_ShouldSupportMultipleSheetsAndMapping`；集成真实文件测试 |
| File/byte[] 便利入口与原子提交 | 已完成 | 复用 Core `IFileExportCommitter` 和 `ExcelStreamExtensions` | `MiniExcelProviderIntegrationTest`；Package Consumer |
| 固定列、动态列、ValueMap、Converter | 已完成 | 复用 Core mapping plan 和 `MiniExcelValueAdapter` | `DynamicColumns_ShouldRoundTripWithNamedConverter`；全量 Unit |
| Attribute/Profile/Fluent/JSON/XML 映射契约 | 已完成 | Provider 不新增映射旁路，使用 `IExcelMappingPlanFactory` | 共享映射回归 739/739；Integration 39/39 |
| Validation、Unique、错误集合 | 已完成 | 复用 Core validation binding/error contract | 全量 Unit/Integration |
| Relations 与自定义比较器 | 已完成 | MiniExcel 导入后复用关系绑定语义 | `Relations_ShouldHonorCustomKeyComparer` |
| 真实异步和取消 | 已完成 | `SaveAsAsync`、`QueryAsync`、真实 Stream async copy；无 `Task.Run`/`.Result`/`.Wait()` | `AsyncRoundTrip_ShouldUseMiniExcelAsyncApi`；`PreCanceledAsync_ShouldNotWriteOrRead` |
| XLSX ZIP/XML 安全预检 | 已完成 | Core `ExcelXlsxZipPreflight`，MiniExcel/NPOI wrapper | NPOI 预检 27/27（双 TFM）；MiniExcel resource test |
| DI 注册与双 Provider 包消费者 | 已完成 | `AddBingOfficesMiniExcel`，首注册生效保持与 NPOI 一致 | DI Unit；net6/net8 PackageReference Consumer |
| 100K 真实 IO 证据 | 已完成 | controlled probe 写入并读回 100,000 行 | `benchmark-miniexcel-probe-100k-final.json` |
| 500K/1M 完整 BenchmarkDotNet | 受限 | 未运行长测；避免以单次受限环境结果冒充门禁 | `benchmark-report.md` |
| 模板、复杂样式、merge/multi-header、图片、批注、Chart | 明确不支持 | Preflight 在写入前抛 `BingOfficesUnsupportedFeatureException` | `UnsupportedPresentationOptions_ShouldFailBeforeWriting`；XLS/Failure Workbook tests |
| Failure Workbook | 明确不支持 | Import preflight 拒绝非 `None` 模式 | `FailureWorkbook_ShouldBeRejectedBeforeParser` |

## 当前限制

- MiniExcel Provider 的导出默认只承诺普通 XLSX、默认表头/数据行位置和无样式数据值。
- 500K/1M 以及 BenchmarkDotNet 自动生成项目被 NuGet SSL/凭证环境阻断；100K controlled probe 可复现，但不构成跨 Provider 性能结论。
- API baseline 的 `approvedBy=task-plan-additive-api` 是本任务记录的 additive 基线标记，不代表人工成员审批。
