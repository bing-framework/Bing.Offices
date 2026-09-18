# Provider 测试矩阵

| 项目 | 角色/TFM | 直接 ProjectReference | PackageReference/fixture | net6 Case | net8 Case | 责任与边界 |
| --- | --- | --- | --- | ---: | ---: | --- |
| `Bing.Offices.Tests` | Common Unit / net6.0/net8.0 | Abstractions; Core; ProfileFixtures | MetadataLoadContext; GenFu; IdUtils; DI | 196 | 196 | CSV/映射/日期/公共 SPI；无 Provider；包含 30 个 ReviewFix Core 回归、3 个 Loader 回归和独立 Core 扩展门禁 |
| `Bing.Offices.Npoi.Tests` | NPOI Unit / net6.0/net8.0 | Abstractions; Core; Npoi; ProfileFixtures; ResourceProbe(net8) | NPOI（PrivateAssets=all）; DI; MetadataLoadContext; GenFu; IdUtils | 539 | 539 | NPOI DOM、HSSF/XSSF、Importer/Exporter、扩展、失败合同；含三组原跨 Provider 日期/物理位置回归的 NPOI 分支；Round7 各 TFM 539/539 |
| `Bing.Offices.MiniExcel.Tests` | MiniExcel Unit / net6.0/net8.0 | Abstractions; Core; MiniExcel | MiniExcel; NPOI package（fixture-only, PrivateAssets=all）; DI | 39 | 39 | MiniExcel 日期/行列/RawDate 全量索引/Relation 首匹配及委托异常边界/Converter/取消/结构化错误；不引用 Bing.Offices.Npoi |
| `Bing.Offices.Tests.Integration` | Aggregate Integration / net6.0/net8.0 | Abstractions; Core; Npoi; MiniExcel | NPOI; MiniExcel; MetadataLoadContext; DI | 29 | 29 | 五个跨 Provider 合同、双顺序 DI 先注册合同、真实公共 RoundTrip、API 分类/快照、CSV 公共 IO |
| `Bing.Offices.Npoi.Tests.Integration` | NPOI Integration / net6.0/net8.0 | Abstractions; Core; Npoi | DI | 28（普通）+2 Large | 28（普通）+2 Large | 真实 Excel 文件/模板/图片/导入提交/XLS/XLSX；富工作簿 reopen 验收（合并、批注、图表）；500K/1M 用 `Category=Large` 单独执行 |
| `Bing.Offices.MiniExcel.Tests.Integration` | MiniExcel Integration / net6.0/net8.0 | Abstractions; Core; MiniExcel | DI | 7（普通）+2 Large | 7（普通）+2 Large | 真实同步/异步文件、多 Sheet/动态列/日期、替换、失败、取消、原子提交和临时文件清理；500K/1M 用 `Category=Large` 单独执行 |

- Common 不直接引用 NPOI/MiniExcel；MiniExcel Unit 的 NPOI 为 test-only fixture 包并标记 `PrivateAssets=all`。
- NPOI/MiniExcel Integration 各自只直接引用自身 Provider；Aggregate 才同时引用两个 Provider。
- `AsyncPipelineTest` 保留 `Excel Async staging` 禁并行 Collection；CSV 方法在 Common 直接构造 Core 实现。
- ResourceProbe 仅在 NPOI Unit net8 条件引用，避免 Common 传递污染。
- 普通职责 TRX 使用 `Category!=Large`；Large workload 明确使用 `Category=Large`，只在 workflow dispatch 的 `run_large=true` Job 执行。

## Round 7 NPOI 分支追溯

| 原跨 Provider 行为 | NPOI 直接测试 | Fixture / 关键断言 | TFM / TRX |
| --- | --- | --- | --- |
| 1900/1904 日期系统、负 serial、serial 60/61 | `NpoiProviderBoundaryRegressionTest.RealXlsxDateSystems_ShouldUseCoreSerialContract` | 真实 XSSF XLSX；完整 DateTime/DateTimeOffset 数组 | net6/net8；`artifacts/tests/review-fix-round7/npoi-net6/npoi-net6.trx`、`npoi-net8/npoi-net8.trx` |
| 跳过正文行后的日期、动态列、来源行和转换/验证物理上下文 | `NpoiProviderBoundaryRegressionTest.DataRowStartIndex_ShouldAlignRawDateSerials` | 4 个 InlineData；serial `61,-0.25,-1.25,1`；SourceRows、converter/validation Row/Column | net6/net8；同上 |
| ReadColumns 范围的绝对错误列 | `NpoiProviderBoundaryRegressionTest.ReadColumnRange_ShouldReportAbsolutePhysicalColumnIndex` | 真实 XSSF XLSX；单个 ValueConversion，Row2/Column3/PropertyName Age | net6/net8；同上 |
