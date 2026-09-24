# Excel Provider 能力与边界

## Provider 选择

`Bing.Offices.Abstractions` 和 `Bing.Offices.Core` 提供公共契约及 Mapping Plan；`Bing.Offices.Npoi`、`Bing.Offices.MiniExcel` 与 `Bing.Offices.ClosedXml` 是并列 Provider。应用启动时选择一个 Provider：

```csharp
services.AddBingOfficesNpoi();
// 或
services.AddBingOfficesMiniExcel();
// 或
services.AddBingOfficesClosedXml();
```

业务层继续注入 `IExcelImporter`、`IExcelExporter`，不应引用 MiniExcel 的原生 Attribute、Workbook 类型或第二套请求模型。两个扩展都通过 `TryAdd` 注册公共服务和 Excel 实现；如果同一个 `IServiceCollection` 同时调用两者，先注册的实现保留。需要切换 Provider 时应在组合根选择注册分支，而不是依赖注册顺序或按请求解析。

## ClosedXML 第一阶段矩阵

ClosedXML 作为富 XLSX Provider 使用同一套 Workbook Request 和 Core Mapping Plan。它适合模板、样式、合并和公式保存场景；DOM 操作仍是单实例串行边界，不是 MiniExcel 的低内存替代品。

| 能力 | ClosedXML 第一阶段 | 说明 |
| --- | --- | --- |
| 常规 XLSX List/Workbook 导入/导出 | P0 | 真实 `XLWorkbook` 和 XLSX 读回测试；支持多 Sheet、名称/索引选择 |
| Mapping/Profile/JSON/XML、Converter、ValueMap、Validation、Unique、Relations | P0/P1 | 由 Core 生成计划；固定列、动态列和 Entity 多 List Region 关系绑定已接入 |
| 基本字体、填充、对齐、数字格式、列宽、行高 | P1 | 行高单位为 point；null 保留模板高度，显式值覆盖 |
| Merge、Formula 保存/读回、模板 Stream/File | P1 | 公式契约是保存、读取和缓存值，不是完整 Excel 计算引擎 |
| 原子文件提交、取消、流所有权、输入大小/ZIP 预检 | P1 | 使用公共文件提交器和共享 XLSX 预检；DOM 创建前拒绝超限输入 |
| Entity Layout | P1 | 支持固定 Cell、多个同 Sheet/跨 Sheet List Region、Merge、Relations 和模板布局读写；Entity List Region 动态列仍显式不支持 |
| Failure Workbook、原生 Workbook Validation | Unsupported | ClosedXML 第一版在 Provider 创建 DOM 前 fail-fast，不生成部分结果 |
| Chart、PivotTable、Image、XLSM 宏保留、XLS | Unsupported | 不静默丢弃结构；请求在 `XLWorkbook` 创建前返回 UnsupportedFeature |

注册方式：

```csharp
services.AddBingOfficesClosedXml();
```

该扩展与其他 Provider 一样使用 `TryAdd`。需要限制 ClosedXML DOM 准入时可传入 `ClosedXmlProviderOptions`：

```csharp
services.AddBingOfficesClosedXml(options =>
{
    options.MaxConcurrentWorkbooks = 1;
    options.MaxQueuedOperations = 64;
});
```

同时注册多个 Provider 时保留 first-registration-wins；业务代码只依赖 `IExcelImporter`/`IExcelExporter`。

## MiniExcel 第一阶段矩阵

| 能力 | MiniExcel 第一阶段 | 说明 |
| --- | --- | --- |
| 常规 XLSX 导入/导出 | P0 | Stream、文件、字节扩展和多 Sheet；使用公共 Workbook Request |
| Mapping/Profile/JSON/XML | P0 | 由 Core 编译 Mapping Plan，Provider 只消费计划 |
| Converter、ValueMap、Validation、Unique | P0 | 复用公共 converter、validation binding 和 `UniqueTracker` |
| 固定列、别名、忽略列、日期/数字/枚举/bool | P0 | 保持现有请求和结果语义 |
| 动态列 | P0 | 保持延迟枚举；不得默认把整个输入 `ToList()` |
| Sheet name/index、header、trim、空行、缺列/未知列策略 | P0 | 通过现有 `ExcelSheetPolicies` 和 `ExcelImportPolicies` 表达 |
| 基本表头和基本数字格式 | P0 | 不承诺 NPOI 的完整样式 DOM |
| 模板导出 | P1 | 只有实际验证通过的占位符、循环和多 Sheet 语义才可标记支持 |
| 简单宽度、对齐、字体、填充、边框 | P1 | 不能满足的配置必须显式报告不支持 |
| 复杂样式、复杂多级表头、图片、批注 | P2/Unsupported | 不得静默丢弃；按 Provider 能力预检 |
| XLS | Unsupported | MiniExcel 请求 XLS 必须快速失败，不得偷偷转换为 XLSX |
| Chart、NPOI Failure Workbook、完整图片复制 | Unsupported/P2 | 第一阶段不复制 NPOI 专属 DOM 能力；需要时继续使用 NPOI |

NPOI 继续承担 XLS/XLSX、复杂模板、图片、批注、图表、复杂样式和 Failure Workbook 等完整 Workbook DOM 能力。Provider 不支持的请求必须在输出或导入前返回现有结构化 UnsupportedFeature 异常/错误，不能返回一个悄悄缺少内容的成功文件。

## 流、异步和文件边界

输入和输出 Stream 始终由调用方拥有，Provider 不负责关闭调用方流。直接写入调用方 Stream 没有回滚保证；文件 API 继续使用公共 `IFileExportCommitter`，通过同目录暂存、flush、取消检查和最终提交保护已有目标文件。

MiniExcel 的 Async API 必须把真实的输入复制、输出复制和 FileStream flush 保持为异步 IO，并在枚举、行处理、复制和提交边界检查 `CancellationToken`。如果底层 MiniExcel 操作本身是同步的，Async 只描述外围 IO，不得使用 `Task.Run`、`.Result` 或 `.Wait()` 伪装异步。取消或失败必须清理暂存文件，并保持已有目标文件不变。

非定位输入流是否需要暂存、模板流是否要求同步读取、输出流的当前位置和 `leaveOpen` 行为，必须由 Provider 测试固定；不能因为第三方 API 接受 Stream 就推断公共 Stream Contract 已满足。

### MiniExcel 日期和验证边界

MiniExcel 导入在第三方解析前读取 `xl/workbook.xml` 的 `workbookPr/@date1904`，并将该标志传入 Core 日期转换合同；日期单元格的原始 numeric serial 也从 worksheet XML 保留后交给 `ExcelDateParser`，因此 1900 闰年兼容边界、负数线性语义和 1904-01-01 起点由同一规则处理。非 numeric 文本继续按显式日期格式解析。因此 1904 parity 由真实日期 fixture 和双 Provider 结构化合同验证，不把单独的 ZIP 标志读取误称为完整能力。

MiniExcel 的 `ExcelImportValidationMode.ConfiguredRules` 和 `Disabled` 已有明确执行语义：前者执行公共 Mapping/Attribute/Unique 规则，后者跳过这些配置校验但仍执行值转换。MiniExcel 无法读取并执行 NPOI Workbook 原生 Data Validation，因此 `WorkbookRules` 和 `ConfiguredAndWorkbook` 在预检阶段返回结构化 `BingOfficesUnsupportedFeatureException`，不会静默忽略原生规则；需要原生校验时应选择 NPOI Provider。

本任务的 100K 受控往返探针写入根 `artifacts/benchmarks/`；BenchmarkDotNet、测试结果、包、Consumer、资源矩阵和 API 日志均必须写入根 `artifacts/`。500K/1M、生产 2C4G 机器和外部 CI 运行未执行时统一标记 `NOT_VERIFIED`。

## 资源与并存部署

MiniExcel 的目标是降低常规 XLSX 的 Workbook DOM 和峰值内存，但不等于零分配或无限数据规模。仍需应用级内存/CPU 限制，并使用 `ExcelResourceLimits`、最大错误数、最大唯一值和文件提交策略。跨行 Unique 的 Dictionary/HashSet 生命周期应随一次导入结束，不能成为进程级无限增长状态。

NPOI 和 MiniExcel 不应在同一个请求上同时构建 Workbook。发布入口如需并发限制，应在调用 Provider 前共享进程级 `SemaphoreSlim`；Provider 选择和并发门禁属于应用组合根职责。性能与资源结论必须来自真实文件/Stream 测试和 Benchmark，不能以第三方库的宣传吞吐替代证据。
