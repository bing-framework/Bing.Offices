# 批量导入与失败工作簿

## 功能与入口

批次接口通过串行回调交付成功实体和结构化错误，适合分批写数据库或消息队列。批次回调不提供数据库事务；此前已经交付的数据不会因后续失败自动回滚。

| 入口 | Provider | 解析和结果语义 |
| --- | --- | --- |
| `IExcelBatchImporter.ImportBatches*` | ExcelDataReader，XLS/XLSX/XLSB | 单 Sheet 前向读取，使用 `ExcelBatchImportRequest<TItem>`，默认批大小 1000 |
| `IExcelWorkbookBatchImporter.ImportWorkbookBatches*` | NPOI、ClosedXML、ExcelDataReader | 按源文件物理 Sheet 顺序交付；NPOI/ClosedXML 先完整 DOM 导入再投影，ExcelDataReader 前向读取 |
| `IExcelEntityBatchImporter.ImportEntityBatches*`、`ImportForTemplateBatches*` | NPOI、ClosedXML | 按 Entity List Region 交付；默认 Buffered + Complete，另有受限 Forward 与 SummaryOnly |

NPOI 的格式为 XLS/XLSX，ClosedXML 为 XLSX。MiniExcel 不注册上述分批导入接口。不能把“提供批次回调”直接理解为“没有 DOM”或恒定内存。

## 用法示例

以下片段使用已经注册的 NPOI/ClosedXML/ExcelDataReader 服务容器 `provider`、调用方输入流 `source` 和取消令牌。`OrderWorkbook`/`OrderRow` 模型见[快速开始](../getting-started.md)，命名空间为 `Bing.Offices.Imports` 和 `Microsoft.Extensions.DependencyInjection`。

```csharp
var request = ExcelImport.Workbook<OrderWorkbook>(book =>
    book.Sheet("订单", model => model.Orders));
var importer = provider.GetRequiredService<IExcelWorkbookBatchImporter>();
var summary = await importer.ImportWorkbookBatchesAsync(source, request,
    (batch, token) =>
    {
        token.ThrowIfCancellationRequested();
        Console.WriteLine($"{batch.SheetName}：{batch.Items.Count} 条，{batch.Errors.Count} 个错误");
        return System.Threading.Tasks.Task.CompletedTask;
    }, cancellationToken);
```

`ExcelWorkbookImportBatch` 包含 `SheetName`、`ItemType`、全局一基 `BatchNumber`、一基 `FirstRowIndex`、`Items`、`Errors`。异构 Sheet 的 `Items` 为对象集合，应先核对 `ItemType` 再转换。每批处理完才继续，不应在回调中启动未等待的后台工作。

ExcelDataReader 单表可构造 `new ExcelBatchImportRequest<OrderRow>("订单") { BatchSize = 500 }`，交给 `IExcelBatchImporter.ImportBatches*`。完整 Workbook 的逐 Sheet 批大小由 Sheet 导入构建器配置。

## Entity 解析模式与资源

`ExcelEntityBatchImportOptions` 同时控制批大小、导入选项、解析模式和结果保留模式：

| 配置 | 实际路径 |
| --- | --- |
| `Buffered` + `Complete`（默认） | 先得到完整实体，再分区投影批次；最终摘要保留完整结果 |
| 显式 `Forward` | 使用受限前向解析器；`Complete` 仍可保留完整结果，不自动成为摘要模式 |
| `SummaryOnly`，即使配置 `Buffered` | 切换受限 Forward 路径，不保留完整列表/关系导航集合；仍需满足 Forward 的布局约束 |

例如摘要模式选项为 `new ExcelEntityBatchImportOptions(1000, new ExcelEntityImportOptions(), ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly)`。布局及模板参数沿用[实体布局](entity-layout.md)。无法归属明细区域的固定字段、结构或关系错误可能通过诊断批次交付，不能只处理有成功实体的批次。

Forward 的固定字段、多区域、命名锚点、分页/分组和关系支持有组合限制；详见[Provider Forward 边界](09-providers.md#forward-entity-增量能力)。不支持的组合在布局预检阶段拒绝；不能据某个单项已实现推断任意组合均支持。XLS 路径还包含 HSSF 格式探测与适配，失败输出可能重新打开原始 DOM；前向导入不等于全流程零 DOM。

## Entity 前向临时磁盘预算

NPOI 和 ClosedXML 支持可选的单次导入临时存储上限：

```csharp
var options = new ExcelEntityBatchImportOptions(
    1000, new ExcelEntityImportOptions(),
    ExcelEntityBatchParsingMode.Forward,
    ExcelEntityBatchResultMode.SummaryOnly,
    maxTemporaryStorageBytes: 256L * 1024 * 1024);
```

预算按同时存活的受管文件逻辑长度合计：输入暂存、共享字符串 UTF-8 正文与定长索引，以及 NPOI XLS 的逐 Sheet XML 和 XLSX 投影。覆盖写不重复扣费，缩小或释放文件会归还额度；超限在扩容前抛出结构化 `ResourceLimitExceeded`，随后释放临时文件。源流保持调用方所有权；已交付批次不回滚。

`null` 保持原行为，零或负数无效。指定预算时共享字符串使用磁盘索引，包括 `Forward + Complete`；Complete 的实体集合仍会保留。`Buffered + SummaryOnly` 沿用前向路径并保留预算。当前不支持预算与 `Buffered + Complete`、模板入口、失败工作簿输出组合，这些请求在读取前返回 `UnsupportedFeature`，不会静默忽略预算。

该预算不包含调用方源/目标流、文件系统块开销、引擎缓存、实体堆和错误行快照，也不替代 `MaxInputBytes`、ZIP/XML、Unique、关系键及失败产物的各项限制。临时文件占用受控不等于全流程恒定内存。

## 失败工作簿

`ExcelImportFailureOptions` 提供 `None`、`ErrorRowsOnly` 和 `AnnotatedOriginal` 策略。启用时必须在 `Destination`（调用方拥有的流）与 `DestinationPath`（原子文件提交）中二选一。Workbook 使用 `.FailureWorkbook(options)`，单表请求用 `FailureOptions`，Entity 用 `ExcelEntityImportOptions` 的五参数构造函数传入。

例如选项为 `new ExcelImportFailureOptions { Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly, DestinationPath = "orders-errors.xlsx", MaxSerializedBytes = 16L * 1024 * 1024 }`。没有导入错误时不会写出失败产物。

- NPOI/ClosedXML 的完整 Workbook 和 Buffered Entity 批次先完成完整导入及可选失败产物提交，再交付投影批次。回调异常不会撤销已提交文件。
- ExcelDataReader 分批入口与 Entity Forward 入口在解析、回调完成后再生成失败产物；回调失败不继续提交。Entity Forward 的 `ErrorRowsOnly` 和 `AnnotatedOriginal` 会额外打开原始工作簿生成输出。
- ExcelDataReader 的 `ErrorRowsOnly` 生成最小 XLSX，保留有限样式、显式列宽与隐藏状态；`AnnotatedOriginal` 仅支持 XLSX 的部件保留与原坐标批注，XLS/XLSB 拒绝。条件格式等细节见[导入校验](import-validation.md#失败工作簿能力边界)与 Provider 矩阵。
- `MaxSerializedBytes` 限制最终序列化字节，不限制原始 DOM、解压数据、实体集合或所有暂存内存。`MaxCandidateErrorRows`、复制单元格/图片、对象估算及输入的 `ExcelResourceLimits` 是不同上限。

## 核心实现、依赖与限制

关键路径：

- `src/Bing.Offices.Abstractions/Bing/Offices/Imports/{Batch,Failures,Resources}` 与 `Entities/Import`：请求、回调、摘要、错误和限额。
- `src/provider-shared/{ExcelWorkbookBatchProjection.cs,ExcelEntityBatchProjection.cs,ExcelEntityForwardRelationTracker.cs,ExcelForwardXlsxReader.cs,ExcelForwardXlsxFailureWorkbookWriter.cs}`：投影、关系与前向解析/失败输出。
- NPOI 的 `Entities/NpoiEntityBatchImporter.cs`、`NpoiForwardXlsxEntityBatchImporter.cs`、`NpoiForwardHssfEntityBatchImporter.cs`，以及 ClosedXML 的 `Entities/ClosedXmlEntityBatchImporter.cs`。
- `src/Bing.Offices.ExcelDataReader/Bing/Offices/ExcelDataReader`：实际分批读取、关系绑定、失败行捕获与 ZIP 批注输出。

依赖对应 Provider 及 Core 映射、验证和文件提交服务。输入流由调用方拥有；模板按 `LeaveOpen` 处理。唯一值与关系键索引可能随数据增长，必须配置合适上限。异步入口保留引擎同步解析阶段；取消及文件原子性见[异步 IO](async-io.md)与[提交器](file-committers.md)。
