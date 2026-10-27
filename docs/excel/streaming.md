# XLSX 低内存导入导出

低内存服务显式注册，使用 SXSSF 行窗口导出、ZIP/XML 逐行导入。现有 `IExcelExporter`、`IExcelImporter` 及 `AddBingOfficesNpoi` 保持原来的 DOM 行为。需要模板、图片、图表或复杂合并时，继续使用原服务。

```csharp
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

var services = new ServiceCollection();
services.AddBingOfficesNpoiStreaming();
using var provider = services.BuildServiceProvider();
var exporter = provider.GetRequiredService<IExcelStreamingExporter>();
var importer = provider.GetRequiredService<IExcelStreamingImporter>();
```

## 导出

复用 `ExcelExport.Workbook`，支持多 Sheet 顺序输出、映射、固定列和预声明动态列。数据源每个 Sheet 只枚举一次；传入惰性 `IEnumerable<T>`，避免调用方先把数据全部装入 List。

```csharp
var request = ExcelExport.Workbook(book => book.AddSheet("Data", ReadRows()));
await exporter.ExportToFileAsync(request, "report.xlsx",
    new ExcelStreamingExportOptions { RowWindowSize = 100 });

IEnumerable<ReportRow> ReadRows()
{
    for (var index = 0; index < 100_000; index++)
        yield return new ReportRow { Code = $"编号-{index}", Count = index };
}
```

`Export`／`ExportAsync` 输出到调用方流；`ExportToFile`／`ExportToFileAsync` 使用暂存文件和现有原子提交器。失败或取消不替换已有目标文件。任意调用方流写入后不提供回滚，流始终保持打开。SXSSF 构建、XML 刷盘和 ZIP 打包为同步阶段，异步出口的最终复制和文件提交使用真实异步 IO，不启动后台 `Task.Run`。

只生成普通 `.xlsx`，使用 inline strings。模板、图片列、图表、合并列、多级表头、模板区域、AutoFit/Adaptive 自动列宽明确拒绝，不回退 DOM。使用固定列宽。样式在创建单元格时应用，写出后不回访已释放的行。

## 导入

```csharp
using var input = File.OpenRead("report.xlsx");
var summary = await importer.ImportAsync(
    ExcelStreamingImport.Sheet<ReportRow>("Data"), input,
    async (row, cancellationToken) =>
    {
        if (row.IsValid)
            await SaveRowAsync(row.Item, cancellationToken);
        else
            RecordErrors(row.Errors);
        return ExcelRowReadDecision.Continue;
    });

public sealed class ReportRow
{
    public string Code { get; set; }
    public int Count { get; set; }
}
```

`SaveRowAsync` 和 `RecordErrors` 为业务自己的持久化与错误记录方法。同步消费者使用 `Import(request, input, row => ExcelRowReadDecision.Continue)`。按零基索引选择 Sheet 时使用 `ExcelStreamingImport.Sheet<ReportRow>(0)`。第三个配置参数复用 Sheet Fluent API，例如 `.Mapping(...)`、`.DynamicColumns(...)`、`.MaxReadColumns(200)`。

输入从流当前位置开始，先暂存到有额度的文件；流始终保持打开。解析实际包关系，支持稀疏行、inline/shared strings、富文本内容、1900/1904 日期系统和公式缓存值。不计算公式；缺少缓存的公式作为错误行交付。仅支持普通 Transitional XLSX，不支持 XLS、加密文件、宏文件或跨 Sheet 关系组装。

日期单元格及日期公式缓存的文本与原 DOM 服务一致，使用 `yyyy-MM-dd HH:mm:ss.fff`；字符串属性、原始值校验和自定义转换器读取相同文本，同时保留结构化日期值。只有参与列绑定的值错误会使实体失败；范围外或未映射列仍接受 XML 结构和文本额度检查。

每次等待回调完成后才读取下一行，没有后台预读队列。回调已完成的数据库写入不回滚；回调异常和取消直接传播。行结果 `RowIndex` 是零基，`ExcelImportError.RowIndex/ColumnIndex` 保持一基。失败行 `Item` 为 null。摘要只保存计数和有限错误样本，不累计实体或源行列表。

| 情况 | IsComplete | IsSuccess | StopReason |
|---|---|---|---|
| 完整读取，无错误 | true | true | None |
| 读完但有普通行错误 | true | false | None |
| 消费者返回 Stop | false | false | ConsumerStopped |
| 达到错误停止阈值 | false | false | ErrorLimit |
| 行处理达到资源额度 | false | false | ResourceLimit |

预检或暂存阶段资源错误抛出异常；行处理中的额度错误返回不完整摘要。错误样本达到上限只截断样本，不会因此停止。配置的 `StopAtFirstEmptyRow` 仍表示按配置完成读取。调用方需要持久化全部错误时，应在回调中及时处理行错误。

已交付行的回调返回 `Stop` 时，`ConsumerStopped` 优先于同一行触发的错误阈值或唯一值额度，已有错误计数和样本仍保留。返回 `Continue` 时按对应阈值停止。取消和消费者异常始终直接传播。

## 默认额度

| 项目 | 默认值 |
|---|---|
| 导出窗口／压缩 | 100 行／启用临时 XML 压缩 |
| 共享字符串及查询缓存 | 共用 8 MiB 计量预算，超出使用磁盘数据及定长索引 |
| 单元格／单行／导出窗口文本 | 32,767／262,144／2,097,152 字符 |
| 列／Sheet／样式／字体 | 100／32／4,096／256 |
| 唯一值 | 100,000 项及累计 8,388,608 字符 |
| 错误停止／摘要样本 | 1,000／100 |
| 库自有临时文件 | 每次调用 512 MiB |
| 输入／ZIP／XML | 沿用 `ExcelResourceLimits` 的有限默认值 |

导出选项为 `ExcelStreamingExportOptions`，导入选项为 `ExcelStreamingImportOptions`。导入列数使用请求的 `MaxReadColumns` 调整。唯一值和错误数量同时受 `ResourceLimits` 同名配置约束，取较小值。基础输入、ZIP 和 XML 额度不能关闭。

内存预算按 UTF-16 文本加每项管理开销计量，不是 CLR 堆或进程工作集硬上限。共享字符串索引也落盘，不保留全量偏移数组。`MaxTemporaryBytes` 限制每次调用的库自有临时文件总量，包括文件出口同时存在的暂存包和提交副本；复制前会为完整副本预留额度，因此文件出口需要容纳两份输出包。调用方目标流不计入该额度。NPOI 自建临时文件仍使用系统 `poifiles` 目录，不计入库自有额度；调用方应保证系统临时盘空间。

正常清理先 Dispose SXSSF，再关闭底层 XSSF；失败时继续尝试每个 writer。没有主异常时，清理错误直接抛出；已有主异常时，清理错误保存在 `exception.Data["Bing.Offices.Streaming.CleanupErrors"]`，不覆盖取消或业务异常。每次调用独立拥有工作簿、索引与缓存，不并发操作同一 Workbook。

## 容量证据与迁移

本次验收覆盖重复／高唯一字符串、1 万／10 万行、每行 10 列的导入和导出，共 8 个独立进程场景。高唯一导入实际触发磁盘存储；每个单元格均验证，正常完成无临时文件残留。测量区间包含暂存、索引、逐行处理、打包、提交和清理。详细数字见仓库 `ai_docs/tasks/BO-EXCEL-LOW-MEMORY-20261027/resource-results.md`；结果不是容量 SLA。

现有用户无需迁移。选择低内存服务时添加注册、把结果集合式导入改成逐行回调，并处理失败行和停止摘要；不要把旧模板配置直接移入新出口。新 API 为兼容增量，未删除公共成员。Abstractions 仍为 netstandard2.0，NPOI 仍为 2.7.4，没有新增 Open XML SDK 或 `IAsyncEnumerable<T>` 依赖。
