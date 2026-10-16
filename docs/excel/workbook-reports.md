# 工作簿、模板与报表内容

## 功能说明

`ExcelExport.Workbook` 构建包含多个异构 Sheet 的导出请求；`ExcelImport.Workbook<TWorkbook>` 把各 Sheet 映射到根模型集合。Workbook 请求还承载元数据、模板、字典列表与透视表。样式、表头、列宽、行高和报表内容在 Sheet 构建器配置。

| 能力 | 公共入口 | 实际实现与限制 |
| --- | --- | --- |
| 多表、元数据、模板 | `AddSheet`、`Metadata`、`UseTemplate` | NPOI/ClosedXML 支持完整工作簿；MiniExcel 的模板等高级请求按矩阵拒绝 |
| 样式与物理布局 | `SheetStyle`、`HeaderStyle`、`BodyStyle`、`HeaderRows`、`ColumnWidth`、`RowHeight` | 能力依 Provider 而异，样式键需有对应配置，不保证第三方引擎任意样式 |
| 表格、筛选、冻结、条件格式、名称、打印 | `Table`、`AutoFilter`、`FreezePane`、`ConditionalFormat`、`NamedRange`、`PrintLayout` | NPOI/ClosedXML 完整模型；SpreadCheetah 仅实现前向可表达子集 |
| 图片、校验与批注 | `Image`、`DataValidation`、`ExcelComment` | 图片/校验在 NPOI、ClosedXML、SpreadCheetah 各有实现；批注及冲突策略以具体入口为准 |
| 图表 | `Chart(ExcelChartDefinition)` | 只有 NPOI XLSX 支持柱状、折线、单系列饼图；其他数据 Provider 拒绝此请求 |
| 透视 | Workbook 的 `PivotTable`、`UpdatePivotTable` | 另见[透视专题](pivot-tables.md)，不计算公式、不运行客户端刷新 |

## 用法示例

下例在已有的 `rows`、`exporter`、`output` 上创建冻结表头和筛选区域；`rows` 至少一行，含 `Code` 和 `Amount` 两个映射列。

```csharp
var request = ExcelExport.Workbook(book => book
    .Metadata(new ExcelWorkbookMetadataOptions { Author = "订单系统" })
    .AddSheet("订单", rows, sheet => sheet
        .FreezePane(new ExcelFreezePaneDefinition { Rows = 1 })
        .AutoFilter(new ExcelAutoFilterDefinition
        {
            Range = new ExcelRangeDefinition
            {
                StartRow = 0, EndRow = rows.Length,
                StartColumn = 0, EndColumn = 1
            }
        })));
exporter.Export(request, output);
```

所需命名空间为 `Bing.Offices.Exports`；完整注册与模型见[快速开始](../getting-started.md)。范围坐标是零基闭区间，因此上例包含表头和全部数据行；图表的区域与锚点也须按对应定义检查边界，不能混用 A1 字符串与数值坐标。

模板通过 `book.UseTemplate(template, leaveOpen: true)` 指定。配置 `UseTemplateRegion`、`TemplateCellOverwrite` 和 `CommentConflicts` 时，先核对模板命名范围、原有单元格及冲突策略；模板保留能力不表示可以编辑任何未知部件。模板流是否释放由 `leaveOpen` 控制。

## 核心实现与关键位置

1. Abstractions 的 Builder 校验名称、范围、枚举和配置冲突，并创建请求快照。
2. Core Mapping Plan 决定最终物理列、标题、转换、动态列和样式布局。
3. Provider 先检查格式和特性，再按最终列写业务数据及报表内容，最后序列化；文件入口交给文件提交器。

关键路径：

- `src/Bing.Offices.Abstractions/Bing/Offices/Exports/{Workbook,Sheets,Reports,Content,Charts}`。
- `src/Bing.Offices.Npoi/Bing/Offices/Exports/{NpoiReportWriter.cs,NpoiSheetContentWriter.cs,NpoiExcelExporter.cs}`。
- `src/Bing.Offices.ClosedXml/Bing/Offices/Exports/{ClosedXmlReportWriter.cs,ClosedXmlSheetContentWriter.cs,ClosedXmlExcelExporter.cs}`。
- `src/Bing.Offices.SpreadCheetah/Bing/Offices/Exports` 与 `src/provider-shared/{ExcelReportPreflight.cs,ExcelSheetContent.cs}`。

## 依赖与配置

选择对应 Provider 包并注册其服务。模型映射可通过 Attribute、Profile、v2 文档或 Sheet 的 `Mapping` 覆盖。隐藏辅助 Sheet 和名称列表由 `LookupList`/`CascadeList` 生成，分别参见[原生校验](import-validation.md)及[两级联动](cascade-lookups.md)；工作表可编辑区域见[保护](sheet-protection.md)。

## 已知限制

- 图片接受调用方提供的字节，不下载 URL 或自动读取文件。图片格式、锚点与 Provider 支持范围须在输出前确认。
- 公式保存与公式计算是独立能力；写入公式不等于得到新缓存。详见[公式与商业扩展](formula-rendering-conversion.md)。
- 直接输出流不提供回滚；文件输出使用提交器保护既有目标。业务数据枚举或回调也可能抛出异常。
- XLS/XLSX 有不同的行列和特性限制；完整矩阵见[Provider 文档](09-providers.md)。没有找到实现的报表特性按“待实现/待确认”处理，不以底层引擎支持推断公共支持。
