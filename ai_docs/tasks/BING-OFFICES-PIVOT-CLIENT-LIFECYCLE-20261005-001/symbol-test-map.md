# 生产符号与验证入口追溯

本阶段没有新增生产成员。下列工具从隔离本地 NuGet 包调用公开契约；维护工具的入口和 COM 断言不是生产 API。

| 最终生产符号 | 验证项目/入口 | 关键行为与产物 |
| --- | --- | --- |
| `ExcelWorkbookExportBuilder.AddSheet` / `PivotTable` | `build/PivotClientProbe` 顶级入口 | 固定 Warehouse/Amount 来源、有界区域、1/4 个透视、同源及不同源；完整来源与缓存记录 |
| `ClosedXmlExcelExporter.ExportToFile` / `ExportToFileAsync` | `build/PivotClientProbe` 顶级入口 | net6/net8 同步/异步真实夹具；六个资源场景、四个枚举取消场景 |
| `ClosedXmlReportWriter.ApplyPivotTables` / `ClosedXmlPivotSourceValidator.Validate` | `build/PivotClientProbe` 来源/缓存断言；`verify-excel.ps1` 的 `Get-PivotSnapshot`、`Assert-PivotSnapshot`、`Assert-PivotMatrix` | 缓存字段、全部记录及真实客户端完整来源、字段角色、Sum、来源范围和汇总矩阵；来源失败职责另见 SOURCE 的符号表 |
| `DefaultFileExportCommitter.Commit` / `CommitAsync` | `build/PivotClientProbe` 取消分支 | 同一令牌、原目标五字节保持、目录无新增暂存；双 TFM 同步/异步 |

## 客户端工具入口

| 工具入口 | 验证内容 |
| --- | --- |
| `verify-excel.ps1` 主流程 / `Assert-SourceMatrix` | 正常只读打开、完整来源行序和值；副本修改、保存及重开来源 |
| `Wait-InitialPivotSnapshot` | 被动观察空来源/矩阵，记录 Ready、CalculationState 和完整快照；不调用刷新/计算，非空错误立即进入失败断言 |
| `Assert-PivotSnapshot` / `Assert-PivotMatrix` | 首次 40/20/60，显式刷新后及重开 45/20/65；来源与字段角色、四行两列全矩阵 |
| 来源编辑后的断言 | 记录 `SourceChangeVisible`，只允许完整旧结果或完整新结果；不假设自动刷新一定发生或一定不发生 |
| `Add-ComObject` / finally | 只跟踪自己的对象、关闭自己的副本与应用；资源清理异常不报告 Passed |

最终客户端清单与原始失败记录汇总见 `artifacts/client-results.json`；资源场景的准确命令、退出码和独立 JSON 见 `artifacts/resources/results.json`。历史 SOURCE 职责、Provider/合同、API 和包消费者结果不冒充本次客户端结果。
