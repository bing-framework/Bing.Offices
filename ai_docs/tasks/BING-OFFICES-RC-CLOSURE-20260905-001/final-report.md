# RC 最终报告

## 结论

- Task ID：`BING-OFFICES-RC-CLOSURE-20260905-001`
- 发布结论：`No-Go`
- 执行状态：`PARTIAL`，本地可安全执行项已完成；正式 API 与性能资源门禁需要维护者输入。
- Git：`master` / `9d78ab76e28891ac9b7e0e558f4df669721a4ea6` / dirty；执行前 clean。

## 已完成与验证

- 异常：Atomic 内容失败原样传播；文件系统提交失败保持 FileCommit；Observer 同实例/单次通知；DI Loader 观察、静态 Loader 纯解析；Failure Workbook 无 sink/sink 失败均有结构化 Trace。
- 日期：DateTimeOffset 默认 invariant `O` 文本；无隐式 Local offset；XLS/XLSX 导入→导出→关闭重开→再导入，UTC/Pacific 两个 TZ 各 6/6；CSV 往返通过。
- Provider：TryAddPicture 参数和 HSSF/XSSF 成功路径；MovePictures null/未知 Sheet 明确失败；七个 NPOI 扩展容器保持 public。
- 资源：MaxInputBytes 默认 128 MiB，可显式关闭；XLS/OLE 输入限制；XLSX ZIP/XML preflight；Failure Workbook 元数据降级诊断。
- Breaking 删除：旧 ExcelMapping builders、DataTable CsvHelper、runtime v1 migration、LegacyCompatibility、v1 benchmark 与专属死代码已移除；仅迁移表保留删除符号名称。
- API：导出类型收敛为 141 User API、9 Provider User API、15 Provider SPI、0 Compatibility、0 Execution detail；`BingOfficesException` 为 abstract；无生产程序集间 IVT。
- 重构：Sheet 与关系泛型委托缓存；Failure Workbook 诊断职责独立；未完成的大类全面拆分保留为 P2。

## 验证摘要

| 门禁 | 最终结果 |
| --- | --- |
| Release build | PASS，0 errors / 28 warnings |
| Unit netcoreapp3.1/net6/net8 | 各 470/471，0 skipped；唯一失败 formal API hash |
| Integration net6/net8 | 各 15/15 |
| Docs | 10/10，全部 Markdown C# fence 编译执行 |
| DateTimeOffset TZ | UTC 6/6；Pacific 6/6 |
| ResourceProbe | Excel 14/14；Mapping/Unique 16/16；尾延迟 4 档 |
| PackageConsumer | netstandard2.0 编译；三个 runtime 均 `package-consumer-ok` |
| Pack identity | `packages-rc-final` 三包；五个包内 DLL 与 Release 输出 hash 全 match |
| API candidate | 三 TFM 一致；formal approval BLOCKED |
| Independent Review | 0 P0 / 2 P1 / 2 P2，`NEEDS_FIX / No-Go` |

分报告：`unit-test-report.md`、`integration-test-report.md`、`docs-test-report.md`、`package-consumer-report.md`、`benchmark-report.md`、`resource-report.md`、`api-diff.md`、`review.md`。原始证据位于本任务 `artifacts/`。

## 性能事实

已运行 Excel 9、CSV 6、Failure Workbook 3 个 1k/10k/100k ShortRun。100k 单次 Allocated：Excel Import 1,684.74 MB、Excel Export 734.51 MB、CSV Import 198.32 MB、CSV Export 1,448.26 MB、Failure Workbook 2,606.14 MB。多个场景发生 Gen2；当前不得描述为低 GC。

## 开放门禁

1. P1 `FIX-002`：仓库缺失正式成员 baseline JSON；当前 API hash 测试三个 TFM 均失败，canonicalizer 也不覆盖 abstract/sealed、约束、默认参数、访问器与关键 attribute。需要维护者恢复/批准成员 baseline，并使用完整 APICompat 工具复验。
2. P1 `FIX-005`：缺少可比历史 baseline、批准预算与批准人；尚缺图片/模板/样式/验证完整资源矩阵、Failure Workbook 双 DOM 独立 PeakWorkingSet/LOH、取消延迟分位数。需要维护者给出预算后在固定环境复验。
3. P2 `FIX-006`：CSV、Importer/Exporter 与 Failure Writer 仍有大类职责拆分遗留，建议独立重构任务处理。
4. P2 `FIX-008`：TryAddPicture 在 AddPicture 后续阶段失败时的部分修改/回滚语义尚未以可控注入测试冻结。
5. Linux/macOS Integration runner 当前不可用；netcoreapp3.1/net6 已 EOL，旧 TFM 依赖支持警告需发布策略确认。

## Breaking 与迁移

Excel 使用 `ExcelImport/ExcelExport + Workbook/Sheet Builder`；映射使用 `ImportMappingBuilder<T>` / `ExportMappingBuilder<T>` 与 v2 Document；CSV 使用 `ICsvImporter/ICsvExporter + Options + Stream Extensions`。完整对照见 `deprecated-removal.md` 与 `docs/excel/nuget-migration.md`。

## 建议提交分组

1. `fix: close exception date and provider contracts`
2. `refactor!: remove deprecated mapping and csv compatibility APIs`
3. `perf: cache generic dispatch and add production benchmarks`
4. `test: add rc regression package and resource evidence`
5. `docs: publish rc contracts migration and no-go report`

未执行 commit、push、PR、tag 或 NuGet publish。
