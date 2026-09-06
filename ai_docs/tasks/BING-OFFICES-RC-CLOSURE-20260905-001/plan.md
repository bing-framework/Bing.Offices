<!-- AI_PLAN_STATUS: READY -->
# Bing.Offices RC 收口、弃用清理与发布证据闭环计划

## 任务元数据

```yaml
task-id: BING-OFFICES-RC-CLOSURE-20260905-001
priority: P0
language: zh-CN
execution-mode: continuous-resumable
breaking-change: allowed-and-required-for-deprecated-api
auto-commit: false
auto-push: false
auto-create-pr: false
auto-tag: false
auto-publish-nuget: false
```

本计划由用户 2026-09-06 提供的完整执行提示词直接落地，是本任务的已批准计划。后续只使用本 task-id，从未验证项继续；不因 Phase 边界、首次构建/测试/Benchmark/Review 失败而停止。

## 目标与非目标

目标是沿真实公共调用链完成 RC 正确性、异常与日期合同、弃用删除、API/SPI 收敛、内部重构、测试、真实 nupkg 消费、性能资源证据、独立 Review 和发布门禁闭环。

不承诺用 `Task.Run` 伪造异步 IO；不新增平行异常或日期解析体系；不以 `EditorBrowsable(Never)` 代替边界治理；不自动 commit、push、PR、tag 或 publish；不回滚用户改动。

## 强制合同

1. 参数异常、取消和致命异常保持原生语义；不可恢复业务失败统一为现有 `BingOfficesException` 家族。
2. 同一失败只在公共边界分类一次；Observer 与调用方获得同一异常实例，每操作最多通知一次，Observer 失败不覆盖主异常。
3. 文件系统 Create/Flush/Replace/Move 才归类 FileCommit，写委托中的配置、转换和序列化失败原样传播。
4. 日期默认确定性支持 `yyyy-MM-dd`；date-only 为午夜和 `DateTimeKind.Unspecified`；Excel/CSV 共用解析合同；禁止隐式使用服务器本地时区生成 `DateTimeOffset`。
5. `CellExtensions`、`CellStyleExtensions`、`FontExtensions`、`RowExtensions`、`SheetExtensions`、`WorkbookExtensions`、`ExcelNpoiServiceCollectionExtensions` 必须保持 public。
6. 删除 `ExcelMapping.For<T>` 旧体系、DataTable `CsvHelper`、runtime v1 migration、LegacyCompatibility、v1 Benchmark 及仅为其服务的代码/测试/文档/API baseline。
7. 生产程序集之间不得新增 `InternalsVisibleTo`；所有 public/API/缓存/Provider 变更必须有职责级直接测试和最终生产符号到测试方法的追踪。

## 执行阶段

### Phase 0：基线与证据台账

- 记录 branch、commit、dirty state、SDK/Runtime/OS/TFM、项目与依赖。
- restore/build/test 建立可复现基线，区分既有失败与本轮回归。
- 生成调用链、异常、日期、public API、弃用反向引用和测试能力矩阵。
- 复核指定最新评审；若文件缺失，记录缺口并以当前源码和最近可用独立评审为次级证据。

### Phase 1：正确性、异常与资源边界

- 拆分写临时内容与文件提交异常边界，修复 Observer 同实例/单次通知。
- 决策并统一 Mapping Configuration Loader 的观察策略。
- 完成 DateTimeOffset 明确导出/导入策略与跨时区往返。
- 修复图片 Try/Throw、未知 Sheet、失败工作簿诊断、XLSX/XLS 资源策略。
- 增加公共边界、Provider 分支和集成测试。

### Phase 2：弃用删除与 API 收敛

- 删除全部确认弃用入口、实现、测试、Benchmark、文档和 baseline 残留。
- Excel 收敛到 Import/Export + Workbook/Sheet Builder；CSV 收敛到 typed importer/exporter + Options + Stream Extensions。
- 逐项治理 execution-detail public surface，同时保留七个 NPOI 用户扩展容器。
- 更新 API 快照、Breaking Change 与迁移示例。

### Phase 3：内部职责重构

- 拆分 failure workbook、CSV pipeline、mapping parse/merge/compile/cache、NPOI import/export 职责。
- 缓存泛型 Sheet 委托，减少热路径反射和临时集合。
- 保持已冻结的异常、取消、顺序、样式、流所有权与序列化行为。

### Phase 4：测试与包消费

- 完成 Unit、Integration、Docs fence、PackageConsumer 独立报告。
- pack 三个真实 nupkg；空目录消费者只使用 PackageReference 与本地产物源。
- 覆盖 netcoreapp3.1/net6/net8，netstandard2.0 以消费者编译验证；不可用矩阵如实标记 Blocked。

### Phase 5：Benchmark 与资源证据

- 建立真实 baseline、MemoryDiagnoser、环境 identity 和原始 artifacts。
- 覆盖 Excel/CSV、mapping 冷热缓存、高基数、模板/样式/验证/图片/failure workbook、并发与取消。
- 记录 Gen0/1/2、Allocated、LOH、PeakWorkingSet、P50/P95/P99；无结果不得宣称低 GC。

### Phase 6：文档、独立 Review 与发布门禁

- 更新 README、docs、XML docs、Breaking Change、迁移、异常、日期、NPOI、资源合同。
- Release build/test/pack/APICompat/package-only consumer 全量验证。
- 由未参与主要实现的 Review Agent 或独立审查上下文检查；修复开放 P0/P1 后重跑。
- 仅当附件列出的 14 项门禁全部满足时结论为 Go，否则为 Conditional Go 或 No-Go。

## 必需证据文件

任务目录持续维护：`execution.md`、`progress.md`、`decisions.md`、`baseline.md`、`call-chain-matrix.md`、`exception-contract.md`、`date-contract.md`、`public-api-ledger.md`、`deprecated-removal.md`、`unit-test-report.md`、`integration-test-report.md`、`docs-test-report.md`、`package-consumer-report.md`、`benchmark-report.md`、`resource-report.md`、`review.md`、`final-report.md`。

状态只使用 `TODO`、`IN_PROGRESS`、`BLOCKED`、`DONE`、`VERIFIED`。代码完成但没有真实验证最多为 `DONE`。

## Definition of Done

- 异常、日期、NPOI public 扩展、弃用删除、API/SPI/IVT、资源限制全部有当前源码与直接测试证据。
- Unit、Integration、Docs、PackageConsumer 分报告通过，P0 零失败零跳过。
- 多目标 Release build/test/pack、APICompat 和独立 nupkg 消费通过。
- Benchmark/Resource 有本轮真实结果、环境身份、基线和预算状态。
- 独立 Review 无开放 P0/P1；最终报告给出可审计 Go/Conditional Go/No-Go。
- 明确记录未 commit、push、PR、tag 或 publish。
