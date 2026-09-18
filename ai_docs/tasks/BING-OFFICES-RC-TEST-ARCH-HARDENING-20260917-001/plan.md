# Bing.Offices RC 测试架构与发布收口计划

Task ID: `BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001`

## 目标与边界

将现有混合测试收敛为 Common、NPOI、MiniExcel、Provider-specific Integration 和 Aggregate Integration 六个职责边界。保持公共 API、TFM、版本文件和先注册优先 DI 行为不变；不引入 Provider Resolver、公共杂项程序集、Pool 或未证实的复杂性能抽象。

当前 Unit 为 `tests/Bing.Offices.Tests`（532 个测试方法声明），Integration 为 `tests/Bing.Offices.Tests.Integration`（32 个声明）；两者均直接引用 NPOI 与 MiniExcel。生产 Core 为 netstandard2.0，Provider 为 net6/net8；生产 IVT 不得出现，Provider 只保留最小测试 Friend。MiniExcel RawDateReader、逐行 Dictionary/反射和 Relation O(P×C) 仅在拆分安全网建立后按基线优化。

## 目标项目与迁移矩阵

目标新增 `Bing.Offices.Npoi.Tests`、`Bing.Offices.MiniExcel.Tests`、`Bing.Offices.Npoi.Tests.Integration`、`Bing.Offices.MiniExcel.Tests.Integration`；保留 `Bing.Offices.Tests` 作为 Common、`Bing.Offices.Tests.Integration` 作为 Aggregate。Common 只引用 Abstractions/Core；Provider Unit 允许自身 Provider；Provider Integration 只直接引用自身 Provider；Aggregate 才同时引用两个 Provider。

迁移前先生成 `baseline.md`、`test-migration-matrix.md`、`provider-test-matrix.md`，逐方法记录原项目、类/方法、Theory 数据、实际依赖、目标、internal、TFM、fixture 和 Collection。Common 保留 CSV、Mapping、Date、Validation、Cache、Core IO、纯公共 SPI；NPOI 迁所有 NPOI 类型、HSSF/XSSF、DOM、扩展、真实文件；MiniExcel 迁 Provider 实现、日期/Row/Column、Query/SaveAs、取消/流所有权；Aggregate 保留五个跨 Provider 合同和公共治理。默认 Removed=0，按 TFM 对账 `After=Before+Added-Removed`。

## 阶段任务

1. `RC-000/010`：捕获 Git、SDK、Solution、ProjectReference、TFM、IVT、CI、版本哈希、嵌套产物；建立逐方法迁移及符号追溯矩阵。
2. `RC-020`：清理 Common Provider 引用；拆 AsyncPipeline、Loader、ReviewFixRegression 的混合方法；CSV 使用 Core 真实实现；双 TFM 独立通过。
3. `RC-030/040`：创建 NPOI/MiniExcel Unit，保留所有成功/失败、日期、物理 Row/Column、Converter、结构化错误和 fixture；MiniExcel fixture-only NPOI 依赖必须明确且不得引用 Bing.Offices.Npoi。
4. `RC-050/060`：创建两个 Provider Integration，使用真实文件/临时目录；NPOI 覆盖 XLS/XLSX、模板、失败工作簿、图表、图片、批注、合并单元格；MiniExcel 覆盖同步/异步、文件提交、取消、多 Sheet、Large Category。
5. `RC-070/080`：Aggregate 收敛公共合同和 Unsupported 无部分输出；更新 Solution、精确 IVT 白名单、扩展覆盖扫描、资源复制和 CI Job。
6. `RC-090/100/110`：先建立 RawDateReader/Relation/Importer before 基线，再实施最小列过滤、typed delegate、表头绑定或 parent index；不改变 1900/1904、negative serial、comparer、重复键、错误顺序、取消和流位置语义；无收益则保留调查而不复杂化。
7. `RC-120`：按 chinese-comments 处理本轮复杂生产/测试 helper、RawDateReader、ValueAdapter 和 Provider 文档；不全仓库机械补注释。
8. `RC-130/140`：运行 Release、职责级双 TFM、Docs、API Snapshot、Pack/PackageReference Consumer、100K 性能、资源矩阵、diff/review；代码完成与正式发布分别判定。

## 验证与发布门禁

使用 `dotnet restore Bing.Offices.sln`、`dotnet build Bing.Offices.sln -c Release --no-restore`、职责级 `dotnet test ... -f net6.0/net8.0` 和 `git diff --check`。TRX 必须写入 `artifacts/tests/{common,npoi,miniexcel,integration-npoi,integration-miniexcel,integration-cross-provider}/{tfm}`；Large 只在显式 Release Gate 执行。Pack 固定写入根 `artifacts/packages`，Consumer 必须从本次 nupkg 通过 PackageReference 恢复，MiniExcel nuspec 不得依赖 NPOI。

性能必须分 before/after、同 workload、独立进程、至少三次重复；elapsed/allocated 独立计算中位数，峰值使用 `Process.PeakWorkingSet64`，原始 JSON 保留候选 identity。覆盖 1K/10K/100K，Release Gate 再执行 500K/1M；未执行为 `NOT_VERIFIED`。资源批准必须有批准人/时间，未批准为 `BLOCKED`。CI 至少拆分 unit-common/npoi/miniexcel、integration-npoi/miniexcel/cross-provider、docs、consumer-net6/net8、api、resource、benchmark。

必须生成 `execution.md`、`baseline.md`、两份迁移矩阵、三份 Unit/Integration 报告、`api-diff.md`、`package-consumer-report.md`、`benchmark-report.md`、`resource-report.md`、`symbol-test-map.md`、`final-report.md` 和独立 `review.md`。不得修改 `version.props`、`version.dev.props`、公共 API baseline；不得 git add/commit/push/tag、PR 或 NuGet publish。代码整改可为 COMPLETED，但资源批准、500K/1M、生产 2 CPU/4 GiB 或外部 CI 未完成时正式发布必须为 PARTIAL/NOT_VERIFIED。

## 验收标准

Common 不再引用任何 Provider；两个 Provider Unit/Integration 项目存在且职责清晰；Aggregate 只承担跨 Provider 合同；迁移前后方法/Case 对账无无理由丢失；生产 IVT=0、测试 IVT 最小化；日期/Row/Column/RawDate 回归保留；报告能区分项目、TFM、方法和 TRX；根目录外无 artifacts/TestResults/BenchmarkDotNet.Artifacts；双 TFM API/Consumer/Release Build 通过；版本文件与运行前一致；未验证外部条件明确标记，不能写成 PASS。
