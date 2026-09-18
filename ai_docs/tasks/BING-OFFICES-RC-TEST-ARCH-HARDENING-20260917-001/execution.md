<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001
AI_EXECUTION_FINISHED_AT: 2026-09-18T09:31:44+08:00

# 实施执行报告

## 执行结论

代码整改状态为 `PARTIAL`（正式发布仍受外部门禁限制）：已完成方法级职责迁移、Provider Integration 安全网、精确 IVT、CI Large/benchmark 入口和 provider-shared preflight 收口；未经可信 baseline 支持的 RawDate 过滤/重复 Query、Relation 缓存优化及 Importer 表头物理列预绑定均已撤回；性能收益不作宣称。

任务整体状态为 `PARTIAL`：API snapshot compare 被候选 identity 不匹配阻断，性能 before/after、500K/1M、维护者资源批准、真实 2 CPU/4 GiB 和外部 CI 尚未形成正式发布证据。因此不能将本轮写成 Release PASS。

## 任务信息

- Task ID：`BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001`
- 基线 HEAD：`ece61d8f9f7c6fcd1efe1dcf68cc8224165c80cc`
- 基线分支：`feat/miniexcel-provider`
- 计划：`ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/plan.md`
- 本轮未执行 `git add`、`git commit`、`git push`、tag 或 PR。

## 计划执行情况

| Phase | 状态 | 结果 |
| --- | --- | --- |
| RC-000 / RC-010 | COMPLETED | 生成 baseline、测试迁移矩阵、Provider 矩阵和符号追溯；记录基线 Git/SDK/TFM/引用/IVT/版本哈希。迁移前 TRX 未在本轮重新执行，标记 NOT_VERIFIED。 |
| RC-020 | COMPLETED | Common 只保留 Abstractions/Core；CSV 测试直接注册 Core 实现；移除 Provider 和 ResourceProbe 传递引用。 |
| RC-030 / RC-040 | COMPLETED | 创建 NPOI Unit、MiniExcel Unit；保留 HSSF/XSSF、日期、物理行列、Converter、取消和结构化错误回归；MiniExcel 的 NPOI 仅为 PrivateAssets fixture 包。 |
| RC-050 / RC-060 | COMPLETED | 创建两个 Provider Integration；NPOI 增加富工作簿真实文件 reopen，MiniExcel 覆盖同步/异步、多 Sheet/日期和取消/原子提交清理。 |
| RC-070 | COMPLETED | Aggregate 收敛五个跨 Provider 合同、两种注册顺序的先注册优先 DI 合同、公共 API 分类/快照和 CSV 公共 IO；Unsupported 仍断言结构化错误及无部分输出。 |
| RC-080 | COMPLETED | Solution、CI matrix、Large dispatch、真实 benchmark probe、资源复制和报告路径更新；IVT 改为逐生产程序集精确批准列表，并覆盖错配/伪造后缀拒绝。 |
| RC-090 / RC-100 / RC-110 | DEVIATED_OK | RawDateReader 已恢复 HEAD 全量索引；Relation 保持旧逐子项扫描以保留公开委托副作用、首匹配和错误顺序；Importer 预绑定已撤回，`ReadColumnRange` 绝对列号由直接回归覆盖；preflight 移至 `src/provider-shared` 且 Core 显式排除。未经可信 baseline 的性能优化已按计划撤回。 |
| RC-120 | COMPLETED | 新增 Reader/Relation/Importer helper 均沿用现有中文 XML/行内注释风格；报告明确区分已验证热点与仍未执行的发布资源证据。 |
| RC-130 | PARTIAL | Release build、双 TFM 测试、Docs、打包、Consumer、资源和探针完成；API compare、外部 CI、生产资源批准、500K/1M 未完成。 |
| RC-140 | COMPLETED | Round 7 已完成 FIX-006/FIX-007 修复记录、双 TFM 回归和最终差异检查；`review.md` 仍由 Reviewer 保持 `NEEDS_FIX`，下一步是独立重新 Review。 |

## 已完成事项

- 新增四个职责项目：`Bing.Offices.Npoi.Tests`、`Bing.Offices.MiniExcel.Tests`、`Bing.Offices.Npoi.Tests.Integration`、`Bing.Offices.MiniExcel.Tests.Integration`，统一 net6.0/net8.0、`IsPackable=false` 和 common props。
- Common、Provider Unit、Provider Integration、Aggregate 的直接 ProjectReference 与 fixture 依赖符合矩阵；测试程序集不互相引用。
- `AsyncPipelineCsvTest` 直接构造 Core 实现；CSV 文件公共 IO 留在 Aggregate；NPOI 专属文件测试不再混入 Common。
- 公共扩展覆盖拆为 Core 与 NPOI 两个职责级扫描门禁；公共 API IVT 校验使用精确白名单并拒绝 `.Tests.Fake`。
- ReviewFixRegression 的 30 个纯 Core 方法与 Loader 的 3 个纯方法已移入 Common；NPOI 仅保留 2 个 Provider 方法和 1 个 DI Loader 方法。
- NPOI Integration 新增合并单元格、批注和图表的真实 XLSX reopen 断言；MiniExcel Integration 新增同步多 Sheet/日期及取消保留目标/清理临时文件回归。
- MiniExcel RawDateReader 使用 HEAD 全量 numeric-cell 索引，Relation 保留逐子项首匹配父项并保持公开委托调用/异常边界，Importer 固定与动态列预绑定已撤回；`ReadColumnRange` 绝对列号由直接回归覆盖，另有直接回归覆盖全量索引、首匹配、空子集和重复 ParentKey 异常语义。
- 共享 `ExcelXlsxZipPreflight` 已移动至 `src/provider-shared/ExcelXlsxZipPreflight.cs`，仅由两个 Provider 链接编译，Core 不再包含该源文件。
- Provider comparison 改为单 Provider 独立进程运行，header 写入进程 ID、候选 assembly SHA-256、providerCount/filter，样本峰值使用 `Process.PeakWorkingSet64`。
- 更新 `.github/workflows/ci.yml` 为职责 × TFM matrix，并使用 checkout 后捕获的版本哈希进行验证，移除固定失配的 `version.dev.props` 哈希。
- 生成 baseline、两份迁移矩阵、Provider/Unit/Integration/API/Package/Benchmark/Resource/Symbol/Review/Final 报告。

## 部分/未完成事项

- API compare 未通过候选 identity 校验：baseline expected `D78F11D98A5D5F9BF7FAA355072B1FB96920C16ACF53E77013BA2ACB3A28DBA1`，当前实际 `0D858201507D680386551CA755D137E0F3A6CEAFBE848461E1BD5734D2312CF9`。未修改 API baseline。
- Provider comparison 仍只有当前工作树 after 样本，不能据此计算跨 Provider 收益；RawDate hotspot 仅为可运行的全量 numeric-cell 调查 workload，不是 HEAD Reader before/after；Relation 只有 1K before 和完整 after，10K/100K before 未验证。Importer/preflight 语义矩阵由直接回归和 provider-shared 编译边界覆盖。正式生产资源和外部候选证据仍未完成。
- 资源矩阵为 1K、本机环境，`approvalStatus=BLOCKED`；没有维护者批准时间，也不能替代生产 2 CPU/4 GiB 证据。
- 500K/1M、外部候选 commit CI 和正式发布审批未执行。

## 修改文件

- Solution/CI：`Bing.Offices.sln`、`.github/workflows/ci.yml`。
- IVT：`src/Bing.Offices.Abstractions/AssemblyInfo.cs`、`src/Bing.Offices.Core/AssemblyInfo.cs`、`src/Bing.Offices.Npoi/AssemblyInfo.cs`、`src/Bing.Offices.MiniExcel/AssemblyInfo.cs`。
- 项目边界：Common/Aggregate csproj 与四个新测试 csproj。
- 测试迁移：NPOI/MiniExcel Unit 和 Integration 文件、Common CSV async 测试、Aggregate CSV/API/Provider contract 测试、Core/NPOI extension coverage 测试。
- 证据：本任务目录下的 baseline、migration/provider matrix、symbol map、各测试报告、api-diff、package-consumer、benchmark、resource、review、final 和本文件；热点原始样本位于 `artifacts/benchmarks/review-round3-hotspot-{before,after}.jsonl`。

## API/数据/配置变化

- 未改变生产 TFM、版本文件、正式 public API baseline、ValidateMode、ExcelImportValidationMode 或先注册优先 DI 语义。
- 只收紧测试友元到职责级精确程序集；没有新增 Resolver、Pool 或公共杂项程序集。
- 版本文件哈希保持不变：
  - `version.props=77FEBEC429F841DC839F84FED57A48AE4159DF2EA92A15AD808BB3027C39B0D1`
  - `version.dev.props=282D9A7C70D79FFC5F354DA57035A479B08A8485D1AF7637B25AF6C815DCAF31`

## 测试结果

六个职责项目的历史 TRX 与 Round 5 定向复验如下，Case 为 TRX 展开结果：

| 职责 | net6.0 | net8.0 | TRX 目录 |
| --- | ---: | ---: | --- |
| Common Unit | PASS 196 | PASS 196 | `artifacts/tests/review-round3/common-net6/common-net6.trx`; `common-net8/common-net8.trx` |
| NPOI Unit | FAIL 532/533 | PASS 533 | `artifacts/tests/review-round3/npoi-unit-net6.0/npoi-unit-net6.0.trx`; `npoi-unit-net8/npoi-unit-net8.trx` |
| MiniExcel Unit | PASS 39 | PASS 39 | `artifacts/tests/review-fix-round5/miniexcel-net6/miniexcel-net6.trx`; `miniexcel-net8/miniexcel-net8.trx` |
| Aggregate Integration | PASS 29 | PASS 29 | `artifacts/tests/review-round3/aggregate-net6/aggregate-net6.trx`; `aggregate-net8/aggregate-net8.trx` |
| NPOI Integration | PASS 28 | PASS 28 | `artifacts/tests/review-round3/npoi-integration-net6/npoi-integration-net6.trx`; `npoi-integration-net8/npoi-integration-net8.trx` |
| MiniExcel Integration | PASS 7 | PASS 7 | `artifacts/tests/review-round3/miniexcel-integration-net6/miniexcel-integration-net6.trx`; `miniexcel-integration-net8/miniexcel-integration-net8.trx` |

Docs contract：本轮未重跑；保留证据 `artifacts/tests/review-fix-round2/docs-final/docs-net8-final.trx`，不将其计入 Round 3 六职责 TRX。

## Build/Typecheck/Lint/Format

- `dotnet build Bing.Offices.sln -c Release --no-restore`：PASS，0 warning、0 error。
- `git diff --check`：PASS。Git 的 LF/CRLF 提示是换行规范提示，不是 diff 错误。
- `src/tests/benchmarks` 下未发现 `TestResults` 或 `BenchmarkDotNet.Artifacts`；误生成的 benchmark build 输出已可恢复地移至根 `artifacts/review-fix-quarantine/review-fix-build*`，原嵌套 `benchmarks/Bing.Offices.Benchmarks/artifacts` 目录已清除。

## 专项验证

- 四个生产包已打包，net6/net8 Consumer 均恢复、构建、运行成功；Consumer 日志在 `artifacts/consumers/rc-test-arch`。
- API identity self-test PASS；正式 API compare 按上文 identity mismatch 保持 NOT_VERIFIED。
- ResourceProbe 36/36 场景 PASS，结果为 BLOCKED approval，不宣称生产批准。
- MiniExcel 100K probe PASS；Provider comparison 100K × 3 按 NPOI/MiniExcel 分别独立进程生成当前工作树 after 证据，Real IO 24 场景 × 3 保留。
- Round 4 Hotspot probe 的历史样本完成 RawDate 100K×3/10/30 列 before/after 与 Relation 1K before、1K/10K/100K after，各 3 次取中位数；这些样本不作为当前候选性能证据。当前探针已适配 HEAD 三参数 Reader，并仅用于全量 numeric-cell 调查 workload；Relation 10K/100K before 仍 NOT_VERIFIED。记录 elapsed、allocated、Gen0/1/2、LOH 快照、`PeakWorkingSet64`、索引数、XLSX 大小、进程 ID 和 candidate identity。

## 计划偏差

Provider comparison 仍只有历史 after 对照，因此未计算 Provider 间收益；RawDate before 是调查性 numeric-cell replay 而非完整历史 Reader，Relation 的完整 before 受旧 O(P×C) 复杂度阻断，不能写成完整性能 PASS。Importer 的列绑定优化已撤回，绝对物理列定位由直接回归覆盖；Relation 保持公开委托逐子项线性扫描，所有语义回归和未执行的生产资源条件均单独记录。

## 基线问题

- 本轮由 `method-identity-audit.py` 以 UTF-8 锚定属性解析并关联方法声明：HEAD 569、当前六职责 592，Added=23、Removed=0；144 unchanged、391 moved、34 renamed/moved。当前职责分项为 Common 166、NPOI Unit 332、MiniExcel Unit 35、Aggregate 28、NPOI Integration 23、MiniExcel Integration 8。完整旧身份去向和 Added 原因见 `method-identity-audit.md`；迁移前双 TFM TRX 未执行，Theory 展开仍以实际 TRX 为准。
- 原有版本 CI 固定 `version.dev.props` 哈希已改为 checkout 后捕获并在验证后比较；版本文件本身未改动。

## 已知问题

- 正式 Release 仍依赖维护者重新批准候选 identity，并补充外部 CI、生产资源和 500K/1M 证据。
- Consumer net6 构建存在 SDK EOL warning，已单独记录，不影响本机 Consumer 运行结果。

## 风险与回归关注点

- 迁移后继续保持 HSSF/XSSF 成功和失败覆盖，MiniExcel fixture-only NPOI 依赖必须继续保持 PrivateAssets，不得变成 Provider ProjectReference。
- CSV Common 测试必须继续直接构造 Core；Aggregate 的 CSV 公共 IO 不得重新引入 Provider 默认实现。
- 任何后续性能优化必须保留现有取消、流位置、first-match、重复键、异常类型和原子提交语义，并以独立子进程 before/after 原始数据证明。

## Reviewer 注意事项

- `review.md` 当前仍为 `NEEDS_FIX`，本文件的 Review 修复记录只记录 Fixer 证据，不改写 Reviewer 结论；下一步必须由独立 Reviewer 重新审查 FIX-001..FIX-009。正式 Release 仍为 `PARTIAL`。
- 本轮未自动提交、推送、创建 PR 或发布 NuGet。

## Git 状态

- 工作区包含本轮测试边界、IVT、CI 和任务文档修改；旧测试文件因迁移显示为删除，新路径为未跟踪文件，待维护者审核后再决定暂存方式。
- 未执行任何破坏性 Git 操作。

## Review 修复记录

### Round 1（fixScope=recommended）

- 来源：`review.md`（Reviewer 结论 `NEEDS_FIX`，2026-09-17T13:03:00+08:00）；本节只记录 Fixer 实际变更和证据，不修改 Reviewer 文件。
- `FIX-001`（MUST_FIX）：已完成。新增 `ReviewFixCoreRegressionTest`，Common 收纳 30 个纯 Core/Abstractions 方法；Loader 拆为 Common 3 个纯方法与 NPOI 1 个 DI 方法；NPOI `ReviewFixRegressionTest` 仅保留 2 个 Provider wiring 方法。Common/NPOI 双 TFM 普通 TRX 分别 PASS 196/533。
- `FIX-002`（MUST_FIX）：已补齐实现和门禁。MiniExcel Integration 现有 4 个普通 Case，NPOI Integration 现有 26 个普通 Case，并分别增加 500K/1M 两个 `Category=Large` Theory Case；普通双 TFM TRX PASS，Large 仅通过显式 workflow dispatch 入口执行，本机未运行 Large。
- `FIX-003`（SHOULD_FIX）：已完成。`PublicApiContractTest` 按生产程序集维护精确 IVT 集合，覆盖合法错配、缺失和 `.Tests.Fake` 拒绝；Aggregate 双 TFM PASS 28。
- `FIX-004`（SHOULD_FIX）：已完成。新增 `ProviderRegistrationOrder_ShouldPreserveFirstRegistration`，验证 NPOI→MiniExcel 与 MiniExcel→NPOI 两种注册顺序；Aggregate 双 TFM PASS 28。
- `FIX-005`（SHOULD_FIX）：已完成实现。CI 普通矩阵使用 `Category!=Large`，`workflow_dispatch.run_large` 启动独立 Large Job；Release gate 执行单 Provider 100K probe、真实 IO probe，nuspec 版本改用 XML API 读取。外部 CI 未运行，仍为 NOT_VERIFIED。
- `FIX-006`（MUST_FIX）：部分完成。Provider comparison 现在每次只运行一个 Provider，header 含 `processId`、候选 benchmark assembly SHA-256、`providerFilter/providerCount`，峰值字段来自 `Process.PeakWorkingSet64`；两份独立 100K×3 after JSON 已生成并与源码一致。RawDateReader/Relation 1K/10K/100K before、Importer/preflight 语义安全矩阵仍未建立，故本项及代码总体不标 COMPLETED。
- `FIX-007`（MUST_FIX）：已完成文档纠偏。迁移矩阵和 symbol map 已改为真实路径/类名/方法名，补充 Large/DI/富工作簿方法和本轮 TRX 路径；迁移前双 TFM TRX 不存在，已明确记录 NOT_VERIFIED，不伪造逐 Case before。

### 本轮收口

- Fixer 可证明的职责、集成、IVT、CI、方法追溯和采样语义修复已落地；未验证项集中在 RawDate/Relation/Importer 性能基线、API candidate identity、维护者资源批准、500K/1M 实际执行、真实生产资源和外部 CI。
- 本轮执行状态保持 `PARTIAL`；`review.md` 仍保持 `NEEDS_FIX`，需要下一轮独立 Reviewer 依据新源码、TRX 和 artifact 重新验收。

### Round 2（fixScope=recommended）

- 来源：本任务 `review.md`（2026-09-17T14:45:00+08:00，状态 `NEEDS_FIX`）。本节只记录本轮 Fixer 的实际修改和验证，不改写 Reviewer 结论。
- `FIX-001`（MUST_FIX）：COMPLETED/RECHECKED。未重复搬迁；复核 Common 的 30 个 Core 方法、Loader 三个纯方法与 NPOI 的两个 Provider/一个 DI 方法，Release build 与 Common net8 复验通过。
- `FIX-002`（MUST_FIX）：COMPLETED。MiniExcel Integration 增加动态列真实文件、已有目标成功替换、非取消失败保留目标/临时清理，并补齐多 Sheet 全字段和 Large 首尾完整字段断言；NPOI Integration 增加模板命名区域/其它 Sheet 与真实图片锚点 reopen，并补齐 Large Count 断言。MiniExcel 普通 net8 7/7、NPOI 普通 net8 28/28；当前二进制 MiniExcel Large 500K/1M 2/2、NPOI Large 500K/1M 2/2 均通过。TRX 目录：`artifacts/tests/review-round2/miniexcel-integration-fix2-net8`、`npoi-integration-fix3-net8`、`miniexcel-large-final`、`npoi-large-final`。
- `FIX-003`（SHOULD_FIX）：COMPLETED/RECHECKED。精确 IVT 集合已存在，本轮未扩大 Friend；保留 reviewer 的已解决判断。
- `FIX-004`（SHOULD_FIX）：COMPLETED。Aggregate Unsupported 合同断言 `Code/Operation/Provider/Stage`，并增加两种注册顺序下的真实公共 RoundTrip；专项 Aggregate net8 3/3 通过，TRX：`artifacts/tests/review-round2/aggregate-contract-net8/aggregate-contract-net8.trx`。
- `FIX-005`（SHOULD_FIX）：COMPLETED（实现层）。CI 拆出 `resource-gate`、`benchmark-gate` 和 `api-gate`；resource 使用明确 100K 参数，benchmark 保留职责独立 100K probe，API 从 release gate 独立观察；外部 GitHub Actions 仍 NOT_VERIFIED。
- `FIX-006`（MUST_FIX）：PARTIAL。Provider probe 现在每个 Provider×sync/async 在独立子进程中运行，生产 Abstractions/Core/Provider/Benchmark 程序集全部纳入 candidate identity，保留原始样本并输出 elapsed/allocated/working-set/rows-per-second 中位数；小型 NPOI 3×2 probe 通过，产物 `artifacts/benchmarks/review-round2-probe.jsonl`。RawDate/Relation 的迁移前基线、语义安全矩阵和 preflight provider-shared 移动仍未完成，未伪造 before/after 收益。
- `FIX-007`（MUST_FIX）：COMPLETED（文档层）。迁移矩阵新增 30 个 Core 方法、2 个 NPOI Provider 方法和 4 个 Loader 方法的完整身份；symbol map 新增生产符号→测试方法成员级追溯，并明确 Theory `format=xls/xlsx`、TFM/TRX 与迁移前 NOT_VERIFIED。

### Round 2 汇总

- MUST_FIX：FIX-001、FIX-002、FIX-007 已完成；FIX-006 为 PARTIAL，原因是没有可合法生成的迁移前 RawDate/Relation 基线及维护者批准的延期。
- SHOULD_FIX：FIX-003、FIX-004、FIX-005 已完成实现并有本地验证；外部 CI/生产资源证据仍保持 NOT_VERIFIED。
- 回归验证：Release build（UseSharedCompilation=false）通过；Provider probe 小 workload 通过；Aggregate contract 3/3、MiniExcel Integration 普通 7/7、NPOI Integration 普通 28/28、两 Provider Large 各 2/2 通过。
- 当前终态：`PARTIAL`；不得将 review.md 改为 PASS，下一步是独立 Review。

### Round 3（fixScope=recommended）

- Review 状态：NEEDS_FIX
- Fix Scope：recommended
- Review 文件：`ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/review.md`
- 来源：本任务 `review.md` Round 3；本节只记录 Fixer 实际变更和证据，不修改 Reviewer 文件。

#### FIX-006

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/provider-shared/ExcelXlsxZipPreflight.cs`
  - `src/Bing.Offices.Core/Bing.Offices.Core.csproj`
  - `src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj`
  - `src/Bing.Offices.MiniExcel/Bing.Offices.MiniExcel.csproj`
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelRawDateSerialReader.cs`
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`
  - `benchmarks/Bing.Offices.Benchmarks/HotspotProbe.cs`
  - `tests/Bing.Offices.MiniExcel.Tests/MiniExcelProviderTest.cs`
- 根因：共享 preflight 位于 Core 路径；RawDateReader 为所有 numeric cell 建索引；Relation 每个 child 重复扫描父集合并重复调用 parent key；Importer 每行重复解析固定/动态物理列。
- 修复：RawDateReader 按目标日期列和首个有效行过滤；Relation 缓存父键并对可稳定哈希 comparer 建立首项索引，不可安全哈希时保留线性扫描；Importer 在表头阶段预绑定固定/动态列，首行缺少动态列时保留原有逐行 fallback；preflight 仅由两个 Provider 链接编译，Core 显式排除。
- 验证：
  - Round 3 的 hotspot 文件保留为历史证据；Round 4 已纠正 Relation before 控制流和 LOH 字段。当前可采信样本见 `benchmark-report.md` 与 `artifacts/benchmarks/review-fix-round4-*`，Relation 10K/100K before 明确为 NOT_VERIFIED。
  - `MiniExcelProviderTest` net6.0/net8.0：36/36 PASS；包含 `RawDateSerialReader_ShouldIndexOnlyRequestedDateColumnsAndRows`、`Relations_ShouldPreserveFirstMatchingParent` 及既有日期、动态列、取消、流位置、结构化错误和 preflight 回归。
  - `dotnet build Bing.Offices.sln -c Release --no-restore /p:UseSharedCompilation=false`：PASS，0 warning，0 error。

#### FIX-007

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/test-migration-matrix.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/provider-test-matrix.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/unit-test-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/integration-cross-provider-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/integration-npoi-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/integration-miniexcel-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/symbol-test-map.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/benchmark-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/final-report.md`
- 根因：执行报告和迁移矩阵仍引用上一轮旧测试目录、旧计数，并遗漏 Round 2/3 新增的真实文件、DI RoundTrip 和热点回归方法。
- 修复：按当前源码逐方法补入 NPOI 模板/图片、MiniExcel 动态列/替换/失败清理/异步 RoundTrip、Aggregate 真实 DI RoundTrip 和 MiniExcel RawDate/Relation 回归；六职责报告统一到 Round 3 TRX 及 196/533/36/29/28/7 展开 Case。
- 验证：
  - Common 196/196、NPOI Unit 533/533、MiniExcel Unit 36/36、Aggregate 29/29、NPOI Integration 28/28、MiniExcel Integration 7/7（net6.0/net8.0）均 PASS；NPOI Unit net6.0 另串行复验 533/533，排除并发临时目录竞争。
  - `git diff --check`：PASS；报告中只保留历史 Review 记录的旧路径，不改写 `review.md`。

### Round 3 汇总

- MUST_FIX：`FIX-006`、`FIX-007` COMPLETED（代码与本地证据）；迁移前双 TFM TRX、API candidate identity compare、维护者资源批准、500K/1M 外部候选 CI、真实 2 CPU/4 GiB 仍为 NOT_VERIFIED/BLOCKED。
- SHOULD_FIX：`FIX-001` 至 `FIX-005` 保持上一轮 COMPLETED/RECHECKED，本轮未回退。
- Round 3 历史快照状态为代码整改 `COMPLETED`，正式发布结论仍 `PARTIAL`；Round 5 已对其中的性能、身份账和 Relation 缓存结论进行修正，当前状态以 Round 5 汇总为准。

### Round 4（fixScope=recommended）

- 来源：本任务 `review.md` Round 4；本节记录 FIX-006、FIX-007、FIX-008 的实际修复与证据，不修改 Reviewer 结论。

#### FIX-006

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：PARTIAL
- 修改文件：
  - `benchmarks/Bing.Offices.Benchmarks/HotspotProbe.cs`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/benchmark-report.md`
- 根因：旧探针 Relation before 省略了 HEAD 的 DynamicInvoke/Comparer 控制流，LOH 字段也误称为 peak；旧 O(P×C) 控制流在 10K/100K 完整 before 场景不可接受地耗时。
- 修复：before Relation 改为 HEAD 源等价的逐 child、首父短路控制流；after 使用当前生产实现；LOH 字段改为 `lohSnapshotBytes`，采样 `GenerationInfo[3].SizeAfterBytes` 并在文档中标为快照；原始样本保留候选 identity、PID、GC、PeakWorkingSet 和中位数。Round 5 进一步确认 RawDate before 只是调查性 replay，不能作为历史 Reader 等价证据。
- 验证：
  - `dotnet build benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore -p:BaseOutputPath=artifacts/review-fix-build2/`：PASS，0 warning/0 error。
  - 当前 after `artifacts/benchmarks/review-fix-round4-hotspot-after2.jsonl`：RawDate 100K×3/10/30、Relation 1K/10K/100K，各 3 次：PASS。
  - before：RawDate 100K×3/10/30、Relation 1K，各 3 次：PASS；Relation 10K/100K 因旧 O(P×C) worker 未在运行窗口内完成，已终止，状态 `NOT_VERIFIED`。因此本 FIX 不能标记 COMPLETED，不能据此宣称完整 Relation 性能收益。

#### FIX-007

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/baseline.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/test-migration-matrix.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/execution.md`
- 根因：旧报告混用属性行、方法声明和 TRX 展开 Case，继续引用 571→582。
- 修复：Round 5 改用 `method-identity-audit.py` 的锚定属性/方法解析，记录 HEAD=569、当前=589、Added=20、Removed=0，并列出六项目分项 166/329/35/28/23/8；Theory 展开仍与方法总账分离，历史迁移前 TRX 仍为 NOT_VERIFIED。
- 验证：审计脚本同时读取 HEAD/current，输出 144 unchanged、391 moved、34 renamed/moved、20 added、0 removed；完整逐身份结果见 `method-identity-audit.md`。

#### FIX-008

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 执行状态：COMPLETED
- 修改文件：
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`
  - `tests/Bing.Offices.MiniExcel.Tests/MiniExcelProviderTest.cs`
  - `benchmarks/Bing.Offices.Benchmarks/HotspotProbe.cs`
- 根因：旧实现提前求值所有父键，改变了空 child、首匹配和未访问父委托的短路语义；直接全量索引又无法保持这些语义。
- 修复：Round 4 的惰性父键缓存在 Round 5 被撤回，生产代码恢复逐 child 的旧线性扫描，保留公开 ParentKey 的重复调用、首匹配、Navigation 和异常边界；空 child、首匹配和重复委托异常均有直接回归。删除会提前求值全体父项以及跨 child 缓存的实现。
- 验证：
  - `Relations_WithNoChildren_ShouldNotEvaluateParentKeys`：通过。
  - `Relations_ShouldStopAfterFirstMatchingParentBeforeEvaluatingLaterParents`：通过。
  - Round 4 历史 MiniExcel Unit net6.0/net8.0：38/38 PASS；Round 5 双 TFM：39/39 PASS，TRX 分别为 `artifacts/tests/review-fix-round5/miniexcel-net6/miniexcel-net6.trx`、`artifacts/tests/review-fix-round5/miniexcel-net8/miniexcel-net8.trx`。

### Round 4 汇总

- MUST_FIX：FIX-008 COMPLETED；FIX-006 PARTIAL（Relation 10K/100K before 无可接受完成证据）。
- SHOULD_FIX：FIX-007 COMPLETED。
- 已完成：Relation 控制流修复、直接回归、统一方法属性对账、LOH 采样命名和可复现 after 探针。
- PARTIAL：真实历史 baseline DLL、完整 Relation before 性能样本、Importer/Export/Roundtrip 同 workload 证据仍缺失。
- BLOCKED：维护者资源批准、外部候选 CI、生产 2 CPU/4 GiB、500K/1M 正式门禁仍按计划保持 NOT_VERIFIED/BLOCKED。
- 回归验证（Round 4 历史）：MiniExcel Unit 双 TFM 38/38；benchmark build PASS；Round 5 已完成完整解决方案 Release build（0 warning/0 error）和 MiniExcel Unit 双 TFM 39/39。
- 下一步：保持执行终态 `PARTIAL`，由独立 Reviewer 重新审查；不自动提交、推送、PR 或发布。

### Round 5（fixScope=recommended）

- 来源：本任务 `review.md` Round 5；本节只记录 Fixer 实际变更和验证，不修改 Reviewer 结论。

#### FIX-006

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：PARTIAL。
- 修正 `benchmark-report.md` 的 RawDate 100K×30 独立中位数：allocated=`2,410,420,240 B`，LOH=`337,756,880 B`；LOH 明确为完整 GC 后 `GenerationInfo[3].SizeAfterBytes` 快照，不是峰值。
- 报告补充 RawDate before 是调查性 numeric-cell replay、不是历史 Reader DLL；Relation 10K/100K before、Importer/Export/Roundtrip 同 workload 仍 `NOT_VERIFIED`，因此不宣称未证明的性能收益。

#### FIX-007

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- 新增 `method-identity-audit.py` 与生成的 `method-identity-audit.md`，同一解析器读取 HEAD 与当前六职责源码，锚定独立 Fact/Theory 属性并关联方法声明。
- 对账结果：HEAD=569，当前=589，Added=20，Removed=0；其中 Unchanged=144、Moved=391、Renamed/Moved=34。明确记录 `DataRowStartIndex...AcrossProviders`、`ReadColumnRange...AcrossProviders` 改名，覆盖门禁拆分和 Round 5 Relation 新测试；迁移前双 TFM TRX 仍 `NOT_VERIFIED`。

#### FIX-008

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED。
- 撤回 Round 4 的惰性父键缓存/索引，恢复逐子项 `ParentKey.DynamicInvoke` 的旧线性扫描，保持公开委托重复调用、首匹配、Navigation 与 `DynamicInvoke` 异常边界。
- 新增 `Relations_ShouldPreserveRepeatedParentKeyEvaluationAndErrorBoundary`，与空子集、首匹配回归一起验证 `keyCalls=2`、首 child 已绑定和一个结构化 Relationship 错误。
- MiniExcel Unit net6.0/net8.0 均 `39/39 PASS`；TRX：`artifacts/tests/review-fix-round5/miniexcel-net6/miniexcel-net6.trx`、`artifacts/tests/review-fix-round5/miniexcel-net8/miniexcel-net8.trx`。

#### FIX-009

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- 确认 `benchmarks/Bing.Offices.Benchmarks/artifacts/review-fix-build*` 仅为本轮生成的 build 输出后，以可恢复方式移动到根 `artifacts/review-fix-quarantine/review-fix-build` 与 `review-fix-build2`，并删除空的原嵌套目录。
- 重新扫描 `src/tests/benchmarks`：无 `artifacts`、`TestResults`、`BenchmarkDotNet.Artifacts` 嵌套目录；quarantine 路径保留原始输出，未执行 Git 清理操作。

### Round 5 汇总

- MUST_FIX：FIX-008 COMPLETED；FIX-006 仍 PARTIAL，原因是可信的 RawDate 历史 Reader、Relation 10K/100K before 和 Importer/Export/Roundtrip 同 workload 证据未形成。
- SHOULD_FIX：FIX-007、FIX-009 COMPLETED。
- 代码/文档验证：MiniExcel Unit 双 TFM 39/39；解决方案 Release build 0 warning/0 error；`method-identity-audit.py` 生成 569→589 的零移除身份账；`git diff --check` 通过（仅换行规范 warning）。
- `review.md` 保持 Reviewer 的 `NEEDS_FIX` 原文；正式发布结论保持 `PARTIAL`，不自动提交、推送、PR 或发布。

### Round 6（fixScope=recommended）

- 来源：本任务 `review.md` Round 6；本节记录 FIX-006/FIX-007 的实际修复和验证，不修改 Reviewer 结论。

#### FIX-006

- 严重程度：HIGH；处理要求：MUST_FIX；执行状态：COMPLETED（按计划撤回未被可信 baseline 证明的优化）。
- 根因：RawDate 的列/行过滤与重复 Query 使用调查性 replay 作为前后对照，缺少真实 HEAD Reader baseline；旧 Relation after 也来自已撤回实现，不能继续作为当前收益证据。
- 修复：恢复 `MiniExcelRawDateSerialReader.Read(Stream,string,CancellationToken)` 的 HEAD 全量 numeric-cell 索引和原始流位置合同；恢复同步/异步导入先读 RawDate、再单次 Query 的控制流。保留 Importer 物理列预绑定仅用于 `ReadColumnRange` 的绝对列错误定位，不宣称性能收益；Relation 仍为逐 child 首匹配扫描。更新 `benchmark-report.md`、`final-report.md`，将历史 JSONL 标为调查资料，纠正 LOH 采样含义并禁止重标旧 identity。
- 修改文件：
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Internals/MiniExcelRawDateSerialReader.cs`
  - `src/Bing.Offices.MiniExcel/Bing/Offices/Imports/MiniExcelExcelImporter.cs`
  - `tests/Bing.Offices.MiniExcel.Tests/MiniExcelProviderTest.cs`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/benchmark-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/final-report.md`
- 验证：
  - MiniExcel Unit net6.0：39/39 PASS，TRX=`artifacts/tests/review-fix-round7/miniexcel-net6/miniexcel-net6.trx`。
  - MiniExcel Unit net8.0：39/39 PASS，TRX=`artifacts/tests/review-fix-round7/miniexcel-net8/miniexcel-net8.trx`。
  - `git diff --check`：PASS（仅换行规范提示）。

#### FIX-007

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- 根因：原跨 Provider 方法迁移后只保留 MiniExcel 分支，方法身份总账无法证明 NPOI 行为仍有覆盖。
- 修复：新增 `NpoiProviderBoundaryRegressionTest`，恢复原 1900/1904 与负/serial60/61 日期数组、四组 DataRowStartIndex Theory（含 dynamic、SourceRows、converter/validation 物理上下文）及 ReadColumns 结构化错误列定位；更新 test-migration-matrix、provider-test-matrix、symbol-test-map，明确每个 NPOI 分支、Case、TFM 与 TRX 去向。重新运行 method identity audit，当前为 HEAD 569、After 592、Added 23、Removed 0。
- 修改文件：
  - `tests/Bing.Offices.Npoi.Tests/NpoiProviderBoundaryRegressionTest.cs`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/test-migration-matrix.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/provider-test-matrix.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/symbol-test-map.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/method-identity-audit.md`
- 验证：
  - NPOI Unit net6.0：539/539 PASS，TRX=`artifacts/tests/review-fix-round7/npoi-net6/npoi-net6.trx`。
  - NPOI Unit net8.0：539/539 PASS，TRX=`artifacts/tests/review-fix-round7/npoi-net8/npoi-net8.trx`。
  - 保持 MiniExcel Unit 双 TFM 39/39 PASS；MiniExcel Unit 不引用 `Bing.Offices.Npoi` ProjectReference。

### Round 6 汇总

- MUST_FIX：FIX-006 COMPLETED（未证明 RawDate 性能优化撤回；Relation 缓存此前已回退；Importer 预绑定仅保留正确性语义）。
- SHOULD_FIX：FIX-007 COMPLETED（NPOI 三组原分支补回并完成矩阵/符号追溯）。
- 回归验证：NPOI Unit 双 TFM 各 539/539；MiniExcel Unit 双 TFM 各 39/39；方法身份审计 569→592、Removed=0；`git diff --check` PASS。
- 任务整体与正式发布：仍为 `PARTIAL`，API candidate identity、维护者资源批准、真实 2 CPU/4 GiB、外部候选 CI 和正式发布条件未验证。
- 下一步：由独立 Reviewer 重新审查；不自动提交、推送、PR 或发布。

### Round 7（fixScope=recommended）

- 来源：本任务 `review.md` Round 7；本节记录 FIX-006 的实际修复，不修改 Reviewer 结论。

#### FIX-006

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- 根因：RawDateReader 的过滤重载已撤回，但 HotspotProbe 仍反射旧五参数签名；当前摘要还误称 Importer 物理列预绑定保留。
- 修复：将探针改为反射当前三参数 `Read(Stream,string,CancellationToken)`，缺少签名时显式抛 `MissingMethodException`；将 before header 改为调查性 numeric-cell replay，明确不等价于 HEAD Reader。更新 benchmark/final/execution 当前摘要，明确 RawDate、Relation 缓存和 Importer 预绑定均已撤回，历史 JSONL 不重标。
- 修改文件：
  - `benchmarks/Bing.Offices.Benchmarks/HotspotProbe.cs`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/benchmark-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/final-report.md`
  - `ai_docs/tasks/BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001/execution.md`
- 验证：
  - `dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --hotspot-worker artifacts/benchmarks/review-fix-round8-raw-date.jsonl after raw-date 3 3`：PASS；3/3 样本，`indexCount=300000`，退出码 0，未再出现悬空反射重载异常。
  - `dotnet build Bing.Offices.sln -c Release --no-restore`：PASS，0 warning、0 error。
  - `dotnet test tests/Bing.Offices.MiniExcel.Tests/Bing.Offices.MiniExcel.Tests.csproj -f net6.0 -c Release --no-build --no-restore`：PASS，39/39；TRX=`artifacts/tests/review-fix-round8/miniexcel-net6/miniexcel-net6.trx`。
  - `dotnet test tests/Bing.Offices.MiniExcel.Tests/Bing.Offices.MiniExcel.Tests.csproj -f net8.0 -c Release --no-build --no-restore`：PASS，39/39；TRX=`artifacts/tests/review-fix-round8/miniexcel-net8/miniexcel-net8.trx`。
  - `git diff --check`：PASS（仅换行规范提示）。

### Round 7 汇总

- MUST_FIX：无。
- SHOULD_FIX：FIX-006 COMPLETED。
- 回归验证：Round 7 NPOI/MiniExcel Unit TRX 与 NPOI 边界定向测试保持通过；Round 8 MiniExcel Unit 双 TFM 各 39/39、hotspot worker 3/3 PASS，解决方案 Release build 0 warning/0 error，`git diff --check` PASS。
- 任务整体与正式发布：仍为 `PARTIAL`，API candidate identity、维护者资源批准、真实 2 CPU/4 GiB、外部候选 CI 和正式发布条件未验证。
- 下一步：由独立 Reviewer 重新审查；不自动提交、推送、PR 或发布。
