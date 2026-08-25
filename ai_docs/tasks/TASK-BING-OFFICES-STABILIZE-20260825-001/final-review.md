# 最终审查

## 执行结论

结论：`PARTIAL`，`no-release`。

本任务已完成核心 P0 正确性修复、API 收敛、Profile Core 迁移、CSV/NPOI 低风险内部拆分、可支持目标框架的 Unit/Integration 验证、Docs Tests、Benchmark smoke、Release build/pack 和本地包元数据审计。未执行 commit、push、tag、PR、NuGet publish 或其它破坏性 Git 操作。

当前不建议发布。原因是批准计划仍有 P1 工作未完成，且本轮没有伪造缺失证据：完整 Stream import/export 性能前后对比、高影响路径优化、NPOI importer/exporter 其余职责拆分、本地包安装消费者链路、Office 互操作和部分旧 runtime 运行验证仍未完成。

## 已完成范围

- ValidationMode 四模式、Continue 错误收集、Unique journal 回滚和失败工作簿 Sheet identity 修复。
- 模板 Preserve/Replace 策略和字段级回归。
- Mapping 单一最终合并边界、方向缺失失败和显式 convention fallback。
- JSON/XML 安全限制、DTD/外部实体防护和 v1 显式迁移。
- Profile 只读 resolver、稳定名称注册重载、Core-only 注册入口和程序集部分加载容错。
- 删除方向不安全 loader facade 和伪异步 CSV API；同步 API baseline、文档和 Docs consumer。
- `NpoiStreamCopier`、`NpoiFailureWorkbookWriter`、`CsvHeaderBinder`、`CsvDynamicTypeResolver`、`CsvPropertyBinding` 内部职责拆分。
- NuGet 2.0.0 本地包审计：三个包均有 `README.md`、`LICENSE`、`icon.png` 和 `.snupkg`；nuspec 声明 README；SourceLink 不进入运行时 dependencies。

## 验证结果

- net8 Unit：213/213。
- net6 Unit：213/213。
- net8 Integration：11/11。
- net6 Integration：11/11。
- Docs consumer：8/8。
- Profile/API 聚焦：17/17。
- Stream/失败工作簿聚焦：133/133；对应 Integration：11/11。
- CSV 聚焦：net8 24/24，net6 24/24。
- Mapping Benchmark smoke：208 个基准执行，无运行时配置异常。
- 固定 Mapping cache-hit 基线：16 个参数组合实际执行。
- ResourceProbe：16/16。
- Release build、pack：通过；未发布。
- 静态错误检查：无错误；`git diff --check` 无尾随空白错误，仅报告既有 CRLF/LF 转换提示。

## PARTIAL 项

### P1-REMAINDER-01：NPOI importer/exporter 与 Mapping loader 继续拆分

已完成流复制、失败工作簿 writer 和部分 CSV 管线拆分；NPOI importer/exporter 的 sheet reader、row materializer、validation pipeline、template/style/relationship 等其余职责仍在原文件中。Mapping configuration loader 的 JSON/XML parser、共享 validator 和 migration 仍未完成进一步拆分。

解除条件：按同样的最小内部协作者策略继续实现，每个抽取点完成直接或公共 API 回归，并通过 net8/net6 可支持测试。

### P1-REMAINDER-02：性能前后对比和高影响路径优化

已完成固定 Benchmark Job、Mapping cache-hit 16 组合基线和 ResourceProbe，但未完成完整 Stream import/export workload、before/after 同负载数据和优化阈值。未对 MemoryStream 整块复制、failure workbook `ToArray()`、cache key 等路径做无证据优化。

解除条件：固定相同输入规模、TFM、Job 和环境，完成 before/after 数据、分配/GC/输出大小记录及回归阈值。

### P1-BLOCKED-03：旧 runtime Unit 执行

本机缺少 .NET 5、.NET 7、.NET Core 3.1 runtime。对应项目已完成编译，但 testhost 无法启动，未安装 runtime，也未绕过测试执行。

解除条件：在具备对应 runtime 的 CI 或受控机器上分别运行 net5/net7/netcoreapp3.1 Unit，并记录原始命令和结果。

### P1-BLOCKED-04：Office 互操作

当前环境未提供可控的 Excel、LibreOffice 或 WPS 互操作执行环境。本轮保留并通过 NPOI reopen 和 xUnit 验证，不宣称 Office 应用兼容性。

解除条件：提供可控安装版本、输入/输出样本和隔离执行环境，完成 Excel/LibreOffice/WPS 打开、保存和关键字段核验。

### P1-REMAINDER-05：Profile 扫描直接异常测试与本地包消费者链路

生产代码已捕获 `ReflectionTypeLoadException` 并保留可加载类型，但由于 Core 没有测试友元，本轮未新增生产 `InternalsVisibleTo` 或测试出口来强行构造异常。稳定 alias 已通过显式 `AddMappingProfile<T>(name)` 注册重载提供；程序集扫描继续使用 FullName 兼容 fallback，未新增未锁定的公开 attribute。Docs consumer 已通过，但尚未完成独立安装最新 nupkg 的本地消费者链路。

解除条件：采用不扩大生产 API/IVT 的测试方案覆盖异常路径，并以临时本地 NuGet 源安装三个最新包后运行消费者编译/测试。

## 发布判定

- 判定：`NO-RELEASE`。
- 版本建议：本任务包含删除方向不安全入口、删除伪异步入口和方向语义变更；后续发布应使用下一个 major 或明确 pre-release，并保留迁移文档。
- 本轮没有执行包发布、版本号修改或外部系统操作。

## 追踪文件

- 计划：[plan.md](plan.md)
- 执行记录：[execution.md](execution.md)
- 验证记录：[verification.md](verification.md)
- 进度记录：[progress.md](progress.md)
- API 治理：[api-governance.md](api-governance.md)
- 性能基线：[performance-baseline.md](performance-baseline.md)
