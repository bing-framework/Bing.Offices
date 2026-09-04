# 执行进度

| Task | Status | Scope | Evidence | Risk | Remaining | Updated |
| --- | --- | --- | --- | --- | --- | --- |
| P0-01 | VERIFIED | 环境、Git、项目和输入核验 | `baseline.md`；`dotnet sln list`；`dotnet nuget list source`；`dotnet --info`；HEAD `54c641d` | 评审输入缺失 | 后续若环境变化需追加基线 | 2026-09-01 |
| P0-02 | PARTIAL | restore、Release build、Unit/Integration/Docs、pack | 受控 TFM 已收敛为 net8/net6/netcoreapp3.1；最终 Release build `0 error / 15 warning`；正式 API compare 退出码 0；Unit net8/net6 各 `384/384` | netcoreapp3.1 runtime 缺失；pack/consumer 全矩阵、长路径 SDK/MSBuild 限制和交付复现仍未收口 | 完成 pack/consumer 矩阵并评估警告；安装缺失 runtime 后补跑 | 2026-09-03 |
| P0-03 | VERIFIED | 弃用、兼容、execution detail、未实现扫描 | `deprecated-removal.md`：受控 UTF-8 扫描、当前生产 `[Obsolete]` 精确命中为 0、71 个生产 EditorBrowsable.Never、3 个窄范围 capability fallback、0 个生产 Result/Wait/Task.Run/TODO、6 个合规测试 IVT；Public API 183 个类型分类及 94 个 execution detail 逐符号附录 | 剩余 compatibility/public execution detail 仍需具名维护者逐项批准；P0-03 VERIFIED 仍只表示扫描/证据闭环 | 对剩余候选按 D-013 补批准或后续 task；Round 5 删除集已完成 API diff 和回归 | 2026-09-02 |
| P1-01 | VERIFIED | Document Factory | 请求配置已按方向合并；回归和 plan 链路通过 | API 仍为 public execution detail | 由 API 治理决定是否 internal 化 | 2026-09-01 |
| P1-02 | VERIFIED | v1 JSON/XML 迁移方向 | 非目标方向保持 `null`；JSON/XML、string/stream、fallback 和 DTD 测试通过 | 完整 TFM 回归受 runtime 限制 | 记录到最终矩阵 | 2026-09-01 |
| P1-03 | VERIFIED | Relation Binder 异常 | 原始异常类型经 `ExceptionDispatchInfo` 重抛；委托边界回归通过 | 无 | 记录到最终矩阵 | 2026-09-01 |
| P1-04 | VERIFIED | CSV 公式策略和 Options | BOM/控制字符/Unicode whitespace 前缀防护、Preserve、RFC4180、culture 和非法 options 通过 | DataTable 兼容层仍是待治理 public API | 完成兼容层 API 裁决 | 2026-09-01 |
| P1-05 | VERIFIED | 错误分类、DI、未实现路径 | `AddBingOfficesNpoi` 链式入口已验证；未知能力改为 `NotSupportedException`；NPOI 2.7.4 HSSF 对 `Hidden`/`Collapsed` 明确抛 `NotImplementedException`、XSSF 三项均支持，Failure Workbook 三处 capability fallback 合法保留 | 旧 exceptions 仍存在并待 API 治理；fallback 仅覆盖对应 row capability | 维护 API 台账；不得将 capability fallback 扩展为 catch-all | 2026-09-01 |
| P1-06 | VERIFIED | 资源限制和 ResourceProbe | 独立 Excel child probe 7/7；另有独立 mapping/unique probe 16/16；未预开同一 Workbook；资源报告已生成 | 不能证明 DOM 前完整内存上限 | 补 Failure Workbook 资源样本或明确 deferred | 2026-09-02 |
| P1-07 | VERIFIED | Failure Workbook 和输出状态 | AnnotatedOriginal/ErrorRowsOnly、MaxSerializedBytes、取消、清理诊断、目标流保护专项 14/14 | 双 DOM 峰值仍需独立资源样本 | 补最终边界和发布阻断 | 2026-09-01 |
| P2-01 | PARTIAL | 删除候选取证和迁移 | Round 6 已删除 `OfficeException` 层级、`ExcelSetting`、`SheetSetting`；类型映射异常迁移为标准异常；正式 API baseline 已更新并通过 compare/Unit | DataTable 显式 `CsvHelper`、其它 execution detail 仍未删除 | 保留剩余候选的真实 PARTIAL/No-Go，后续继续逐符号治理 | 2026-09-03 |
| P2-02 | PARTIAL | API 分层 | `ExcelMappingPlanFactory`、类型映射、绑定解析器、默认 loader、CSV concrete 已 internal；NPOI/Benchmark/ResourceProbe 已改用公开 Provider SPI；DI 通过公开接口注册；正式 baseline 已记录 Provider `cacheCapacity` | `MappingConfigurationMerger`、剩余 public execution detail 仍为跨程序集公开边界 | 完成剩余逐符号治理和跨程序集边界收敛 | 2026-09-03 |
| P2-03 | VERIFIED | DI/入口收敛 | `AddNpoi` 已删除；`AddBingOfficesNpoi` 返回同一 `IServiceCollection`；Unit/Integration/Docs/consumer 通过 | 正式 API hash 未更新 | 维护者批准后更新 formal baseline | 2026-09-01 |
| P3-01 | VERIFIED | NPOI import 拆分 | `NpoiImportSheetExecutor` 是唯一 Sheet 列绑定/动态列/表头执行实现；删除 `NpoiExcelImporter` 中未调用的重复实现；Release build 和缓存/StreamPipeline 回归通过 | 完整 Unit 受 API hash 阻断；DOM 取消延迟仍受 NPOI 限制 | 纳入 Round 8 最终矩阵，保持 No-Go | 2026-09-02 |
| P3-02 | PARTIAL | CSV/Failure 拆分 | 新增 `CsvPipelineSupport.cs`，拆出 Reader/Writer/LimitedStream/异常支持职责；Docs consumer `11/11`，StreamPipeline net6/net8 各 `90/90` | Failure Workbook 产物构建/临时提交尚未拆分 | 补充专项回归，继续保持最小拆分 | 2026-09-02 |
| P3-03 | VERIFIED | Mapping hot path | `ExcelMappingPlanCacheKey` 已提取；Provider SPI 支持显式 cache capacity；Tenant eviction 基准实际创建对应容量并断言淘汰，新增 capacity 回归；Release build 通过 | rule-index/dynamic compiler 尚未拆出；完整 Benchmark 尚未批准 | 以统一 Benchmark 矩阵和预算决定后续拆分 | 2026-09-02 |
| P3-04 | VERIFIED | 异步/所有权审计 | `src/` 生产路径无 `.Result`/`.Wait()`/`Task.Run`；Round 8 决策记录同步 API、DOM 取消延迟、Workbook/Stream/temp/enumerator 所有权；既有取消和流回归通过 | Failure Workbook 双 DOM 峰值和完整取消延迟量化仍未形成 RC 门禁 | 纳入最终矩阵，保持 No-Go | 2026-09-02 |
| P4-01 | PARTIAL | 测试矩阵和职责覆盖 | 关键 P0/P1 方法已与 net6/net8/Docs/Integration 结果绑定；失败工作簿专项 14/14；正式 baseline 后 Unit net6/net8 各 `384/384` | 旧 TFM 未运行；完整方法追溯仍需补齐 | 补完整方法追溯和缺失 TFM 回归 | 2026-09-03 |
| P4-02 | PARTIAL | isolated nupkg consumer | `artifacts/package-consumer-rerun2` 无 ProjectReference；Round 5 `packages-round5` 的 2.0.0 nupkg 在短路径 `C:\nupkg-cache-round5` restore/build/run 退出码均为 0，输出 `package-consumer-ok`；任务深路径缓存仍受 `MSB3106` 阻断 | 仅 net8 consumer；长路径 SDK/MSBuild 限制；正式包身份仍待不可变治理 | 记录最新包 hash；后续补其它 TFM/正式 feed 或解决环境限制 | 2026-09-02 |
| P4-03 | VERIFIED | 四类测试报告 | Unit/Integration/Docs/Package consumer 报告已创建并包含真实退出码和限制 | Unit 为 PARTIAL，不得写全绿 | API baseline 裁决后追加最终结果 | 2026-09-01 |
| P5-01 | PARTIAL | Benchmark 基线 | StreamPipeline 9/9 ShortRun 完成；Round 8 修复 TenantPlanCache 实际容量与 `CACHE_EVICTION` 证据，并有 net8 直接回归；补充基准产物已读取 | CSV/取消/完整矩阵与正式预算未完成 | 补计划场景并统一提交身份 | 2026-09-02 |
| P5-02 | DONE | 证据驱动优化 | 本轮未做无证据性能优化；保留真实分配和异常计数 | 需要正式预算和 before/after 才能优化 | 由维护者批准后继续 | 2026-09-01 |
| P5-03 | VERIFIED | 独立资源报告 | Excel ResourceProbe 7/7；mapping/unique probe 16/16；LOH/workset、尾延迟和 DOM 边界已记录 | budgetStatus 为 `UNAPPROVED` | 补 Failure Workbook 样本或明确不纳入 RC | 2026-09-02 |
| P6-01 | PARTIAL | 文档/XML/包同步 | DOM、ownership、MaxSerializedBytes、AddBingOfficesNpoi 已同步部分文档 | API 候选和正式 baseline 尚未收口 | 完成文档/API/包内容交叉检查 | 2026-09-01 |
| P6-02 | BLOCKED | 独立 Review/发布门禁 | Round 9 独立 Review 保持 `BLOCKED`；正式 API baseline 和可运行 TFM Unit 已收口；RC No-Go 仍由缺失 runtime、未批准性能预算、ResourceProbe 边界、剩余 API 治理、consumer 环境和 clean-clone 证据组成 | 独立 Reviewer 尚未对本次 baseline 收口再次验收；仍无具名延期/waiver | 完成剩余门禁后交回独立 Reviewer；保留 `BLOCKED` | 2026-09-03 |
