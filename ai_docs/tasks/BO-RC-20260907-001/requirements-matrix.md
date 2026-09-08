# 需求矩阵

| 需求 | 当前状态 | 证据/说明 |
| --- | --- | --- |
| NPOI/Unit/Integration net6 + net8 | COMPLETED locally | `framework.props`、测试 csproj、双 TFM build/test |
| User/Provider User/SPI/Internal 分层 | PARTIAL | 公开 concrete provider、NPOI 扩展 namespace 已完成；正式 baseline 待审批 |
| NPOI namespace 漂移 | COMPLETED locally | 扩展迁移到 `Bing.Offices.Npoi.Extensions`；docs/tests/consumer 通过 |
| Excel ImportAsync | COMPLETED locally | `CopyAsync` 输入、同步 DOM、Async Failure staging、直接测试 |
| Excel ExportAsync/FileAsync | COMPLETED locally | 同步 NPOI staging + `CopyToAsync`/`CommitAsync`、直接测试 |
| CSV ImportAsync | COMPLETED locally | CsvHelper `ReadAsync`、异步 parser callback、直接测试 |
| CSV ExportAsync/FileAsync | COMPLETED locally | CsvHelper `NextRecordAsync`/`FlushAsync`、`CommitAsync`、直接测试 |
| Task.Run 伪异步 | COMPLETED | 生产代码未新增，Async 生产路径无 Task.Run/Result/Wait |
| Cancellation/ownership | PARTIAL | pre-cancel、AsyncOnly、Failure staging 已覆盖；完整大文件/锁矩阵待回归 |
| Sync/Async parity | PARTIAL | 直接 CSV/Excel happy path 与错误边界已覆盖，完整动态列/转换/校验矩阵待扩展 |
| 双 TFM Unit/Integration | COMPLETED locally | 两 TFM Unit 各 509 pass + API approval block；Integration 各 15/15 |
| PackageReference-only consumer | COMPLETED locally | 任务专属 feed/cache，net6/net8 均 `package-consumer-ok` |
| Benchmark workload 修复 | COMPLETED locally | Unique 输入 GlobalSetup、Dynamic factory 分离、PeakWorkingSet 移出普通 benchmark |
| Sync/Async benchmark | COMPLETED smoke | CSV 1K/10K/100K/1M Dry 原始结果已保存；正式预算未批准 |
| API approval | BLOCKED external | 需要维护者提供 `approvedBy/approvedAt` 和成员 diff 审批 |
| 性能/资源预算 | BLOCKED external | smoke/resource probe 有证据，尚无批准阈值 |
