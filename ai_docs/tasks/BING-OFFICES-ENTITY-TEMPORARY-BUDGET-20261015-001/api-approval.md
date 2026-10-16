# 公共 API 候选与迁移

状态：`BLOCKED_APPROVAL`。用户授权继续实现剩余 TODO；本文件记录具体新增成员供审阅，不代替成员级签署，不修改正式 baseline，也不沿用此前成员审批推定本成员获批。

| 新增成员 | 用途 | 兼容性 |
| --- | --- | --- |
| `ExcelEntityBatchImportOptions(int, ExcelEntityImportOptions, ExcelEntityBatchParsingMode, ExcelEntityBatchResultMode, long?)` | 显式指定临时存储聚合预算 | 保留全部既有构造函数，旧四参数构造委托新构造并传 null |
| `ExcelEntityBatchImportOptions.MaxTemporaryStorageBytes { get; }` | 查询调用选项 | 新增只读属性，不修改既有签名与枚举值 |

迁移无需改动既有调用；选择五参数构造启用预算。首版仅支持 NPOI/ClosedXML Forward 或 SummaryOnly 的非模板、无失败输出导入。使用说明见 `docs/excel/batch-import.md`。构造函数拒绝非正值，不设置默认容量阈值；调用方根据磁盘配额决定额度。

双 TFM 候选快照与包消费者验证结果在 execution.md 记录；发布审批与实现正确性分别报告。历史 API 待审批项不在本轮自动批准。

2026-10-15 实测：双 TFM NuGet consumer 通过；候选见 api-candidate.json。完整 Abstractions 快照相对工作区 baseline 各新增 29 条、删除 0 条，其中本轮为以上两成员，另 27 条属于此前候选。本轮未更新正式 baseline。
