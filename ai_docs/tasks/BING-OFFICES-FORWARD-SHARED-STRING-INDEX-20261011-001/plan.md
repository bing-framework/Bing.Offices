# Forward 摘要模式共享字符串磁盘索引

**Task ID：** `BING-OFFICES-FORWARD-SHARED-STRING-INDEX-20261011-001`

## 父级待办

本任务是 F-006“端到端低内存 Entity 导入”的一个独立实现项，依据 `BING-OFFICES-ENTITY-FORWARD-BATCH-20261006-001` 的开放 TODO。完成本子项不代表 F-006 全部完成。

## 问题与目标

`ExcelForwardXlsxSharedStrings.ReadDiskBacked` 已将共享字符串文本写入 `DeleteOnClose` 临时文件，但仍以 `List<long>`、`List<int>` 保存每项偏移和长度，并在返回前复制为两个数组。大量短或空字符串会使索引元数据随条目数增长，构造尾段还会同时保留列表容量和数组。

在不改变公开 API、`IReadOnlyList<string>` 读取语义、内存模式和 Provider 行为的前提下，将磁盘模式索引也移入定长记录临时文件，使内存索引状态仅保留计数和固定缓冲。

## 实施内容

1. 共享字符串正文仍写入独立临时文件；每项以 8 字节偏移和 4 字节长度写入顺序索引文件。
2. 按索引号定位定长记录，再读取正文；对索引、偏移和长度执行边界检查，并串行保护两个文件流的位置状态。
3. 成功对象负责释放两个 `DeleteOnClose` 文件；解析取消、损坏 XML、写入失败和构造失败均关闭并清理两者。
4. 为职责测试提供内部临时目录重载，不增加公共成员；NPOI/ClosedXML 测试项目链接同一测试源码。
5. 保持 `FromValues` 内存模式与 Complete 导入路径不变；沿用 sharedStrings XML 字节限制，不新增配置或依赖。

## 验收

- 大量空、短、Unicode 和富文本共享字符串可按任意索引读取并完整枚举。
- 共享字符串数量为 N 时，磁盘索引大小为 `12 * N` 字节；索引不再使用按条目数增长的托管列表/数组。
- 正常释放、预取消和 XML 解析异常后，测试专用临时目录均无残留文件。
- NPOI 与 ClosedXML 的 Forward `SummaryOnly` 现有行为在 net6.0、net8.0 保持通过；生产 Provider 双 TFM 编译通过。
- 无公共 API、包依赖、CI workflow/job/matrix、API snapshot 或业务结果变化。

## 边界

该子项降低共享字符串索引的托管内存，不构成 DOM、完整结果实体集合、所有 Provider 临时状态或整体进程峰值的端到端内存保证。F-006 的剩余项继续开放；不执行 Excel 客户端验证。
