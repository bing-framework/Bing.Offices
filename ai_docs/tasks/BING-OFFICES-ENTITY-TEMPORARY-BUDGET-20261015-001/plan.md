# Entity 分批导入临时存储预算

Task ID：`BING-OFFICES-ENTITY-TEMPORARY-BUDGET-20261015-001`

## 目标与范围

继续本地 TODO F-006 的资源约束代码工作，新增可选的 Entity 批次临时磁盘预算。现有行数、错误数、唯一值、关系键和 ZIP/XML 限制已有实现，不重复实现。不把本阶段等同于 F-006 的全部内存预算完成。

## 实施

1. 保留既有构造函数，增量增加 `ExcelEntityBatchImportOptions` 五参数构造函数与只读 `MaxTemporaryStorageBytes`。null 保持原行为，正数限制单次导入同时存活的受管临时文件逻辑长度总和。零、负数拒绝。公共成员记录为候选，不能自行签署审批或覆盖正式 API baseline。
2. 使用共享内部预算与 FileStream 实现，写入/扩容前预留增长量；覆盖覆盖写、定位后写、同步/异步写、SetLength、缩小及释放。超限抛出带 Provider/Import/Open 元数据的 ResourceLimit 异常。各次调用使用独立实例，不使用静态环境状态。
3. NPOI/ClosedXML Forward 与 SummaryOnly 接入。预算覆盖输入暂存、共享字符串正文和索引；NPOI XLS 还覆盖转换 XLSX 和逐 Sheet XML。启用预算时共享字符串使用磁盘索引，Complete 仍保留其约定的完整实体结果。
4. Buffered+Complete、不受控的模板预检与失败工作簿输出在指定新预算时预检 Unsupported，防止漏算。调用方源/目标流、文件系统块开销、引擎缓存及实体堆不属于该预算，不宣称恒定总内存。
5. 测试覆盖边界、聚合、释放、并发调用隔离、真实异步写、取消、异常清理与所有权；NPOI XLS/XLSX 分支均覆盖。维护使用说明、成员追溯与生产符号测试映射。

## 验证与影响

受影响项目：Abstractions、NPOI、ClosedXML 及对应测试；TFM：抽象 netstandard2.0，Provider net6.0/net8.0。无依赖、打包拓扑或版本变化。风险 MEDIUM：新增选择性资源限制及临时 IO。

先运行预算与 Entity Forward 定向测试，再运行受影响双 TFM 测试；公共 API 候选检查与 consumer 编译，阶段收口执行构建/测试与差异、严格 UTF-8/BOM/EOL 检查。环境或既有 API 审批阻塞单独记录，不通过改 baseline 消除。无需重跑未受影响的百万行性能矩阵。

## 完成标准

新预算进入真实 Provider 主调用链；超限前停止增长且清理临时文件，源流保持调用方所有权；默认行为兼容；不支持组合在读取或回调前拒绝；直接测试和执行报告可追溯。既有未审批成员、Excel 客户端验证及 F-006 堆内存预算仍保持原状态。
