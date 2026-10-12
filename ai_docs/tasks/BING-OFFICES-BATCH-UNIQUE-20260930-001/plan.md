# ExcelDataReader 固定列分批导入跨行唯一性

## 目标

按照 Excel 能力差距审查中的 P1 顺序，补齐 ExcelDataReader 固定列分批导入的跨行唯一性。支持单 Sheet `IExcelBatchImporter` 和 Workbook 多 Sheet `IExcelWorkbookBatchImporter`，不改变既有公共签名、批次交付顺序、错误收集和输入流所有权。

## 实施范围

1. 使用现有 `Bing.Offices.Providers.UniqueTracker` 管理每个导入请求的唯一值状态；当前行先暂存，实体成功后提交，转换/校验失败时回滚。
2. `MaxTrackedUniqueValues` 作为固定列批次的可用资源限制，不再在预检阶段拒绝；唯一值比较沿用 `ExcelResourceLimits.UniqueComparison`。
3. Workbook 多 Sheet 批次为每个物理 Sheet 保持独立唯一状态；同一 Sheet 的跨批次重复必须被拒绝，失败批次不能污染后续状态，避免不同 Sheet 的同名属性意外互相冲突。
4. 保留动态列、跨 Sheet 关系、Failure Workbook、Workbook 原生校验和 Entity Layout 的现有 Unsupported 边界。
5. 增加同步/异步、单 Sheet/多 Sheet、重复值、失败回滚、资源上限、取消和双 TFM 职责测试；更新 Provider 文档、符号测试映射和执行记录。

## 验收

- 固定列模型中的 `[ExcelUnique]` 或配置 Unique 规则跨批次生效。
- 同值按 `UniqueComparison` 判重；空值策略保持既有语义。
- 当前行存在其他错误时，已暂存唯一值回滚，后续合法行可继续使用该值。
- `MaxTrackedUniqueValues` 超限返回结构化 `ResourceLimit`，不调用不应交付的成功实体；同步和异步行为一致。
- Workbook 多 Sheet 按物理 Sheet 保持独立唯一状态，错误上下文保留实际 Sheet 和物理行。
- 现有不支持能力仍在预检阶段拒绝，旧批次构造及消费者代码继续编译。
