# 基于 Utopa 业务场景增强实体布局能力

## 目标与范围

读取 `D:\MyWorking\Utopa\Utopa.Erp.IE` 的真实使用方式，增强 `Bing.Offices` 的 Entity Layout 公共能力。外部 Utopa 仓库只作只读样本，不修改、不复制其业务代码或模板；不自动提交、推送、发布或修改版本号。

外部样本集中体现普通列表、动态商品字段、固定单元格单据头、可变明细、合计尾部、多 Sheet 汇总和模板校验等场景。当前基础库已经具备固定单元格、列表区域、合并、关系、模板、动态列和结构化校验，但 NPOI Entity 动态列仍拒绝，且缺少多动态字典和可变明细尾部抽象。

## P1 公共契约

1. 在 `Bing.Offices.Entities` 增加 `ExcelEntityCellAttribute`，构造函数接收工作表名和 A1 地址，提供可选 `ConverterName`。增加 `ExcelEntity.LayoutFromAttributes<TEntity>` 和 `ExcelEntityLayoutBuilder<TEntity>.CellsFromAttributes()`；按属性名稳定扫描公开实例属性，可与现有 Fluent Cell/Merge/ListRegion 混用，保留既有校验和 Mapping Plan 语义。
2. 在 `ExcelEntityListRegionBuilder<TItem>` 增加可重复的 `DynamicColumnGroup` 和 `UnknownDynamicValues`。各组的 groupKey、动态 Key、标题及别名全局唯一；导出按已有布局规则合并，导入写回对应字典；未配置分组时保持现有单字典兼容行为。显式分组与 Mapping 隐式动态列冲突时在 Plan 阶段结构化失败。
3. 增加 `Footer(string markerText, Action<ExcelEntityListFooterBuilder<TItem>>)` 及尾部构建器。支持 `GapRows`、常量 Cell、基于明细快照的聚合 Cell、相对 Merge 和 MarkerStyle。Footer 从明细结束后的首行开始，marker 位于相对 A1；导入按规范化 marker 截断明细，不把 Footer 当作数据。
4. 新增 Provider 需要读取的描述对象时采用公开、`EditorBrowsable(Never)` 的 SPI 类型，不能通过生产程序集 `InternalsVisibleTo` 绕过边界。所有既有 API、构造函数和单字典调用保持源码及二进制兼容。

## P2 Provider

1. NPOI 移除 Entity Layout 动态列拒绝分支，复用现有 Mapping Plan、动态列规划、转换、校验和错误包装；实现单字典、多分组动态列、属性固定单元格、Footer、相对合并及 marker 导入边界；XLS/HSSF 与 XLSX/XSSF 均覆盖。
2. ClosedXML 将已有单字典实体动态列接入分组描述，并实现同语义 Footer、marker、预检和错误分类，保留 DOM 准入、取消、模板和文件提交边界。
3. MiniExcel、SpreadCheetah、ExcelDataReader 不增加 Entity Layout 伪实现，继续通过能力测试明确拒绝或不注册该接口。

## P3 测试

增加与 Utopa 等价但独立的测试模型和夹具：商品档案（三个动态字典）、采购单（属性头、模板、明细、合计）、赠品单（无表头明细和空明细）、盘点单（多 Sheet/多区域）、导入校验与结构化错误。覆盖动态 Key 冲突、物理列冲突、Footer 重叠/marker 缺失、行列边界、转换器异常、预取消/中途取消、旧目标保护、临时文件清理、流所有权、同步异步一致性及 NPOI XLS/XLSX、ClosedXML XLSX。

保留并验证旧 Fluent Layout 和单动态字典用法。新增 net6.0/net8.0 包消费者示例，维护生产符号到测试方法的追溯表。

## P4 文档与交付

- 补齐新增公共成员中文 XML 注释，新增约束放入 `remarks`。
- 更新 Entity Layout、动态列、Provider 矩阵和迁移文档，补充旧式 NPOI 单据迁移示例，修正 ClosedXML 动态列能力说明。
- 生成 net6.0/net8.0 API snapshot 和差异，确认无公共成员删除或未批准变更。
- 维护 `execution.md`、API approval、双 TFM 快照和生产符号追溯；不自动 commit/push/PR。

## 验证与验收

按“职责级测试 → 受影响 Provider 测试 → 集成/Consumer → API 检查 → Release 构建与 `Category!=Large` 全量检查 → 编码/diff 检查”执行。无关 500K/1M 性能矩阵不重跑；仅在共享热路径出现实际性能问题时增加定向资源探针。

验收条件：属性式固定单元格、动态列分组和 Footer 在 NPOI/ClosedXML 公开 Entity API 中可用；NPOI XLS/XLSX 与 ClosedXML XLSX 的业务合同通过；旧 API 无删除；net6.0/net8.0 包消费者成功；生产友元仍仅为测试程序集。

## 后续建议

模板命名锚点、行号/计算列 DSL、分页小计和签字区、可复用 Entity Provider 合同测试包、Roslyn 地址冲突分析器、Source Generator 预编译布局和旧版 API 迁移分析器另立任务；不在本次实现中扩展公式引擎或隐式 Provider 切换。
