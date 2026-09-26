# Bing.Offices 开源 Provider 业务能力补齐计划

## Task

`BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001`

## 目标

按用户本轮批准范围，优先补齐开源 Provider 业务能力；Aspose、公式重新计算、高级 Custom 校验、透视表延期，不修改其实现或宣称验收。

## 阶段

1. F01：SpreadCheetah 复用 Core 映射计划、ValueMap、命名 Converter；固定/动态列顺序、空值和转换异常一致。静态配置在输出前检查，不预枚举数据；绑定/键集合每请求复用。默认批次 1000，串行背压，数据只枚举一次。
2. F02：Sheet 新增 Images/DataValidations 和 Builder。PNG/JPEG 字节内容、零基锚点、像素尺寸/偏移，逐项写入，不使用 System.Drawing 编码。整数/小数/日期/文本长度八运算符及显式列表，空值与输入/错误提示。SpreadCheetah/NPOI/ClosedXML 实现，MiniExcel 预检拒绝。复杂公式/名称列表/跨 Sheet/外部引用延期。
3. F03：NPOI XLSX 补双色最小最大色阶、最小最大数据条、三交通灯图标集，与 ClosedXML 等价；XLS 无法表达时预检 Unsupported。
4. F04：无效样式、重复名称、格式范围越界、冲突筛选提前拒绝；同范围 Table/AutoFilter 合并；名称 Workbook/Sheet 作用域；ShowTotals 只显示汇总行；无法表达的冻结可视起点拒绝；打印比例/页数冲突拒绝。静态预检先于枚举/写入。
5. F05：独立 Provider/操作/格式/场景 Profile；图片关系锚点/校验 XML/重新打开结构断言；流式 0/1/999/1000/1001、背压、单次枚举、真实中途取消、IO/枚举异常、临时文件/旧目标/所有权；consumer 实际执行 SpreadCheetah。
6. F06：固定 SDK 摘要、net6 runtime、字体来源；非 root、只读输入、独立临时目录、显式内存限制。Docker 运行受影响测试及缺字体负向场景；最终 100K/500K/1M 记录耗时、峰值内存、文件大小和批次影响，不自行批准阈值。

## 公共契约约束

- 保留现有 IExcelImporter/IExcelExporter 和既有枚举值。
- 新能力使用独立接口，不通过运行时自动换 Provider 或静默降级。
- Abstractions 继续兼容 netstandard2.0；流式导出使用串行批次 Sink/回调，不引入 IAsyncEnumerable。
- 新格式 Unsupported 必须预检失败；文件目标继续复用原子提交器。
- 所有新增公共成员更新双 TFM API snapshot、API 审批、package consumer 和生产符号追溯。

## 验证

- 逐阶段执行 L0/L1/L2/L3；Phase 收口执行受影响项目全量测试。
- 流式导出覆盖 0/1/边界/100K/500K/1M、背压、取消、写入失败和清理。
- 本轮图片、校验、报表和映射使用职责级直接测试与独立结构检查；NPOI XLS/XLSX 分别覆盖成功和拒绝。
- 验证顺序：定向 → 受影响双 TFM → API/package consumer → 全量构建/测试 → Docker/容量；最后严格 UTF-8/BOM/EOL 和 git diff --check。
- F01–F06 均需实现和直接证据；独立 review 仅 OPEN_ACTIONABLE > 0 进入修复。通用 Stream 失败允许部分输出，路径复用原子提交器。
- 最终执行双 TFM 构建、全量测试、API/package consumer、Docker/Linux、编码和 diff 检查。
- 不自动 commit、push、发布、升级既有依赖或批准性能/容量阈值。

## 明确限制

- Aspose.Cells 为独立可选商业包，Core 不引用；缺少许可证时预检失败。
- XLSM 只保留宏，不执行或编辑 VBA；默认宏策略为 Reject。
- ODS 与 XLSX 转换提供损失报告，不承诺完全无损。
- 透视表、Office Automation、宏执行、签名修改和 VBA 编辑不在本任务范围。
