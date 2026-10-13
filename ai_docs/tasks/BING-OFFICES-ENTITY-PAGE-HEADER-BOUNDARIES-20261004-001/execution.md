# Entity 重复页眉边界执行记录

Task ID：`BING-OFFICES-ENTITY-PAGE-HEADER-BOUNDARIES-20261004-001`

## 完成情况

- DESIGN-001：页眉 marker 文本作为标记列保留值。NPOI 和 ClosedXML 在明细完成转换、动态列与计算列写入后检查实际单元格；冲突返回原始坐标，且不改写调用方流。无列表表头时，最终物理明细列不得包含 marker 列。
- DESIGN-002：实际页眉的 Fluent `Cell` 可写入相同标题并调整样式，但不能改写最终映射标题；`Merge` 不得触及列表表头。首页和后续页均检查。导入时核对复制页表头，篡改后返回原始坐标的结构化配置错误。
- NPOI XLS/HSSF、XLSX/XSSF 与 ClosedXML XLSX 使用相同边界。未新增公开成员，旧 Fluent 签名不变。

## 验证

- NPOI 重复页眉职责测试：net6.0、net8.0 各 66 项通过，覆盖两种格式的 marker 冲突、Fluent 表头与合并、复制页表头篡改、无表头 marker 列冲突和合法往返。
- ClosedXML 重复页眉职责测试：net6.0、net8.0 各 42 项通过，覆盖固定、动态、计算列写出值以及导入篡改坐标。
- `dotnet test Bing.Offices.sln -c Release --no-restore -m:1 --filter Category!=Large` 通过；该次全量运行在最后新增的两项 ClosedXML 动态/计算列测试之前，新增测试随后按双 TFM 定向通过。
- `dotnet build Bing.Offices.sln -c Release --no-restore -m:1` 为 0 错误。串行 `dotnet pack` 成功；使用本任务本地包的 net6.0、net8.0 消费者均输出 `package-consumer-ok`。
- 双 TFM API 成员哈希和成员数与审批基线相同。只更新候选源码清单及 NPOI、ClosedXML 包身份；正式 API snapshot 比较通过。
- 最终 Release 解决方案构建 0 错误；`git diff --check` 无输出。本任务涉及的 12 个文本文件严格 UTF-8、BOM、LF 和末尾换行检查通过。保留既有 net6.0 依赖支持提示。

## 剩余 TODO

- DESIGN-003：明确富文本、超链接、批注、校验、条件格式、隐藏行和锁定状态的复制或拒绝契约。
- DESIGN-004：用稀疏大表和大 `End` 测量扫描成本及取消响应，再决定是否引入按块取消。
- DESIGN-005：提取复用的 Provider 页眉合同测试，同时保留引擎特有的真实文件断言。
- DESIGN-006：公共透视表、公式计算与真实打印渲染单独设计。

保留现有未提交工作。未修改版本号，未 commit、push 或发布。
