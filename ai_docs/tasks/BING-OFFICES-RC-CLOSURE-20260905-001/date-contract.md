# 日期合同

| 来源/目标 | 当前证据 | 冻结目标 | 状态 |
| --- | --- | --- | --- |
| 文本 -> DateTime | 共享 `ExcelDateParser` 默认 exact `yyyy-MM-dd` | 午夜、Unspecified、culture-independent | VERIFIED |
| 显式格式 -> DateTime | Attribute/Mapping formats | 只使用显式 exact formats | VERIFIED |
| Excel numeric/formula cache | `ExcelCellValue` 含 CellKind/date1904 | 正确传递 1900/1904 | VERIFIED |
| 文本 explicit offset -> DateTimeOffset | parser 保留 offset | 值与 offset 保留 | 基线测试存在 |
| 无 offset文本 -> DateTimeOffset | 仅固定 offset policy 允许 | 禁止 Local 隐式策略 | 基线测试存在 |
| DateTime -> DateTimeOffset | 无固定 offset 时显式失败 | 必须有明确固定 offset/时区策略 | VERIFIED |
| DateTimeOffset 导出 Excel | XLSX/XLS 默认写 invariant `O` 文本 | 默认 ISO offset 字符串 | VERIFIED，双 TZ 两阶段往返 |
| DateTimeOffset 导出 CSV | invariant `O` 文本并可再导入 | 明确 ISO round-trip 格式 | VERIFIED，双 TZ 往返 |
| DateOnly | netstandard2.0 不统一支持 | 本轮不公开，文档明确 | DONE（范围决策） |

验收必须含 nullable、blank、invalid、leap year、boundary、1900/1904、formula cache，并至少在两个不同 `TZ` 环境下验证 Excel/CSV 导入→导出→再导入。

## 跨时区证据

2026-09-06 在 `TZ=UTC` 与 `TZ=Pacific Standard Time` 两个独立 testhost 环境运行 DateTimeOffset 专项，各 6/6。矩阵包含 XLS/XLSX 导入→导出→关闭重开→再导入、单元格字符串类型、CSV 往返和 fixed-offset parser。TRX：`artifacts/test-results/timezone/reviewfix-datetimeoffset-utc.trx`、`reviewfix-datetimeoffset-pacific.trx`。
