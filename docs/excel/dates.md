# 日期与 DateTimeOffset

Excel 与 CSV 共用确定性日期解析合同。默认文本格式是 `yyyy-MM-dd`，不使用服务器当前区域性猜测输入；仅年月日转换为 `DateTime` 时，时间为午夜且 `Kind` 为 `Unspecified`。需要其他格式或区域性时，通过 `ExcelDateAttribute`、Fluent 或 Mapping 文档显式配置精确格式和 Culture。

Excel 数值日期、公式缓存日期和文本日期按实际 CellKind 处理。工作簿的 1900/1904 date windowing 会传入解析器；nullable、空白和非法日期仍按列校验与错误收集策略处理。

`DateTimeOffset` 默认以 invariant ISO round-trip (`O`) 文本导出到 Excel 和 CSV，从而保留值与 offset，不依赖本机时区。带显式 offset 的 ISO 文本可直接导入。无 offset 文本或 `DateTime` 不能隐式转换为 `DateTimeOffset`；调用方必须选择 `ExcelDateOffsetPolicy.UseFixedOffset` 并设置 `OffsetMinutes`。默认策略 `RequireExplicitOffset` 会拒绝缺少 offset 的值。

跨时区部署不改变上述结果。本项目的专项测试分别在 `TZ=UTC` 与 `TZ=Pacific Standard Time` 下执行 XLS、XLSX 和 CSV 往返，并断言相同值和 offset。

`DateOnly` 不属于当前统一公共合同：Abstractions/Core 保持 netstandard2.0，NPOI 发布目标为 net6.0 与 net8.0；当前不通过伪兼容公开 DateOnly。
