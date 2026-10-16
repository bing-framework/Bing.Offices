# 日期与 DateTimeOffset

## 功能与解析逻辑

Excel 与 CSV 共用确定性日期解析合同。默认文本格式是 `yyyy-MM-dd`，不根据服务器当前区域性猜测；只有年月日的输入转换为 `DateTime` 时取午夜，`Kind` 为 `Unspecified`。

Excel 数值日期、公式缓存日期和文本日期按实际单元格种类处理，并传入工作簿的 1900/1904 日期系统。MiniExcel 适配先读取 `xl/workbook.xml` 中 `workbookPr/@date1904`，再保留工作表原始数字序列值交由 Core 处理边界与负数语义；不能取得序列值时才走文本或原生日期路径。空值和非法值仍服从列校验及错误收集策略。

`DateTimeOffset` 默认以 invariant ISO 往返格式 `O` 导出到 Excel 和 CSV，保留值与偏移，不依赖本机时区。带显式偏移的 ISO 文本可直接导入；缺少偏移的文本或 `DateTime` 不会隐式采用服务器时区。

## 用法与配置

在模型日期属性上使用 `Bing.Offices.Attributes.ExcelDateAttribute`，例如 `[ExcelDate("yyyy-MM-dd")] public DateTime Date { get; set; }`。其他精确格式或区域性可通过该特性的 `Format`、`CultureName`，或方向 Profile、JSON/XML 映射显式提供，见[映射文档](mapping-json-xml.md)。

若业务允许把无偏移输入解释为固定偏移，配置 `ExcelDateOffsetPolicy.UseFixedOffset` 并设置 `OffsetMinutes`；默认 `RequireExplicitOffset` 拒绝缺少偏移的值。`OffsetMinutes` 是可空配置属性，不应直接写成 C# 特性的命名实参；通过 Fluent/配置对象或映射文档设置。

## 核心实现与依赖

- 解析器及偏移策略：`src/Bing.Offices.Core/Bing/Offices/Dates`。
- 特性：`src/Bing.Offices.Core/Bing/Offices/Attributes/Filters/ExcelDateAttribute.cs`。
- 列配置、转换绑定与计划：Core 和 Abstractions 的 `Bing/Offices/Configurations`、`Bing/Offices/Mappings`。
- MiniExcel 日期系统和原始值适配：`src/Bing.Offices.MiniExcel/Bing/Offices/Internals`。

依赖 Core 的映射计划与所选 CSV/Excel Provider，没有独立时区数据库服务。统一解析测试及各 Provider 日期测试位于 `tests` 对应项目；历史跨时区测试证据不能代替当前工作区重跑结果。

## 限制与注意事项

跨时区部署不应改变显式日期格式与偏移策略；业务时区换算由调用方完成。格式不匹配不会回退到当前文化的宽松猜测。1900/1904 数值不是 Unix 时间戳。

`DateOnly` 统一公共合同为**待实现/待确认（待实现）**：Abstractions/Core 仍以 `netstandard2.0` 为目标，不能由 `DateTime` 支持推断公开支持 `DateOnly`。
