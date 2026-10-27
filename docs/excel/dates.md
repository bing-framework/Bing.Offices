# 日期与 DateTimeOffset

Excel 与 CSV 共用确定性日期解析合同。默认文本格式是 `yyyy-MM-dd`，不使用服务器当前区域性猜测输入；仅年月日转换为 `DateTime` 时，时间为午夜且 `Kind` 为 `Unspecified`。固定 CLR 属性需要其他文本格式或区域性时，在属性上使用 `ExcelDateAttribute`：`Format` 设置精确格式，`CultureName` 设置解析区域性。

```csharp
using System;
using Bing.Offices.Attributes;

public sealed class ImportRow
{
    [ExcelDate("dd/MM/yyyy", CultureName = "fr-FR")]
    public DateTime CreatedAt { get; set; }
}
```

导入 Sheet 的 `.Culture(CultureInfo)` 设置数字和日期转换所用的通用区域性；固定属性上的 `ExcelDateAttribute.CultureName` 优先于该值。Sheet Culture 不设置日期格式或 `DateTimeOffset` offset。当前没有固定属性专用的 Fluent 日期格式/offset 配置入口。

`ExcelDateAttribute.OffsetMinutes` 虽是公开属性，但其类型为 `int?`，C# 不允许在特性声明中为 nullable 属性赋值。因此不能通过 `[ExcelDate(...)]` 为无 offset 的 `DateTimeOffset` 文本指定固定偏移。此类输入需通过已注册的自定义 `INamedExcelValueConverter` 显式解析并赋予所需 offset，再用列映射的 `.HasConverter("名称")` 选择。带显式 offset 的 ISO 文本可按默认日期合同导入。

`ExcelMappingDynamicValidationConfiguration` 的 `date` 规则可通过 `Format` 和 `CultureName` 配置动态列校验，适用于映射文档中的动态列，不等同于固定 CLR 属性的日期配置。导出映射的 `.HasFormatter("格式")` 仅设置导出列格式，不控制导入日期解析。详见[导入校验](import-validation.md)、[映射 Profile](mapping-profile.md)和[动态列](dynamic-columns.md)。

Excel 数值日期、公式缓存日期和文本日期按实际 CellKind 处理。工作簿的 1900/1904 date windowing 会传入解析器；nullable、空白和非法日期仍按列校验与错误收集策略处理。

`DateTimeOffset` 默认以 invariant ISO round-trip (`O`) 文本导出到 Excel 和 CSV，从而保留值与 offset，不依赖本机时区。无 offset 文本不会隐式采用服务器时区；默认策略 `RequireExplicitOffset` 会拒绝缺少 offset 的值。

跨时区部署不改变上述结果。本项目的专项测试分别在 `TZ=UTC` 与 `TZ=Pacific Standard Time` 下执行 XLS、XLSX 和 CSV 往返，并断言相同值和 offset。

`DateOnly` 不属于当前统一公共合同：Abstractions/Core 保持 netstandard2.0，NPOI 发布目标为 net6.0 与 net8.0；当前不通过伪兼容公开 DateOnly。
