# Mapping Profile

Mapping Profile 直接表达方向：`IImportMappingProfile<TImport>`、`IExportMappingProfile<TExport>`、同模型 `IMappingProfile<TModel>` 或异模型 `IMappingProfile<TImport, TExport>`。双向 Profile 的两个方向仍使用独立配置快照。

```csharp
public sealed class OrderProfile : IMappingProfile<OrderImport, OrderExport>
{
    public void Configure(FluentSetting<OrderImport, OrderExport> setting)
    {
        setting.Import.Property(order => order.Code)
            .HasHeader("订单号")
            .HasAlias("旧订单号");
        setting.Export.Property(order => order.DisplayName)
            .HasHeader("客户名称")
            .HasFormatter("@");
    }
}

```

配置优先级为 `Convention/Default < Attribute < Profile < JSON/XML Document < Request`。构建后的配置是独立快照，修改输入集合不会改变已构建映射。

Registry 支持显式注册和程序集扫描。重复的 `(ProfileName, Direction, ModelType)` 注册会失败，不由 DI 注册顺序覆盖：

```csharp
var explicitServices = new ServiceCollection();
explicitServices.AddMappingProfile<OrderProfile>();
var scannedServices = new ServiceCollection();
scannedServices.AddMappingProfiles(typeof(OrderProfile).Assembly);
```

Registry 的唯一键是 `(ProfileName, Direction, ModelType)`。Profile 名称默认使用具体类型的 `FullName`；需要读取配置时通过 `TryGetDescriptor` 指定方向和模型，不能用一个双向 snapshot 代替单方向 descriptor。
