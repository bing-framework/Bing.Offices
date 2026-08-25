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

Registry 支持显式注册和程序集扫描。重复的 `(ProfileName, Direction, ModelType)` 注册会失败，不由 DI 注册顺序覆盖。推荐显式提供稳定名称，类型 `FullName` 仅作为兼容 fallback：

```csharp
var explicitServices = new ServiceCollection();
explicitServices.AddMappingProfile<OrderProfile>();
var scannedServices = new ServiceCollection();
scannedServices.AddMappingProfiles(typeof(OrderProfile).Assembly);
```

应用程序可以使用 `AddMappingProfile<OrderProfile>("orders")` 显式注册稳定名称；上面的无名称形式仅用于兼容 FullName。

Registry 的唯一键是 `(ProfileName, Direction, ModelType)`。需要读取配置时通过只读 `IMappingProfileResolver.TryGetDescriptor` 指定方向和模型，不能用一个双向 snapshot 代替单方向 descriptor。导入请求只匹配导入 Profile，导出请求只匹配导出 Profile；方向缺失时不会静默回退。
