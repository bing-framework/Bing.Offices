# Mapping Profile

`ExcelMappingProfile<TImport, TExport>` 将 Import 和 Export 配置分开保存。两边可以使用不同 DTO，也可以使用同一个 DTO。

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

var profile = new ExcelMappingProfile<OrderImport, OrderExport>(new OrderProfile());
```

配置优先级为 `Convention/Default < Attribute < Profile < JSON/XML Document < Request`。构建后的配置是独立快照，修改输入集合不会改变已构建映射。

Registry 支持显式注册和程序集扫描。重复的 `(ProfileName, Direction, ModelType)` 注册会失败，不由 DI 注册顺序覆盖：

```csharp
services.AddSingleton<OrderProfile>();
services.AddMappingProfile<OrderProfile, OrderImport, OrderExport>("orders");
services.AddMappingProfilesFromAssembly(typeof(OrderProfile).Assembly);
```

旧版 `ExcelMappingProfile<T>` 在当前 major 保留并标记为 Obsolete，迁移期间仍可读取。
