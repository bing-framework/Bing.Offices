# 映射 Profile

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

## 核心实现与依赖配置

Abstractions 的 `Configurations` 定义方向契约、Profile Registry 和配置快照；Core 的 `Extensions/MappingProfileServiceCollectionExtensions.cs` 注册显式或扫描得到的 Profile，`Mappings/ExcelMappingPlanFactory.cs` 合并配置并构建最终读写计划。完整路径前缀分别为 `src/Bing.Offices.Abstractions/Bing/Offices` 和 `src/Bing.Offices.Core/Bing/Offices`。

配置文档引用的 `profile` 名称必须与 Registry 中的方向和模型匹配；`modelAlias` 由 `ExcelModelAliasRegistry` 约束。`IExcelValueConverter`、`INamedExcelValueConverter`、`IExcelValidationRule` 和 `INamedExcelValidationRule` 可由宿主注册后进入默认计划工厂。命名转换器仅按显式名称绑定，不能从列标题猜测。

计划工厂默认最多缓存 256 项，以模型、方向、Profile/文档元数据及规范化配置参与键生成。缓存属于工厂实例；动态委托和配置隔离语义由计划实现负责，不能用自制的“只按模型类型”缓存覆盖它。构建失败不会让调用方的原配置变成半合并状态。CSV 与 Excel 均复用该计划，但样式、模板等 Excel 特性不会因此对 CSV 生效。
