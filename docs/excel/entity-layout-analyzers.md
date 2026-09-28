# Entity Layout 编译期诊断

`Bing.Offices.Analyzers` 是可选的 Roslyn 分析器包。它检查源码中能确定的布局常量；工作簿模板、变量地址和运行时生成的动态列仍由 Entity Layout 构建与 Provider 预检验证。

本仓库可先运行 `dotnet pack src/Bing.Offices.Analyzers/Bing.Offices.Analyzers.csproj -c Release` 生成本地包。第三方项目把该包加入自己的 NuGet 源后，按所生成的包版本引用：

```xml
<PackageReference Include="Bing.Offices.Analyzers" Version="2.0.0" PrivateAssets="all" />
```

分析器只参与编译，不向业务程序添加运行时依赖。当前规则如下：

| 规则 | 级别 | 检查范围 |
| --- | --- | --- |
| `BOE001` | Error | `[ExcelEntityCell]` 和 Fluent `Cell` 的常量地址不是有效的 XLSX A1 单元格地址 |
| `BOE002` | Error | 同一类型的属性特性或同一 Fluent 调用链在同一 Sheet 中使用重复固定地址 |
| `BOE003` | Info | `[ExcelEntityCell]` 标注无 setter 属性；属性可导出，导入不能写回 |
| `BOE004` | Error | 同一 Fluent 调用链中，内联 `DynamicColumnGroup` 的组标识、动态列 Key、标题或别名重复 |
| `BOE005` | Error | 同一内联布局中的合并区域、显式有界列表区域及固定单元格发生可确定的占位冲突；属性式和 Fluent 固定单元格同址、多个固定单元格落入同一合并锚点也在此检查 |
| `BOM001` | Warning | 已绑定到旧版 `Bing.Offices.Exports.IExcelExportService` 或 `ExportOptions<T>` 的类型引用需要人工迁移 |
| `BOM002` | Warning | 已绑定到旧版 `Bing.Offices.Attributes.HasDynamicColumnAttribute` 的标记需要人工迁移 |

比较地址时忽略大小写和 `$`；动态名称忽略大小写并修剪首尾空白。`BOE005` 只在 `Layout` 或 `LayoutFromAttributes` 的单条 Fluent 链中比较常量位置，列表区域需要内联常量 `End` 才能形成确定的边界；合并区域左上角的单个固定单元格允许使用。分析器不追踪变量、工厂返回值、跨语句构建器状态或模板命名锚点解析。`BOE003` 为提示级别，因为只读属性用于纯导出是合法的。运行时校验继续负责未知配置、模板位置、动态列表范围和 Footer 冲突。

命名参数可以调整 `Cell`、`Merge`、`DynamicColumnGroup` 的实参顺序，分析器按参数名称识别其含义。目标类型推断的内联 `new()` 定义可参与检查；条件表达式和其他无法静态确定的定义不会被误判。仓库中的 `Bing.Offices.Analyzers.Consumer` 从本地生成的 NuGet 包验证正常编译，并通过 `BOE_NEGATIVE`、`BOE_AREA_NEGATIVE` 编译符号分别验证 `BOE001`、`BOE005` 会阻止无效布局。

`BOM001` 和 `BOM002` 面向仍引用 1.x 程序集、能够解析旧类型的项目。两个规则只报告已绑定的完整类型，不对未解析名称或业务项目中的同名类型猜测；不提供自动代码修复。`ColumnNameAttribute` 和 `DynamicColumnAttribute` 在 2.x 仍可使用，单独出现不会触发迁移提示。迁移路径见[NuGet 迁移文档](nuget-migration.md)。
