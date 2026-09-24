# `src/*` 中文 XML 注释治理计划

## Summary

目标文件：`ai_docs/plans/chinese-comments-plan.md`。

治理仅覆盖 `src/**`，只允许修改 XML 文档注释，不改变业务逻辑、签名、可见性、API Snapshot 或运行时行为。源码修改继续保持仓库现状的 UTF-8 BOM、LF；计划文件使用 UTF-8 无 BOM、LF。

## 批次状态

| 批次 | 状态 | 说明 |
|---|---|---|
| CC-100 | COMPLETED | 上游接口、抽象类和公共契约已完成并通过复核。 |
| CC-200 | COMPLETED | DTO、Entity、Options 和枚举已完成并通过复核。 |
| CC-300 | COMPLETED | 构造函数和属性文案已完成并通过复核。 |
| CC-400 | COMPLETED | 实现类、重写成员和重载方法已完成并通过复核。 |
| CC-500 | COMPLETED | private、internal、static 辅助成员及 MiniExcel 标签已完成并通过复核。 |
| CC-600 | COMPLETED | 字段、常量、缓存键和配置键已完成并通过复核。 |
| CC-700 | COMPLETED | 本轮完成声明级审计、编码检查、方案构建和差异检查。 |

当前 `OPEN_ACTIONABLE=0`，下一步为 `STOP`。后续如需新增注释范围，应另行建立计划或明确新的批次。

## 扫描范围与统计

排除 `bin`、`obj`、`*.g.cs`、`*.generated.cs`、`*.Designer.cs`、Migration、ModelSnapshot、自动生成客户端、代理代码和第三方源码。

| 区域 | 文件数 | 声明数 |
|---|---:|---:|
| Bing.Offices.Abstractions | 91 | 1,156 |
| Bing.Offices.Core | 59 | 668 |
| Bing.Offices.Npoi | 47 | 703 |
| Bing.Offices.MiniExcel | 11 | 151 |
| Bing.Offices.ClosedXml | 14 | 253 |
| provider-shared | 1 | 18 |
| 合计 | 223 | 2,949 |

声明构成：322 个类型、38 个枚举、196 个枚举成员、10 个委托、143 个构造函数、1,112 个方法、824 个属性、250 个字段、54 个常量。当前所有声明均已有三行 `<summary>` 或 `<inheritdoc />`，缺失数为 0；现有摘要 2,694 处，`inheritdoc` 255 处。

已确认问题：

- 25 个构造函数不符合统一初始化文案，分布在 11 个文件，集中于 Entity、ClosedXML、NPOI Entity 和共享预检。
- MiniExcel 4 个文件中，17 个声明缺少合计 62 个 `<param>`，4 个泛型声明缺少 `<typeparam>`，10 个非空返回方法缺少 `<returns>`。
- 84 个摘要未使用规范结束标点，其中 48 个集中在 `Core/Styles/Color.cs`，其余集中于旧 Core Attribute 和 NPOI Extension。
- 8 个摘要属于冗长候选；重点复核动态列支持类型、估算范围、模板/工作簿策略是否应下沉到 `<remarks>`。
- 116 个摘要属于极短文案候选，其中 99 个是含义明确的枚举成员；枚举名称准确时不机械扩写，只统一必要标点。
- 33 组重载使用不同摘要。多数差异用于区分文本/流、配置来源或参数能力，需逐组确认后仅统一真正同语义的重载。
- 6 个属性被机械规则标记，但均为 `internal get / private set`，现有“获取……”符合外部可见访问契约，不修改。
- 255 个 `inheritdoc` 未发现确定性误用；未确认 override 或显式接口实现遗漏。私有 `Create` 与接口同名属于误报，不得改为 `inheritdoc`。
- 304 个字段和常量均已有注释，缓存键和配置键没有确定性缺失，后续只做语义质量复核。

## 分批实施

### CC-100 接口、抽象类和公共契约

检查 Abstractions 下的 `Configurations`、`Csv`、`Conversions`、`Entities`、`Exports`、`Imports`、`IO`、`Providers`、`Styles`、`Validations`、`Exceptions`，以及 Attribute 抽象基类。

先完善上游契约的参数、返回值、空值、单位和所有权语义。公共摘要只保留核心用途；资源限制、流所有权、兼容性等扩展信息移入 `<remarks>`。本批完成后，下游实现才能决定是否使用 `inheritdoc`。

### CC-200 DTO、Entity、Options 和枚举

处理 Abstractions 的请求、结果、配置、样式、Entity、错误类型，以及 `ClosedXmlProviderOptions`、Core Metadata 和枚举。

统一属性访问文案；布尔属性仅在 `true/false` 含义、reset/clear 或默认行为不明显时补充解释。枚举成员保留简短领域名称，不为满足字数机械扩写。重点复核 `ExcelEntityLayout.cs`、`ExcelEntityResults.cs`、`ExcelDynamicColumnDefinition.cs`、`ExcelImportPolicies.cs` 和 `Color.cs`。

### CC-300 构造函数和属性文案

将 25 个构造函数统一为：

```csharp
/// <summary>
/// 初始化一个 <see cref="当前类型"/> 类型的实例。
/// </summary>
```

原摘要中的“已验证快照”“零基坐标”“指定准入器”等有效差异移入 `<remarks>`，不得丢失契约信息。主要文件为 Entity 两个文件、ClosedXML 的 Entity/Importer/Exporter/Plan/Admission/Preflight、NPOI Entity Executor 和共享 ZIP 预检。

复核全部属性的访问器可见性；`private set` 属性按调用方可见契约使用“获取……”，`init` 使用“获取或初始化……”。

### CC-400 实现类、重写成员和重载方法

覆盖 Core、NPOI、MiniExcel、ClosedXML 的实现类。接口实现、显式接口实现、抽象成员实现和 override 在上游契约有效时改用 `/// <inheritdoc />`；仅通过同名方法推测继承关系是禁止的。

逐组复核 33 个重载候选。输入介质或能力确有差异时保留差异摘要；语义一致时统一核心用途，并把流所有权、模板策略等差异放入参数或 remarks。

### CC-500 private、internal 和 static 辅助方法

优先修复以下 MiniExcel 文件的确定性标签缺口：

- `MiniExcelExcelImporter.cs`
- `MiniExcelRelationCoordinator.cs`
- `MiniExcelRowMaterializer.cs`
- `MiniExcelSheetPlanBuilder.cs`

补齐 62 个 `<param>`、4 个 `<typeparam>` 和 10 个 `<returns>`，其中 `bool` 返回值明确真假含义，可空返回值明确 `null` 条件。同步改善 Core 的旧 Attribute、`TypeExtensions`、`PropertyInfoExtensions`、`InternalHelper` 和 `MergedRegionInfo` 中过短或机械化参数说明。

### CC-600 字段、常量、缓存键和配置键

复核现有 250 个字段和 54 个常量，不新增无意义重复注释。重点检查 Core/NPOI `InternalConst`、Mapping Cache Key、NPOI Style Cache、ClosedXML Admission 状态、Provider 配置默认值和共享预检限制。

注释应说明用途、作用域、默认值或生命周期；缓存键需说明区分维度，布尔字段说明真假含义。确认无问题的声明保持不动。

### CC-700 最终审计和构建验证

重新执行声明级 Roslyn 扫描，确认不存在缺失注释、构造文案偏差、标签缺口和错误 `inheritdoc`。人工复核 33 组重载、84 个标点候选、8 个冗长候选以及所有实际修改文件。

只运行注释变更所需的 L0 验证；L1-L5 测试不适用，除非最终 diff 意外包含业务代码。

## 验证命令

```powershell
git status --short
git diff --check
git diff --stat
git diff --unified=0 -- 'src/**/*.cs'
git ls-files --eol -- 'src/**/*.cs' 'ai_docs/plans/chinese-comments-plan.md'
dotnet build .\Bing.Offices.sln -c Release --no-restore --no-incremental
```

额外验证所有修改后的 `.cs` 均可严格 UTF-8 解码、包含 BOM、仅使用 LF、无 Mixed EOL 且保留末尾换行；生成的 `Bing.Offices*.xml` 文档必须可解析。最终 diff 中 C# 变化只能位于 `///` 文档行。

## 风险、待确认项与验收标准

- 工作区存在大量未提交改动；每批开始前重新读取目标文件，不回退、覆盖或格式化无关注释和代码。
- `inheritdoc` 必须依据真实接口或基类契约，不能依据名称匹配。
- `<see cref>`、泛型引用和 XML 特殊字符必须合法；编译不得新增 XML 文档警告。
- 极短枚举说明可能已经准确，不以长度作为修改依据。
- 当前没有阻塞性待确认项；无法从源码和测试确定的语义使用客观描述，不编造规则。

验收时要求：2,949 个声明持续全部有文档；25 个构造函数完成统一；MiniExcel 标签缺口归零；确认的文案和 inheritdoc 问题归零；所有 C# diff 仅含注释；编码和行尾保持不变；全方案 Release 构建 0 error，且不新增 warning。
