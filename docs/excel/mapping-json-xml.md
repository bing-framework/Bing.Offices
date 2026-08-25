# JSON/XML Mapping

v2 文档统一使用以下结构，JSON 和 XML 会先归一化为 `ExcelMappingDocument`：

```json
{
  "version": 2,
  "import": {
    "profile": "orders",
    "modelAlias": "order-import",
    "columns": [{ "propertyName": "Code", "title": "订单号" }]
  },
  "export": {
    "profile": "orders",
    "modelAlias": "order-export",
    "columns": [{ "propertyName": "Code", "title": "订单编号" }]
  }
}
```

```csharp
var document = ExcelMappingConfigurationLoader.FromJsonDocument(json);
var importRequest = ExcelImport.Workbook<OrdersWorkbook>(builder =>
    builder.Sheet("订单", workbook => workbook.Items, sheet => sheet.Mapping(document)));
```

XML 使用同名元素：`ExcelMappingDocument/Version/Import/Profile/ModelAlias/Export/Profile/ModelAlias`。Profile 和 alias 只能出现在对应方向节点；根节点不再承载方向业务元数据。v2 主加载器拒绝旧版平铺 v1 配置。

限制包括 UTF-8 文档大小、JSON 深度、字符串、列、别名、校验规则和 XML 字符数。未知 JSON 字段和 XML 节点拒绝并报告路径。XML 禁止 DTD 与外部实体，禁止使用程序集限定 CLR 类型名；`modelAlias` 是稳定业务别名。

加载器不拥有调用方提供的流，读取完成后流保持可用。生产输出应使用 v2，不应把 v1 作为新的写出格式。

方向配置的合并顺序固定为：`Convention < Attribute < Profile < Document < Request`。高优先级配置的 Patch 操作含义如下：

| 状态 | 标量字段 | 集合字段 | DynamicColumns | Style/Layout |
| --- | --- | --- | --- | --- |
| 未设置 | 保留低优先级值 | 保留低优先级值 | 不改变低优先级集合 | 不改变对应字段 |
| 设置值 | 覆盖低优先级值 | 按字段规则合并 | `replace` 替换，`append` 按稳定 `key` 更新或追加 | 按字段覆盖 |
| `clear/reset` | 清除或恢复为未设置 | 清空或移除指定项 | `clearDynamicColumns` 清空，`dynamicColumnKeysToRemove` 按 `key` 移除 | `resetStyle`/`resetLayout` 整体重置；字段级 clear/reset 只影响指定字段 |

例如，方向节点可以表达动态列增量和字段级重置：

```json
{
  "version": 2,
  "import": {
    "dynamicColumnMergeMode": 1,
    "dynamicColumnKeysToRemove": ["legacy"],
    "dynamicColumns": [{ "key": "region", "title": "区域" }],
    "style": { "clearHeaderStyleKey": true },
    "layout": { "resetColumnIndex": true }
  }
}
```

v1 迁移可以保留诊断信息：

```csharp
var document = ExcelMappingConfigurationLoader.MigrateV1Json(
  "{\"columns\":[{\"propertyName\":\"Code\",\"title\":\"编码\"}]}",
  MappingDirection.Import,
  out var diagnostics);
// diagnostics 中包含 V1_MIGRATED，document 已归一化为 Version = 2 的 Import 方向文档。
```
