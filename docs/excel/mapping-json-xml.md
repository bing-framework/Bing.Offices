# JSON/XML Mapping

v2 文档统一使用以下结构，JSON 和 XML 会先归一化为 `ExcelMappingDocument`：

```json
{
  "version": 2,
  "profile": "orders",
  "modelAlias": "order-row",
  "import": { "columns": [{ "propertyName": "Code", "title": "订单号" }] },
  "export": { "columns": [{ "propertyName": "Code", "title": "订单编号" }] }
}
```

```csharp
var document = ExcelMappingConfigurationLoader.FromJsonDocument(json);
var importRequest = ExcelImport.Workbook<OrdersWorkbook>(builder =>
    builder.Sheet("订单", workbook => workbook.Items, sheet => sheet.Mapping(document)));
```

XML 使用同名元素：`ExcelMappingDocument/Version/Profile/ModelAlias/Import/Export`。旧版平铺 `columns` 或缺省 `version` 的 v1 配置仍可读取，旧 Loader 方法返回 v2 文档的 Import 配置。

限制包括 UTF-8 文档大小、JSON 深度、字符串、列、别名、校验规则和 XML 字符数。未知 JSON 字段和 XML 节点拒绝并报告路径。XML 禁止 DTD 与外部实体，禁止使用程序集限定 CLR 类型名；`modelAlias` 是稳定业务别名。

加载器不拥有调用方提供的流，读取完成后流保持可用。生产输出应使用 v2，不应把 v1 作为新的写出格式。

v1 迁移可以保留诊断信息：

```csharp
var document = ExcelMappingConfigurationLoader.FromJsonDocument(
  "{\"columns\":[{\"propertyName\":\"Code\",\"title\":\"编码\"}]}",
  out var diagnostics);
// diagnostics 中包含 V1_MIGRATED，document 已归一化为 Version = 2。
```
