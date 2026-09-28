# 源码目录与命名空间

`Bing.Offices.Abstractions` 按能力组织契约源码。目录用于定位实现；公开类型的命名空间由兼容性要求决定，不随文件夹层级变化。

| 目录 | 内容 | 公开命名空间 |
|---|---|---|
| `Entities/Layout` | 实体布局、固定单元格、地址与布局校验 | `Bing.Offices.Entities` |
| `Entities/Lists` | 列表区域及其计算列、动态列、Footer、分组小计 | `Bing.Offices.Entities` |
| `Entities/Import`、`Entities/Export` | 实体布局导入、导出契约 | `Bing.Offices.Entities` |
| `Exports/Workbook`、`Exports/Sheets` | 工作簿与 Sheet 导出请求 | `Bing.Offices.Exports` |
| `Exports/DynamicColumns`、`Exports/Reports`、`Exports/Charts` | 动态列、报表与图表 | `Bing.Offices.Exports` |
| `Exports/Content`、`Exports/Streaming` | 内容与流式导出 | `Bing.Offices.Exports` |
| `Imports/Workbook`、`Imports/Sheets`、`Imports/Batch` | 工作簿、Sheet 与分批导入 | `Bing.Offices.Imports` |
| `Imports/Failures`、`Imports/Validation`、`Imports/Resources` | 失败处理、校验与资源约束 | `Bing.Offices.Imports` |

`IExcelExporter` 和 `IExcelImporter` 留在各自能力根目录。新增契约先按职责选择现有子目录；同一职责有多个类型时，一个非 `partial` 类型对应一个源码文件。确需新目录时，先说明它与现有目录的边界。内部类型可以使用对应能力的子命名空间；调整已公开类型的命名空间或签名时须走公共 API 审批与双 TFM 快照检查。

`Bing.Offices.Npoi.Entities` 包含 NPOI 实体布局内部执行类型。其他 Provider 和 Core 各自保持现有组织；不要根据 Abstractions 的目录直接推导它们的命名空间。

文档测试项目中的目录检查会验证这三个契约目录的归类和公开命名空间。未来如需调整归类，应同步更新该检查、本文和 API 验证证据。

## 下一阶段盘点

| 候选范围 | 现状 | 建议 |
|---|---|---|
| NPOI `Imports` | 25 个文件，混有导入执行、失败工作簿与预检校验 | 优先按职责评估目录拆分和内部引用影响 |
| NPOI `Exports` | 9 个文件，包含规划、写入、报表与样式缓存 | 与导入目录一同评估，保持公开契约兼容 |
| ClosedXML `Internals` | 11 个文件，覆盖规划、准入、预检与适配 | 评估是否按内部职责分目录 |
| Core `Mappings` | 17 个文件，仍属于同一映射域 | 观察维护成本后再决定是否拆分 |

Core `Csv` 的 13 个文件目前构成完整的 CSV 能力目录。其余 Provider 的源码较少，暂不因文件数量单独拆分。上述候选范围只是下一阶段的盘点，不改变当前实现。
