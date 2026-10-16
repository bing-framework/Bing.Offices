# 源码目录与命名空间

生产代码由 9 个项目组成：`Bing.Offices.Abstractions`、`Bing.Offices.Core`、6 个 Excel Provider/扩展项目和 `Bing.Offices.Analyzers`。项目职责、依赖和公开入口见[项目模块与实现边界](project-modules.md)。

`Bing.Offices.Abstractions` 按能力组织契约源码。目录用于定位实现；公开类型的命名空间由兼容性要求决定，不随文件夹层级变化。

| 目录 | 内容 | 公开命名空间 |
|---|---|---|
| `Entities/Layout` | 实体布局、固定单元格、地址与布局校验 | `Bing.Offices.Entities` |
| `Entities/Lists` | 列表区域及其计算列、动态列、Footer、分组小计 | `Bing.Offices.Entities` |
| `Entities/Anchors` | 逐页名称及过期名称策略 | `Bing.Offices.Entities` |
| `Entities/Import`、`Entities/Export` | 实体布局导入、导出契约 | `Bing.Offices.Entities` |
| `Exports/Workbook`、`Exports/Sheets` | 工作簿与 Sheet 导出请求 | `Bing.Offices.Exports` |
| `Exports/DynamicColumns`、`Exports/Reports`、`Exports/Charts` | 动态列、报表与图表 | `Bing.Offices.Exports` |
| `Exports/Content`、`Exports/Streaming` | 内容与流式导出 | `Bing.Offices.Exports` |
| `Imports/Workbook`、`Imports/Sheets`、`Imports/Batch` | 工作簿、Sheet 与分批导入 | `Bing.Offices.Imports` |
| `Imports/Failures`、`Imports/Validation`、`Imports/Resources` | 失败处理、校验与资源约束 | `Bing.Offices.Imports` |

`IExcelExporter` 和 `IExcelImporter` 留在各自能力根目录。新增契约先按职责选择现有子目录；同一职责有多个类型时，一个非 `partial` 类型对应一个源码文件。确需新目录时，先说明它与现有目录的边界。内部类型可以使用对应能力的子命名空间；调整已公开类型的命名空间或签名时须走公共 API 审批与双 TFM 快照检查。

`Bing.Offices.Npoi.Entities` 包含 NPOI 实体布局内部执行类型。其他 Provider 和 Core 各自保持现有组织；不要根据 Abstractions 的目录直接推导它们的命名空间。

文档测试项目中的目录检查会验证这三个契约目录的归类和公开命名空间。未来如需调整归类，应同步更新该检查、本文和 API 验证证据。

## 其他实现目录

| 目录 | 核心职责 | 定位建议 |
|---|---|---|
| NPOI `Imports` / `Exports` | 导入物化、失败产物、报表写入、透视缓存及预检 | 先按类型职责查找，不按旧文件数定位 |
| NPOI / ClosedXML `Entities` | 固定字段、列表、分页、Forward 和实体批次 | 见[实体布局](excel/entity-layout.md)和[批量导入](excel/batch-import.md) |
| ClosedXML `Internals` | 规划、DOM 准入、行/关系键预检、值和样式适配 | 对照 `ClosedXmlProviderOptions` 查资源配置 |
| Core `Mappings`、`Dates`、`Configurations` | 映射缓存/合并、日期转换及 JSON/XML 加载 | 公开入口和内部实现类型分开使用 |

Core `Csv` 包含 Reader/Writer、异步与取消适配、列绑定和资源约束。`src/provider-shared` 是由 Provider 项目链接编译的共享内部源码，不是第十个生产项目。测试、构建、客户端探针和包消费者见[构建与验证](development/build-validation.md)。目录本身不增加任何对外能力。
