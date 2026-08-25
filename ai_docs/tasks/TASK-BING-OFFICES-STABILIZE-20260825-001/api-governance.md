# API 治理记录

## 当前决策

| API/区域 | 当前状态 | 本任务决策 |
| --- | --- | --- |
| Workbook Request | 公开推荐主路径 | 保留，继续作为导入/导出主契约 |
| `ExcelTemplateCellOverwritePolicy` | 本轮新增 | 公开枚举和 Builder 方法，默认保持兼容行为 |
| `ExcelMappingConfigurationLoader.FromJson/FromXml` | 方向不安全 facade | 已删除；迁移到 `FromJsonDocument`/`FromXmlDocument` |
| `CsvStreamExtensions.ExportToBytesAsync` | 伪异步 API | 已删除；使用同步 `ExportToBytes` |
| `IMappingProfileResolver` | 新增只读解析契约 | Plan compiler 和业务消费方依赖只读 resolver；`IMappingProfileRegistry` 保留启动期 Register |
| Profile 稳定名称 | Core 新增注册重载 | 推荐 `AddMappingProfile<T>("stable-name")`；FullName 仅兼容 fallback |
| Profile 扫描 | Core 扩展提供程序集扫描与注册 | 已增加 `ReflectionTypeLoadException` 部分加载容错；扫描默认使用 FullName，稳定 alias 通过显式注册重载提供 |

## 约束

- 未执行版本号修改、包发布、commit、push 或 PR。
- 任何 breaking change 需先完成成员级引用清单和迁移说明。
