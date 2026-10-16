# 本轮文档交付清单

整理日期：2026-10-15。本轮新增 8 篇，更新或翻译 38 篇。功能及修订前状态见[对照清单](documentation-inventory.md)，阅读入口见[总索引](README.md)。操作类型“更新、翻译”表示同步校正实现说明，并把英文标题或面向阅读者的说明改为中文；API、路径、配置键与机器状态保持原文。

## 新增文档

| 文件路径（仓库相对路径） | 操作 | 内容摘要 |
| --- | --- | --- |
| `docs/documentation-inventory.md` | 新增 | 按模块记录四类初始状态、源码依据、本轮处理和待实现/待确认项 |
| `docs/getting-started.md` | 新增 | 包和 TFM、NPOI 注册、完整 Excel 往返示例及配置优先级 |
| `docs/excel/workbook-reports.md` | 新增 | 多 Sheet、模板、元数据、样式、报表内容、图片与图表 |
| `docs/excel/entity-layout.md` | 新增 | 固定字段、多明细、命名位置、实体往返与真实 DI 入口 |
| `docs/excel/batch-import.md` | 新增 | 三类批次接口、Buffered/Forward/SummaryOnly 和失败产物提交顺序 |
| `docs/excel/formula-rendering-conversion.md` | 新增 | 公式与缓存语义、商业计算、许可证/字体、PDF/页面和格式转换 |
| `docs/development/build-validation.md` | 新增 | 构建、测试、包消费者、API、Golden、资源与客户端工具 |
| `docs/documentation-changes.md` | 新增 | 逐文件操作、摘要与本轮验证范围 |

## 更新与翻译

| 文件路径（仓库相对路径） | 操作 | 内容摘要 |
| --- | --- | --- |
| `README.md` | 更新、翻译 | 修正失败工作簿、透视、公式缓存和 TFM 说明，整理包表与中文入口 |
| `docs/README.md` | 更新、翻译 | 按入门、功能、Provider 和维护建立完整中文索引与一句话简介 |
| `docs/project-modules.md` | 更新 | 对齐 ExcelDataReader 能力、文档测试覆盖及共享源码定位 |
| `docs/source-layout.md` | 更新 | 补 Anchors 与共享源码，移除过时文件数和尚待拆分推断 |
| `docs/csv/README.md` | 更新 | 补依赖、注册、源码与职责测试定位 |
| `docs/excel/README.md` | 更新、翻译 | 接入新专题，修正实体 Forward 结果保留、合同包和示例覆盖说明 |
| `docs/excel/09-providers.md` | 更新 | 对齐实际依赖、注册、批次、原生校验、失败工作簿和 Forward 边界 |
| `docs/excel/10-third-party-components.md` | 更新 | 对齐已接入版本、源码与配置，区分未接入候选和未确认授权 |
| `docs/excel/mapping-profile.md` | 更新、翻译 | 补方向、转换器、缓存隔离、实现及依赖配置 |
| `docs/excel/mapping-json-xml.md` | 更新、翻译 | 补加载器、观察器入口、解析限制和 v1 迁移缺口 |
| `docs/excel/import-validation.md` | 更新、翻译 | 补规则、资源限制、错误截断及失败产物入口 |
| `docs/excel/dynamic-columns.md` | 更新、翻译 | 补动态列/行计划实现、配置和批量布局关联 |
| `docs/excel/dates.md` | 更新 | 补调用与定位、显式偏移、日期系统和 DateOnly 缺口 |
| `docs/excel/exceptions-and-observers.md` | 更新、翻译 | 补观察器真实签名、DI、分发及取消传播 |
| `docs/excel/npoi-extensions.md` | 更新 | 补单元格用法、注册文件和对象生命周期 |
| `docs/excel/async-io.md` | 更新、翻译 | 明确外围异步、同步引擎与历史部署配置的适用范围 |
| `docs/excel/file-committers.md` | 更新 | 补提交算法与路径，将已落地合同包和模板移出后续建议 |
| `docs/excel/sheet-protection.md` | 更新 | 补保护写入、样式、配置及与文件加密的区别 |
| `docs/excel/cascade-lookups.md` | 更新 | 补隐藏表/名称生成、依赖与字典来源边界 |
| `docs/excel/pivot-tables.md` | 更新 | 补定义、缓存和模板实现路径，保留客户端待确认边界 |
| `docs/excel/entity-layout-analyzers.md` | 更新 | 补两个分析器、Roslyn 依赖、测试与静态分析限制 |
| `docs/excel/entity-provider-contracts.md` | 更新 | 对齐合同包、TFM、依赖、接入及 Forward 不在合同范围的边界 |
| `docs/excel/nuget-migration.md` | 更新、翻译 | 移除不存在的具名 Mapping 方法建议，修正异常替代与历史验证语义 |
| `docs/ai-workflow/README.md` | 更新、翻译 | 工作流实现总览、Node 依赖、全部十一篇入口和执行边界 |
| `docs/ai-workflow/QUICKSTART.md` | 更新、翻译 | 对齐初始化、状态、自检和目标生成用法 |
| `docs/ai-workflow/ROUTES.md` | 更新、翻译 | 对齐三类运行环境协作路线和实际脚本职责 |
| `docs/ai-workflow/CODEX.md` | 更新、翻译 | 区分仓库适配与宿主原生能力，说明目标文本生成边界 |
| `docs/ai-workflow/COPILOT.md` | 更新、翻译 | 对齐 Hook、状态、通知、依赖及不保证宿主能力的限制 |
| `docs/ai-workflow/ANTIGRAVITY.md` | 更新、翻译 | 对齐停止钩子、配置和原生目标入口条件 |
| `docs/ai-workflow/GOAL.md` | 更新、翻译 | 明确脚本不启动 Goal，补 GOAL_RULES、可行动项及循环停止规则 |
| `docs/ai-workflow/PROTOCOL.md` | 更新、翻译 | 对齐任务文档协议、修复范围与自动修复前提 |
| `docs/ai-workflow/AGENT-PROFILES.md` | 更新、翻译 | 说明配置生成、同步位置及配置值与真实模型可用性的区别 |
| `docs/ai-workflow/CHINESE-COMMENTS.md` | 更新 | 移除特定机器 H 盘依赖，改当前技能入口与外部安装待确认说明 |
| `docs/ai-workflow/TROUBLESHOOTING.md` | 更新、翻译 | 对齐 Hook、状态、通知、技能与目标执行的排查边界 |
| `templates/AGENTS.workflow.snippet.md` | 更新、翻译 | 中文规则说明及目标生成、可行动项与用户级规则约束 |
| `templates/excel-provider-starter/README.md` | 翻译 | 将模板标题中文化，保留基础 XLSX 与 Entity 未实现边界 |
| `tests/Bing.Offices.Testing/Resources/Golden/README.md` | 翻译 | 将夹具说明译为中文，补离线生成器与哈希维护路径 |
| `src/Bing.Offices.Analyzers/AnalyzerReleases.Unshipped.md` | 翻译 | 翻译七条规则说明，保留分析器发布跟踪格式所需字段与状态 |

`build/PivotClientProbe/README.md` 已核对：原文为中文，已有真实命令、包依赖、环境和客户端边界，本轮未改写。`docs/plans`、`ai_docs/tasks` 和原始测试日志保留历史语义。任务开始前已有的生产/API 基线及历史记录变更不属于本次文档交付。

## 验证与格式

- 现有文档测试：`dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release --framework net8.0 --verbosity minimal -m:1 -p:UseSharedCompilation=false -p:RestoreIgnoreFailedSources=true -p:NuGetAudit=false`，19 项通过、0 失败、0 跳过；覆盖固定文档列表中的 27 个 C# 示例代码块及目录约束，不等于执行了全部 Markdown。
- 从新增文档提取入门往返、报表、Workbook 批次、实体往返和公式读取五组示例，以临时 net8.0 消费者编译运行，全部通过。实体示例实跑发现 NPOI 未直接注册完整实体 SPI 后，已改为常规接口配合 Core 扩展。临时消费者在验证后清理；Aspose 商业配置片段没有作为真实渲染运行。
- 本轮 46 篇文档的 211 个本地链接及对应标题锚点、仓库源码路径、严格 UTF-8、BOM、混合行尾和末尾换行检查通过；`git diff --check` 通过。外部站点状态及上游仓库内部路径不属于本地路径校验。
- 新增文档使用 UTF-8 无 BOM/LF；原有统一 CRLF 文档保持原样。`dates.md` 原有混合行尾在完整修订时统一为 LF；没有执行全仓格式转换。
- 构建出现现有 `CS1573`（XML 参数注释）和 `CS8632`（可空上下文）警告，未修改生产代码。商业许可证/字体、真实 Office 客户端、线上包发布及生产容量仍为待实现/待确认（待确认），本轮未把历史证据记为当前通过。
