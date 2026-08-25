# 执行进度

## 当前状态

- 任务：`TASK-BING-OFFICES-STABILIZE-20260825-001`
- 执行器：Copilot plan-execution
- 状态：IN_PROGRESS
- 禁止操作：未执行 commit、push、tag、PR、NuGet publish 或破坏性 Git 操作。

## 已完成

- Phase 0：已读取批准的 `plan.md`，启动任务运行状态，记录工作区状态。
- TASK-1.1：`AddNpoi()` 改为使用 `ExcelValidationRules.CreateDefault()` 注册默认规则，包含 `MaxValueExcelValidationRule`。
- TASK-1.2：配置校验和 Workbook 原生校验按四种 `ExcelImportValidationMode` 分流；Continue 模式可收集两类错误，失败行不会提交。
- TASK-1.3：失败工作簿按实际解析 Sheet 名称关联请求，修复 `ByIndex` 与非零 `HeaderRowIndex`。
- TASK-1.4：新增 `ExcelTemplateCellOverwritePolicy`，默认保留模板样式/批注，显式 Replace 清除样式/批注；表头、自定义表头、正文统一接入。
- TASK-2.1：Builder/DocumentFactory 保留分层快照，Plan Factory 负责 Document/Request 最终合并；CSV/NPOI 主路径已接入独立快照。
- TASK-2.2：缺失方向默认失败，只有显式 `UseConventionFallback` 才允许约定映射；删除方向不安全 loader facade。
- TASK-2.3 主要完成：新增只读 `IMappingProfileResolver`、稳定名称注册重载和程序集扫描部分加载容错；Profile DI 扩展与 descriptor factory 已迁移到 Core，Npoi 不再公开 Profile 注册 API。仍缺少真正构造 `ReflectionTypeLoadException` 的直接测试和扫描 alias 元数据支持。
- TASK-3.2：删除 `ExportToBytesAsync`、`FromJson`/`FromXml` 方向不安全 facade，并更新调用方和 API 快照。
- TASK-6.1 主要治理：命名规则基准使用真实注册规则；计划构建拆分 cold/cache hit/cache miss；主动 LOH 负载从 Mapping 基准移除并保留在隔离 ResourceProbe。
- TASK-6.1 继续完成固定基线配置：两个 Benchmark 类固定 `launch=1/warmup=2/iteration=3`；代表性 `DynamicPlanBuildCacheHit` 16 个参数组合已实际运行并输出 Mean、Allocated、Gen0/Gen1。
- TASK-7.2 发布元数据审计部分完成：三个 2.0.0 包均包含 README、LICENSE、icon 和 `.snupkg`；nuspec 声明 README；SourceLink 依赖已限定为私有构建资产，不进入消费者依赖。未执行发布。
- TASK-4.1 小步完成：新增 `NpoiStreamCopier` 内部协作者，统一 importer/exporter 的流复制、取消和输入大小限制路径；主类其余职责仍待继续拆分。
- TASK-4.1 继续完成低风险拆分：新增 `NpoiFailureWorkbookWriter`，隔离失败工作簿的错误摘要、ErrorRowsOnly 行复制、批注、合并区域、数据验证、图片和输出流写入逻辑；主 importer 仅保留调用边界。
- TASK-4.2 小步完成：新增 `CsvHeaderBinder` 和 `CsvColumn` 内部协作者，隔离 CSV 表头/按位置绑定、重复表头、动态列和 HeaderMatch 校验；CSV 公共 API 与行级转换、校验、Unique journal 行为保持不变。
- TASK-4.2 继续完成：新增 `CsvDynamicTypeResolver`，统一 CSV 导入/导出动态列的受限类型解析和错误语义；未改变公共 API。
- TASK-4.2 继续完成：将 `CsvPropertyBinding` 移至独立 internal 文件，实体管线文件继续聚焦导入/导出流程；未改变成员可见性和行为。
- Round 4：新增 `NpoiRelationBinder`，隔离 NPOI 关系绑定、导航属性写入、错误定位和取消处理。
- Round 4：新增共享 `NpoiWorkbookPlanKeyBuilder`，统一 Import/Export 的 workbook plan 分组键。
- Round 4：新增 `ExcelMappingTextReader`，隔离 Mapping Loader 的有界文本读取、UTF-8 大小校验和流所有权边界。
- Round 4：`ExcelMappingPlanFactory` cache key 改用 `JsonSerializer.SerializeToUtf8Bytes`，移除 JSON 字符串到 UTF-8 的额外中间转换；字段和 SHA256/Base64 契约保持不变。
- Round 5：新增 `NpoiImportPlanBuilder` 和 `NpoiExportPlanBuilder`，隔离 Importer/Exporter 的 workbook plan 分组、泛型反射创建和异常解包；主流程行为与公共 API 保持不变。
- Round 6：新增 `NpoiWorkbookValidationPipeline`，隔离 Importer 的 Workbook 原生 Data Validation 解析、日期/数值/列表比较和错误收集；保留原有校验顺序、UnsupportedFeaturePolicy 与 StopOnFirstFailure 语义。
- Round 6：Mapping benchmark 新增同 payload 的 cache-key 序列化对照，比较旧式 JSON 字符串转 UTF-8 与直接 `SerializeToUtf8Bytes` 路径；固定 Job 下代表性分配约从 `5.89 KB` 降至 `2.52 KB`。

## 本轮验证

- P0 ValidationMode 矩阵：4/4 通过。
- DI 默认规则一致性测试：通过。
- `ErrorRowsOnly` ByIndex + 非零表头测试：1/1 通过。
- 模板现有回归：5/5 通过。
- 模板 Preserve/Replace 字段级回归：2/2 通过。
- 相关命令均为 `dotnet test ... -f net8.0 -c Release --no-restore`；构建伴随既有多目标框架兼容性和过时 API 警告，无新增编译错误。
- net8 Unit 全量：213/213 通过。
- net8 Integration 全量：11/11 通过；net6 Integration 全量：11/11 通过。
- Docs consumer net8：8/8 通过。
- Profile/API 聚焦回归：17/17 通过。
- `StreamPipelineTest` 聚焦回归：79/79 通过，包含 `NpoiStreamCopier` 直接测试。
- 失败工作簿与流管线聚焦回归：104/104 通过；真实导入 Integration：11/11 通过。
- Mapping Benchmark Dry smoke：成功执行 208 个基准，无配置异常；仅有 Dry job 迭代时间偏短提示。
- Round 5 plan builder 抽取后的 net8 核心专项：132/132 通过；包含 Mapping cache、Workbook Request、关系绑定和 Stream import/export 回归。
- Round 6 Workbook validation 拆分后的 net8 核心专项：157/157 通过；net8/net6 Unit 各 216/216 通过；net8/net6 Integration 各 12/12 通过。
- Round 6 cache-key 对照 benchmark：64 个参数组合完成；旧/新路径代表性分配约 5.89 KB/2.52 KB，Mean 约 2.3 us/2.15 us；Dry 迭代偏短，仅作结构证据。
- Benchmark 项目 net8 Release 编译：通过。
- 解决方案 Release build：通过；pack：通过，未发布。
- ResourceProbe：16/16 场景通过。
- net6 Unit 全量：213/213 通过；net5/net7/netcoreapp3.1 仅编译通过，运行被缺失 runtime 阻断。

## 待执行

- 缺失 .NET 5/7/3.1 runtime 的 Unit 运行验证。
- Phase 4 importer/exporter/mapping/CSV 内部协作者拆分。
- Round 5 后仍需继续拆分 Importer/Exporter/Loader 的核心行处理、单元格写入/样式、JSON/XML 解析和业务校验职责；本轮只完成 workbook plan builder 边界。
- Round 6 后仍需继续拆分 Importer 的行 materialization/配置校验、Exporter 的 cell/style/chart 写入和 Loader 的 JSON/XML parser/validator；本轮不宣称 Phase 4 完成。
- Profile 扫描的真正 `ReflectionTypeLoadException` 直接测试和 alias 扫描属性支持。
- 完整性能基线表、资源探针独立运行证据和高影响路径优化对比。
- Office 互操作、老目标 runtime 运行验证、完整 Stream import/export 性能对比、NuGet 本地包消费者安装链路、最终 `final-review.md` 和合法终态 `execution.md`。
