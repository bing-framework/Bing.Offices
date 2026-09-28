<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-PROVIDER-STARTER-20260928-001
AI_EXECUTION_FINISHED_AT: 2026-09-28T12:36:39.000Z
AI_EXECUTION_STARTED_AT: 2026-09-28T12:20:09.442Z
AI_EXECUTION_RESULT: PASS

# 第三方 Excel Provider 最小项目模板：执行记录

## 实现

- `templates/excel-provider-starter` 是独立的 net6.0/net8.0 示例项目，从 NuGet 包引用 Abstractions/Core，直接使用 ClosedXML 引擎实现公开 `IExcelExporter` 与 `IExcelProviderFeatureDescriptor`。
- 支持基础标量列表、多 Sheet、同步/异步流与文件导出；文件路径经过可注入的 `IFileExportCommitter`。对 Entity、模板、映射、高级布局及其他格式在目标输出前返回结构化不支持错误。单 Sheet 行数和暂存输出字节数设有显式上限。
- DI 示例保留宿主已注册的文件提交器和异常观察器。README 说明受限能力及何时启用独立 Entity Provider 合同；CI 包消费者阶段新增双 TFM 模板测试。
- 本阶段没有修改生产程序集、公共 API、API snapshot 或版本号。

## 验证

- 使用模板自带 `NuGet.Config` 从 `artifacts/packages` 还原，独立包缓存还原成功。
- `Starter.Provider.Tests`：net6.0 11/11、net8.0 11/11。测试覆盖真实 XLSX 两个 Sheet 的完整读回、真实异步目标 IO、能力声明、DI 自定义提交器、不支持请求预检、Entity 拒绝、构建失败异常观察、文件替换和预取消目标保护。
- Docs Markdown 代码围栏测试（net8.0）1/1。`git diff --check` 通过；新增和本阶段修改文件的严格 UTF-8、BOM、行尾及末尾换行检查通过。工作流步骤结构已人工核对，尚未在 GitHub Actions 中执行。
- 文档入口已链接模板；现有生产源码未改动，因此本阶段不重新生成 API 快照或运行无关的全量性能矩阵。

## 后续 TODO

- 实现 Entity Layout 接口后，按实际能力启用独立合同测试包的七项入口。
- 旧版 1.x 迁移分析器需要独立设计，不在本阶段增加未验证的自动转换。
- Source Generator 预编译布局待性能证据后评估。

未提交、推送或发布。
