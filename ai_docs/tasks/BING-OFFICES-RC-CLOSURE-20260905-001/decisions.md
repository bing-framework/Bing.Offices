# 决策记录

## D-001：缺失的 2026-09-05 评审输入

- 状态：`BLOCKED`（仅该证据源，不阻塞源码实施）。
- 旧结论：附件要求优先读取 `ai_docs/codebase-analysis/bing-offices-implementation-review-20260905.md`。
- 当前证据：HEAD `9d78ab7` 中不存在该文件；全仓库只发现此前任务的 review 文件。
- 决定：不伪造或推断缺失报告；以用户附件约束、当前源码/配置/测试为主证据，以 `BING-OFFICES-RC-HARDENING-20260904-001` 的 review/final-report 为次级历史证据。
- 验收：若后续文件出现，逐条复核并记录差异；最终报告保留缺失事实。

## D-002：异步 API 范围

- 状态：`DONE`。
- 决定：本轮不承诺新增异步 API。NPOI 核心为同步 DOM/序列化；禁止以 `Task.Run`、`.Result` 或 `.Wait()` 包装。
- 验收：公开文档说明同步 Provider 限制；源码不得新增伪异步入口。

## D-003：配置加载观察策略

- 状态：`VERIFIED`。
- 决定：采用方案 B。DI loader 为推荐观察入口并接 Observer；静态 API 仅保留 v2 纯解析表面，public v1 migration 已删除。
- 验收：独立 `ExcelMappingConfigurationLoaderTest` 4/4，覆盖 DI、直接构造、静态纯解析、Observer 同实例/单次通知及 Observer 失败。

## D-004：DateTimeOffset 输出

- 状态：`VERIFIED`。
- 决定：默认导出 invariant ISO 8601 round-trip 文本以保留值和 offset；无 offset 文本或 `DateTime` 转 `DateTimeOffset` 必须显式固定 offset，禁止本机 Local。
- 验收：XLS/XLSX 导入→导出→关闭重开→再导入、CSV 往返和 parser 在两个 TZ 下各 6/6。

## D-005：XLS/OLE 资源策略

- 状态：`VERIFIED`。
- 决定：所有 Excel 输入默认受 128 MiB `MaxInputBytes` 限制，显式 null 可关闭；XLS/OLE 在完整缓冲和 DOM 前按字节拒绝，不声称具备 ZIP 内部结构预检。XLSX 在相同输入预算之外继续执行 entry/XML/压缩比预检。
- 验收：XLS 与 XLSX Provider 直接测试、不可寻址流测试、XLSX 预检测试。
