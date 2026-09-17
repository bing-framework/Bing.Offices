# 资源与安全矩阵

| 场景 | 生产行为 | 证据 |
|---|---|---|
| 输入流读取 | 先复制到受预算约束的 MemoryStream，再做 XLSX preflight 和 Query | `MiniExcelExcelImporter.CopyToMemory/CopyToMemoryAsync`；100K probe |
| 输入字节预算 | 默认 `MaxInputBytes=128 MiB`，超限在第三方解析器前抛资源异常 | `ResourceLimit_ShouldFailBeforeMiniExcelParser` |
| ZIP 炸弹/路径/重复 entry | 检查 entry 数、单 entry/总解压、压缩比、重复/路径穿越 | Core `ExcelXlsxZipPreflight`；NPOI 预检 27/27 双 TFM |
| XML 安全 | DTD 禁止、深度/字符预算、sharedStrings/styles/worksheet 单项和总量预算 | Core 预检测试矩阵 |
| 取消 | 复制、预检、行枚举和 MiniExcel async API 传播 `CancellationToken`；取消不翻译为业务错误 | `PreCanceledAsync_ShouldNotWriteOrRead`；预检取消用例 |
| 导出文件提交 | 复用 Core `IFileExportCommitter`，先临时文件后原子替换；失败不覆盖旧目标 | MiniExcel integration 真实路径；既有 `ExcelAsyncFileIntegrationTest` 回归 |
| 流所有权 | Provider 不关闭调用方 destination/source；失败工作簿目标也由调用方拥有 | 接口测试与全量回归 |
| 失败工作簿 | MiniExcel 不复制/写入失败工作簿，非 None 模式在 preflight fail-fast | `FailureWorkbook_ShouldBeRejectedBeforeParser`，目标流长度保持 0 |
| 并发/缓存 | plan invoker 为静态按类型缓存；mapping factory 负责有界 plan cache；请求状态为局部对象 | 双 TFM full Unit/Integration；无静态请求数据 |

MiniExcel Provider 当前不是完整的端到端 streaming importer：导入缓冲会带来额外内存峰值，这是性能报告和部署容量评估中的已知限制。未执行 10/50/100 并发长矩阵，因此不宣称无限并发或 1M 安全通过。

本次运行的 CI 资源矩阵输出为 `scenarios=36 status=passed approval=BLOCKED`：36 个场景全部通过；命令未提供该历史资源任务的 `approvedBy/approvedAt` 参数，因此审批字段按工具契约保持 BLOCKED，不把它解释为失败或人工批准。
