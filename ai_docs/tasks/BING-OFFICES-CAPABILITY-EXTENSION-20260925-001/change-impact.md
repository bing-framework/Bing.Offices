# Change Impact Analysis

| 项目 | 影响 |
|---|---|
| 生产项目 | `Bing.Offices.Abstractions` 新增兼容的 `ExcelFormat.Xlsb`、细粒度能力描述和 `IExcelBatchImporter`；新增 `Bing.Offices.ExcelDataReader` Provider。 |
| 测试项目 | ExcelDataReader 直接职责测试、只读 Provider 合同、共享 API/消费者门禁；既有 NPOI、MiniExcel、ClosedXML 合同继续运行。 |
| 公共 API | 仅新增成员和新包；保留现有枚举值、接口签名和 Provider 包，不删除或改名既有成员。双 TFM snapshot 已更新并获本任务 `api-approval.md` 批准。 |
| 运行时路径 | 新 Provider 对输入先同步/异步暂存，再由 ExcelDataReader 前向解析；完整导入共享 Workbook 行预算，分批导入按单 Sheet 串行交付。 |
| 资源与所有权 | 输入流由调用方拥有；暂存文件由 Provider 清理。完整导入超限时不返回此前 Sheet 的部分实体；分批导入保留此前已交付批次。 |
| 依赖与发布 | 仅新增 ExcelDataReader 3.9.0 及必要的 CodePages 支持；不升级既有引擎、版本、提交、推送或发布。 |
| 未覆盖环境 | Linux/容器字体、生产规模容量和外部 CI 仍需在对应部署环境验证，不以本机测试替代。 |

风险等级：高（公共 API 与新 Provider），通过双 TFM 构建、公共 API snapshot、真实 nupkg consumer、Provider Contract 和直接职责测试降低风险。
