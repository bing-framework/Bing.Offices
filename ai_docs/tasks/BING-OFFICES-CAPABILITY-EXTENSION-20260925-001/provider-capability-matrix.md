# Provider 能力矩阵

本矩阵只记录 Bing.Offices 当前公开并验证过的能力。底层第三方组件拥有的能力，不会因为列入 `docs/excel/10-third-party-components.md` 就自动成为 Offices 合同。

| Provider | 读取 | 写入 | 完整 Workbook 导入 | 分批导入 | 公式语义 | 明确边界 | 直接证据 |
|---|---|---|---|---|---|---|---|
| NPOI | XLS/XLSX | XLS/XLSX | 支持 | 未公开 | 按 Provider 合同 | 受现有资源、Failure Workbook 和格式矩阵约束 | `tests/Bing.Offices.ProviderContract.Tests`、`tests/Bing.Offices.Npoi.Tests` |
| MiniExcel | XLSX | XLSX | 支持 | 未公开 | 按 Provider 合同 | XLS、复杂 Workbook、Failure Workbook 等在预检中拒绝 | `tests/Bing.Offices.ProviderContract.Tests`、`tests/Bing.Offices.MiniExcel.Tests` |
| ClosedXML | XLSX | XLSX | 支持 | 未公开 | Workbook 规则与配置规则按合同执行 | 不支持语义结构化返回 `UnsupportedFeature`，不静默降级 | `tests/Bing.Offices.ProviderContract.Tests`、`tests/Bing.Offices.ClosedXml.Tests` |
| ExcelDataReader | XLS/XLSX/XLSB | 无 | 支持 | 支持，单 Sheet | 只读取引擎提供的缓存值，不计算公式 | XLSX 的 `MaxCells/MaxColumnsPerSheet` 按 ZIP 物理 `<c>` 预检；XLS/XLSB 请求这两项限制显式 Unsupported；另有固定列、无导出、Entity、动态列、关系、图片、原生 Workbook 校验和 Failure Workbook 边界 | `tests/Bing.Offices.ExcelDataReader.Tests/ExcelDataReaderProviderTest.cs`、`ExcelDataReaderBoundaryTest.cs`、`ImportOnlyProviderContractTest.cs` |

## 分批语义

- `IExcelBatchImporter` 使用串行回调；异步回调完成前不会读取下一批。
- 批次只返回成功实体和当前批次错误；失败行不会作为成功实体交付。
- 跨行唯一性、跨 Sheet 关系等需要全局状态的规则在首版拒绝，而不是返回不完整的成功结果。
- 资源超限会停止后续交付；已交付批次不回滚。输入流始终由调用方拥有，Provider 的暂存文件在成功、失败、取消和回调异常后清理。

## 能力选择规则

能力预期按“操作 + 格式 + 场景”由合同显式维护；不会从生产能力标志自动推导测试预期。旧 Provider 请求 XLSB 时必须明确拒绝，不发生自动路由或自动降级。
