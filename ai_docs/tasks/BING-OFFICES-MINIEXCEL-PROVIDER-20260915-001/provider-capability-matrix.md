# Provider 能力矩阵

| 能力 | MiniExcel 结果 | 证据/边界 |
|---|---|---|
| XLSX 普通导入导出 | 支持 | 真实 MiniExcel roundtrip、包消费者、集成真实路径 |
| 多 Sheet / 异构 DTO | 支持 | `RoundTrip_ShouldSupportMultipleSheetsAndMapping` |
| Stream / File / byte[] | 支持 | 接口 Stream、Core 原子提交、`ExcelStreamExtensions`、消费者输出 |
| 固定列 / 动态列 | 支持 | Core mapping plan + 延迟字典行适配器 |
| Converter / ValueMap / Formatter | 支持 | `MiniExcelValueAdapter` 与 named converter 测试 |
| Validation / Unique / Relations | 支持 | Core validation/relation contract 回归 |
| XLS | 不支持，fail-fast | `Format != Xlsx` 在写入前拒绝 |
| 模板 | 不支持，fail-fast | `Template`/`TemplateRegion` 在 preflight 拒绝 |
| 复杂样式、数字格式、列宽、隐藏 Sheet | 不支持，fail-fast | 请求级与最终 mapping plan 均检查，避免静默丢失 |
| merge / multi-header | 不支持，fail-fast | `HeaderRows` 非空即拒绝 |
| 图片 | 不支持，fail-fast | 非默认 `ImageMultiplicity` 即拒绝 |
| 批注 | 不支持，fail-fast | HeaderRows/非默认冲突策略即拒绝 |
| Chart | 不支持，fail-fast | `Charts` 非空即拒绝 |
| Failure Workbook | 不支持，fail-fast | `FailureOptions.Mode != None` 即拒绝 |

MiniExcel 公共调用方只依赖 Bing.Offices 接口、Request、Result 和异常契约；项目源码不引用 NPOI 类型，包 nuspec 也不含 NPOI 依赖或标签。
