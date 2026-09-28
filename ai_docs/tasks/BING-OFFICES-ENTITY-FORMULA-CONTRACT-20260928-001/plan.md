# Entity Footer 公式 Provider 合同与进度索引

## 目标

在现有独立 `Bing.Offices.EntityProviderContracts` 包增加第七项公开合同入口，覆盖显式 Footer 公式和连续明细求和公式。统一 Utopa 任务当前阶段索引，避免旧阶段 TODO 被当作当前缺口。

## 实施

1. 增加 `VerifyFooterFormulas(IExcelEntityExporter, IExcelEntityImporter)`，使用公开 Entity 接口导出 XLSX，按 Open XML 解析并精确断言原生公式、单元格地址、空明细、`GapRows`、marker、导入明细边界和调用方流所有权。
2. 验证最终 Footer 与中间小计使用连续求和辅助方法时，在布局阶段、输出流进入调用链之前拒绝。
3. NPOI、ClosedXML 通过同一合同入口；第三方包消费者从本地合同包执行该入口。更新合同使用文档、生产符号追溯和主任务顶部进度索引。

## 验证

先运行双 TFM Provider 合同定向测试，再打包合同包并用隔离本地 NuGet 源运行双 TFM 包消费者；检查文档、`git diff --check` 和修改文件编码。此阶段不改八个生产程序集的公开 API，因此不更新其 API snapshot。不提交、推送、发布或修改版本。
