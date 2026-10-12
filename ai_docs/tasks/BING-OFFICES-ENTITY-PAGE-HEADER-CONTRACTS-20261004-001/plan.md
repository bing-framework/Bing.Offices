# Entity 重复页眉共享 Provider 合同

Task ID：`BING-OFFICES-ENTITY-PAGE-HEADER-CONTRACTS-20261004-001`

## 目标

完成 DESIGN-005。将重复页眉的跨 Provider 行为提取到可复用 Entity Layout 合同包，官方 NPOI 与 ClosedXML 同时运行；引擎特有的 XLS/HSSF、XLSX/XSSF、样式、合并和隐藏行结构断言继续留在各自测试项目。

## 实施

1. 在现有 `EntityLayoutProviderContractSuite` 增加重复页眉验证入口，合同实现独立成文件并随测试包发布；只依赖公开 Abstractions。
2. 可复用合同覆盖空/单/多页往返、页眉标记与页码、复制页表头、目标占用保护和结构化错误原始坐标。
3. 资源选项不在 `IExcelEntityImporter` 接口上；官方 Provider 合同测试用小型真实工作簿及各自公开的 options 重载，统一核对输入资源失败。第三方 Provider 可在声明支持实际重复页眉时选择运行共享入口。
4. 更新合同文档与生产符号到测试方法映射。执行双 TFM 职责测试、常规门禁、合同包消费者、API 检查、编码与差异检查。不修改生产逻辑，不提交、推送或发布。
