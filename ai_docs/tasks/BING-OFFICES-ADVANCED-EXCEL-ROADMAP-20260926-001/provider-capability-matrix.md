# Provider Capability Matrix

| Provider | 读取 | 写入 | 流式新建 | 报表核心 | 公式 | 渲染/转换 | 明确边界 |
|---|---|---|---|---|---|---|---|
| NPOI | XLS/XLSX | XLS/XLSX | 否 | Filter/Freeze/CellValue/Formula/Name/Print；XLSX 另支持 Table/ColorScale/DataBar/IconSet | 文本和缓存读取；重新计算 Unsupported | 不在本任务新增 | XLSB/XLSM/ODS、完整计算、渲染不由现有合同承诺 |
| ClosedXML | XLSX | XLSX | 否 | Table/Filter/Freeze/CellValue/Formula/ColorScale/DataBar/IconSet/Name/Print | 文本和缓存读取；重新计算 Unsupported | 不在本任务新增 | XLS/XLSM、完整 Formula Engine、渲染不支持 |
| MiniExcel | XLSX | XLSX | 历史数据写入 | 仅既有公共子集 | 缓存值读取 | 不支持 | XLS、模板高级结构和 Workbook 原生校验按既有边界拒绝 |
| ExcelDataReader | XLS/XLSX/XLSB | 无 | 无 | 无 | 缓存值 | 无 | 只读、固定列、单 Sheet 批量导入 |
| SpreadCheetah | 无 | 新建 XLSX | 是 | 前向可表达子集 | 公式文本写入，不负责计算 | 无 | 不读模板、不回溯、不支持 AutoFit/透视表/图表/批注/宏/加密 |
| Aspose.Cells | XLS/XLSX/XLSM/ODS | XLS/XLSX/XLSM/ODS | 无 | 工作簿型高级能力 | 公式文本/缓存/重新计算 | PDF、PNG、JPEG、TIFF | 商业许可证、字体、宏策略和转换损失需宿主配置/验证 |

## 证据规则

本轮 F01–F06 只验收开源 Provider，商业行保留历史说明，不作为本轮已验收能力。NPOI XLSX 本轮补充 ColorScale/DataBar/IconSet；XLS 对三者与 Table 提前拒绝。

| 新增业务子集 | NPOI XLS | NPOI XLSX | ClosedXML XLSX | SpreadCheetah XLSX | MiniExcel XLSX |
| --- | --- | --- | --- | --- | --- |
| PNG/JPEG 字节图片、像素锚点 | 支持 | 支持 | 支持 | 支持，图片请求磁盘暂存 | 预检拒绝 |
| 整数/小数/日期/文本长度八种比较 | 支持 | 支持 | 支持 | 支持 | 预检拒绝 |
| 显式列表、空值/输入/错误提示 | 支持 | 支持，修正引擎引号序列化 | 支持 | 支持 | 预检拒绝 |
| 高级条件格式：双色、数据条、交通灯 | 预检拒绝 | 支持 | 支持 | 预检拒绝 | 预检拒绝 |

固定/动态映射通过 `SpreadCheetahMappingContractTest` 验证 ValueMap、命名 Converter、物理列顺序和预检；图片/校验由 `SheetContentContractTest` 使用独立预期及 NPOI/OOXML 读取验证。

- `Supported` 仅在对应 Provider 的职责级直接测试和双 TFM 结果存在时使用。
- `Conditional` 必须随请求预检和结构化 Unsupported 行为；不能根据其他 Provider 自动降级。
- Aspose 无许可证时只验证预检失败；真实商业能力需要许可证、固定字体容器和 Golden 文件。
