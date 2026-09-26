# Bing.Offices 能力补齐与第三方引擎扩展

## 目标

在保持既有 NPOI、MiniExcel、ClosedXML 公共契约兼容的前提下，补齐 Provider 合同扩展基础，并新增开源优先的 ExcelDataReader 只读 Provider。首批支持 XLS/XLSX/XLSB 固定列导入、完整 Workbook 导入和单 Sheet 回调式分批导入。

## 实施范围

- 新增细粒度、可选的 Provider 能力描述；新增 `ExcelFormat.Xlsb`，旧成员值不变。
- 新增回调式 `IExcelBatchImporter`，不在 netstandard2.0 公共抽象中引入 `IAsyncEnumerable`。
- 新增 `Bing.Offices.ExcelDataReader` Provider，复用 Core 映射计划、转换器和校验绑定。
- 只读 Provider 独立注册，不伪造导出器；现有 Provider 合同继续保持导入/导出往返路径。
- 更新 Provider 文档、能力矩阵、API 追溯和执行证据。

## 明确边界

- ExcelDataReader 首批不支持导出、Entity 布局、动态列、关系、图片、原生 Workbook 校验和 Failure Workbook。
- 分批导入仅支持单 Sheet 固定列；已交付批次不回滚，跨行唯一性和跨 Sheet 关系拒绝执行。
- 输入流由调用方拥有；Provider 为非定位流创建临时文件暂存，异步入口使用真实异步复制。
- 不自动路由、降级、升级依赖、修改版本、提交或发布。

## 验收

- 新 Provider 在 net6.0/net8.0 构建并通过职责级 XLS/XLSX/XLSB、转换、校验、资源限制、取消和分批测试。
- 既有测试、双 TFM API/consumer gate 不回归；`OPEN_ACTIONABLE=0` 后停止自动修复。
