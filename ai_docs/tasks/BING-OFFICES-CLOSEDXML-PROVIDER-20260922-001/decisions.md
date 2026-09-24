# Decisions

1. ClosedXML 固定为 `0.105.1`，目标 TFM 与现有 Provider 一致，为 `net6.0;net8.0`。
2. 保持既有 capability bit 数值；Styles、Formula、Comment、Chart 等没有独立公共语义时不扩展 Abstractions enum。
3. ClosedXML DOM 只在 Provider 内部使用；公共 Request/Result/SPI 不引用 `ClosedXML.Excel`。
4. 输入先复制到受限可定位流，再执行共享 ZIP/XML 预检，之后才创建 `XLWorkbook`。ClosedXML 导入额外在 DOM 创建前执行 worksheet 物理行预算预检；图片限制通过共享预检的 Provider 内部开关启用，避免改变 NPOI/MiniExcel 未绑定图片时的既有语义。
5. Async API 只对外围 Stream/File IO 使用真实 async；没有 `Task.Run`、`.Result` 或 `.Wait()`。
6. 第一轮已验证基础 List/Workbook、映射、动态列、样式/边框/reset、自定义表头/Comment、公式写入、模板流所有权、资源预检、原子文件提交、基础 Entity Layout、基础 Relations 和 Unique。Entity 的复杂关系/资源矩阵、Failure Workbook、原生 Workbook Validation、Chart/Pivot/Image 创建和完整 Formula Engine 仍未达到发布证据要求，必须保持明确边界。
7. 模板中的已存在结构由 ClosedXML 读取/保存；无法证明无损的高级部件不宣称支持。
8. `xl/workbook.xml` 的 `workbookPr@date1904` 按本地元素名读取，并在转换器边界将 `DateTime`/`DateTimeOffset` 结果归一化；1904 数值序列已有真实 XLSX 读回测试。
9. ClosedXML 的 `MaxRows` 采用 fail-fast 资源边界：发现选定 Sheet 的物理行预算超限时返回空 Workbook 和结构化资源错误，不保留部分行；这与其必须在 DOM 创建前控制峰值的 Provider 约束一致。
