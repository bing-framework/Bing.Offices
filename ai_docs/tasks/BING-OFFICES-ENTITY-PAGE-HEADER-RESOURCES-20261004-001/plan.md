# Entity 重复页眉大跨度扫描与取消

Task ID：`BING-OFFICES-ENTITY-PAGE-HEADER-RESOURCES-20261004-001`

## 范围

完成 DESIGN-004：验证稀疏大表、显式大 `End`、资源限额与取消响应。不增加公共 API，不改变 marker、分页、分组或明细终点语义。

## 实施

1. 将既有取消令牌传入 NPOI/ClosedXML 的重复页眉末行回溯及共享 marker 扫描，并在长循环中分块检查。
2. 复制页表头校验也响应取消。保留原有行扫描策略；只有测试显示实际性能回归时才改用已用单元格扫描。
3. 用 XLS/HSSF、XLSX/XSSF、ClosedXML XLSX 的真实稀疏工作簿核对大 `End`、非列表远端单元格、资源限制及取消。
4. 执行双 TFM 职责测试、Release 常规验证、API 成员与包身份比较、编码和差异检查。不提交、推送或发布。
