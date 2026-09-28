# Continuation 11：最终 Footer 命名锚点

## 契约

新增 `FooterNamed(anchorName, markerText, configure)`；导出工作簿中的命名范围指向本次最终 Footer marker 的单元格，明细变长时随之移动。空明细也生成最终 Footer。模板中同名且同 Sheet 的单格名称更新地址；缺失名称由 Provider 创建。非单格、跨 Sheet、重名或布局内与其他锚点冲突在写入目标流前报 Plan 配置错误。导入仍由 marker 精确判断边界，不依赖名称元数据。

PageSubtotal 与 GroupSubtotal 可能重复，不能绑定同一个单格名称；本阶段仅针对最终 Footer。旧 `Footer` 调用保持原样，既有 API 不删除。

## 变更影响与验证

- ChangedProjects：Abstractions、NPOI、ClosedXML；两 Provider 职责测试、Docs、API/Consumer。
- ChangedPublicContracts：一个 Fluent 方法和一个隐藏 Provider SPI 属性；需双 TFM API 差异、成员追溯及维护者审批记录。
- ChangedRuntimePaths：Entity 导出模板/非模板，XLS/HSSF、XLSX/XSSF、ClosedXML XLSX；导入仅验证名称不妨碍 marker 路径。
- ChangedBuildPackaging：NuGet Consumer 编译运行；ChangedBenchmarkHarness：无。RiskLevel：HIGH。
- 先 L1 两 Provider 真工作簿，再 L2 双 TFM Provider，L3 Contract/API/Consumer/Docs，阶段完成时 solution Release 构建与常规检查。无需重跑 500K/1M 性能矩阵。
