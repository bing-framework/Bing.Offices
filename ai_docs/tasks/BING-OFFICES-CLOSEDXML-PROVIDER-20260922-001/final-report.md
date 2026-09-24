# ClosedXML Provider Final Report

## Outcome

本轮完成了一个可编译、可直接测试的 `Bing.Offices.ClosedXml` 基础 Provider：

- `ClosedXmlExcelExporter` / `ClosedXmlExcelImporter` 和 `AddBingOfficesClosedXml`；
- net6/net8 项目、解决方案、版本属性和 PackageReference 接入；
- List/Workbook XLSX、多 Sheet、名称/索引、Core Mapping、动态列、基础样式、Merge、Formula、模板、流所有权、取消、真实外围异步 IO、输入预检和原子文件提交；
- ClosedXML Unit 74 × 2 TFM、Integration 4 × 2 TFM 通过；固定列混合显式索引、动态列稀疏物理布局/显式物理索引/冲突/物理列宽、列级样式优先级、行高、重复表头、隐藏行列/合并锚点/Blank/Empty/Error 边界、MaxRows/MaxSheets/MaxColumns/MaxCells 在 `XLWorkbook` 创建前拒绝并返回结构化资源错误，seekable/non-seekable `MaxInputBytes`、导入取消和图片数量/大小限制在 ClosedXML DOM 前拒绝；模板公式、Comment、Data Validation、Conditional Formatting、Merge、Style、Comment conflict/overwrite 真实读回通过；Entity 同 Sheet/跨 Sheet 多 List Region、Relations、缺父项/异步取消、模板异步、真实文件提交失败清理、缺 Sheet/缺 Merge/重叠 Region 和映射计划工厂/缓存类型配置动态隔离与失败恢复均有直接测试；Cross-provider 合同 10 × 2 覆盖 Mapping、Formula 数值/布尔缓存值及字符串公式策略、动态列、Relations、RowHeight、Validation、Sheet/Column/Cell resource limit；公共 API Snapshot 因未批准 baseline 缺 ClosedXML 而阻塞；最新 nupkg 的 net6/net8 Package Consumer 实际执行 ClosedXML 导出/导入，ThirdParty Provider Consumer 通过；1K/10K/100K ClosedXML 普通 round-trip 及独立 export/import/file/multisheet/style/template/formula sync/外围 async Probe 已记录真实中位数。
- README、Provider 能力矩阵、任务矩阵和验证报告更新；全量方案回归共 1944 tests，成功 1942、失败 2，唯一失败类别为未批准 API snapshot。

生产符号到直接测试方法的映射见 [`symbol-test-map.md`](symbol-test-map.md)。

## Release status

`PASS_WITH_ISSUES / BLOCKED`。本轮独立 Review 已将 FIX-001（混合固定索引和稀疏列宽）、FIX-002（普通异步导出 admission 生命周期）和 FIX-003（职责级/Cross-provider 证据）全部关闭，`OPEN_ACTIONABLE=0`。全量方案回归除批准 API snapshot 外通过；Failure Workbook、原生 Workbook Validation、Chart/Pivot/Image/XLSM 宏保留、完整公式计算、复杂模板结构、500K/1M、完整大容量资源矩阵、成员级 API baseline 审批以及跨 OS/字体验证仍未完成；其中 API/性能/外部环境项目需要维护者批准或外部环境，不能由执行 Agent 自行标记为发布 `PASS`。没有自动提交或发布 NuGet。
