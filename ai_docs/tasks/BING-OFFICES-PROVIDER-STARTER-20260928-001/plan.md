# 第三方 Excel Provider 最小项目模板

Task ID: `BING-OFFICES-PROVIDER-STARTER-20260928-001`

## 目标

提供可独立构建和测试的第三方 Provider 起点，演示公开 `IExcelExporter`、方向化能力声明、DI 注册、结构化错误、文件提交、取消和调用方流所有权。模板只声明实际完成的 XLSX 基础列表/多 Sheet 导出能力，不假称 Entity、模板或高级布局。

## 实施范围

1. `templates/excel-provider-starter/` 提供双 TFM Provider 项目、测试项目和 README。示例直接使用 ClosedXML 底层库写出真实 XLSX，不依赖官方 ClosedXML Provider 的实现。
2. 完整支持基础公共属性列、多 Sheet、同步/异步流与文件导出；先写暂存流再写调用方目标。高级请求按公开 `BingOfficesUnsupportedFeatureException` 在输出前拒绝。
3. DI 使用 `TryAdd` 注册导出器及能力描述，并保留宿主自定义 `IFileExportCommitter` 与异常观察器。
4. 测试覆盖真实工作簿数据、双 TFM、同步/异步、文件替换与预取消、流所有权、能力描述、不支持能力和异常观察。文档示范 Entity 合同包在真正实现 Entity 后的启用方式，并明确当前不执行该合同。
5. 从本地 NuGet 包及离线缓存还原并运行模板测试；更新主任务阶段索引和追溯。生产源码、公共 API、快照及版本号不变。

## 验收

模板无需引用仓库内源码即可从 NuGet 包构建；net6.0/net8.0 测试通过；真实 XLSX 内容完整；不支持请求在目标流写入前失败；文件取消不破坏已有目标；能力声明与实际路径一致。
