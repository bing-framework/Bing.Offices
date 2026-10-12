# 公共 API 成员审批与迁移

- taskId: BING-OFFICES-REPEATED-PAGE-ANCHORS-20261003-001
- approvedBy: user（要求按建议顺序继续实现 TODO 能力）
- approvedAt: 2026-10-03 (Asia/Shanghai)
- 范围：只增加分页小计逐页命名入口及 Provider 可读取的描述属性；不删除公共成员或修改既有签名。

## 批准的新增成员

- `ExcelEntityListRegionBuilder<TItem>.PageSubtotalNamed(string anchorPrefix, string markerText, Action<ExcelEntityListFooterBuilder<TItem>> configure = null)`：逐页小计名称的公开 Fluent 入口。
- `ExcelEntityListRegion<TEntity>.PageSubtotalAnchorPrefix`：隐藏于 IntelliSense 的只读 Provider SPI 描述。

## 兼容性与迁移

现有 `PageSubtotal`、`FooterNamed`、`GroupSubtotal` 与构造调用保持原样。需要外部公式或程序定位中间页小计时，改用 `PageSubtotalNamed("PageTotal", "页小计", ...)`；输出仅给非末页小计命名为 `PageTotal_001` 起的顺序名称。模板页数缩短时须清理未使用的同系列旧名称，否则导出在写入目标前失败。导入仍依据 marker，不依赖名称。

双 TFM 成员差异：各新增 2 个、删除 0 个，仅涉及 `Bing.Offices.Abstractions`。职责测试通过后已更新基线，并以本次构建和 NuGet 包再次比较，net6.0、net8.0 均通过。
