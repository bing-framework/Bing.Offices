# ClosedXML Unit Test Report

## Result

`PASS`：`Bing.Offices.ClosedXml.Tests` 在 `net6.0` 和 `net8.0` 各通过 74 个测试，共 148 个测试用例。

执行命令：

```text
dotnet test tests/Bing.Offices.ClosedXml.Tests/Bing.Offices.ClosedXml.Tests.csproj --no-restore -v:minimal
```

覆盖范围包括：DI 与 first-registration-wins、选项重载、能力声明、多 Sheet、固定列 HasColumnIndex 顺序与混合显式/默认索引、动态物理布局/显式物理索引/空隙/冲突和列级样式优先级、行高、基础样式和边框/重置语义、公式写入与读回、固定列往返、动态列往返、Merge、1900/1904 日期系统、隐藏行列/合并锚点/Blank/Empty/Error 及不兼容目标转换边界、XLS/failure-workbook fail-fast、MaxInputBytes/MaxRows/MaxSheets/MaxColumns/MaxCells DOM 前拒绝/导入取消/图片数量与大小资源预检、seekable/non-seekable 异步输入边界、模板流所有权、不支持模板部件的 DOM 前预检、自定义表头跨度与 Comment/Validation/Conditional Formatting 保留、Comment conflict 与 template overwrite 策略、Entity 固定 Cell/List Region/Merge、多个 List Region/Relations/缺父项/异步取消、Entity Template 同步/异步、Entity 文件 API/真实提交失败清理、缺 Sheet/缺 Merge/重叠 Region 负向场景、映射计划工厂调用和类型/配置/动态列/方向缓存隔离与失败恢复、异步 admission WaitAsync/取消/DOM 后释放、异步取消、异步异常观察器单次通知、命名 Converter、ValueMap、结构化 Validation、Relations 和 Unique 资源上限。

## Remaining gaps

Failure Workbook、原生 Workbook Validation、Chart/Pivot/Image/XLSM 以及完整 validation/relations 策略矩阵和复杂 Entity 布局尚未达到发布证据要求。
