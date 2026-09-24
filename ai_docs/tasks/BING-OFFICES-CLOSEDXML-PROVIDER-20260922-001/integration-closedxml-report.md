# ClosedXML Integration Test Report

## Result

`PASS`：`Bing.Offices.ClosedXml.Tests.Integration` 在 `net6.0` 和 `net8.0` 各通过 4 个测试，共 8 个测试用例。

执行命令：

```text
dotnet test tests/Bing.Offices.ClosedXml.Tests.Integration/Bing.Offices.ClosedXml.Tests.Integration.csproj --no-restore -v:minimal
```

测试使用真实 `XLWorkbook`、`MemoryStream`、`FileStream` 和临时 `.xlsx` 文件，覆盖异步文件提交、同步文件替换、生成失败时保留已有目标文件、真实文件读回、异步导入和取消时保留已有目标文件。

## Remaining gaps

多平台字体/AutoFit、超大文件、ZIP bomb、图片/结构化模板拒绝矩阵和并发 admission control 尚未运行。
