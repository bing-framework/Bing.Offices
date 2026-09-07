# Integration 与文档测试报告

## Integration

命令：`dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore --no-build`

- `15 passed / 0 skipped / 0 failed`
- 覆盖真实 XLS/XLSX/CSV、DI、映射、文件锁/原子提交、失败工作簿、取消和资源错误。

## Docs

命令：`dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -f net8.0 -c Release --no-restore --no-build`

- `10 passed / 0 skipped / 0 failed`
- 已同步 NPOI 发布 TFM 和日期合同文案。

## 未覆盖矩阵

- Linux/macOS runner 未在本执行环境中提供。
- net6/netcoreapp3.1 已从支持矩阵移除，不再运行旧 TFM 测试。
- Office/WPS/LibreOffice 互操作未纳入本地证据。
