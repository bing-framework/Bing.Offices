# 集成测试报告

## 已执行结果

| TFM | 项目 | 总数 | 通过 | 失败 | 退出码 | TRX |
| --- | --- | ---: | ---: | ---: | ---: | --- |
| net8.0 | `tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj` | 15 | 15 | 0 | 0 | `artifacts/integration-net8-rerun/integration-net8-rerun.trx` |
| net6.0 | `tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj` | 15 | 15 | 0 | 0 | `artifacts/integration-net6-rerun/integration-net6-rerun.trx` |

命令：

```powershell
 dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -c Release --no-restore -f net8.0 --logger "trx;LogFileName=integration-net8-rerun.trx"
 dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -c Release --no-restore -f net6.0 --logger "trx;LogFileName=integration-net6-rerun.trx"
```

## 覆盖范围

已通过的 30 个测试覆盖：

- XLS 与 XLSX 真实工作簿往返；
- NPOI DI 注册、服务替换、配置 loader、CSV converter、named validation rule；
- 自定义值转换器、真实 workbook 验证和 Fluent/XML 配置；
- 非 seekable 输入、预取消导入/导出和中途取消；
- 文件导出原子替换、锁定目标、失败工作簿临时目录冲突与目标复制失败；
- 调用方 Stream 保持打开和失败后目标内容保护。

## 外部依赖与限制

测试使用本地构造的 XLS/XLSX 和临时文件，不依赖真实数据库、Redis、公网或生产数据。Windows 文件锁相关结果只代表当前 Windows 环境；Linux 行为未在本环境验证。netcoreapp3.1、net5.0、net7.0 Integration 目标不存在或无可运行 runtime，不伪造结果。

## 状态

`VERIFIED`（net6/net8 两个可执行 TFM）。
