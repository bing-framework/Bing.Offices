# F01–F06 集成验证报告

## 最终冻结代码门禁

`artifacts/open-provider-20260926/final-build.log`：Release 构建退出 0，0 错误、5 个既有依赖兼容性/版本解析警告。

`artifacts/open-provider-20260926/final-test.log` 和 `final-tests/`：完整 Solution 测试退出 0，25 个 TRX 合计 **2,820/2,820，通过，无失败和跳过**。未过滤既有 Large 测试，NPOI/MiniExcel 真实大文件往返包括在内。

| 测试项目 | net6.0 | net8.0 |
| --- | ---: | ---: |
| Bing.Offices.Tests | 204 | 204 |
| Tests.Integration | 30 | 30 |
| Npoi.Tests | 602 | 602 |
| Npoi.Tests.Integration | 30 | 30 |
| MiniExcel.Tests | 44 | 44 |
| MiniExcel.Tests.Integration | 9 | 9 |
| ClosedXml.Tests | 115 | 115 |
| ClosedXml.Tests.Integration | 4 | 4 |
| ExcelDataReader.Tests | 32 | 32 |
| ProviderContract.Tests | 276 | 276 |
| SpreadCheetah.Tests | 57 | 57 |
| AsposeCells.Tests（仅既有预检回归） | 2 | 2 |
| Docs.Tests | — | 10 |

实际命令：

```text
dotnet build Bing.Offices.sln --no-restore -c Release -v:minimal -m:1
dotnet test Bing.Offices.sln --no-restore -c Release -v:minimal -m:1 --logger "trx;LogFilePrefix=final" --results-directory artifacts/open-provider-20260926/final-tests
```

## Linux 功能和容量

`docker-final/docker-smoke.log`：双 TFM 各 SheetContent 83、NPOI Report 25、ClosedXML Report 11、SpreadCheetah 57；合计 **352/352**，命令退出 0。固定镜像、非 root、只读源、独立临时目录、4 GiB 限制均由同一脚本执行。

`docker-final/fontless-npoi-result.json`：NPOI AutoFit 在无字体环境结构化失败，SpreadCheetah 成功完整读回中文；不以安装字体检查替代真实操作。

`docker-final/streaming-capacity.jsonl`：100K/500K/1M × 批次 100/1000/10000，9/9 成功，详见 `resource-report.md`。正式阈值未批准，只报告单次实测。

## API 与包消费者

双 TFM 只追加 3 类型、35 成员，旧成员删除数 0；最终八包、隔离缓存消费者及包身份见 `member-api-diff.md`、`package-consumer-report.md`。全新缓存的 net6.0/net8.0 消费者 restore/build/run 全部退出 0，实际调用 SpreadCheetah 并通过 NPOI 检查 Sheet 和完整数据行。

## 证据边界

- 根目录旧 `full-tests`、`full-test.log`、`f06-summary.json` 和旧容量数据不属于最终候选；不将旧源码混合失败掩盖为通过。
- Aspose 只随全量运行既有预检回归，不代表商业渲染、加密、格式保真验收；这些能力本轮 Deferred。
- 远程 CI 未触发，本地 Docker 不能冒充远程运行。现有 net6 兼容性/版本解析警告保留，未升级依赖绕过。
- 代码冻结后仅报告更新，不因文档调整重复运行全量或容量样本。
