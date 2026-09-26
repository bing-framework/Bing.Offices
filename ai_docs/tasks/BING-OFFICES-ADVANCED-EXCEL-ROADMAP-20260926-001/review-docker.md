# F06 Docker / Linux 独立审查

## 审查结论

| Finding | 状态 | 证据 | 处理 |
|---|---|---|---|
| F06-D01 无字体负向场景只检查 `fc-list` 不存在，没有执行真实 NPOI AutoFit 路径 | RESOLVED | `scripts/verify-docker.sh` fontless probe；`build/StreamingProbe/Program.cs` | fontless 镜像删除字体文件后，通过公共 `NpoiExcelExporter` + `ExcelColumnWidthMode.AutoFit` 执行并记录实际 outcome。 |
| F06-D02 Docker smoke 未进入 CI，容量结果未由脚本持久化 | RESOLVED | `.github/workflows/ci.yml`; `scripts/verify-docker.sh` | smoke 作为必跑 Linux gate 并上传证据；容量保持 `workflow_dispatch` 手动触发，产生 JSONL/log。 |
| F06-D03 容器脚本覆盖系统 `HOME` | RESOLVED | `scripts/verify-docker.sh` | 保留用户 home 语义，为 `/home/bingoffices` 提供可写 tmpfs，仅设置 `DOTNET_CLI_HOME` 和 `NUGET_PACKAGES`。 |
| F06-D04 容量单元共享 `/tmp` 且结果只在 stdout | RESOLVED | `scripts/verify-docker.sh` | 每个 rows/batch 单元使用独立 `TMPDIR`；宿主端保存完整日志并提取严格 9 行 JSONL。 |

## 已有证据复用

`artifacts/open-provider-20260926/f06-summary.json` 和
`streaming-capacity.jsonl` 已记录 100K/500K/1M x 100/1000/10000 的 9 个容量单元，
全部成功。本次不重跑未变化的容量样本；旧 `f06-summary.json` 中的报表合同失败
保留为历史快照，不代表当前合同结果；新 final report 必须将该段标记为 stale。

## 当前验证

- fontless 镜像内传统字体目录文件数为 `0`。
- 真实公共 `NpoiExcelExporter` + `ExcelColumnWidthMode.AutoFit` 路径已执行；实际结果为
  `ExportFailed / NPOI / Export / Write / BingOfficesExportException`。同一无字体容器内，
  SpreadCheetah 真实生成 XLSX 并通过工作簿读回中文数据，输出 `2822` 字节。证据位于
  `artifacts/docker/fontless-npoi-result.json`。此结果是环境行为记录，不伪造为成功降级。
- 最终受影响合同 Docker smoke 因生产 review 正在确认共享预检参数边界而暂缓；
  本次已在进入 fonted restore 后主动中止，不将不完整日志宣称为 PASS。

## 保留边界

- 基础 SDK 使用 digest，.NET 6 runtime 使用 SHA-512，Debian snapshot 和字体包版本均已固定。
- 生产源码以只读 volume 提供；构建副本位于非 root 用户可写 tmpfs。
- 容量数据是测量证据，不自动批准产品阈值。
- NPOI AutoFit 依赖可用字体；部署环境必须安装并校验批准字体。无字体时当前契约是结构化写入失败，不是自动降级。
