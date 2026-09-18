# Resource 报告

- 命令：`dotnet run --project tests/Bing.Offices.ResourceProbe/Bing.Offices.ResourceProbe.csproj -c Release --no-build --no-restore -- --staging-matrix artifacts/resource/rc-test-arch-resource-matrix.jsonl 1000`。
- 场景：36/36，status=passed；真实 artifact：`artifacts/resource/rc-test-arch-resource-matrix.jsonl`，报告：同目录 `.md`。
- approvalStatus：BLOCKED。未提供维护者批准人/时间；矩阵内的 2 CPU/4 GiB 字段是目标配置，不是本机生产证据。
- 500K/1M、真实 2 CPU/4 GiB 和外部 CI：NOT_VERIFIED。

