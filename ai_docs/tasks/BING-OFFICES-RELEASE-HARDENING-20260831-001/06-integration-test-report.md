# Integration Test Report

状态：`VERIFIED`

待覆盖/复验：XLS/XLSX 真实磁盘、selector、原子目标锁定、Failure Workbook 目录冲突和锁定目标、失败产物重开、独立资源探针。Windows 专属场景在非 Windows 环境标记 `NOT_VERIFIABLE`，不计为 PASS。

当前证据：`dotnet test ...Bing.Offices.Tests.Integration.csproj -c Release -f net6.0 --no-restore -v:q` 和 net8.0 同命令均成功退出；当前环境为 Windows，因此 Windows locked target 场景已执行。日志包含既有 netcoreapp3.1 第三方包 TFM 警告，无测试失败摘要。
