# Package Consumer 报告

## 结论

状态：`VERIFIED`。本轮三个 `2.0.0` nupkg 已由独立消费者从任务本地产物源恢复；消费者没有 `ProjectReference`。`netstandard2.0` 合同项目编译通过，`netcoreapp3.1`、`net6.0`、`net8.0` 运行消费者均输出 `package-consumer-ok`。

该结论只证明当前工作树生成的本地包可消费，不代表同版本包的不可变发布身份，也不解除正式 API baseline、性能预算和独立 Review 门禁。

## 身份与结构

- Git：`master` / `9d78ab76e28891ac9b7e0e558f4df669721a4ea6` / dirty（全部变更属于本任务，执行前干净）。
- SDK：`10.0.400`；Windows；Release。
- 最终本地源：`artifacts/packages-rc-final`；第三方源：`https://api.nuget.org/v3/index.json`。
- 最终隔离缓存：`artifacts/.packages-rc-final` 与 `artifacts/.http-cache-rc-final`。
- 运行消费者：`artifacts/package-consumer/runtime/Consumer.csproj`，目标 `netcoreapp3.1;net6.0;net8.0`。
- 合同消费者：`artifacts/package-consumer/contract/ContractConsumer.csproj`，目标 `netstandard2.0`。
- 两个项目均只含 `PackageReference`；assets 中没有 `type=project` 的 Bing 依赖。

## 包身份

| 包 | SHA-256 | 关键内容 |
| --- | --- | --- |
| `Bing.Offices.Abstractions.2.0.0.nupkg` | `5DF62DCCBB57C258FF0909BCCA130BABEC4BB58A06C2702A4F0DFF7FD5C91DD9` | netstandard2.0 DLL/XML、nuspec、LICENSE、README |
| `Bing.Offices.Core.2.0.0.nupkg` | `6CDE11FE299E10F3D8D9C399611C910B17CB280592A2D607B69D35F3B5ED0503` | netstandard2.0 DLL/XML、nuspec、LICENSE、README |
| `Bing.Offices.Npoi.2.0.0.nupkg` | `8FA16959AA74637D83636E5A3692E6E31FEF21F5A5D7D80ABD92680FECB79964` | netcoreapp3.1/net6.0/net8.0 DLL/XML、nuspec、LICENSE、README |

运行消费者 assets 将 Abstractions/Core/Npoi `2.0.0` 全部解析为 `type=package`；合同消费者将 Abstractions/Core `2.0.0` 解析为 `type=package`。Npoi 包不声明 netstandard2.0 资产，因此 netstandard2.0 使用 Core/Abstractions 合同消费者验证，不伪造 Npoi 兼容性。

## 覆盖与结果

运行消费者实际验证：`AddBingOfficesNpoi` 链式注册；四个 Excel/CSV 服务解析；方向化 Export Mapping；JSON v2 Mapping；XLSX 导出、重开和导入；typed CSV 往返；七个 NPOI 扩展容器；`TryAddPicture` 参数失败无修改；未知 `ISheet` 通过统一 `BingOfficesException` catch 获得 UnsupportedFeature/Export/Write 分类。合同消费者编译使用 Core/Abstractions 的 Mapping Loader、typed CSV、`ExcelRequired` 与 `ExcelDate`。

包内 DLL 与当前 `output/release` 的 SHA-256 已逐项相等：Abstractions/Core netstandard2.0，以及 Npoi netcoreapp3.1/net6.0/net8.0 共五项均 `match=True`。

| 目标 | Restore | Build | Run | 结果 |
| --- | --- | --- | --- | --- |
| `netstandard2.0` | PASS | PASS，0 warning/0 error | 不适用 | 合同编译通过 |
| `netcoreapp3.1` | PASS | PASS | PASS | `package-consumer-ok` |
| `net6.0` | PASS | PASS | PASS | `package-consumer-ok` |
| `net8.0` | PASS | PASS | PASS | `package-consumer-ok` |

netcoreapp3.1/net6.0 已 EOL；构建还报告部分 8.0 第三方包不正式支持 netcoreapp3.1。这些是支持矩阵风险，不是本轮包消费失败。

## 可复现命令

```powershell
$env:NUGET_PACKAGES = '<task>/artifacts/.packages-rc-final'
$env:NUGET_HTTP_CACHE_PATH = '<task>/artifacts/.http-cache-rc-final'
dotnet restore <task>/artifacts/package-consumer/runtime/Consumer.csproj --configfile <task>/artifacts/package-consumer/NuGet.Config --force-evaluate --no-cache
dotnet restore <task>/artifacts/package-consumer/contract/ContractConsumer.csproj --configfile <task>/artifacts/package-consumer/NuGet.Config --force-evaluate --no-cache
dotnet build <task>/artifacts/package-consumer/runtime/Consumer.csproj -c Release --no-restore
dotnet build <task>/artifacts/package-consumer/contract/ContractConsumer.csproj -c Release --no-restore
dotnet <task>/artifacts/package-consumer/runtime/bin/Release/netcoreapp3.1/Consumer.dll
dotnet <task>/artifacts/package-consumer/runtime/bin/Release/net6.0/Consumer.dll
dotnet <task>/artifacts/package-consumer/runtime/bin/Release/net8.0/Consumer.dll
```

## 失败记录

首次将 Npoi 消费者错误加入 `netstandard2.0`，restore 正确返回 `NU1202`，随后拆分为运行消费者和 Core 合同消费者。首次两个项目同目录共享 `obj` 导致 assets 冲突，已通过物理子目录隔离。首次 CSV 场景复用了带 Excel 日期特性的模型并触发校验错误，已改为独立 CSV 模型；失败进程经命令行精确确认后终止。最终证据均来自修复后的重新 restore/build/run。
