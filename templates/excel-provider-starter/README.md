# Excel Provider Starter

这是仓库内的最小 Provider 项目模板，目标是展示如何实现 Bing.Offices 公共导出接口并验证真实 XLSX 输出。模板使用 ClosedXML 作为底层引擎；它不是已发布的 NuGet 包，也不代表官方 Provider 的完整能力。

## 项目结构

- `Starter.Provider`：实现公开 `IExcelExporter` 与 `IExcelProviderFeatureDescriptor`，使用 ClosedXML 写基础列表和多 Sheet 工作簿，并通过 `AddStarterExcelProvider` 注册服务。
- `Starter.Provider.Tests`：net6.0/net8.0 职责测试工程，当前 11 个 xUnit 用例使用真实 XLSX 文件验证导出结果、能力声明、资源边界、文件提交和异常观察。

此 starter 只声明基础 XLSX 完整工作簿导出能力。它不提供 `IExcelEntityExporter`、`IExcelEntityImporter` 或 `IExcelStreamingExporter`；模板写入、Entity Layout、合并和其他高级布局必须在请求到达 ClosedXML 前以结构化 `UnsupportedFeature` 拒绝。不要为了让请求“成功”而静默忽略这些配置。示例实现每个工作表最多接受 10,000 行、每个工作簿最多序列化 16 MiB，并在输出前将工作簿暂存于内存；生产 Provider 应根据实际资源预算调整这些限制。

## 本地构建

在仓库根目录执行以下 PowerShell 命令。模板不依赖 Entity 合同包，因此初次构建只需打包 Abstractions 与 Core。`$BingOfficesPackageVersion` 必须与本地 Bing.Offices 包版本一致；仓库当前版本为 `2.0.0`。若仓库包版本已变更，将变量改为对应版本，不要覆盖 `VersionPrefix`。

```powershell
$BingOfficesPackageVersion = "2.0.0"
dotnet pack src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj -c Release -o artifacts/packages
dotnet pack src/Bing.Offices.Core/Bing.Offices.Core.csproj -c Release -o artifacts/packages
dotnet restore templates/excel-provider-starter/Starter.Provider.Tests/Starter.Provider.Tests.csproj --configfile templates/excel-provider-starter/NuGet.Config --packages artifacts/excel-provider-starter/nuget-cache -p:BingOfficesPackageVersion=$BingOfficesPackageVersion
dotnet test templates/excel-provider-starter/Starter.Provider.Tests/Starter.Provider.Tests.csproj --no-restore -c Release --framework net6.0 -p:BingOfficesPackageVersion=$BingOfficesPackageVersion
dotnet test templates/excel-provider-starter/Starter.Provider.Tests/Starter.Provider.Tests.csproj --no-restore -c Release --framework net8.0 -p:BingOfficesPackageVersion=$BingOfficesPackageVersion
```

前两条命令把 Abstractions 和 Core 包写入本地 `artifacts/packages` feed；模板 NuGet.Config 将这两个包映射到本地 feed，并将其余依赖映射到 nuget.org。restore 使用独立的 `artifacts/excel-provider-starter/nuget-cache`，随后两个 TFM 顺序测试且不再次还原。离线环境需预先准备依赖包的本机 NuGet 缓存。将模板复制到外部仓库时，应调整 NuGet.Config 中指向 `artifacts/packages` 的本地包路径。

## 替换底层引擎

1. 在 `StarterExcelExporter` 中替换 ClosedXML 的工作簿创建、Sheet/列写入和序列化代码，继续从 `ExcelWorkbookExportRequest` 读取工作表、数据和格式；此 starter 示例不使用 Core Mapping Plan，只导出未标注映射特性的公开标量属性。
2. 保留 `IExcelExporter` 的同步/异步 Stream 和文件入口。文件入口通过注入的 `IFileExportCommitter` 提交，不能直接截断已有目标文件；Stream 由调用方拥有，直接写入失败可能留下部分输出。
3. 保留 `IExcelProviderFeatureDescriptor` 的方向、格式和完整工作簿能力声明。引擎或请求能力不足时，在写目标前返回结构化 Unsupported 异常，并通过 `IBingOfficesExceptionObserver` 报告支持的异常事件。
4. 为新引擎补充真实工作簿读取断言、格式边界、预取消、失败提交和资源上限测试。只测返回值或 mock 调用不能代替 XLSX 结构证据。

## 启用 Entity Layout 合同

当前 starter 不声明 Entity 能力，因此不要运行 Entity 合同七项，也不要通过跳过或忽略失败将其描述为已支持。实现并验证 Entity Layout 后，再按实际支持范围逐项增加能力声明及独立测试：

- `VerifyCore`：动态字段分组、最终 Footer、命名锚点和明细往返。
- `VerifyAttributeAndNamedAnchors`：属性式固定单元格及模板命名锚点往返。
- `VerifyNamedListFixedCellCollision`：模板中命名列表与固定单元格冲突时，导入/导出配置拒绝及目标流保护。
- `VerifyPageSubtotal`：分页小计、总计 marker 和导入边界。
- `VerifyGroupSubtotal`：连续分组小计、总计 marker 和导入边界。
- `VerifyFooterFormulas`：原生 Footer 公式、明细求和范围和导入边界。
- `VerifyAsync`：异步往返、取消及调用方流所有权。

测试工程添加 `Bing.Offices.EntityProviderContracts` 包引用后，在每项能力的独立测试中传入 Provider 的 `IExcelEntityExporter` 和 `IExcelEntityImporter`：

```csharp
EntityLayoutProviderContractSuite.VerifyCore(entityExporter, entityImporter);
EntityLayoutProviderContractSuite.VerifyPageSubtotal(entityExporter, entityImporter);
EntityLayoutProviderContractSuite.VerifyGroupSubtotal(entityExporter, entityImporter);
EntityLayoutProviderContractSuite.VerifyFooterFormulas(entityExporter, entityImporter);
EntityLayoutProviderContractSuite.VerifyAttributeAndNamedAnchors(entityExporter, entityImporter);
EntityLayoutProviderContractSuite.VerifyNamedListFixedCellCollision(entityExporter, entityImporter);
await EntityLayoutProviderContractSuite.VerifyAsync(entityExporter, entityImporter);
```

这些合同只覆盖 XLSX；XLS/HSSF、底层引擎特有能力和失败行为仍需 Provider 自己的职责测试。未实现的能力应保持未声明，并用独立测试验证其拒绝行为，不能靠跳过合同制造通过结果。
