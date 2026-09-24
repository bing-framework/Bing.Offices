# Package Consumer Report

## Status

`PASS`。已为 `ThirdPartyProvider.Consumer`、`Consumer.Net6` 和 `Consumer.Net8` 增加 `Bing.Offices.ClosedXml` PackageReference，并保留 Provider-only 类型不进入公共 API。最新源码变更后重新执行 `dotnet pack src/Bing.Offices.ClosedXml/Bing.Offices.ClosedXml.csproj -c Release --no-restore`，生成根目录本地 nupkg；net6/net8 nuspec 均声明 `Bing.Offices.Core`、`ClosedXML [0.105.1]` 和对应 DI Abstractions。

使用任务专属本地 nupkg feed 和隔离缓存完成无 ProjectReference 的还原、编译和运行：最终源码重新打包后的 nupkg SHA-256 为 `A75EE5FA91021DF3D6B860FEFB251F281479EFA86DA3E5BA7003EB716F970C37`。`Consumer.Net6`（.NET 6.0.36）输出 `package-consumer-ok ... closedXmlBytes=6296`，`Consumer.Net8`（.NET 8.0.30）输出 `package-consumer-ok ... closedXmlBytes=6296`；两个 Consumer 均执行了 ClosedXML DI 的真实 XLSX 导出/导入读回；`ThirdPartyProvider.Consumer` 的 net6/net8 均输出 `third-party-public-only-provider-ok`。网络 NuGet 源的审计索引不可用，但本地包闭包验证通过。
