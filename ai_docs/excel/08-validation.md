# Excel 08：验证与限制

本地验证命令：

```powershell
dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 -c Release --no-restore
dotnet test tests/Bing.Offices.Tests.Integration/Bing.Offices.Tests.Integration.csproj -f net8.0 -c Release --no-restore
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet pack src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj -c Release --no-restore
```

真实 Excel/LibreOffice 模板、图表和跨应用互操作性需要在安装相应应用的环境中补充验证。当前实现不宣称 XLS 图表支持。
