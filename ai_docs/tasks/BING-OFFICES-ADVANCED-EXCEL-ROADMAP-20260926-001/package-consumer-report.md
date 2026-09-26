# Package Consumer Gate

## Scope

- Task: `BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001`
- Package version: `2.0.0` (no version upgrade)
- Isolated feed: `artifacts/open-provider-20260926/packages-final`
- Isolated NuGet configuration/cache: `artifacts/open-provider-20260926/consumer-final`

## Packed Packages

The following eight packages were packed from the current Release outputs:

- `Bing.Offices.Abstractions.2.0.0.nupkg`
- `Bing.Offices.Core.2.0.0.nupkg`
- `Bing.Offices.Npoi.2.0.0.nupkg`
- `Bing.Offices.MiniExcel.2.0.0.nupkg`
- `Bing.Offices.ClosedXml.2.0.0.nupkg`
- `Bing.Offices.ExcelDataReader.2.0.0.nupkg`
- `Bing.Offices.SpreadCheetah.2.0.0.nupkg`
- `Bing.Offices.AsposeCells.2.0.0.nupkg`

The consumer projects reference the six Provider packages; Abstractions and Core are restored through the package dependency graph. This task adds an actual SpreadCheetah export and verifies the generated Sheet and complete data row with NPOI. Package source mapping pins Bing.Offices packages to the final local feed; the fresh cache prevents reuse of older same-version packages. No global NuGet source is modified.

## Results

| Consumer | Restore/build | Runtime | Evidence |
|---|---|---|---|
| `Bing.Offices.Consumer.Net6` | Passed | Passed (.NET 6.0.36) | `artifacts/open-provider-20260926/consumer-final/net6-run.log` |
| `Bing.Offices.Consumer.Net8` | Passed | Passed (.NET 8.0.30) | `artifacts/open-provider-20260926/consumer-final/net8-run.log` |

Both consumers reported `package-consumer-ok` and verified generated content, including the packaged SpreadCheetah workbook's sheet and row values. The net6 run emitted the expected SDK/net6 end-of-support and `Microsoft.Bcl.Memory` compatibility warnings; it had zero errors.

Commands used:

```text
dotnet restore tests/Bing.Offices.Consumer.Net6/Bing.Offices.Consumer.Net6.csproj --configfile artifacts/open-provider-20260926/consumer-final/NuGet.Config --packages artifacts/open-provider-20260926/consumer-final/nuget-cache -p:BingOfficesPackageVersion=2.0.0 --force-evaluate --no-cache
dotnet restore tests/Bing.Offices.Consumer.Net8/Bing.Offices.Consumer.Net8.csproj --configfile artifacts/open-provider-20260926/consumer-final/NuGet.Config --packages artifacts/open-provider-20260926/consumer-final/nuget-cache -p:BingOfficesPackageVersion=2.0.0 --force-evaluate --no-cache
dotnet build <consumer-project> --no-restore -c Release -p:BingOfficesPackageVersion=2.0.0 -v:minimal
$env:BING_OFFICES_PACKAGE_VERSION='2.0.0'; dotnet run --project <consumer-project> -c Release --no-build --no-restore
```
