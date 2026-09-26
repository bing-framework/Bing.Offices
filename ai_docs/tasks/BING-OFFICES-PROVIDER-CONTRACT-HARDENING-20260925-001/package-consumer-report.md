# Package Consumer Report

## Current Candidate Consumer Gate (follow-up)

Consumers now restore product packages exclusively from `artifacts/api-candidate`, using isolated cache `artifacts/consumer/cache-provider-followup` and source mapping in `artifacts/consumer/provider-followup.NuGet.Config`. All five cached product nupkg files were checked byte-for-byte against that candidate feed (SHA-256 equality, 5/5). This supersedes the historical feed below; it does not claim any package was published.

Build/run: Consumer.Net6 and Consumer.Net8 returned `package-consumer-ok`; ThirdPartyProvider.Consumer returned `third-party-public-only-provider-ok` on net6.0 and net8.0. All four exited 0. Output binaries are in `artifacts/provider-followup-final/consumer/{net6,net8,third-net6,third-net8}`. net6 build emits the existing NETSDK1138 warning; dependencies and versions were unchanged.

Restore command template: `dotnet restore <consumer.csproj> --packages artifacts/consumer/cache-provider-followup --configfile artifacts/consumer/provider-followup.NuGet.Config -p:BingOfficesPackageVersion=2.0.0 -p:NuGetAudit=false`. Build with `--no-restore -c Release -f <TFM> -p:BingOfficesPackageVersion=2.0.0`; run with `BING_OFFICES_PACKAGE_VERSION=2.0.0`. The local dependency feed excludes all Bing.Offices packages to prevent reuse of stale same-version packages.

Initial restore attempts exposed inherited source mapping and missing net6 dependencies; the isolated config and existing net6/net8 dependency packages resolved them. The first consumer launch lacked its required version environment variable and was not counted as PASS. The final four runs include it.

## Historical Package Identity

Consumers were rebuilt against the fresh approved feed `artifacts/packages/entity-api-approved-20260925` at package identity `2.0.0`; no version or dependency was changed. The five consumer-cache nupkg files match the fresh approved feed byte-for-byte. The API candidate records the canonical package identities and is verified by the dual-TFM API comparison.

| Consumer | Build | Run |
| --- | --- | --- |
| `Consumer.Net6` | PASS, expected `NETSDK1138` plus offline `NU1900` advisory warning | PASS: `package-consumer-ok`, .NET 6.0.36; `npoiExtensions=ok` |
| `Consumer.Net8` | PASS, offline `NU1900` advisory warning | PASS: `package-consumer-ok`, .NET 8.0.30; `npoiExtensions=ok` |
| `ThirdPartyProvider.Consumer` net6 | PASS, expected `NETSDK1138` plus offline `NU1900` advisory warning | PASS: `third-party-public-only-provider-ok` |
| `ThirdPartyProvider.Consumer` net8 | PASS, offline `NU1900` advisory warning | PASS: `third-party-public-only-provider-ok` |

The consumer project files reference product packages only; neither references `Bing.Offices.Testing` nor `Bing.Offices.ProviderContract.Tests`. The third-party fixture continues to implement only the legacy public entity importer/exporter SPI, confirming that the approved resource SPI is opt-in and does not break existing implementations.
