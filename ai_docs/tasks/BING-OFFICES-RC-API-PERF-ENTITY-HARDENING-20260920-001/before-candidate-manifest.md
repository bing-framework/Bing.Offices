# Before Candidate Manifest

## Identity

- Task: `BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001`
- Source commit: `94bb52e84ffc70634067b04541857433ef7af9df`
- Source role: clean base extracted with `git archive`; the extracted tree was restored and built independently before running probes.
- Configuration: `Release`, `.NET 8.0.30`, Windows `10.0.22631`, processor count `22`.
- Build result: `dotnet build Bing.Offices.sln -c Release --no-restore --nologo` passed with `0` warnings and `0` errors in the extracted base tree.
- Resource budget: `UNAPPROVED`; this is not a 2CPU/4GiB or production-machine run.

## Probe Artifacts

| Probe | Artifact | Status |
|---|---|---|
| Provider comparison, 100K, 3 repetitions | `provider-comparison-100k-before-final2.json` | PASS; source commit and physical assembly hashes recorded in header |
| Real IO, 1K/10K/100K, 24 scenarios | `real-io-before-final.json` | PASS; source role and commit recorded; budget `UNAPPROVED` |
| RawDate/Relation hotspot | `hotspot-before-final.json` | PARTIAL; RawDate 3/10/30 and Relation 1K/10K complete; Relation 100K stopped before a complete result |
| Relation before baseline Job Object bounded probe | `windows-job-hotspot-before-baseline-60s-final.json` and `hotspot-before-baseline-60s-final.json` | PARTIAL/NOT_VERIFIED; independently built from an isolated tree at source commit `94bb52e84ffc70634067b04541857433ef7af9df`; RawDate 3/10/30 and Relation 1K/10K completed at 3 repetitions each; Relation 100K reached `runner-timeout` at 60 seconds with exit `-1` and no complete statistics; `TEMP` cleanup `0/0/deleted` |
| DynamicPlan microbenchmark, 8 combinations | `dynamic-plan-before-after.json` plus `dynamic-plan-before/results/*` | PASS as directional base sample; InProcess ShortRun, budget `UNAPPROVED` |
| PropertyAccessor microbenchmark, 4 methods | `property-accessor-before-after.json` plus `property-accessor-before/results/*` | PASS as directional base sample; InProcess ShortRun, budget `UNAPPROVED` |
| GenericSheetDispatch microbenchmark, 4 combinations | `generic-dispatch-before-after.json` plus `generic-dispatch-before/results/*` | PASS as directional base sample; InProcess ShortRun, budget `UNAPPROVED` |
| GenericSheetDispatch exception probe | `generic-dispatch-exception-probe.json` | Current candidate direct FirstChanceException probe; caught NPOI `XSSFFactory.CreateDocumentPart` MissingMethodException with successful import; budget `UNAPPROVED` |

## Physical SHA-256

The following values are physical Release file hashes recorded by the before probes. They are not API canonical identity hashes.

| Assembly | SHA-256 |
|---|---|
| `Bing.Offices.Benchmarks.dll` | `ED59CFF92AC0964E8F4E669297064AA752378B3A0EBDFC9932393DD020F0A723` |
| `Bing.Offices.Abstractions.dll` | `259D4CE12612B05AB14D87C8985497E85DBE1AEFF947CFBFBEF72BA9048DF3AC` |
| `Bing.Offices.Core.dll` | `98BE39A21E02525AFC7D9E37A4E6016E504BE183501C73F222D6E5C4F216AC37` |
| `Bing.Offices.Npoi.dll` | `CD4114D5E6F83D2F8B57EC5257679316893CAE461E1EDB645B9AE9B924061261` |
| `Bing.Offices.MiniExcel.dll` | `88226DED66126FCCE125A9B85F092FEBC81AD8835CE789C5F450A129826A0278` |

## Restricted Rerun Physical SHA-256

The two bounded-probe artifacts above record the following physical Release assembly hashes in their headers. These are from a different independent build than the existing `before-final` hashes in the preceding table; they do not replace or merge with those values. The artifact header is authoritative for this bounded rerun.

| Assembly | SHA-256 in bounded-probe artifact header |
|---|---|
| `Bing.Offices.Benchmarks.dll` | `65AEA685D40829C81F3221B88DE813EC4B782521EAFE4E616873B5736B63BF25` |
| `Bing.Offices.Abstractions.dll` | `56FC07A1B5DB95C1EB4850D3180675D55C54D4E86B9B19AB0461A2008E310617` |
| `Bing.Offices.Core.dll` | `48E62C1E21049EC7624E7DD85BDF705E4A1E4C05E8A355AAA97BD48ADA1B75CA` |
| `Bing.Offices.MiniExcel.dll` | `C4AC8A25DFE3221CE88F68C7D9D79EEFAAECF6F4B332BC3F977C0459FF0E9AB2` |

## Limitations

- The before tree was used to establish a clean baseline and to run the completed Provider/Real IO probes; it does not prove every A-I workload.
- The old Relation 100K hotspot was deliberately stopped after the 10K workload became impractically long. No throughput or improvement claim is made for that missing cell.
- The bounded Relation before rerun only fixes a baseline capacity boundary: its Relation 100K cell timed out under the 60-second Job Object runner and has no complete before statistics. It remains `PARTIAL/NOT_VERIFIED` and does not establish a performance gate or replace the complete before-final identity.
- No threshold was approved or applied. These files are evidence for independent review, not a release approval.
