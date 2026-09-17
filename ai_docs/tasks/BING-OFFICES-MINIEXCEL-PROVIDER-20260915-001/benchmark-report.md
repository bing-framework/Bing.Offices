# Benchmark 与受限环境报告

## 环境

- Windows PowerShell，.NET SDK 10.0.401。
- Runtime：.NET 8.0.30；目标 Provider：MiniExcel 1.46.0。
- Benchmark 类型：真实 XLSX IO controlled probe；每行通过延迟数据集合导出，再从生成字节导入并校验行数。

## Benchmark 发现

`dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --list flat` 已列出：

- `MiniExcelRealIoBenchmarks.ExportImportSync`，RowCount=1000/10000/100000。
- `MiniExcelRealIoBenchmarks.ExportImportAsync`，RowCount=1000/10000/100000。

## 100K controlled probe

原始文件：`artifacts/benchmark-miniexcel-probe-100k-final.json`。

| 模式 | elapsed | allocated | 输出 |
|---|---:|---:|---:|
| sync | 2013.8095 ms | 1,284,034,768 B | 2,078,323 B |
| async | 1473.6153 ms | 1,281,865,400 B | 2,078,321 B |

两种模式均报告 `rowCount=100000`，说明当前链路可完成 100K 真实写入/读取；分配量显示该实现仍有明显的行字典、内存缓冲和 MiniExcel 内部开销，不能据此宣称低内存或优于 NPOI。

## 未完成项

- BenchmarkDotNet Dry 发现了 6 个 MiniExcel benchmark，但自动生成项目因 NuGet `NU1301`（api.nuget.org SSL/凭证）无法构建，0 次有效 BDN 迭代。
- 500K/1M 长测未运行；当前证据只覆盖 100K controlled probe。部署前应在可访问 NuGet、稳定磁盘和固定机器的环境补跑，并记录 Gen0/1/2、峰值工作集、P95 和 rows/s。
