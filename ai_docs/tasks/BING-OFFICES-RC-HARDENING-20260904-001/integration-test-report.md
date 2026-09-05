# Integration Test Report

## 环境与边界

- OS：Windows 10 `10.0.19045.6466`，x64。
- TFM：`net8.0`、`net6.0`。
- Provider：NPOI `2.7.4`；CSV 使用 Core pipeline。
- 依赖：本地临时流/临时文件；不依赖数据库、Redis、公网或生产资源。
- Windows 文件锁、临时文件和原子提交语义仅代表 Windows 结果，未推广到 Linux/macOS。

## 结果

| TFM | 总数 | 通过 | 失败 | 退出码 | TRX |
| --- | ---: | ---: | ---: | ---: | --- |
| `net8.0` | 15 | 15 | 0 | 0 | `tests/Bing.Offices.Tests.Integration/TestResults/rc-hardening-integration-net8-final.trx` |
| `net6.0` | 15 | 15 | 0 | 0 | `tests/Bing.Offices.Tests.Integration/TestResults/rc-hardening-integration-net6-final.trx` |

## 覆盖内容

- Excel/CSV 真实流导入导出、DI 与直接构造路径。
- 统一异常类型和 inner exception：目标锁定、Failure Workbook 临时目录冲突、目标复制失败及清理行为。
- 取消、流所有权、原子文件提交、Failure Workbook 输出和原目标内容保护。
- NPOI HSSF/XSSF 基础格式和 workbook pipeline。

## 当前限制

- ZIP preflight 的细粒度矩阵主要由 Unit 和独立 ResourceProbe 证明；Integration 结果不宣称覆盖任意恶意压缩输入。
- 1900/1904、公式缓存日期的完整真实文件矩阵仍小于计划中的全量矩阵。
- 未运行外部数据库或外部缓存集成测试，本模块不需要这些依赖。

## 结论

当前可运行 Integration matrix：`PASS 30/30`；跨平台和完整资源边界：`PARTIAL`。
