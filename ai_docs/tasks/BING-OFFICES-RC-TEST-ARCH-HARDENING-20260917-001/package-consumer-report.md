# Package / Consumer 报告

- Pack：四个生产包均生成 `2.0.0` nupkg/snupkg，输出目录：`artifacts/packages`。
- MiniExcel nuspec 不含 NPOI 依赖（本次包检查）。
- Consumer net6：PASS，`package-consumer-ok tfm=.NET 6.0.36 packages=Bing.Offices.Npoi+Bing.Offices.MiniExcel/2.0.0 csvBytes=31 excelBytes=4272 miniExcelBytes=4208 npoiExtensions=ok`。
- Consumer net8：PASS，`package-consumer-ok tfm=.NET 8.0.30 packages=Bing.Offices.Npoi+Bing.Offices.MiniExcel/2.0.0 csvBytes=31 excelBytes=4271 miniExcelBytes=4208 npoiExtensions=ok`。
- net6 restore/build 有 SDK EOL warning `NETSDK1138`，无编译错误；版本文件未变。
- 证据：`artifacts/consumers/rc-test-arch/net6-build.log`、`net8-build.log`、`net6-run.log`、`net8-run.log`。

