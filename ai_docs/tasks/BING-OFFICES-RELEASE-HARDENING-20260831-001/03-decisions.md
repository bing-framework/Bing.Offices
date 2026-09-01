# 决策记录

## D-001 输入证据缺失

用户指定的 20260831 评审文件和 merged snapshot 当前不在工作区。执行结论以当前源码、测试、包和 Benchmark 运行证据为准；缺失输入在最终报告中保留 `INPUT_MISSING`。

## D-002 锁文件保护

`dotnet restore --locked-mode` 报 `NU1004` 时不修改 `packages.lock.json`，不使用 `--force-evaluate` 作为发布证据。只在已有资产可用时继续 `--no-restore` 验证。

## D-003 selector 结果复用

在 Workbook DOM 建立后一次解析全部请求 selector，保留请求、物理索引和物理名称；计划构建、导入循环和 Failure Workbook 映射复用该结果。缺失 selector 不进入计划构建，重复物理 Sheet 在计划执行前确定性失败。

## D-004 发布判定

实现完成不等于发布通过。P0、当前 TFM 回归、包消费、Benchmark、文档和互操作证据分别记录；客户端不可用时标记 `NOT_VERIFIABLE`，不包装为 PASS。
