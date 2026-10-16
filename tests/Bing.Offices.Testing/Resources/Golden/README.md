# Golden XLSX 标准夹具

这些文件是通过离线 `build/GoldenFixtures` 生成器创建的、与 Provider 无关的 Open XML 输入。运行时合同测试验证 manifest 中必需的元数据与哈希，不会通过某个 Provider 重新生成输入。

只有经过审查的 `UpdateReason` 才允许更新夹具；更新时同步重新生成对应 SHA-256 条目。生成器入口是 `build/GoldenFixtures/Program.cs`，依赖与参数以该项目为准，不能用修改期望输出的方式掩盖 Provider 回归。
