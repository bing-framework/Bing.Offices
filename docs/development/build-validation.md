# 构建、测试与维护工具

## 功能与项目结构

本仓库通过解决方案构建、职责测试、文档示例编译、API 快照、包消费者与有界资源探针验证公开能力。工具通过不等于功能在任意 Office 客户端或生产机器中通过。

关键入口为 `Bing.Offices.sln`、`framework.props`、`common.props`、`common.tests.props`、`Directory.Build.targets` 与 `.github/workflows/ci.yml`。Provider 使用 `.NET 6/.NET 8`；Core、Abstractions、Analyzer 为 `netstandard2.0`，文档测试项目自身为 `net8.0`，不能把它描述成双 TFM 测试。

## 本地用法与依赖

在仓库根目录的 PowerShell 执行；先安装与项目匹配的 .NET SDK/运行时，并准备 NuGet 源。源码构建的版本来自 `version.props` 与 `version.dev.props`，不要为了通过构建覆盖版本文件。

```powershell
dotnet restore Bing.Offices.sln
dotnet build Bing.Offices.sln -c Release --no-restore
dotnet test tests/Bing.Offices.Docs.Tests/Bing.Offices.Docs.Tests.csproj -c Release --framework net8.0
```

业务改动应选择对应测试项目和 TFM，例如 NPOI 的 `tests/Bing.Offices.Npoi.Tests`。测试是否可跳过还需遵守仓库 `AGENTS.md` 中关于默认实现、Provider 分支、缓存隔离、文件提交和公共 API 的直接覆盖要求。

`common.props` 将常规构建产物写入根 `output/debug` 或 `output/release`，项目中也有 `obj`、生成 XML 文档等产物。自定义 Probe 或测试可能指定独立输出路径；清理时不要把 `ai_docs/tasks`、Golden 输入或 API 基线当作缓存。

## 核心验证工具

| 入口 | 核心逻辑与配置 | 限制 |
| --- | --- | --- |
| `tests/Bing.Offices.Tests` 与各 `*.Tests` | Core/Provider 的职责级单元测试，按相应 csproj 运行 | 完整集合和 Large 类别按任务影响选择 |
| `tests/Bing.Offices.ProviderContract.Tests` | 以公共接口验证格式、流、资源和拒绝行为 | 不替代引擎特有测试 |
| `tests/Bing.Offices.Docs.Tests` | `DocsConsumerTest.cs` 提取固定文档列表中的 C# 代码块，编译并执行；`SourceLayoutTests.cs` 检查契约归类 | 不是自动执行所有 Markdown；新增文档不自动进入固定清单 |
| `tests/Bing.Offices.Consumer.Net6`、`Consumer.Net8`、`ThirdPartyProvider.Consumer` | 通过项目配置的包消费路径验证公开 API 可用性 | 须先生成匹配版本/来源的本地包，不能将项目引用测试当成包发布证明 |
| `build/ApiSnapshot` | `Program.cs`、`PublicApiSnapshot.cs` 生成/校验 API 与候选身份；基线为 `build/api-snapshot-baseline.json` | API 变化需双 TFM、成员追溯和审批，不自动刷新基线 |
| `build/GoldenFixtures` | 离线生成 Provider 无关 Open XML 输入 | 更新原因与 manifest 哈希必须同步，见[夹具说明](../../tests/Bing.Offices.Testing/Resources/Golden/README.md) |
| `tests/Bing.Offices.ResourceProbe`、`build/StreamingProbe` | 受控工作负载、暂存与资源数据 | 需绑定候选、规模、环境和预算，不能外推任意容量 |
| `build/PivotClientProbe` | 生成透视夹具；Python/PowerShell 验证专用 Excel 实例生命周期、首开/刷新/重开 | Windows/真实客户端为单独依赖，见[使用说明](../../build/PivotClientProbe/README.md) |
| `scripts/verify-docker.sh`、`docker/verify/Dockerfile` | Linux 容器验证入口与字体准备 | 依赖 Docker，不能用未运行的脚本宣称容器验证成功 |

## CI、配置与限制

CI 定义 Release 构建、分角色测试与包/API/资源相关作业；手动输入 `run_large`、`run_capacity` 分别控制额外大规模职责/容量工作负载。具体过滤器、依赖顺序和 TFM 以 `.github/workflows/ci.yml` 为准。资源探针、商业引擎和 Office 客户端各有独立依赖，不能用无许可证或无 Office 的测试结果外推真实渲染/客户端数值。

仓库 Agent 工具使用 Node.js，脚本和配置位于 `.agents`，见[中文工作流索引](../ai-workflow/README.md)。这些脚本不属于 Bing.Offices 运行时 API。

对纯文档修改，优先检查链接、路径、UTF-8、BOM/行尾及 `git diff --check`；涉及已有可执行示例时运行文档测试。生产容量、线上包发布和外部 CI 当前状态均为**待实现/待确认（待确认）**，需要相应环境证据。
