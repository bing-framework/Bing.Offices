# Goal 目标推理

## 功能与入口

Goal 是 Harness 的自主执行模式，不是第五个业务阶段。四个公共阶段仍由 Skills 定义：`create-plan`、`execute-plan`、`review-code`、`fix-review`。Skill 规定职责、范围、输入和完成条件；Goal 延长该阶段的持续推理与执行过程。

`.agents/scripts/goal-task.mjs` 根据阶段、`taskId`、Harness、Profile 和执行策略生成目标文本：

```powershell
node .agents/scripts/goal-task.mjs execute-plan <taskId> --harness codex
node .agents/scripts/goal-task.mjs fix-review <taskId> --harness antigravity --profile balanced
```

支持阶段 `create-plan`、`execute-plan`、`review-code`、`fix-review`；Harness 为 `copilot`、`codex`、`antigravity`。`--profile` 选择角色 Profile，`--json` 输出结构化 JSON，`--workspace` 指定项目根目录。脚本只向标准输出写出目标文本，不会启动原生 Goal、执行 Skill、修改任务文件或自行运行 Review/Fix 循环。调用者需将文本交给对应 Harness 的原生 Goal 入口。

## 规则优先级与停止条件

开始 Goal 或 Plan 执行前，先读取用户级 `GOAL_RULES.md`：优先 `$CODEX_HOME/GOAL_RULES.md`，未设置 `CODEX_HOME` 时读取 `~/.codex/GOAL_RULES.md`。用户当前指令、仓库 `AGENTS.md` 和任务 `plan.md` 的优先级高于该公共规则；不得静默忽略更高优先级的约束。

只有存在 `OPEN_ACTIONABLE` finding 才允许开始下一轮自动 Fix。没有该状态时必须停止，即使还存在审批、外部环境、已接受限制、验证边界或延后事项。单个 Task 的自动 Review/Fix 循环最多三轮；同类实验最多一次诊断和一次正式确认。达到无进展停止条件时，不通过重试相同 Fix、延长 timeout 或重跑未受影响的证据继续循环。

## 本地策略配置

`.agents/workflow-policy.json` 提供阶段运行时间、`fixScope`、无进展停止和 Git 安全策略；`.agents/agent-profiles.json` 提供各 Harness 的模型、思考等级和 `executionMode`。当前 `reviewLoop.maxRounds` 配置值为 `5`，不覆盖用户级 `GOAL_RULES.md` 的三轮上限。`goal-task.mjs` 不读取 `reviewLoop`，也不在进程中强制这些时限或循环上限；它只把阶段策略字段写入目标文本，执行限制仍须由 Harness 与操作者遵守。

`fixScope=recommended` 默认选取 `MUST_FIX` 与 `SHOULD_FIX`，跳过 `OPTIONAL`；它不改变只有 `OPEN_ACTIONABLE` 才能进入自动修复的条件。`maxFixAttemptsPerItem` 是本地策略指示值，不替代公共无进展、轮次和审批规则。

## 实现、依赖与边界

目标文本由 `.agents/scripts/goal-task.mjs` 生成，读取 `.agents/agent-profiles.json` 和 `.agents/workflow-policy.json`；阶段契约位于 `.agents/skills/<stage>/SKILL.md`。脚本使用 Node.js ES Modules 和内置 `fs`、`path`、`process` 模块；仓库没有 npm 依赖清单或已声明的最低 Node.js 版本。`executionMode=goal` 只表达角色策略，实际是否能启动原生 Goal 取决于所用 Harness 和当前环境。
