# Codex 使用说明

V4.4 将 Codex 提升为完整闭环 Harness。

## 四个公共 Skill

Codex 可直接使用：

```text
$create-plan
$execute-plan
$review-code
$fix-review
```

对应：

```text
Plan
→ Execute
→ Review
→ Fix
→ Re-review
```

公共文件：

```text
.agents/skills/create-plan/SKILL.md
.agents/skills/execute-plan/SKILL.md
.agents/skills/review-code/SKILL.md
.agents/skills/fix-review/SKILL.md
```

## 创建计划

```text
$create-plan
taskId=<taskId>
需求：...
```

输出：

```text
ai_docs/tasks/<taskId>/plan.md
```

只规划，不实施代码。

## 执行计划

```text
$execute-plan
taskId=<taskId>
```

状态注册：

```powershell
node .agents/scripts/task-state.mjs start <taskId> --source codex
```

完成后：

```powershell
node .agents/scripts/task-finish.mjs <taskId>
```

`task-finish.mjs` 校验 `execution.md` 的终态机器元数据并收口任务；按 `.env.local`/进程环境中的通知配置，它也可能发送阶段通知。若要禁止该次通知，可传 `--no-notify`。

## 代码 Review

```text
$review-code
taskId=<taskId>
```

输出：

```text
ai_docs/tasks/<taskId>/review.md
```

未解决 MUST_FIX 或 SHOULD_FIX 时保持 NEEDS_FIX。

## 修复 Review

```text
$fix-review
taskId=<taskId>
```

默认：

```text
fixScope=recommended
= MUST_FIX + SHOULD_FIX
```

修复完成后再次：

```text
$review-code
```

仅当 `OPEN_ACTIONABLE > 0` 时才开始新的自动 Fix。没有可执行 Finding、仅剩审批/外部阻塞或达到轮次上限时停止；自动 Review/Fix 最多三轮。

## Goal 目标文本

V4.4 把 Codex Goal 作为正式 executionMode。

`agent-profiles.json` 中 Codex 角色默认可配置：

```json
{
  "executionMode": "goal"
}
```

生成阶段目标文本：

```powershell
node .agents/scripts/goal-task.mjs create-plan <taskId> --harness codex
node .agents/scripts/goal-task.mjs execute-plan <taskId> --harness codex
node .agents/scripts/goal-task.mjs review-code <taskId> --harness codex
node .agents/scripts/goal-task.mjs fix-review <taskId> --harness codex
```

将生成的目标文本交给当前 Codex 的 Goal 入口。

`goal-task.mjs` 只读取 Profile 和 `.agents/workflow-policy.json`，向标准输出写出目标文本（`--json` 可输出结构化结果）；它不会启动 Codex 原生 Goal、运行 Skill 或执法运行时间/Review 轮次。开始 Goal 或 Plan 执行前先读取用户级 `GOAL_RULES.md`（`$CODEX_HOME/GOAL_RULES.md`；未设置时为 `~/.codex/GOAL_RULES.md`）。该文件规定只有 `OPEN_ACTIONABLE` 才能进入下一轮自动 Fix，自动 Review/Fix 最多三轮；本地策略中的 `reviewLoop.maxRounds: 5` 不得覆盖此限制。

原则：

```text
Skill = WHAT / RULES / BOUNDARY
Goal  = 自主执行 / 持续时长 / 目标推理
```

Goal 不替代 Skill。

## Agent Profile

同步：

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced --target codex
```

生成：

```text
.codex/config.toml
.codex/agents/*.toml
.agents/runtime-profiles/codex.json
```

Codex 角色 TOML 会记录模型、model_reasoning_effort，并以注释记录 executionMode。

## 实现与运行边界

公共契约位于 `.agents/skills/`，任务状态由 `.agents/scripts/task-state.mjs`、`task-finish.mjs` 管理；Profile 同步由 `sync-agent-profiles.mjs` 生成 Codex TOML。脚本使用 Node.js ES Modules 和内置模块；仓库未声明最低 Node.js 版本或 npm 依赖。`executionMode=goal` 是同步出的角色策略，不会替代在 Codex 中手动启动原生 Goal。
