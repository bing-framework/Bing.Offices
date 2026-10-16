# 工作流程闭环路线

Universal Agent Workflow V4.4 支持两条闭环线路，并统一使用四个公共 Skill：

```text
create-plan
execute-plan
review-code
fix-review
```

共享任务目录：

```text
ai_docs/tasks/<taskId>/
├─ plan.md
├─ execution.md
└─ review.md
```

## 路线一：纯 Copilot / Codex 自闭环

```text
Copilot / Codex
      ↓
Create Plan
      ↓
Execute
      ↓
Review
      ├─ PASS → 完成
      └─ NEEDS_FIX
             ↓
            Fix
             ↓
          Re-review
             ↓
            PASS
```

Copilot：

```text
/create-plan
/execute-plan
/review-code
/fix-review
/review-code
```

或严格 Custom Agent Adapter：

```text
/draft-plan
/run-plan
/inspect-code
/repair-review
/inspect-code
```

Codex：

```text
$create-plan
$execute-plan
$review-code
$fix-review
$review-code
```

Codex 可按角色配置 `executionMode=goal`。`.agents/scripts/goal-task.mjs` 为指定阶段输出目标文本，调用者仍需将其交给 Codex 原生 Goal 入口；该脚本不会启动 Goal。

## 路线二：Copilot / Codex + Antigravity

职责：

```text
Copilot / Codex
= Planner + Reviewer

Antigravity
= Executor + Fixer
```

闭环：

```text
Copilot / Codex
Create Plan
      ↓
Antigravity Goal
Execute
      ↓
Copilot / Codex
Review
      ↓
NEEDS_FIX
      ↓
Antigravity Goal
Fix
      ↓
Copilot / Codex
Re-review
      ↓
PASS
```

复杂 Antigravity 阶段推荐：

```powershell
node .agents/scripts/goal-task.mjs execute-plan <taskId> --harness antigravity
node .agents/scripts/goal-task.mjs fix-review <taskId> --harness antigravity
```

这些命令也只输出目标文本，需由调用者交给 Antigravity 原生 Goal 入口。运行前先读取用户级 `GOAL_RULES.md`（`$CODEX_HOME/GOAL_RULES.md`；未设置时为 `~/.codex/GOAL_RULES.md`）；只有存在 `OPEN_ACTIONABLE` 才进入自动 Fix，自动 Review/Fix 最多三轮。项目配置的 `reviewLoop.maxRounds: 5` 不得放宽该上限。

## 共享原则

无论哪条路线：

- taskId 不变；
- plan.md / execution.md / review.md 协议不变；
- MUST_FIX + SHOULD_FIX 默认进入 Review Fix；
- OPTIONAL 默认不自动修；
- 不自动 commit/push/PR；
- Planner / Reviewer 与 Executor / Fixer 可以来自不同 Harness；
- Goal 只控制自主执行深度，不改变 Skill 的任务边界。

## 实现、依赖与边界

四阶段定义位于 `.agents/skills/`；Codex、Copilot、Antigravity 的适配分别见 `.agents/skills/`、`.github/agents/`/`.github/hooks/ai-workflow.json`、`.agents/workflows/`/`.agents/hooks.json`。任务状态由 `.agents/scripts/task-state.mjs` 和 `task-finish.mjs` 管理；路线与执行策略使用 `.agents/agent-profiles.json` 和 `.agents/workflow-policy.json`。脚本要求 Node.js ES Modules；仓库未声明最低版本，也没有 npm 依赖清单。路线只约定职责和共享产物，不会自动切换 Harness 或启动另一方的 Goal。
