# Universal Agent Workflow V4.4 快速开始

脚本依赖 Node.js ES Modules 与 Node 内置模块；仓库没有 npm 依赖清单，也未声明最低 Node.js 版本。运行前确认当前目录位于包含 `.agents` 的项目中。

## 1. 初始化 Agent Profile

首次：

```powershell
node .agents/scripts/init-agent-profiles.mjs
```

后续同步：

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced
```

## 2. 四个公共 Skill

```text
create-plan
execute-plan
review-code
fix-review
```

### Copilot

```text
/create-plan
/execute-plan
/review-code
/fix-review
```

严格 Custom Agent Adapter：

```text
/draft-plan
/run-plan
/inspect-code
/repair-review
```

### Codex

```text
$create-plan
$execute-plan
$review-code
$fix-review
```

### Antigravity

```text
/create-plan
/execute-plan
/review-code
/fix-review
```

## 3. Goal 目标文本

Codex：

```powershell
node .agents/scripts/goal-task.mjs create-plan <taskId> --harness codex
node .agents/scripts/goal-task.mjs execute-plan <taskId> --harness codex
node .agents/scripts/goal-task.mjs review-code <taskId> --harness codex
node .agents/scripts/goal-task.mjs fix-review <taskId> --harness codex
```

以上命令只把 Goal 指令输出到标准输出；不会启动 Codex 原生 Goal，也不会自动执行阶段任务。检查输出后，将目标文本交给 Codex 或 Antigravity 的原生 Goal 入口。目标文本读取 `.agents/agent-profiles.json` 与 `.agents/workflow-policy.json`；可用 `--profile <name>` 选择 Profile，`--json` 输出 JSON。

Antigravity：

```powershell
node .agents/scripts/goal-task.mjs execute-plan <taskId> --harness antigravity
node .agents/scripts/goal-task.mjs fix-review <taskId> --harness antigravity
```

## 4. Review Fix 默认范围

```text
fixScope=recommended

MUST_FIX      ✅
SHOULD_FIX    ✅
OPTIONAL      ⏭
```

`fixScope` 决定纳入哪些审查级别，不代表每个 Finding 都可自动修复。开始 Goal 或 Plan 执行前，先读取用户级 `GOAL_RULES.md`（`$CODEX_HOME/GOAL_RULES.md`；未设置时读取 `~/.codex/GOAL_RULES.md`）。只有 `OPEN_ACTIONABLE > 0` 才允许开始下一轮自动修复；自动 Review/Fix 最多三轮。项目 `.agents/workflow-policy.json` 中的 `reviewLoop.maxRounds: 5` 不覆盖该规则。

## 5. 两条路线

路线一：

```text
Copilot / Codex
Plan → Execute → Review → Fix → Re-review
```

路线二：

```text
Copilot / Codex：Plan + Review
Antigravity：Execute + Fix
```

详情：

```text
docs/ai-workflow/ROUTES.md
docs/ai-workflow/GOAL.md
```

## 6. 自检

```powershell
node .agents/scripts/workflow-doctor.mjs
```

四阶段规则位于 `.agents/skills/`；状态命令由 `task-state.mjs` 和 `task-finish.mjs` 实现；Agent 配置由 `init-agent-profiles.mjs`、`sync-agent-profiles.mjs` 维护。初始化或同步命令会写入配置文件及对应 Harness 文件，确认目标范围后再运行。
