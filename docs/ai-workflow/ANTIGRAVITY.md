# Antigravity IDE 使用说明

V4.4 使用项目四阶段 Skill、Antigravity Stop Hook 和 Goal 目标文本生成器。

## Hook cwd

当前项目已验证：

```text
processCwd=<project>/.agents
```

因此 `.agents/hooks.json` 保持：

```json
"command": "node ./scripts/stop-guard.mjs"
```

不要改成：

```text
node ./.agents/scripts/stop-guard.mjs
```

## 四阶段 Skill

```text
/create-plan <taskId>
/execute-plan <taskId>
/review-code <taskId>
/fix-review <taskId>
```

路线二默认仍建议：

```text
Copilot / Codex：Plan + Review
Antigravity：Execute + Fix
```

## Goal

复杂 Execute：

```powershell
node .agents/scripts/goal-task.mjs execute-plan <taskId> --harness antigravity
```

复杂 Review Fix：

```powershell
node .agents/scripts/goal-task.mjs fix-review <taskId> --harness antigravity
```

把输出目标交给 Antigravity 原生 Goal。

`goal-task.mjs` 只向标准输出写目标文本，不会启动原生 Goal，也不会执行计划。Goal 的职责是持续推理、执行、测试和修复到当前阶段完成；它不得改变 `plan.md` / `review.md` 的任务边界。开始 Goal 或 Plan 执行前先读取用户级 `GOAL_RULES.md`（`$CODEX_HOME/GOAL_RULES.md`；未设置时为 `~/.codex/GOAL_RULES.md`）；只有 `OPEN_ACTIONABLE > 0` 才能开始下一轮自动 Fix，自动 Review/Fix 最多三轮。`.agents/workflow-policy.json` 中 `reviewLoop.maxRounds: 5` 不覆盖更严格的用户级规则。

## Stop Guard

Stop Hook 继续负责：

1. `IN_PROGRESS + model_stop` → 有限继续；
2. 已写合法终态但遗漏 `task-finish` → 兜底收口/通知；
3. 已 `active=false` → 正常停止，不重复通知。

## Agent Profile

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced --target antigravity
```

模型写入 Workspace Custom Agent。

思考等级仍按当前 Antigravity 会话/版本能力应用。

`executionMode` 会写入生成 Agent 的说明区，用于提示是否采用 Goal。

## 实现、配置与边界

Antigravity Hook 定义位于 `.agents/hooks.json`，Stop 逻辑位于 `.agents/scripts/stop-guard.mjs`；任务状态由 `.agents/scripts/task-state.mjs` 和 `task-finish.mjs` 管理，Profile 由 `.agents/scripts/sync-agent-profiles.mjs` 写入 `.agents/agents/`。配置读取 `.agents/agent-profiles.json`、`.agents/workflow-policy.json`；命令行脚本依赖 Node.js ES Modules 和内置模块，仓库未声明最低版本或 npm 依赖。Hook 是否触发取决于 Antigravity 加载项目 Hook 的运行环境，原生 Goal 必须由用户在 Harness 中启动。
