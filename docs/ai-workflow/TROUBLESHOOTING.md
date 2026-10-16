# 故障排查

## 先运行

```powershell
node .agents/scripts/workflow-doctor.mjs
```

## 飞书测试

Dry run：

```powershell
node .agents/scripts/notify-feishu.mjs --dry-run --workspace .
```

发送测试：

```powershell
node .agents/scripts/notify-feishu.mjs --send-test --workspace .
```

`--send-test` 会向已配置的飞书 Webhook 发送真实测试消息；`workflow-notify.mjs ... --force` 也会尝试绕过去重并发送通知。执行前确认环境文件和接收目标，不要把这两条命令当作只读诊断。

## 查看任务状态

```powershell
node .agents/scripts/task-state.mjs status
```

## task-finish 报错

检查 execution.md 前三行是否严格：

```text
AI_EXECUTION_STATUS
AI_TASK_ID
AI_EXECUTION_FINISHED_AT
```

并确认 `AI_EXECUTION_FINISHED_AT >= current-task.startedAt`。

## Antigravity Hook 不工作

当前项目必须保持：

```text
.agents/hooks.json
```

Stop command：

```text
node ./scripts/stop-guard.mjs
```

需要诊断时临时启用：

```json
"hook-diagnostic": {
  "enabled": true
}
```

然后看：

```text
.agents/runtime/hooks.log
```

诊断完成后再关闭。

## Copilot 找不到 Skill

确认：

```text
.agents/skills/execute-plan/SKILL.md
.agents/skills/fix-review/SKILL.md
```

并使用 VS Code Chat Customizations / Diagnostics 检查发现状态。

## Codex 找不到 Skill

确认从 Git 仓库内启动 Codex，并且 `.agents/skills` 位于当前目录到仓库根目录的扫描链上。

必要时重启 Codex。


## Copilot 每一步完成但不通知

先确认：

```text
.github/hooks/ai-workflow.json
```

被 VS Code 加载。

执行：

```text
/hooks
```

或者：

```text
Chat: Configure Hooks
```

并查看：

```text
Developer: Show Agent Debug Logs
```

确认 `UserPromptSubmit` / `Stop` 都有执行。

再检查：

```dotenv
AI_WORKFLOW_NOTIFY=true
FEISHU_ENABLED=true
FEISHU_WEBHOOK_URL=...
```

手工测试：

```powershell
node .agents/scripts/notify-feishu.mjs --send-test --workspace .
```

Plan 测试：

```powershell
node .agents/scripts/workflow-notify.mjs plan-created <taskId> --source copilot --force
```

Review 测试：

```powershell
node .agents/scripts/workflow-notify.mjs review-completed <taskId> --source copilot --force
```

## SHOULD_FIX 没有执行

查看：

```powershell
node .agents/scripts/task-state.mjs status
```

Review Fix 应看到：

```json
"fixScope": "recommended"
```

如果为：

```text
must
```

重新：

```powershell
node .agents/scripts/task-state.mjs review-fix <taskId> --source copilot --fix-scope recommended
```


## Codex 看不到 create-plan / review-code Skill

确认文件存在：

```text
.agents/skills/create-plan/SKILL.md
.agents/skills/review-code/SKILL.md
```

重新打开/刷新 Codex 的项目上下文后再尝试：

```text
$create-plan
$review-code
```

## Goal 配置存在但执行不符合预期

生成阶段目标文本：

```powershell
node .agents/scripts/goal-task.mjs execute-plan <taskId> --harness codex
```

此命令只向标准输出生成目标文本，不会启动 Goal 或执行任务。检查输出是否包含：

```text
Skill 路径
taskId
Profile
Model
Effort
executionMode
安全上限
停止条件
```

将目标文本手动交给 Codex 或 Antigravity 的原生 Goal 入口。开始 Goal 或 Plan 执行前先读取用户级 `GOAL_RULES.md`（`$CODEX_HOME/GOAL_RULES.md`；未设置时为 `~/.codex/GOAL_RULES.md`）。只有 `OPEN_ACTIONABLE > 0` 才允许新的自动 Fix；自动 Review/Fix 最多三轮。`.agents/workflow-policy.json` 中的 `reviewLoop.maxRounds: 5` 不覆盖更严格的用户级限制。Goal 不应绕过 Skill；如果目标文本扩大范围，以 Skill、仓库 `AGENTS.md` 和用户级规则为准。

## 实现、依赖与配置

本页涉及的状态诊断入口为 `.agents/scripts/workflow-doctor.mjs`、`.agents/scripts/task-state.mjs`；Hook 逻辑位于 `.agents/scripts/copilot-hook.mjs` 和 `.agents/scripts/stop-guard.mjs`；通知由 `.agents/scripts/workflow-notify.mjs`、`notify-feishu.mjs` 处理。配置入口包括 `.github/hooks/ai-workflow.json`、`.agents/hooks.json`、`.agents/workflow-policy.json`、`.env.local` 或 `AI_WORKFLOW_ENV_FILE` 指定的环境文件。脚本依赖 Node.js ES Modules 和内置模块，仓库未声明最低版本或 npm 依赖；通知还需要可用的全局 `fetch`、`FEISHU_WEBHOOK_URL`，以及可选 `FEISHU_SECRET`。排查命令只在确认副作用后运行；本页未要求执行任何通知或配置写入脚本。
