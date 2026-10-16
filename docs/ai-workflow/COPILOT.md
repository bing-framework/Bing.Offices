# Copilot 使用说明（V4.4）

## 1. Copilot Hook

新增：

```text
.github/hooks/ai-workflow.json
```

VS Code Workspace Hook 使用：

```json
{
  "hooks": {
    "UserPromptSubmit": [...],
    "Stop": [...]
  }
}
```

它和 Antigravity `.agents/hooks.json` 是两套不同的 Harness Adapter。

## 2. UserPromptSubmit

当提交：

```text
/draft-plan 或 /create-plan
/inspect-code 或 /review-code
/run-plan
/execute-plan
/repair-review
/fix-review
```

Hook 会写入：

```text
.agents/runtime/copilot-stage.json
```

用于记录当前阶段。

## 3. Stop

Copilot Stop Hook 会：

### 计划

检测新的 `plan.md`，调用：

```powershell
workflow-notify.mjs plan-created
```

### 执行与修复

如果 runtime 仍：

```text
active=true
execution.md=IN_PROGRESS
```

第一次 Stop 会返回：

```text
decision=block
```

要求 Agent 继续。

如果 execution.md 已合法终态：

```text
task-finish.mjs
```

按通知配置执行兜底收口和飞书通知。

### 审查

检测新的 `review.md`，调用：

```powershell
workflow-notify.mjs review-completed
```

## 4. 防止无限循环

VS Code Stop Hook 会检查：

```text
stop_hook_active
```

如果已经因为 Stop Hook 续跑过一次，则不会再次无限 block。

## 5. 通知幂等

`workflow-notify.mjs` 使用：

```text
.agents/runtime/notification-ledger.json
```

去重。

因此 Prompt 显式通知和 Stop Hook 兜底同时存在，也不会重复发同一阶段消息。

## 6. 调试 Hook

VS Code 可使用：

```text
Developer: Show Agent Debug Logs
```

查看 Hook 输入/输出。

也可：

```text
/hooks
```

检查 Hook 是否被发现。

## 7. 修复范围

默认：

```text
fixScope=recommended
```

修：

```text
MUST_FIX + SHOULD_FIX
```

不修：

```text
OPTIONAL
```


## V4.2 Agent Profile

Copilot 自定义 Agent 的 `model:` 由：

```powershell
node .agents/scripts/sync-agent-profiles.mjs <profile> --target copilot
```

统一维护。

不要再逐个 Agent 手工改 model。

Copilot effort 当前按 `remembered` 模式记录期望值，见：

```text
.agents/runtime-profiles/copilot.json
```


## V4.4 公共四阶段 Skill

Copilot 也可以直接使用：

```text
/create-plan
/execute-plan
/review-code
/fix-review
```

严格 Custom Agent Prompt Adapter：

```text
/draft-plan
/run-plan
/inspect-code
/repair-review
```

之所以把 Prompt Adapter 改名，是为了避免与公共 Skill 的 Slash Command 同名冲突。

## 实现、配置与边界

Hook 入口由 `.github/hooks/ai-workflow.json` 注册，逻辑位于 `.agents/scripts/copilot-hook.mjs`；阶段通知和任务收口分别使用 `workflow-notify.mjs`、`task-finish.mjs`。Agent Profile 来自 `.agents/agent-profiles.json`，同步目标是 `.github/agents/`；可选通知读取 `.env.local` 或 `AI_WORKFLOW_ENV_FILE` 指定的环境文件，`AI_WORKFLOW_NOTIFY` 优先，`STOP_GUARD_NOTIFY` 仅作兼容回退。脚本使用 Node.js ES Modules 与内置模块，仓库未声明最低版本或 npm 依赖。

此适配要求 VS Code 实际加载 Workspace Hooks；未加载时不会触发阶段检测。Stop Hook 只对仍处于 `IN_PROGRESS` 的活动任务请求有限续跑，重复 Stop 受 `stop_hook_active` 限制。Copilot Profile 的思考等级是记录的期望值，由 IDE/会话应用；本适配不启动原生 Goal。
