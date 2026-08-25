# Universal Agent Workflow V4.1 快速开始

## 1. V4.1 两个关键变化

默认 Review Fix：

```text
fixScope = recommended
```

即：

```text
MUST_FIX       ✅ 修
SHOULD_FIX     ✅ 修
OPTIONAL       ⏭ 默认不修
```

Copilot 增加确定性 Workspace Hooks：

```text
.github/hooks/ai-workflow.json
```

用于：

```text
UserPromptSubmit → 识别当前阶段
Stop             → 通知 / Finalize / 未完成保护
```

## 2. 安装

把压缩包内容合并到项目根目录。

保留真实：

```text
.agents/.env.local
```

建议：

```dotenv
AI_WORKFLOW_NOTIFY=true
FEISHU_ENABLED=true
FEISHU_MESSAGE_MODE=card
```

## 3. 自检

```bash
node .agents/scripts/workflow-doctor.mjs
```

Copilot Hook 可以通过 VS Code：

```text
Chat: Configure Hooks
```

或：

```text
/hooks
```

确认加载。

## 4. Copilot

```text
/create-plan
```

完成后：

```text
📋 Plan Created
```

执行：

```text
/execute-plan
```

或者严格自定义 Agent：

```text
/run-plan
```

完成后：

```text
✅ Implementation Completed
```

Review：

```text
/review-plan
```

完成后飞书会显示：

```text
PASS
PASS_WITH_ISSUES
NEEDS_FIX
BLOCKED
```

以及：

```text
MUST_FIX 数量
SHOULD_FIX 数量
OPTIONAL 数量
```

如果 NEEDS_FIX：

```text
/fix-review
```

或：

```text
/repair-review
```

默认：

```text
MUST_FIX + SHOULD_FIX
```

修完再次：

```text
/review-plan
```

## 5. 自定义修复范围

只修 MUST_FIX：

```bash
node .agents/scripts/task-state.mjs review-fix <taskId> --source copilot --fix-scope must
```

默认：

```bash
--fix-scope recommended
```

全部连 OPTIONAL 都修：

```bash
--fix-scope all
```

## 6. Antigravity

仍使用：

```text
/execute-plan <taskId>
/fix-review <taskId>
```

Antigravity 的 `.agents/hooks.json` 保持原 schema。

不要把 `.github/hooks/ai-workflow.json` 复制成 Antigravity Hook。

## 7. Codex

```text
$execute-plan
$fix-review
```

fix-review 默认同样是：

```text
recommended
```
