# 通用 Agent 工作流

公共四阶段 Skill：

- 创建计划：`.agents/skills/create-plan/SKILL.md`
- 执行计划：`.agents/skills/execute-plan/SKILL.md`
- 独立审查：`.agents/skills/review-code/SKILL.md`
- 修复审查问题：`.agents/skills/fix-review/SKILL.md`

默认任务文件：

```text
ai_docs/tasks/<taskId>/plan.md
ai_docs/tasks/<taskId>/execution.md
ai_docs/tasks/<taskId>/review.md
```

在支持 Goal / 目标推理的运行环境中：

```bash
node .agents/scripts/goal-task.mjs <create-plan|execute-plan|review-code|fix-review> <taskId> --harness <codex|antigravity|copilot>
```

Goal 只负责持续推理/执行到阶段完成，不得绕过 Skill 的职责、边界或 Git 安全规则。

任务执行/修复开始：

```text
node .agents/scripts/task-state.mjs ...
```

execution.md 写入合法终态后：

```text
node .agents/scripts/task-finish.mjs <taskId>
```

不要自动 git commit、git push 或创建 PR。
不要修改 plan.md/review.md 伪造执行或验收状态。

执行自动循环前读取用户级 `GOAL_RULES.md`（优先 `$CODEX_HOME/GOAL_RULES.md`，未设置时为 `~/.codex/GOAL_RULES.md`）。只有 `OPEN_ACTIONABLE > 0` 才进入下一轮修复；`PARTIAL`、`BLOCKED` 或发布状态未通过不能单独触发重试。仓库脚本生成目标，不代替当前运行环境实际启动或执行目标。
