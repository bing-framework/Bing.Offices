# Universal Agent Workflow V4.4

本工作流为 GitHub Copilot、Google Antigravity IDE 与 OpenAI Codex 提供统一的四阶段任务约定：创建计划、执行计划、独立审查、修复审查。各 Harness 复用公共 Agent Skills、Node.js 状态脚本和任务产物格式。

## 功能与任务产物

任务文件位于 `ai_docs/tasks/<taskId>/`：

```text
plan.md       实施计划
execution.md  执行记录与终态
review.md     独立审查结论
```

Copilot 的 VS Code Hook 与 Antigravity Stop Hook 负责更新阶段状态、有限续跑、任务收口和可选通知；Codex 可直接使用公共 Skills。Goal 是原生 Harness 的执行模式，不是第五个阶段。

## 实现与依赖

| 职责 | 关键实现 |
| --- | --- |
| 四阶段规则 | `.agents/skills/{create-plan,execute-plan,review-code,fix-review}/SKILL.md` |
| 任务状态与收口 | `.agents/scripts/task-state.mjs`、`task-finish.mjs` |
| Copilot Hook | `.github/hooks/ai-workflow.json`、`.agents/scripts/copilot-hook.mjs` |
| Antigravity Hook | `.agents/hooks.json`、`.agents/scripts/stop-guard.mjs` |
| Agent Profile 与阶段策略 | `.agents/agent-profiles.json`、`.agents/workflow-policy.json` |
| 可选飞书通知 | `.agents/scripts/workflow-notify.mjs`、`notify-feishu.mjs` |

脚本使用 Node.js ES Modules 和 Node 内置模块；仓库未提供 `package.json` 或声明最低 Node.js 版本。通知使用运行时全局 `fetch`，因此运行环境需提供该 API。Profile、Hook 和阶段策略都由上述 JSON 文件配置；脚本不会自动安装 Harness，也不会替代 IDE 自身的 Goal 入口。

## 文档索引

| 文档 | 内容 |
| --- | --- |
| README.md（本页） | 工作流总览、实现入口与文档索引 |
| [QUICKSTART.md](QUICKSTART.md) | 首次初始化、四阶段入口和自检命令 |
| [ROUTES.md](ROUTES.md) | Copilot、Codex 与 Antigravity 的协作路线 |
| [CODEX.md](CODEX.md) | Codex Skills、任务状态和 Goal 目标文本 |
| [COPILOT.md](COPILOT.md) | VS Code Hook、阶段状态和通知流程 |
| [ANTIGRAVITY.md](ANTIGRAVITY.md) | Stop Hook、Agent Profile 与 Goal 使用边界 |
| [GOAL.md](GOAL.md) | Goal 目标生成、全局规则和自动停止条件 |
| [PROTOCOL.md](PROTOCOL.md) | `execution.md`、`review.md` 与修复范围协议 |
| [AGENT-PROFILES.md](AGENT-PROFILES.md) | 模型、思考等级和执行模式配置 |
| [CHINESE-COMMENTS.md](CHINESE-COMMENTS.md) | 当前仓库 `chinese-comments` 技能入口与边界 |
| [TROUBLESHOOTING.md](TROUBLESHOOTING.md) | Hook、状态、通知及技能发现排查 |

## 重要边界

开始 Goal 或 Plan 执行前先读取用户级 `GOAL_RULES.md`（`$CODEX_HOME/GOAL_RULES.md`；未设置时为 `~/.codex/GOAL_RULES.md`）。只有存在 `OPEN_ACTIONABLE` 才能开始下一轮自动修复；自动 Review/Fix 循环最多三轮。项目 JSON 中的 `reviewLoop.maxRounds: 5` 不覆盖更严格的用户级规则。

`.agents/scripts/goal-task.mjs` 只根据任务、Profile 和策略输出目标文本，不会启动 Codex/Antigravity 原生 Goal，不会执行代码或自行执法循环上限。目标文本需由操作者交给对应 Harness 的原生 Goal 入口。
