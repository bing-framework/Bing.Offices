# Agent Runtime Profiles 使用说明

V4.4 将 **模型选择 + 思考等级策略 + 执行模式** 统一收敛到：

```text
.agents/agent-profiles.json
```

并新增初始化器：

```powershell
node .agents/scripts/init-agent-profiles.mjs
```

## 1. 推荐使用方式

第一次接入项目时，不建议直接面对整份 JSON 手工修改。

直接执行：

```powershell
node .agents/scripts/init-agent-profiles.mjs
```

初始化器会使用中文交互完成：

```text
选择默认 Profile
→ 可选读取现有 Copilot / Codex / Antigravity 模型配置
→ 可选逐角色调整 model / effort
→ 生成 agent-profiles.json
→ 可选立即同步三个 Harness
```

## 2. 非交互初始化

使用均衡模式初始化：

```powershell
node .agents/scripts/init-agent-profiles.mjs --preset balanced --yes
```

如果已有配置并确认重新生成：

```powershell
node .agents/scripts/init-agent-profiles.mjs --preset balanced --yes --force
```

从现有下游 Agent 配置读取模型/思考等级后初始化：

```powershell
node .agents/scripts/init-agent-profiles.mjs --preset balanced --from-existing --force
```

只生成，不立即同步：

```powershell
node .agents/scripts/init-agent-profiles.mjs --preset balanced --yes --force --no-sync
```

## 3. 配置职责

```text
init-agent-profiles.mjs
        ↓
首次生成 / 重新初始化
        ↓
.agents/agent-profiles.json
        ↓
后续只维护这一份
        ↓
sync-agent-profiles.mjs
        ↓
Copilot / Codex / Antigravity
```

`agent-profiles.json` 是模型与思考等级策略的 **Source of Truth**。

不要长期分别手工维护：

```text
.github/agents/*.agent.md
.codex/config.toml
.codex/agents/*.toml
.agents/agents/*/agent.md
```

否则下一次同步可能覆盖这些模型绑定配置。

## 4. 中文元数据规范

机器协议字段保持英文：

```text
fast
balanced
quality
max-quality

plan-writer
plan-executor
code-reviewer
review-fixer

copilot
codex
antigravity

model
effort
capabilities
workflowRoles
```

给人阅读的元数据统一中文：

```text
displayName
description
notes
CLI 提示
CLI 校验错误
使用说明
```

例如：

```json
{
  "balanced": {
    "displayName": "均衡模式",
    "description": "在执行质量、推理深度和模型使用成本之间取得平衡，适合作为日常开发默认配置。",
    "notes": "推荐作为默认 Profile。"
  }
}
```

## 5. 内置 Profile

| ID | 中文名称 | 用途 |
|---|---|---|
| `fast` | 快速模式 | 简单改动、低风险维护 |
| `balanced` | 均衡模式 | 日常开发默认 |
| `quality` | 质量优先 | 复杂实现、重构、关键 Review |
| `max-quality` | 最高质量 | 架构设计、复杂根因分析、关键代码审查 |

查看：

```powershell
node .agents/scripts/sync-agent-profiles.mjs --list
```

## 6. 同步

预览：

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced --dry-run
```

正式同步三边：

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced
```

只同步 Copilot：

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced --target copilot
```

只同步 Codex：

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced --target codex
```

只同步 Antigravity：

```powershell
node .agents/scripts/sync-agent-profiles.mjs balanced --target antigravity
```

## 7. 从现有配置导入

`--from-existing` 会尽力读取：

```text
Copilot
.github/agents/<role>.agent.md
→ model

Codex
.codex/agents/<role>.toml
→ model
→ model_reasoning_effort

Antigravity
.agents/agents/<role>/agent.md
→ model
→ 已生成说明中的期望思考等级
```

导入结果只覆盖当前选择的默认 Profile，不会尝试猜测其他 Profile 应该如何配置。

## 8. 平台差异

### Copilot

同步 Agent `model`。

思考等级保存在 Profile 中作为期望值；当前 IDE 仍按模型/会话能力应用。

### Codex

同步：

```toml
model = "..."
model_reasoning_effort = "high"
```

模型和思考等级都可按角色原生应用。

### Antigravity

同步 Custom Agent `model`。

Profile 中保留期望 `effort`，但不会写入未经当前版本确认的 Agent frontmatter 字段。

## 9. 两条工作路线

Profile 与工作路线是两个维度：

```text
Route
= 哪个 Harness 负责哪个阶段

Profile
= 每个 Role 在某个 Harness 下使用什么模型/effort
```

因此同一个 `balanced` / `quality` 可以用于：

```text
路线一：Copilot / Codex 自闭环
```

也可以用于：

```text
路线二：Copilot / Codex Plan + Review，Antigravity Execute + Fix
```

详见 `ROUTES.md`。


## V4.4 executionMode

角色配置新增：

```json
{
  "model": "...",
  "effort": "high",
  "executionMode": "goal"
}
```

推荐语义：

```text
copilot
  agent

codex
  goal

antigravity
  plan-writer/code-reviewer = interactive
  plan-executor/review-fixer = goal
```

`sync-agent-profiles.mjs` 会把 executionMode 写入运行时报告；Codex TOML 使用注释记录，不写入未经当前环境确认的额外配置字段。

Goal 阶段目标文本由：

```powershell
node .agents/scripts/goal-task.mjs ...
```

该脚本只输出目标文本，不会启动原生 Goal、执行 Skill 或创建任务产物；调用者需将文本提交给对应 Harness。开始 Goal 或 Plan 执行前需读取用户级 `GOAL_RULES.md`；自动 Fix 仅在 `OPEN_ACTIONABLE > 0` 时允许，自动 Review/Fix 最多三轮。项目策略 `reviewLoop.maxRounds: 5` 不覆盖该限制。

## 实现、依赖与限制

配置结构见 `.agents/agent-profiles.json`，初始化由 `.agents/scripts/init-agent-profiles.mjs` 负责，同步由 `.agents/scripts/sync-agent-profiles.mjs` 生成 `.github/agents/`、`.codex/` 和 `.agents/agents/` 的 Harness 文件，并写运行时报告到 `.agents/runtime-profiles/`。Profile 和阶段安全策略分别配置在 `.agents/agent-profiles.json`、`.agents/workflow-policy.json`。脚本使用 Node.js ES Modules 和内置模块；仓库未声明最低 Node.js 版本或 npm 依赖。模型名和思考等级必须由目标 Harness 实际支持；Profile 中的 `effort` 可能只是期望值，`executionMode` 也不会自行启动 Goal。
