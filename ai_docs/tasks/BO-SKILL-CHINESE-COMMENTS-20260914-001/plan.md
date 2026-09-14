# `chinese-comments` 四种 CLI 通用技能计划

## 摘要

基于现有 `.github/skills/chinese-comments/SKILL.md`，制作一份可维护、可跨项目安装的技能，支持 Codex CLI、Gemini CLI、GitHub Copilot CLI 和 Claude Code。

按用户选择，交付仓库技能和安装脚本；本次实施不自动安装到个人目录。

## 实施内容

- 保留 `chinese-comments` 现有 C#/.NET 中文 XML 注释规则，并补充工作范围、补全与审查边界、UTF-8、XML 转义、参数一致性和不编造语义要求。
- 将规范源迁移到 `.agents/skills/chinese-comments/SKILL.md`，提供 Claude 的 `.claude/skills/chinese-comments/SKILL.md` 项目级入口，避免维护不同规则。
- 新增无依赖 Node.js 安装脚本，支持 `--agent codex|gemini|copilot|claude|all`、`--scope project|user`、`--target`、`--dry-run`、`--check` 和 `--force`；默认当前项目、全部平台。
- 提供安装与使用说明，包含项目级/用户级安装、重复安装、冲突备份和四种 CLI 的刷新/发现方式。

## 验收

- 校验 frontmatter、UTF-8、技能名称、链接及共享规则一致性。
- 在临时目录验证中文和空格路径、项目级/模拟用户级安装、重复安装、冲突、备份、dry-run 和 check。
- 使用独立 C# 样例验证继承注释、泛型异步返回值、字段常量、XML 特殊字符、生成文件排除和只审查边界。
- 对本机可用 CLI 进行最小发现验证；缺失或未认证的平台只记录未实测。
- 完成 `git diff --check`，不修改生产 API、CI 或此前 API 快照修复。

## 默认决策

- 使用支持 Agent Skills 的 CLI 版本，不新增旧式 commands、hooks 或插件包装。
- 保留自动发现并支持显式调用，四个平台共享同一套规则，仅安装路径不同。
