# `chinese-comments` 多 CLI 技能任务总结

- Task ID：`BO-SKILL-CHINESE-COMMENTS-20260914-001`
- 原计划日期：2026-09-14
- 原计划实施状态：`COMPLETED`
- 文档复核日期：2026-10-14

## 目标与原验收范围

基于 C#/.NET 中文 XML 文档注释规范，交付 Codex、Gemini CLI、GitHub Copilot CLI 和 Claude Code 可使用的 Agent Skill。技能须保持统一规则，区分补全和审查、明确处理范围，并覆盖泛型、异步返回值、字段常量、`inheritdoc`、XML 转义、UTF-8、生成文件排除和不得编造业务语义等要求。

原计划还包含跨平台项目/用户安装器、冲突保护与备份行为、使用说明和独立样例验证；本任务未安装到个人目录，不修改生产 API、CI 或依赖。

## 原执行结果

执行报告将规范源、Claude 同步入口、Node 安装器、安装器自检、内容契约自检和使用文档均记为完成。记录的验证包括：

- 安装器契约自检、内容契约自检和项目级 `--check` 通过。
- Node 语法检查和 `git diff --check` 通过。
- C# 样例覆盖接口实现继承注释、泛型异步返回值、字段常量、XML 特殊字符和生成文件排除。
- 四种 CLI 的真实认证会话当时不可用，因此未进行真实交互验证；该限制不记为平台实测通过。
- 原执行文档记录了安装器 ACL 路径调整及脚本放置位置，相关工具后来随规范源迁移到源仓库。

## 后续来源调整与当前状态

完成原计划后，技能规范源改由 `awesome-copilot` 维护；当前项目只消费安装生成的副本，使用说明位于 `docs/ai-workflow/CHINESE-COMMENTS.md`。本次只读核对确认：`.agents/skills/chinese-comments/SKILL.md` 与 `.claude/skills/chinese-comments/SKILL.md` 相同，但两者与源仓库 `skills/dotnet/chinese-comments/SKILL.md` 的 SHA-256 不同。

因此，原技能建设任务已完成并归并；当前未闭合的是从源仓库重新安装/同步副本。此项是后续来源更新，不回写为原计划的验收结果。本次未修改技能源、项目副本或安装文档。
