# `chinese-comments` 技能

本仓库只使用 `chinese-comments`，不维护技能规范源。规范源和安装器位于：

```text
H:\Working\ME\Git\awesome-copilot\skills\dotnet\chinese-comments\SKILL.md
H:\Working\ME\Git\awesome-copilot\scripts\install-chinese-comments-skill.mjs
```

`.agents/skills/chinese-comments/SKILL.md` 和 `.claude/skills/chinese-comments/SKILL.md` 是安装生成的使用副本。规则变更只在 `awesome-copilot` 更新，再从规范源重新安装；不要在本仓库手工维护另一份规则。

## 安装到当前项目

从技能源仓库执行：

```powershell
Set-Location 'H:\Working\ME\Git\awesome-copilot'
node .\scripts\install-chinese-comments-skill.mjs --agent all --scope project --target 'H:\Bing_Framework\Bing.Offices'
```

脚本会把 Codex、Gemini CLI 和 Copilot CLI 使用的入口放在 `.agents/skills`，把 Claude Code 使用的入口放在 `.claude/skills`。同一个共享入口只写入一次。

## 安装到其他项目

从技能源仓库运行，并传入其他目标项目根目录：

```powershell
Set-Location 'H:\Working\ME\Git\awesome-copilot'
node .\scripts\install-chinese-comments-skill.mjs --agent all --scope project --target 'D:\Projects\Example App'
```

用户级安装会写入当前用户目录下的共享 Agent Skills 目录，适用于之后打开的多个项目：

```powershell
node .\scripts\install-chinese-comments-skill.mjs --agent all --scope user
```

实际写入前可查看计划，或只检查是否一致：

```powershell
node .\scripts\install-chinese-comments-skill.mjs --agent all --scope project --target 'H:\Bing_Framework\Bing.Offices' --dry-run
node .\scripts\install-chinese-comments-skill.mjs --agent all --scope project --target 'H:\Bing_Framework\Bing.Offices' --check
```

目标文件内容不同时，脚本默认失败并保留原文件；确认需要替换时使用 `--force`，原文件会先备份为同目录下的 `SKILL.md.bak.<时间戳>`。不要求 Windows 符号链接权限，也不会覆盖目录、符号链接或其他类型的目标。

## 四种 CLI 中使用

安装后重新启动 CLI，或使用对应的刷新命令：

- Codex：运行 `/skills` 或在提示中使用 `$chinese-comments`。
- Gemini CLI：运行 `/skills list`，需要时运行 `/skills reload`，也可使用 `gemini skills list`。
- GitHub Copilot CLI：运行 `/skills reload`，使用 `/skills info chinese-comments` 检查。
- Claude Code：运行 `/skills` 查看，或直接使用 `/chinese-comments`。

也可以直接提出任务，例如：“请使用 `chinese-comments` 检查当前改动中的 C# XML 注释”。自动发现由各 CLI 根据技能描述决定；显式调用可以避免触发判断差异。

规则依据各 CLI 的 Agent Skills 入口：

- [Codex Skills](https://learn.chatgpt.com/docs/build-skills)
- [Gemini CLI Agent Skills](https://geminicli.com/docs/cli/skills/)
- [GitHub Copilot CLI Agent Skills](https://docs.github.com/en/copilot/how-tos/copilot-cli/customize-copilot/add-skills)
- [Claude Code Skills](https://code.claude.com/docs/en/skills)

## 边界

补全模式只修改注释和任务明确要求的文档；审查模式只报告问题。未指定路径时优先处理当前 Git 改动中的 C# 文件；没有可确定范围时先询问，不默认扫描整个仓库。`bin`、`obj`、生成代码、迁移、代理代码和第三方源码不在处理范围内。

## 交付与维护状态

原任务 `BO-SKILL-CHINESE-COMMENTS-20260914-001` 的计划日期为 2026-09-14，原计划实施状态为 `COMPLETED`。任务总结于 2026-10-14 复核，并于 2026-10-15 合入本说明；本说明承接原总结的交付结论、验证限制和后续事项。

原交付范围是统一的 C#/.NET 中文 XML 注释规范及四种 CLI 入口、跨平台项目/用户安装器、冲突保护与备份、使用说明和独立样例验证。规范覆盖泛型、异步返回值、字段常量、`inheritdoc`、XML 转义、UTF-8、生成文件排除及不得编造业务语义，并区分补全与审查模式及其处理范围。该任务未安装到个人目录，未修改生产 API、CI 或依赖。

### 原验证记录与限制

原执行报告将规范源、Claude 同步入口、Node 安装器、安装器自检、内容契约自检和使用文档记为完成，记录了以下验证：

- 安装器契约自检、内容契约自检和项目级 `--check` 通过。
- Node 语法检查和 `git diff --check` 通过。
- C# 样例覆盖接口实现继承注释、泛型异步返回值、字段常量、XML 特殊字符和生成文件排除。
- 四种 CLI 的真实认证会话当时不可用，未进行真实交互验证；不将格式或契约自检表述为四平台实测通过。

上述结果是原执行记录，本次文档归并没有重新运行这些验证。原执行文档还记录了安装器 ACL 路径调整及脚本放置位置；相关工具后来随规范源迁移到 `awesome-copilot`，当前维护入口见本文开头。

### 待确认的副本同步

2026-10-14 复核及 2026-10-15 只读核验均确认：本项目 `.agents/skills/chinese-comments/SKILL.md` 与 `.claude/skills/chinese-comments/SKILL.md` 内容相同，但与 `awesome-copilot` 中的规范源内容不同。原技能建设任务已完成，后续仍需确认是否从源仓库重新安装或同步；该事项不回写为原计划的验收结果。

本次仅归并说明文档，没有修改规范源、项目技能副本或安装脚本，也没有执行安装或同步。
