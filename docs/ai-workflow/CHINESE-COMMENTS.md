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
