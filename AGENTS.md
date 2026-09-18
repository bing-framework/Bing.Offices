# Agent 全局执行规范

本文件用于约束 Copilot、Codex、Gemini、Claude 等 AI 工具在当前项目中的执行方式。

## 默认行为

- 默认使用简体中文回答。
- 生成代码时，优先补充必要的中文注释。
- 所有命令、脚本、文件读写、日志导出、Markdown 生成，都必须优先考虑 Windows + VS Code + PowerShell 环境下的中文乱码问题。
- 默认编码统一使用 UTF-8。

## Office/CSV 测试与发布门槛

- 修改 Excel/CSV 公共接口、默认实现、Provider 分支、映射、缓存键、文件提交或异步 IO 时，必须同步增加职责级直接测试。
- 每个受影响的默认实现必须有独立测试；禁止只依赖综合测试间接覆盖。
- CSV/Excel 输出测试必须断言完整结果、文件内容或结构化错误，不得只断言片段。
- 修改映射缓存键、计划缓存或配置合并时，必须覆盖隔离、命中、未命中、重复渲染和失败后的对象状态。
- NPOI Provider 的 XLS/HSSF 与 XLSX/XSSF 分支都必须有成功和失败场景；未知实现必须固化实际异常类型。
- 异步 API 必须使用真实异步 IO；生产代码禁止用 `Task.Run`、`.Result` 或 `.Wait()` 伪装异步。
- Excel DOM 操作的并发限制、文件提交和临时文件清理必须有资源矩阵或真实文件证据；取消不得损坏已有目标文件。
- API 变化必须更新双 TFM snapshot、成员级追溯和迁移记录；未经成员级审批不得删除公共成员。
- 提交前必须维护“最终生产符号 -> 测试方法”的可追溯映射；映射应包含关键行为、测试项目和方法名。

## UTF-8 规则（强制）

- 所有文本文件读取必须显式指定 `UTF-8`。
- 所有文本文件写入必须显式指定 `UTF-8`。
- 禁止使用 PowerShell 默认编码直接写入源码、Markdown、XML、Gradle、Kotlin、Java、YAML、JSON、Properties 文件。
- 优先使用 Python 进行文件读写：

```python
from pathlib import Path

text = Path("input.txt").read_text(encoding="utf-8")
Path("output.txt").write_text(text, encoding="utf-8", newline="\n")
```

- 在 shell 命令中修改中文内容时，避免直接内联中文大段文本；优先使用 Python 组装字符串后以 UTF-8 写入。
- 终端出现 `????` 时，不要直接判断文件已损坏；应优先用编辑器或 `unicode_escape` 检查文件真实内容。

## PowerShell 编码规则

在生成 PowerShell 脚本时，必须优先设置控制台编码：

```powershell
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
```

写入文件时必须显式指定编码：

```powershell
Set-Content -Path $path -Value $content -Encoding utf8
Add-Content -Path $path -Value $content -Encoding utf8
Out-File -FilePath $path -Encoding utf8
Export-Csv -Path $path -Encoding utf8 -NoTypeInformation
```

禁止在未确认编码的情况下，直接使用：

```powershell
"中文内容" > file.md
"中文内容" >> file.md
```

## Python 编码规则

生成 Python 脚本时，所有文本文件读写必须显式指定编码：

```python
from pathlib import Path

content = Path("input.txt").read_text(encoding="utf-8")
Path("output.txt").write_text(content, encoding="utf-8", newline="\n")
```

禁止使用不带 `encoding` 的文件读写：

```python
open("input.txt").read()
open("output.txt", "w").write("内容")
```

## .NET / C# 编码规则

生成 .NET / C# 文件读写代码时，必须显式指定 UTF-8：

```csharp
using System.Text;

var text = File.ReadAllText(path, Encoding.UTF8);
File.WriteAllText(path, text, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
```

控制台程序如需输出中文，优先设置：

```csharp
Console.InputEncoding = Encoding.UTF8;
Console.OutputEncoding = Encoding.UTF8;
```

## Node.js 编码规则

生成 Node.js / TypeScript 文件读写代码时，必须显式指定 `utf8`：

```ts
import { readFileSync, writeFileSync } from "node:fs";

const text = readFileSync("input.txt", "utf8");
writeFileSync("output.txt", text, "utf8");
```

## 文件类型规则

以下文件必须按 UTF-8 处理：

- `.cs`
- `.csproj`
- `.sln`
- `.props`
- `.targets`
- `.json`
- `.yaml`
- `.yml`
- `.xml`
- `.md`
- `.sql`
- `.ps1`
- `.bat`
- `.cmd`
- `.js`
- `.ts`
- `.vue`
- `.java`
- `.kt`
- `.gradle`
- `.properties`

## 乱码排查规则

如果用户反馈乱码，优先按以下顺序排查：

1. VS Code 当前文件编码是否为 UTF-8。
2. PowerShell Profile 是否设置 UTF-8。
3. `[Console]::OutputEncoding` 是否为 UTF-8。
4. `$OutputEncoding` 是否为 UTF-8。
5. `chcp` 是否为 `65001`。
6. 文件写入命令是否显式指定 `-Encoding utf8`。
7. Python / Node.js / .NET 是否显式指定 UTF-8。
8. 是否由旧文件本身就是 GBK / ANSI 编码导致。

## 禁止行为

- 禁止默认依赖 Windows ANSI / GBK 编码。
- 禁止在未确认编码时批量替换中文内容。
- 禁止使用未指定编码的文件迁移脚本。
- 禁止生成可能导致中文变成 `????` 的命令。
- 禁止把终端显示乱码直接等同于文件内容损坏。

## 推荐处理方式

当需要批量修改文件内容时，优先生成 Python 脚本，并使用：

```python
from pathlib import Path

path = Path("target-file.md")
text = path.read_text(encoding="utf-8")
text = text.replace("旧内容", "新内容")
path.write_text(text, encoding="utf-8", newline="\n")
```

当需要验证文件真实内容时，可以使用：

```python
from pathlib import Path

text = Path("target-file.md").read_text(encoding="utf-8")
print(text.encode("unicode_escape").decode("ascii"))
```

# Text Encoding and Line Ending Rules

## 统一规则与现有例外

- 文本源码和配置默认使用 **UTF-8 without BOM + LF**，文件末尾保留一个换行，不产生 Mixed Line Endings。
- 保留现有 `.editorconfig` 契约：`.cs`、`.csx`、`.vb`、`.vbx` 使用 UTF-8 BOM；`.csproj`、`.sln`、`.ps1` 使用 CRLF。
- `.bat`、`.cmd` 使用 UTF-8 + CRLF，供 Windows Command Processor 使用。其他源码和配置默认 LF。
- UTF-8 文件类型还包括 `.jsx`、`.tsx`、`.kts`、`.html`、`.css`、`.scss`、`.less`、`.sh`、`.py`、`.toml`、`.ini`。二进制文件不得按文本处理。
- 本节补充前文：所有写入示例须同时遵守目标文件的 BOM 和 EOL 例外；示例中的无 BOM / LF 是默认值，不覆盖例外。
- 修改前依次读取 `.editorconfig`、`.gitattributes`，再检查目标文件真实字节、BOM 和行尾。规则冲突时先报告，不擅自变更已有契约。

## 编辑器与 Git

- Visual Studio、Rider 使用根级 `.editorconfig`；VS Code 须启用支持 EditorConfig 的扩展，保存时遵守文件匹配规则，不用用户级格式化设置覆盖仓库约定。
- Git 使用 `.gitattributes` 控制文本行尾；Git 不负责把 GBK 转为 UTF-8，也不自动校验 BOM。显式 `eol` 规则优先于 `core.autocrlf`，无需修改用户全局 Git 配置。
- 不运行全仓库格式化，不自动执行 `git add --renormalize .`，不自动 Commit、Push 或创建分支。
- 存量编码或行尾偏差单独报告；全仓库转换须由用户批准下一阶段 `Repository Line Ending Normalization`。

## 程序化读写

- 所有程序化文本读写必须显式指定 UTF-8，禁止依赖 ANSI、GBK、系统 Code Page 或语言默认编码。
- 优先精确 Patch，避免读取整文件后为了几行修改重写整文件。确需脚本转换时，仅处理明确批准的目标，并显式控制编码、BOM 和行尾。
- Python 默认写入使用 `Path.write_text(text, encoding="utf-8", newline="\n")`（Python 3.10+）；使用 `open()` 时明确 `encoding="utf-8", newline="\n"`。BOM 例外使用 `utf-8-sig`；CRLF 例外对逻辑 LF 内容使用 `newline="\r\n"`，避免重复转换。
- PowerShell 文本命令前设置本文件规定的三项 UTF-8 控制台编码。`-Encoding utf8` 在不同版本中 BOM 行为不同，且不会保证 LF；精确写入优先使用 Python，或显式 `.NET UTF8Encoding(false)` 与已确定行尾的内容。禁止默认重定向覆盖源码。
- .NET 默认写入使用 `new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)`；BOM 例外显式使用 `true`。跨平台源码生成默认使用 `"\n"`，不得依赖 `Environment.NewLine`，除非目标明确要求随系统变化。指定编码并不会自动统一内容行尾。
- Node.js / TypeScript 读取和写入明确使用 `"utf8"`；生成源码内容默认使用 LF，无意转换 CRLF 或增删 BOM 均禁止。

## 最小 Diff 与验证

- 不为了局部修改改变整文件编码、BOM、行尾或无关空白；若目标存量格式不符，先报告或单独批准转换，不混入业务修改。
- 每次修改后执行 `git diff --check`、`git diff --stat`；必要时执行 `git diff --ignore-space-at-eol`（Git 实际选项）和 `git ls-files --eol`。
- 修改文件必须进行严格 UTF-8 字节检查，核对 BOM、LF/CRLF 例外、末尾换行和无 Mixed Line Endings；不能仅以编译成功为依据。
- 若小改动显示为整文件变化，立即检查 Encoding、BOM、EOL、Formatter 和 Git autocrlf，修复自身引入的无关变化，不还原他人的既有改动。

## 乱码排查顺序

1. 检查实际文件字节、严格解码结果及 BOM；再检查 VS Code 当前编码、`.editorconfig`、`.gitattributes`、Git `core.autocrlf`。
2. 检查 PowerShell 版本、Profile、`[Console]::InputEncoding`、`[Console]::OutputEncoding`、`$OutputEncoding`，必要时检查 `chcp`。
3. 检查写入程序与 Python / Node.js / .NET 是否显式指定 UTF-8，再验证旧文件是否确实为 GBK / ANSI。
4. 区分 Terminal Rendering、Console Encoding、File Encoding 与 Actual File Bytes。终端 `????` 不等于文件损坏；用前文 `unicode_escape` 示例或字节检查验证真实内容。
