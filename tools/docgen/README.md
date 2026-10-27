# docgen · 文档自动化流水线

把仓库里的**原始产物**转成**可交付文档**。

当前工作区包含文档生成工具、文档工作流和测试项目中的示例标记；生成文档与图表由 Git 忽略，CI 会将生成结果作为 Artifact 归档。

```
TestResults/*.trx                                      ┐
TestResults/<collector-guid>/coverage.cobertura.xml   │
BenchmarkDotNet.Artifacts/                            │
output/release/*.xml                                  ├─► 采集层 ─► Document 中间模型 ─► 渲染层 ─► md / docx / xlsx / pptx / pdf
build/api-snapshot-baseline                           │
git log / tags                                        ┘
                                                    ▲
                                        治理层：gate（门禁）
                                                docs-check（防腐化）
```

完整方案设计见 [`docs/doc-automation/README.md`](../../docs/doc-automation/README.md)。

---

## 快速开始

```bash
# 一条命令（首次运行会自动建虚拟环境并装依赖）
python tools/docgen/run.py all

# 只出测试与基准报告
python tools/docgen/run.py tests --format docx,xlsx

# 只出汇报 PPT
python tools/docgen/run.py deck --format pptx

# 对外交付文档（客户抬头 + 条款页 + 签章页）
python tools/docgen/run.py delivery --format docx,pdf

# 质量门禁：注释覆盖率 / 提交规范率（未通过返回非 0，适合挂 CI）
python tools/docgen/run.py gate

# 文档腐化校验：文档示例必须来自可编译源码且逐字一致
python tools/docgen/run.py docs-check

# CI 自检模式：只采集 + 校验，不写文件
python tools/docgen/run.py all --check
```

也可以不使用 `run.py`，自行管理环境：

```powershell
python -m venv tools/docgen/.venv
tools/docgen/.venv/Scripts/python -m pip install -r tools/docgen/requirements.txt
$env:PYTHONPATH = "tools/docgen"
tools/docgen/.venv/Scripts/python -m docgen all
```

---

## 命令参考

| 参数 | 说明 |
|---|---|
| `target` | 文档：`tests` / `release` / `deck` / `api` / `delivery` / `all`；治理：`gate` / `docs-check`，可多选 |
| `-f, --format` | `md,docx,xlsx,pptx,pdf` 逗号分隔，`all` 表示全部（默认） |
| `--check` | 只做采集与文档自检，不写出文件；有异常时返回非 0，适合挂 CI 门禁 |
| `--min` | 仅 `gate` 有效：临时覆盖注释覆盖率总阈值（百分比） |

`docs-check` 需要先生成 Markdown；它会将 `tests/Bing.Offices.Docs.Tests/DocsExamples.cs` 中的 6 个可编译 `// docgen:sample` 标记与生成文档中的示例双向比较。文档 CI 工作流先构建并重新生成 Markdown，再执行此检查；它验证示例与可编译源码一致，不比较 Git 中提交的生成文档（`docs/generated/` 已被忽略）。执行 `gate` 前应准备对应版本的 Release XML 与 API 基线。

### 五类文档

| target | 产物 | 主要数据源 |
|---|---|---|
| `tests` | 测试与基准测试报告 | `.trx` + BenchmarkDotNet CSV + Cobertura + 基准历史归档 |
| `release` | 发布说明与变更日志 | `git log` + API 快照 |
| `deck` | 项目进展汇报（PPT 优先） | 上述全部 |
| `api` | 公共 API 参考 | XML 注释 + API 快照 + Docs.Tests 可编译示例 |
| `delivery` | 对外交付说明与质量证据 | 上述全部 + `build/delivery-profile.json` |

### 两个治理动作

| target | 做什么 | 未通过时 |
|---|---|---|
| `gate` | 计算 XML 注释覆盖率（分母取公共 API 快照的 public+protected 成员，剔除无参构造）与提交规范符合率，与 `build/docgen-gates.json` 中的阈值比对 | 退出码 1 |
| `docs-check` | 扫描 `tests/Bing.Offices.Docs.Tests` 中 `// docgen:sample` 标记的可编译示例，与 `docs/generated/*.md` 内嵌示例双向比对（缺失 / 不一致 / 引用不存在的命名空间） | 退出码 1 |

---

## 产物位置

| 类型 | 目录 | 是否入库 |
|---|---|---|
| Markdown | `docs/generated/` | 否（本地生成，已 gitignore） |
| DOCX / XLSX / PPTX / PDF | `artifacts/docs/` | 否（构建产物，已 gitignore） |
| Markdown 图表 | `docs/generated/assets/<文档名>/` | 否（随生成文档一起忽略） |
| 基准历史快照 | `artifacts/benchmark-history/` | 否（本地保留，按需归档） |

---

## 目录结构

```
tools/docgen/
├─ run.py                    自举运行入口（自动建 venv）
├─ requirements.txt
├─ docgen/
│  ├─ model.py               统一中间模型 Document / Section / Block
│  ├─ theme.py               品牌主题（改这里即可换配色与字体）
│  ├─ repo.py                仓库路径解析（从任意 cwd 运行都可用）
│  ├─ charts.py              matplotlib 位图图表（DOCX / Markdown 用）
│  ├─ reports.py             组装层：把采集结果编排成文档
│  ├─ gates.py               治理层：注释覆盖率 / 提交规范率门禁
│  ├─ samples.py             治理层：示例抽取 + 文档腐化双向校验
│  ├─ cli.py                 命令行入口
│  ├─ collectors/            采集层
│  │  ├─ tests.py            TRX 解析
│  │  ├─ benchmark.py        BenchmarkDotNet CSV 解析
│  │  ├─ archives.py         基准历史归档 + 趋势对比
│  │  ├─ coverage.py         Cobertura 覆盖率解析
│  │  ├─ api.py              XML 注释 + API 快照解析
│  │  └─ changelog.py        git 历史（Conventional Commits）解析
│  └─ renderers/             渲染层
│     ├─ markdown.py
│     ├─ docx.py
│     ├─ xlsx.py
│     ├─ pptx.py
│     └─ pdf.py
```

---

## 扩展指南

### 新增一种输出格式

在 `renderers/` 下实现 `render(document, out_path) -> Path`，然后在
`renderers/__init__.py` 的 `RENDERERS` 里注册。**不需要改任何采集器**。

### 新增一类文档

在 `reports.py` 里写 `build_xxx(layout) -> Document`，然后在 `BUILDERS` 里注册，
并在 `cli.py` 的 `TARGETS` 中加上名字。**不需要改任何渲染器**。

### 换品牌样式

只改 `theme.py` 里的 `Theme` 默认值（颜色、字体、字号）。渲染器里没有硬编码样式。

---

## 已知限制

- **提交规范符合率**决定变更日志质量。近 152 条提交实测 71.71%（109 条可自动归类），
  不符合规范的提交只能整条列在「其他」类，无法按范围聚合。
  本地 `gate` 可检查该指标（示例配置阈值为 70%）；提交规范 CI 不在本次交付范围。
- **代码覆盖率**：CI 的 net8.0 Unit 测试单次运行同时生成顶层 `TestResults/ci-tests-net8.trx` 与收集器 GUID 子目录中的 Cobertura XML；VSTest 还可能在 `TestResults/<run-id>/In/<machine>/coverage.cobertura.xml` 留下逐字节相同的附件副本。门禁按路径排序，要求至少有一份非空且可解析的 Cobertura XML；存在多份时，只有所有副本与首份逐字节相同才通过。docgen 也按路径排序读取首份，因此重复副本不会改变报告内容。Artifact 保留整个 `TestResults/` 树，以保全 TRX 和其相对引用的附件。本地可用 `dotnet test tests/Bing.Offices.Tests/Bing.Offices.Tests.csproj -f net8.0 --logger "trx;LogFileName=ci-tests-net8.trx" --collect:"XPlat Code Coverage" --results-directory TestResults` 生成同一布局。
- **基准趋势**：从首次生成起开始积累 `artifacts/benchmark-history/` 快照，
  第二次及以后（不同 commit）才能做方法级趋势对比。
- DOCX / Markdown 的图表是位图；需要可编辑图表请用 XLSX / PPTX 产物。
- `--check` 只校验文档模型的结构一致性，不校验排版观感。

## 阶段二新增能力

- `DOCGEN_STABLE=1` 环境变量：将 generated_at 固定为「由数据源确定」。来源提交、测试结果等仍随输入变化，不保证跨提交内容相同；生成物不参与 Git 漂移检查。
- `--check` 模式下不写入基准归档（只读自检）。
- 基准归档幂等：同一 commit 重复运行覆盖而非累积。

## 阶段三新增能力（治理）

- **PDF 输出**：`renderers/pdf.py`（reportlab）。中文用 reportlab 内置的
  Adobe CID 字体 `STSong-Light`，CI 无需安装字体文件。目录采用「多趟构建 + 页码回填」，
  页码为真实页码而非占位符。
- **对外交付模板**：`Document.delivery`（`model.Delivery`）+ DOCX 的
  客户抬头 / 交付说明页 / 条款页 / 签章页；PDF 与 Markdown 同步支持。
  客户档案在 `build/delivery-profile.json`，缺失时回落内置默认条款，流水线永远可运行。
- **注释覆盖率门禁**：`docgen gate`。分母取公共 API 快照的 public+protected 成员数
  （XML 文件只含已写注释的成员，不能当分母），剔除无法注释的无参构造函数。
  阈值在 `build/docgen-gates.json`。实测 95.20%，阈值 90/85。
- **文档腐化防护**：`docgen docs-check`。文档示例只能来自
  `tests/Bing.Offices.Docs.Tests` 中带 `// docgen:sample` 标记的可编译方法，
  由 `samples.py` 抽取后嵌入生成文档，并做双向逐字比对。
  改 API → 编译失败；改示例不重新生成 → `docs-check` 失败。
