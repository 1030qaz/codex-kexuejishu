# AGENTS.md

本项目是 Python 脚本与生成文档项目，用于抓取、整理和分析 NGA 论坛帖子内容。Codex 在本仓库工作时遵守以下长期规则。

## 工作前

- 每次修改前先执行 `git status --short --branch`，确认当前分支和未提交改动。
- 每次开始任务前建议执行 `git pull --rebase`，除非当前没有远程仓库或用户明确不希望联网同步。
- 先阅读相关脚本和已有输出约定，再做小范围修改。
- 不要大规模重构，不要改动与当前任务无关的脚本、文档或数据。

## 修改规则

- 优先修改源码、配置、说明文档和维护脚本。
- 不修改自动生成文件，除非用户明确要求重新生成文档或数据。
- 不提交密钥、Cookie、token、证书、账号密码、真实 `.env`、私有配置。
- 不提交二进制大文件、缓存、构建产物、生成的 Word/Markdown 文档。
- `selected_users_posts.json` 是跨电脑共享的已抓取发言快照，允许提交；`stock_name_cache.json` 是个股简称缓存，允许提交。
- 对 FPGA、嵌入式或硬件工程，如果后续加入本仓库，优先修改源码、约束、脚本、文档，谨慎修改 IDE 工程文件。

## 验证规则

- 修改 Python 脚本后至少运行：

```powershell
python -m py_compile (Get-ChildItem -Filter *.py).FullName
```

- 如修改抓取或生成流程，在不暴露 Cookie 的前提下，按任务需要运行对应脚本。
- 如果需要联网抓取，使用环境变量 `NGA_COOKIE`，不要把 Cookie 写入文件或提交。

## 提交前

- 提交前展示变更摘要和 `git status --short`。
- 确认 `.gitignore` 已排除生成物、缓存和敏感配置。
- 建议按功能拆分提交，提交信息说明“为什么改”和“改了什么”。

## Skill 使用规则

- 项目专用 skill 放在 `.agents/skills/`，随仓库同步；当前项目专用 skill 是 `.agents/skills/wolf-perspective/`。
- 不要把用户级通用 skill、系统内置 skill 或第三方临时 skill 直接复制进本仓库，除非用户明确要求。
- 新增或更新 skill 时，必须保留完整目录；如果有 `scripts/`、`references/`、`assets/`，要随 `SKILL.md` 一并保留。
- 每个 skill 必须包含 `SKILL.md`，且 frontmatter 中必须有 `name` 和 `description`。
- 提交前检查 skill 中是否包含 Cookie、token、API key、证书、账号密码、真实 `.env`、本地绝对路径或私有服务器地址。
- `wolf-perspective` 只用于论坛发言、每日总结、板块轮动和交易纪律的学习复盘，不输出确定性荐股结论。

## 日常开发短 Prompt

开始一个新任务前，请先执行 `git status` 和 `git pull --rebase`，确认当前分支和远程同步状态。然后基于最新代码完成以下任务：

```text
【任务描述】
<在这里写你的具体开发任务>
```

要求：
1. 修改范围尽量小；
2. 不做无关重构；
3. 不提交密钥等敏感内容；
4. 修改后运行必要测试或构建；
5. 最后给出变更摘要、测试结果、建议 commit message；
6. 等我确认后再 commit/push。
