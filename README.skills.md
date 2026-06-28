# Skill 同步说明

本文件说明本项目如何同步 Codex skill，方便第二台电脑 clone 后继续使用同一套项目分析规则。

## 当前结论

- 项目专用 skill：`wolf-perspective`
- 项目专用同步位置：`.agents/skills/wolf-perspective/`
- 用户个人通用 skill：保留在本机 `~/.agents/skills/` 或 `~/.codex/skills/`，不放进本项目仓库。
- 系统内置 skill：不复制、不提交，例如 `~/.codex/skills/.system/`。
- 当前未发现 `CODEX_HOME` 环境变量。

## 项目专用 Skill

### wolf-perspective

路径：

```text
.agents/skills/wolf-perspective/SKILL.md
```

用途：

- 用于论坛发言分析、每日总结、板块轮动、交易纪律和复盘学习。
- 强调只做学习和复盘，不做荐股。
- 分析时优先结合项目本地数据和东方财富接口，再读取论坛发言文本。

检查结果：

- 已包含 `SKILL.md`。
- `SKILL.md` 已包含 `name` 和 `description`。
- 当前没有 `scripts/`、`references/`、`assets/` 依赖目录。
- 未发现 Cookie、token、密钥、账号密码、证书或本地私有路径。

## 用户级 Skill 清单

以下 skill 位于本机用户目录，建议作为个人环境保留，不建议直接放入本项目仓库：

```text
algorithmic-art
andrej-karpathy-perspective
brainstorming
brand-guidelines
canvas-design
claude-api
dispatching-parallel-agents
doc-coauthoring
docx
elon-musk-perspective
executing-plans
feynman-perspective
find-skills
finishing-a-development-branch
frontend-design
huashu-nuwa
ilya-sutskever-perspective
internal-comms
mcp-builder
mrbeast-perspective
munger-perspective
nature-academic-search
nature-citation
nature-data
nature-figure
nature-paper-to-patent
nature-paper2ppt
nature-polishing
nature-reader
nature-response
nature-reviewer
nature-writing
naval-perspective
openclaw-medical-skills
patent-disclosure-skill
paul-graham-perspective
pdf
pptx
receiving-code-review
requesting-code-review
research-paper-writing
skill-creator
slack-gif-creator
steve-jobs-perspective
subagent-driven-development
sun-yuchen-perspective
systematic-debugging
taleb-perspective
template-skill
test-driven-development
theme-factory
trump-perspective
using-git-worktrees
using-superpowers
verification-before-completion
web-artifacts-builder
webapp-testing
writing-plans
writing-skills
x-mastery-mentor
xlsx
zhang-yiming-perspective
zhangxuefeng-perspective
```

同步建议：

- 如果第二台电脑也想使用这些个人通用 skill，建议单独建立一个私有 dotfiles 或 personal-skills 仓库。
- 复制 skill 时要复制整个目录，不能只复制 `SKILL.md`。如果目录里有 `scripts/`、`references/`、`assets/`，也要一起保留。
- 不要把 API key、Cookie、token、证书、真实 `.env`、本地绝对路径或私有服务器地址放进 GitHub。

## 第二台电脑启用步骤

1. 拉取项目到你自己的项目目录：

```powershell
git clone https://github.com/1030qaz/codex-kexuejishu.git <你的项目目录>
cd <你的项目目录>
```

2. 确认项目 skill 已存在：

```powershell
Test-Path .agents\skills\wolf-perspective\SKILL.md
```

3. 在 Codex 中打开本项目目录，开始新对话或重启当前项目会话，让 Codex 重新读取项目上下文。

4. 如果 Codex 没有自动识别项目内 skill，可以临时复制到用户级目录：

```powershell
New-Item -ItemType Directory -Force "$HOME\.agents\skills" | Out-Null
Copy-Item -Recurse -Force ".agents\skills\wolf-perspective" "$HOME\.agents\skills\wolf-perspective"
```

5. 后续更新：

```powershell
git pull --rebase
```

## 提交前检查

提交 skill 相关改动前，建议执行：

```powershell
git status --short
git diff -- .agents/skills README.skills.md AGENTS.md
rg -n "ngaPassport|ngacn0com|HMACCOUNT|Cookie\s*[:=]|sk-[A-Za-z0-9]{20,}|AKIA[0-9A-Z]{16}|-----BEGIN (RSA |OPENSSH |EC )?PRIVATE KEY-----" .agents README.skills.md AGENTS.md
```

如果敏感信息扫描有命中，先暂停，不要提交，改成占位符或改为环境变量读取。
