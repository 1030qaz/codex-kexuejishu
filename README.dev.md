# 开发维护说明

本项目适合用 Git 在两台电脑之间同步维护。仓库建议只提交源码、维护说明和必要的小型配置；抓取数据、缓存、Word 文档和真实 Cookie 不进入 Git。

## 项目类型

- 类型：Python 脚本 + 文档生成。
- 主要依赖：`requests[socks]`、`python-docx`。
- 常用验证：`python -m py_compile (Get-ChildItem -Filter *.py).FullName`。
- 常用更新入口：`update_current_month.py`，用于全用户增量抓取并只重建当月文档。

## 第一台电脑：初始化、提交、推送

```powershell
cd D:\科学技术打头阵
git status --short --branch
python -m py_compile (Get-ChildItem -Filter *.py).FullName
git add .gitignore AGENTS.md README.dev.md .env.example *.py post_analysis_schema.json 科学技术打头阵_逐条发言分析框架.md
git status --short
git commit -m "Initialize project maintenance workflow"
```

如果已有远程仓库地址，先设置远程：

```powershell
git remote add origin <REMOTE_URL>
git remote -v
```

推送前请先确认远程地址和将要提交的文件无敏感信息：

```powershell
git push -u origin master
```

如远程默认分支使用 `main`，可改为：

```powershell
git branch -M main
git push -u origin main
```

## 第二台电脑：clone 和继续开发

```powershell
cd D:\
git clone <REMOTE_URL> 科学技术打头阵
cd D:\科学技术打头阵
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m py_compile (Get-ChildItem -Filter *.py).FullName
```

如需抓取论坛内容，在本机临时设置环境变量，不要写入仓库：

```powershell
$env:NGA_COOKIE = "粘贴自己的 NGA 登录 Cookie"
python .\update_current_month.py
```

## 每次开工

```powershell
cd D:\科学技术打头阵
git status --short --branch
git pull --rebase
git switch -c feature/简短任务名
```

如果只是很小的文档或维护改动，也可以在当前功能分支继续；不建议长期直接在 `master` 或 `main` 上开发。

## 每次收工

```powershell
python -m py_compile (Get-ChildItem -Filter *.py).FullName
git status --short
git add <本次真正需要提交的文件>
git diff --cached --stat
git commit -m "简短说明本次变更"
git push -u origin 当前分支名
```

`selected_users_posts.json` 是共享的已抓取发言快照，`stock_name_cache.json` 是个股简称缓存，默认可以提交，用于避免另一台电脑重新抓取。生成的 `monthly_docs/`、`wolf_posts.json`、`market_cache_*.json`、`*.docx` 默认不提交。需要在两台电脑共享生成文档时，优先用网盘或手动复制；确实要纳入 Git 前先确认文件大小和隐私风险。

## 冲突处理原则

- 开工前先 `git pull --rebase`，减少冲突。
- 冲突时优先保留双方对源码和说明文档的有意修改。
- 不用 `git reset --hard` 或强推覆盖另一台电脑的工作，除非明确确认。
- 对生成文件冲突，通常不要解决并提交；重新生成或保留本机版本即可。
- 冲突解决后运行 `python -m py_compile (Get-ChildItem -Filter *.py).FullName`，再提交。

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
