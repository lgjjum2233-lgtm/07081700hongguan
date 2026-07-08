# GitHub 使用建议

## 初始化仓库

当前环境暂未检测到 `git` 命令。安装 Git 后，可在项目目录执行：

```powershell
git init
git add .
git commit -m "Initial macro analysis project structure"
```

## 推送到 GitHub

创建 GitHub 仓库后：

```powershell
git remote add origin <your-repo-url>
git branch -M main
git push -u origin main
```

## Excel 文件管理

本项目包含 `.xlsx` 二进制文件。当前文件体积不大，可以直接提交到 GitHub。

如果后续文件明显变大，建议使用 Git LFS：

```powershell
git lfs install
git lfs track "*.xlsx"
git add .gitattributes
```

## 提交建议

建议提交信息格式：

```text
docs: add project handoff notes
data: update daily workbook 20260527 V1
scripts: add macro summary matching workflow
log: record 20260527 update
```

## 不建议提交

- WPS 临时锁文件：`~$*.xlsx`
- 临时脚本和探索文件：`codex_tmp/`
- 截图缓存：`*.png`
- 本地环境文件：`.env`

