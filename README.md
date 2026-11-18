# PPT2405 安装教程

PPT2405 用蓝白学术模板制作可编辑 PPT。可以上传文献、粘贴大纲或笔记，也可以提供大纲和多篇文献，让 Codex 按章节补查资料、提取原图或生成示意图。

## 安装

在 Codex 中发送以下指令：

```text
使用 $skill-installer 安装 https://github.com/jarviscccwl/PPT2405/tree/v2.0.0/ppt2405
```

也可以下载 [v2.0.0 压缩包](https://github.com/jarviscccwl/PPT2405/archive/refs/tags/v2.0.0.zip)，解压后把完整的 `ppt2405` 文件夹复制到 `$CODEX_HOME/skills/`。默认位置是 Windows 的 `%USERPROFILE%\.codex\skills\` 或 macOS/Linux 的 `~/.codex/skills/`。已有旧版时先移出备份。

确认安装后路径为 `skills/ppt2405/SKILL.md`，同级保留 `assets/`、`scripts/`、`references/`、`agents/` 和 `requirements.txt`。新建会话后输入 `$ppt2405` 调用。

## 环境

使用 Python 3.10+。可优先使用 Codex 配套运行时；自行配置时，在安装后的 `ppt2405` 目录执行：

```bash
python -m pip install -r requirements.txt
pdftoppm -v
```

PDF 提图依赖 Poppler。若 `pdftoppm` 不在 PATH，将 `PPT2405_POPPLER` 指向可执行文件或其目录。推荐用 Windows 上的 PowerPoint 或 WPS 检查成品；模板使用微软雅黑、Times New Roman 和 Cambria Math。其他平台可按[组合接口](ppt2405/references/composition-api.md)配置字体。联网检索和 AI 生图使用当前会话可用的工具。

## 调用示例

- **论文**：「使用 $ppt2405 汇报这篇论文。讲清基本模型、两个扩展模型和实测误差，保留原图、可编辑公式及结果表。」
- **纯大纲**：「使用 $ppt2405 做交通检测入门汇报。按检测原理、观测量、适用场景三节组织，检索官方手册，并生成一张地感线圈示意图。」
- **混合材料**：「使用 $ppt2405 按我的大纲组织这三篇文献。缺少依据的地方补查一手资料，把各页出处放入备注。」

模板在 `ppt2405/assets/`。可打开 PPTX 或用 POTX 新建文件，复制所需样例页，在普通视图直接编辑；同目录 PDF 可预览 26 种版式。
