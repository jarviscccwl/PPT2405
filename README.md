# ppt2405 安装教程

ppt2405 是用于制作中文文献汇报 PPT 的 Codex Skill，配套蓝白学术模板已包含在安装目录中。安装后，提供论文 PDF 或文字大纲即可调用。

## 1. 安装 Skill

### 方法一：让 Codex 从 GitHub 安装

在 Codex 中发送：

```text
使用 $skill-installer 安装这个 Skill：
https://github.com/jarviscccwl/PPT2405/tree/main/ppt2405
```

安装完成后，在下一轮消息中调用 `$ppt2405`。如果已安装同名旧版，先将旧版移出 skills 目录备份，再安装新版；不要将新旧文件混合。

### 方法二：手动安装

1. 在本仓库点击 **Code → Download ZIP**，然后解压。
2. 将里面完整的 `ppt2405` 文件夹复制到 Codex 的 skills 目录。默认位置：

   - Windows：`%USERPROFILE%\.codex\skills\ppt2405`
   - macOS / Linux：`~/.codex/skills/ppt2405`
   - 若设置过 `CODEX_HOME`：使用该目录下的 `skills/ppt2405`。

3. 确认 `SKILL.md` 直接位于 `skills/ppt2405/` 内，避免多套一层文件夹。
4. 在 Codex 下一轮消息中输入 `$ppt2405`；若当前客户端尚未刷新技能列表，新建会话后使用。

安装后目录应包含：

```text
ppt2405/
├── SKILL.md
├── requirements.txt
├── agents/
├── assets/
│   ├── ppt2405.pptx
│   ├── ppt2405.potx
│   └── ppt2405-preview.pdf
├── references/
└── scripts/
```

`references` 和 `scripts` 是 Skill 的运行资源，需要与模板一起保留。

## 2. 准备运行环境

推荐在 **Windows + WPS 或 Microsoft PowerPoint** 下使用，以保持模板字体和原生公式的效果。

- Python 3.10 或更高版本。
- Python 依赖：Pillow、pdfplumber、python-pptx、lxml。
- Poppler 的 `pdftoppm`，用于整页渲染及完整图表提取。
- 字体：微软雅黑、Times New Roman、Cambria Math。包内不分发字体文件。

Codex 桌面版若提供配套运行时，Skill 会先定位已有 Python 和工具；没有运行时查询工具时，会检查 PATH 和已有运行时目录。以下命令中的 `python` 应替换为实际可用的 Python 路径。可先验证依赖：

```bash
python -c "import PIL, pdfplumber, pptx, lxml; print('Dependencies OK')"
```

需要自行安装 Python 依赖时，在 `ppt2405` 目录运行：

```bash
python -m pip install -r requirements.txt
```

确认 `pdftoppm` 可用：

```bash
pdftoppm -v
```

如果 Poppler 没有加入 PATH，可将环境变量 `PPT2405_POPPLER` 指向 `pdftoppm` 可执行文件或其所在目录。macOS / Linux 使用者还需在组版计划的 `font_files` 中指定有权使用的字体文件，配置格式见 Skill 的组合接口说明。

## 3. 调用

提供论文 PDF，并发送：

```text
使用 $ppt2405，根据这篇论文制作中文文献汇报 PPT。
沿用内置模板，重点文字用红色加粗，保留完整图表和可编辑公式。
请交付 PPTX 和 PDF 预览。
```

也可以直接给出大纲：

```text
使用 $ppt2405，按照下面的大纲制作汇报 PPT。
请检索相关真实材料和图片，记录出处，再按配套模板适配排版。
大纲：……
```

素材检索需要当前 Codex 环境可访问的网络检索工具；AI 配图需要可用的图像生成工具。Skill 本身不附带这些服务，也不需要把 API 密钥写入安装目录。

## 4. 使用配套模板

自动制作时，Skill 使用 `assets/ppt2405.pptx` 和对应版式目录。手动制作时，用 WPS / PowerPoint 打开 `assets/ppt2405.potx`；`assets/ppt2405-preview.pdf` 可浏览 26 类页面样式。

制作新报告时输出到独立项目目录，保留安装目录中的原模板。
