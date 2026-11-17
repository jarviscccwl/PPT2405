# 完整论文图表：整页渲染、区域提取、视觉验收

图表以论文的 **figure/table整体** 为单位：保留所有分面、坐标轴、刻度、共享与独立图例、子图标记及图注追溯。一个figure可能由许多位图、矢量线、文字和裁切路径组成；PDF的image对象不是“完整图片”。例如纯矢量曲线页可以没有任何嵌入位图。不要用`extract_image`之类的对象提取结果冒充完整figure。

`scripts/paper_figures.py`提供三个可复用命令：`inspect`、`render`、`extract`。它不自动判断figure边界，不自动判断完整性，不自动接受视觉检查。适用于有可渲染页面的PDF；扫描PDF可渲染，但图注可能没有可提取文字。

## 运行环境

若提供`load_workspace_dependencies`，先用它获取当前设备的Python和工具位置；否则从PATH和已有运行时目录定位，并验证依赖可导入。使用带`pdfplumber`、`Pillow`的Python；渲染需要Poppler的`pdftoppm`。脚本不下载、不安装依赖，也不保存本机用户目录。

Poppler查找顺序：命令参数`--pdftoppm`，环境变量`PPT2405_POPPLER`，`PATH`，当前Python附近的标准runtime相邻目录。参数/环境变量可以是可执行文件或其bin目录。明确指定的路径不存在时直接失败，不静默换另一版本。

PowerShell调用示例（把运行时与skill路径换成实际已找到的路径）：

```powershell
$figurePython = '实际运行时/python.exe'
$figureHelper = '实际skill目录/scripts/paper_figures.py'
$figurePoppler = '实际Poppler目录/bin'
& $figurePython $figureHelper inspect '论文/研究论文.pdf' --output 'work/inspection.json'
& $figurePython $figureHelper render '论文/研究论文.pdf' --pages '3,8-9' --dpi 220 --pdftoppm $figurePoppler --out-dir 'work/page-review'
& $figurePython $figureHelper extract --plan 'work/figures-plan.json' --dpi 300 --pdftoppm $figurePoppler --out-dir 'work/figures-v1'
```

空格和中文路径通过参数传递，保持引用；不把PDF路径拼成shell命令。每次选择新的输出文件/目录，已有输出不会被覆盖。失败退出码为2并输出诊断；先修正路径、依赖或plan，不能把失败当成空图继续排版。

## 1. Inspect：候选调查

```text
paper_figures.py inspect PAPER.pdf --output inspection.json [--pages 1,3-5]
```

不指定`--pages`时调查全部页。输出包含源文件SHA256、总页数、每页尺寸/旋转/MediaBox/CropBox、字符/词/位图/线/曲线/矩形数量，以及带文字坐标的图注候选。图形对象统计是**整页数量**，不能当成单个figure的数量。

图注候选兼容`Fig.1`、`Fig. 1`、`Figure 1`、`Table I`、`图1`、`表1`等形式，但可能误识别正文引用；双栏抽取也可能把两栏文字合并。候选的`candidate_only: true`不是最终图注判断，更不是figure包络框。回到整页确认图号、图注、分面及相邻正文关系。

## 2. Render：先看完整页

```text
paper_figures.py render PAPER.pdf --pages 3,8-9 --dpi 220 --out-dir page-review
```

输出整页PNG、便于查看的整页预览及`render-manifest.json`。使用PDF **MediaBox**，没有自动去页边；画面按PDF旋转后的显示方向计算。如果渲染尺寸与PDF尺寸不一致，脚本停止，不在未知坐标系上继续裁切。

视觉确认以下内容后再决定区域：

- 图注所述面板和实际子图数量是否一致；是否有(a)/(b)、上下行、左右列、共享图例。
- 轴名、刻度、色条、单位、外置图例、连线和标注是否伸出主要绘图区。
- 同页多个figure、正文与表格的边界在哪里；跨栏/跨页内容是否归属同一figure。

候选区域可以由图注位置、栏界和空白带辅助提出，但不能只从最近的位图位置推断。通常在最外可见内容外留4–8 pt安全边；紧邻正文时逐图调整，不按固定页比例盲裁。

## 3. Extract：使用明确的JSON plan

```json
{
  "source_pdf": "../论文/研究论文.pdf",
  "figures": [
    {
      "id": "method-figure-3",
      "page": 3,
      "printed_page": "3",
      "figure_number": "Fig. 3",
      "bbox_pt": [45, 49, 568, 262],
      "caption_bbox_pt": [48, 197, 564, 260],
      "expected_panels": [
        "输入观察与输出动作",
        "(a) General formulation",
        "(b) CNN-based",
        "(c) Transformer-based"
      ]
    }
  ]
}
```

这些坐标仅为格式示例，不能复制到另一篇论文。字段约定：

| 字段 | 含义 |
|---|---|
| `source_pdf` | 顶层默认PDF；每条figure可覆盖，支持同一plan混合多篇论文。相对路径以plan文件目录为基准，JSON中Windows路径宜用`/`。 |
| `id` | 唯一文件标识，1–100个ASCII字母/数字/`.`/`-`/`_`，开头须为字母或数字。大小写不同视为同名；不允许末尾`.`或Windows设备名。目录和PDF文件名仍可为中文。 |
| `page` | **PDF文件页码，1起算**；不能用论文印刷页码替代。 |
| `printed_page` | 可选印刷页号，仅追溯记录，不控制PDF选页。 |
| `figure_number` | 必填图号/表号；单面板也填写。 |
| `bbox_pt` | 必填`[left, top, right, bottom]`，单位pt，1 pt=1/72英寸，原点为旋转后MediaBox显示页的左上角；不是PDF底部原点，也不是像素。 |
| `caption_bbox_pt` | 可选图注文字框，同一坐标系；完整图注区域能在整图框内时一并保留。 |
| `caption_text` | 不提供图注框时可存完整图注转录；扫描页需逐字核对原页。 |
| `expected_panels` | 必填非空描述列表。单面板写一个描述；列表须据原页和图注形成，不能凭习惯猜测。 |

所有bbox必须有限数值、正面积且完全位于该页内。倒置、零面积、负值、越界、NaN及非法页码均失败；脚本先验证整个plan，再创建输出。若图注框不在整图框内，仍保留图注转录与坐标，但`caption_included_in_crop`为false，视觉验收时确认追溯充分。数学上下标可能被文字抽取拆行；渲染图注仍是核对依据。

裁切前将每个相关页完整渲染一次，再根据实际图像宽高按轴映射坐标、向外取整裁切，保留所有矢量与文字渲染内容。脚本不使用嵌入位图导出，不生成或补画缺失数据。

每个figure单独置于`figures/<id>/`，原页存于`pages/`，避免图号与预览文件名互相覆盖。每个figure输出：

- 完整分辨率裁切PNG和裁切预览。
- 未修改的整页PNG，以及标红裁切框的整页预览。
- `manifest.json`：PDF路径与SHA256、PDF页码/印刷页号、图号、点/像素坐标、DPI、预期面板、图注、图片SHA256、页面对象统计与视觉检查状态。

四侧2 px边带的墨迹数只提示边缘是否可能截断；**零值不能证明完整**，因为整块漏掉的子图不会触碰裁切边缘。此检查也可能被正常页边内容触发，仍需视觉判断。

## 4. 必须对照整页与输出后才接受

每次提取的manifest固定输出`verified: false`、`visual_review.status: pending`。即使plan写了`verified`或视觉review字段，工具也不会继承它们。

打开**完整原页**及**最新裁切图**，逐项核对：所有预期面板；坐标轴/刻度；共享及独立图例；标签与注释；图号及完整图注追溯；最外侧内容无裁断。不存在坐标轴等项目时，在review备注注明不适用及原因，不能把“没看”记成“通过”。

只有真实完成上述查看后，才由负责检查的agent/人员在manifest或单独review记录中填写检查结果、检查者、时间、已查看文件及其SHA256，随后将该条状态设为`verified`。整份清单只有所有figure通过时才可整体标记verified。工具自身不会执行这一步。修改bbox、DPI或源文件后重新提取，验收重新回到pending。

不确定边界时保留更大区域或整页进入待审，不静默裁掉分面。跨页figure保持同一图号分组和明确续图关系，分别记录页码/区域，不无说明拼接。读取/裁切论文原图属于忠实素材处理；AI生成只能用于明确标注的解释性示意，不能补画论文数据、轴、图例或缺失分面。

## 放入PPT

存档版可包含完整图形和原文图注。展示版通常只包含**完整图形区域**：全部面板、坐标轴、刻度、图例、颜色条、面板标题、图内结果表和科学注释。连续的正文图注可与图形分开，完整文本与 `caption_bbox_pt` 留在manifest/备注，页面放短中文读图提示，出处放对应页备注。这不等于裁掉图的科学内容。不要把长英文图注作为必须上屏的图片组成，从而缩小真正需要讲解的图；原图署名或许可标记不作为额外来源页脚裁除。

展示版使用第二份明确坐标的提取计划，保存原图/存档版关联、`graphic_bbox_pt` 和独立审核记录。重新对照原页后才能标为通过；不能从缩略图凭比例裁去底部。若图注与图形重叠或边界不清，则保留并另选更大版式。

使用等比例`contain`适配，不能`cover`填框。复杂图太小则单独放大到整页；若需要讲解局部，保留完整总图，另页明确标为“局部放大”并保留对应图号/子图标签。不要为塞入版式删掉图例或轴。

提取通过只说明已查看图表的素材准备完成，不代表论文理解或整份报告已经通过检查。每次正式PPT仍需按实际幻灯片尺寸检查图片及文字可读性。
