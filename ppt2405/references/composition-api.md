# ppt2405 确定性模板填充 API

`scripts/compose_deck.py` 复制所选模板页，填充原生对象，保留母版、布局、校标与默认样式。它不进行论文理解或自动选图，也不把实际渲染验收替换为程序检查。

## 运行

依赖 Python 3.10+、Pillow、实际可用字体文件，无 COM、PowerPoint 自动化或网络依赖。

```text
python scripts/compose_deck.py --plan plan.json --output report.pptx --report report.audit.json
```

默认模板为 `assets/ppt2405.pptx`，目录为 `references/layout-catalog.json`，均相对 skill 根目录解析。可用 `--template`、`--catalog` 覆盖。应使用 Codex 返回的 bundled Python 绝对路径，避免 Windows 的 Python 应用商店占位命令。已有输出默认拒绝覆盖，显式 `--overwrite` 才替换已生成文件。源模板始终不改。

计划中的 `template_sha256` 必填，必须等于本次已验收模板的 SHA-256。建议同时填写 `catalog_sha256`。模板变化后重新读取目录、检查版式并更新计划绑定，不能绕过哈希检查。

## JSON 计划

```json
{
  "version": 1,
  "template_sha256": "替换为实际64位SHA-256",
  "catalog_sha256": "可选：目录文件的实际SHA-256",
  "mode": "template",
  "source": {"paper": "论文名", "doi": "如有"},
  "slides": [
    {
      "layout": "06",
      "text": {
        "title": "2.1 整体方法架构",
        "stage-explanation-1": {"text": "输入观测数据\n构建特征表示", "align": "center", "vertical_align": "middle"}
      },
      "images": {
        "framework": {
          "path": "figures/figure-3-complete.png",
          "source": {"paper": "论文名", "figure": "Fig. 3", "page": 4},
          "alt": "论文完整方法架构图",
          "fit": "contain"
        }
      },
      "notes": ["来源：论文名，Fig. 3，第4页。", "汇报时解释三个阶段的依赖关系。"]
    },
    {
      "layout": "06",
      "text": {"title": "2.2 第二种方法配置"}
    }
  ]
}
```

上例用于解释接口，包含未填写的模板内容，只能作为 `template` 模式草稿。正式报告使用 `strict`：任何未填的方括号占位、图框、数学示例、原生示例表或 AI 样例图都会失败，并写入报告。原本无占位含义的固定标题、面板名称和学校资产可以保留。

`slides` 的数组顺序就是最终放映顺序。`layout` 是目录中的两位 `id`，可重复选择同一个版式。每次复制具有独立幻灯片、备注、关系文件和对象创建标识。旧模板页不会出现在最终顺序；内部未引用 parts 保留以避免误删共享资源，所以文件体积不会按所选页数线性减小。

每页提供 `notes` 时，用其内容替换模板讲稿备注，再追加该页来源记录。`strict` 模式未提供 `notes` 时默认清空模板讲稿，避免正式报告带入模板使用说明；应主动填写真实讲稿。`template` 模式未提供 `notes` 时保留模板原备注。图片/表格的 `source`、每页 `source` 以及旧来源槽赋值会自动追加到对应页备注，不写到放映画面；素材清单仍需保留。

顶层仅支持 `version/template_sha256/catalog_sha256/mode/font_files/slides/source`。每页仅支持 `layout/text/math/images/tables/notes/source/omit`。拼错字段或 slot 明确报错。

若某个可选内容槽本次不用，明确填写例如 `"omit": ["assumptions"]`，使用该版式真实存在的名称。脚本清空该槽内容及它自身的填充/边框，保留占位匹配；原生表格、独立 AI 图片则从当前页移除。省略项记录到 `report.omitted`，`strict` 接受这种明确处理。写空字符串仍视为未填，同一槽不能既 omit 又赋值。该操作不删除母版/布局上的固定框线或固定标签；若整个区域不合用途，应选另一版式或调整模板后重新验收。

`source-*` 或 catalog 标为 `deliveryRole: "notes-source"` 的来源槽自动清空隐藏，无需填充或逐一 `omit`。兼容旧计划在 `text` 中赋给这些槽的来源文字会转入备注，不参与正文适配。不要将实际内容填入来源槽。最新版的 `metric-direction`（版式15）和 `comparison-conditions`（版式19）分别用于可见指标方向和比较条件，应按当前内容填写；它们不是来源槽。此规则不移除图片内部标签、原图署名、AI 标识或封面书目信息。

## 找到正确 slot

读取目录中对应版式的 `slots`。键可用短名，例如 `equation-main`，或完整名称 `ppt2405.08.equation-main`。脚本检查模板中恰有一个同名对象，不靠猜测文本内容或写死占位索引。

图片占位只通过 `images` 填充；文本与公式通过 `text` 或 `math` 填充。一个对象不能在同一页被重复指定。

原生表的稳定名称：

| 版式 | 表名 |
| --- | --- |
| 15 | `results-table` |
| 16 | `error-table` |
| 18 | `ablation-table` |

AI 概念图位于版式 25，名称 `ai-image`。兼容选择器 `table-1`、`image-1` 按该源页原生对象的 XML 顺序计数；优先使用稳定名称。

## 原生文字与实际字体测量

字符串表示普通文字，`\n` 表示分段。对象形式可指定样式，但字号只能位于 catalog 的 `minFontSizePt` 与 `maxFontSizePt` 范围内；范围按槽位用途设定，以 catalog 为准。不会低于该范围来强行塞入内容。

```json
{
  "text": {
    "principle-readout": {
      "text": "第一条解释\n第二条解释",
      "font_size_pt": 18,
      "min_font_size_pt": 18,
      "align": "justify",
      "vertical_align": "top"
    },
    "variable-definition-1": {
      "paragraphs": [
        [
          {"math": {"sub": ["x", "t"]}},
          {"text": " 表示时刻 "},
          {"math": "t"},
          {"text": " 的输入量"}
        ]
      ]
    }
  }
}
```

`paragraphs` 是“段落数组 → run 数组”。每个 run 只能含 `text` 或 `math` 之一。普通 `text` run 支持 `font/bold/color`，外层支持 `font/font_size_pt/min_font_size_pt/bold/color/align/vertical_align`；局部 run 可覆盖外层字体/颜色样式。对齐作用于整个文本对象，包括行内数学，不能写在 run 内。

`align` 可为 `left/center/right/justify`，`vertical_align` 可为 `top/middle/bottom`；分别写入原生段落对齐与文本框锚点。未指定时保留模板，helper 不根据字数猜测对齐。`justify` 是两端对齐，末行保持自然，不使用逐字分散对齐。短标签常用 `center/middle`，宽框长中文可用 `justify/top`；实际选择依 [版式规则](style-and-layouts.md) 并检查渲染，窄框/英文拉出大空隙时改 `left`。表格单格不支持这两个字段，沿用表格模板对齐。

要只强调短语，拆分为普通 run 与深红加粗 run，不给整段统一设置强调样式：

```json
{
  "curve-conclusion": {
    "paragraphs": [[
      {"text": "在相同设置下，"},
      {"text": "关键差异", "bold": true, "color": "C00000"},
      {"text": "应结合评价条件解释。"}
    ]]
  }
}
```

`color` 接受六位十六进制，可带 `#`。普通 run 的粗体参与真实字体测量；写入时按 DrawingML 规定排列描边、填充、字体等子元素，避免 Office 忽略颜色。行内 `math` run 的颜色可以继承或指定，数学字体保持 Cambria Math、当前 helper 固定为常规数学字重，不能用 `bold` 声称生成了粗体向量。数学结构与普通文字强调分别核验。

测量使用本地真实字体的字形宽度与墨迹高度，并计算文本框内边距。普通文字明确写入固定点值行距 `字号 × 1.16`（例如 18pt → 20.88pt），段前、段后均为 0；不继承 Office 的百分比行距和段距。较高的结构公式或行内数学会增大该固定行距，测量和实际 OOXML 使用完全相同的点值。中文允许逐字换行，英文尽量按词换行；公式作为完整行内单元测量。若在最低字号仍超出边界，构建失败并给出 `overflow` 和拆页建议，不裁字、不开启隐藏自动缩小。

Windows 默认读取微软雅黑、微软雅黑粗体、Times New Roman 和 Cambria Math 的系统字体文件。字体不存在则失败；不会静默替换。可在计划顶层明确指定本地字体：

```json
{
  "font_files": {
    "微软雅黑": {"regular": "fonts/msyh.ttc", "bold": "fonts/msyhbd.ttc", "index": 0},
    "Cambria Math": {"regular": "fonts/cambria.ttc", "index": 1}
  }
}
```

图片和字体的相对路径均以 `plan.json` 所在目录为基准。字体文件应由有权使用该字体的环境提供，skill 不打包系统字体。

## 原生公式

公式以 Office OMML 写入，可在 PowerPoint 中编辑。`math` 填充独立公式框；`text.paragraphs` 的 `math` run 用于变量释义中的行内数学。独立 `math` 对象接受 `expr/font_size_pt/min_font_size_pt/align/vertical_align`；对齐取值与普通文本相同，默认保留模板，改变位置不改变数学结构。

```json
{
  "math": {
    "equation-main": {
      "expr": [
        {"hat": "y"}, "=",
        {"frac": [[{"sub": ["x", "1"]}, "+", {"sub": ["x", "2"]}], "2"]}
      ]
    },
    "variable-symbol-1": {"expr": {"sub": ["x", "t"]}}
  }
}
```

支持以下结构；数组表示顺序连接，字符串/数字为字面 Unicode 数学内容：

| 结构 | 含义 |
| --- | --- |
| `{"sub":[base,index]}` | 下标 |
| `{"sup":[base,power]}` | 上标 |
| `{"subsup":[base,index,power]}` | 同时上下标 |
| `{"frac":[numerator,denominator]}` | 分式 |
| `{"sqrt":expr}` | 根式 |
| `{"hat":expr}` | 帽号 |
| `{"text":"m/s"}` | 正体文字、单位 |
| `{"delim":{"left":"(","right":")","body":expr}}` | 成对定界符 |
| `{"nary":{"symbol":"∑","sub":"i=1","sup":"n","body":expr}}` | 求和等大型运算符；上下限可省略 |

这不是 LaTeX 解析器。含反斜杠的 LaTeX 字符串会明确失败；不支持的算子也会失败。矩阵、分段多行公式等暂不支持；需要可编辑表达式时，应扩展 AST 或构建对应原生 OMML 并验证，无法完成时明确报告限制。原文公式图片只能作为有来源的辅助证据，不能替代用户要求的可编辑公式。结构公式的尺寸是保守估计，最终必须查看渲染结果。

## 完整图片与原生表格

`images` 仅接受单帧 PNG/JPEG。保留输入图像完整像素，不生成 `srcRect` 裁切参数，按 `min(slot宽/图宽, slot高/图高)` 等比例居中。`cover` 等模式直接报错。EXIF 旋转不明确的图片要求先规范方向。多图分别填入各自图片 slot，不能为铺满区域而截掉子图。

该功能保证“放入 PPT 时不裁切”，不能证明输入 PNG 已经包含完整论文图；必须先完成独立的原页对照和完整性检查。

```json
{
  "tables": {
    "results-table": {
      "rows": [
        ["方法", "指标A", "指标B", "指标C", "实验条件"],
        ["方法甲", "论文数值", "论文数值", "论文数值", "同一设置"],
        ["方法乙", "论文数值", "论文数值", "论文数值", "同一设置"]
      ],
      "source": {"paper": "论文名", "table": "Table 2", "page": 6}
    }
  }
}
```

保留原生 `a:tbl`、列数/列宽、背景色、边框和行样式。可改变行数，至少一行表头和一行数据；原表总高度不变，行高均分，每格重新测量。列数变化、合并单元格不在当前范围；内容过多应拆成多页，不能缩到不合模板的字号。原有字符串/数值单元格继续可用；需要强调的单格可写为对象：

```json
{"text": "论文已核验数值", "bold": true, "color": "C00000"}
```

单格对象仅支持必填 `text`（字符串或数值）和可选 `bold`（布尔值）、`color`（六位十六进制，可带 `#`）。它只改变该格文字，不影响相邻格、背景或边框；未指定的样式沿用模板。该接口不支持单格内多 run、公式计算或合并单元格，不是 Excel 公式引擎。

`strict` 模式要求每个填入的图片和表格都提供非空 `source`，它会自动写入对应页备注和生成报告，不添加可见来源小字。外部图片记录来源页、作者/机构和素材 ID，其他论文图再记录论文/图号/页码；提供尽可能可追溯的真实字段。AI 图的 source 应标明 `kind: "ai-generated"`、生成工具和提示词/记录路径，并保留页面“AI生成示意图，非论文原图”的声明。

## 报告与验证

成功或失败均写出 JSON 报告，包含模板/目录/计划 SHA、输出 SHA、选页顺序、每个文本框的字号与实际字体文件、图像输入哈希和完整 contain 尺寸、表格大小、未填对象、溢出和来源记录。`sources_to_notes` 记录移入备注的来源类型、槽位与内容；`text_fits` 记录显式 `align/vertical_align`，空值表示沿用模板。所有内部关系必须解析到实际 package part，否则构建失败。

构建固定 ZIP 时间戳，并用输入派生创建标识；相同模板、计划、目录、图片和字体环境产生相同 PPTX 字节。报告中的绝对路径不写入展示正文。

每次生成后仍需逐页渲染检查，重点查看公式、长中文、表格、图例和图片边缘。程序结构验证与隔离 case 测试不等于论文生成效果验收。

## 可选：嵌入真实演示视频

`scripts/embed_media.py` 在正常组版后，将指定的命名图片替换为可点击播放的本地MP4。先用完整视频帧作为 `images` 的PNG海报进行排版；图片可读、宽高比合适后再嵌入。它使用 python-pptx 和 Pillow（从当前运行时定位），不下载媒体。视频尺寸须从媒体元数据读取，不能猜测。

```text
python scripts/embed_media.py --input report-static.pptx --plan media-plan.json --output report.pptx --report media.audit.json
```

计划格式（所有字段必填，文件路径相对计划文件）：

```json
{
  "input_sha256": "report-static.pptx的实际SHA256",
  "videos": [{
    "slide": 5,
    "slot": "ppt2405.20.qualitative-sequence",
    "path": "../assets/external-media/demo.mp4",
    "poster": "../assets/external-media/demo-t01.png",
    "width_px": 1920,
    "height_px": 1080,
    "source": {"url": "实际媒体URL", "page": "实际来源页URL", "poster_timestamp_s": 1}
  }]
}
```

`slide`是最终放映顺序的1起页码；`slot`必须匹配该页唯一完整对象名。工具按目标图片区域等比 contain，不拉伸/裁剪，保留堆叠顺序，核对每张已使用幻灯片的原生数学与表格结构前后相同。写入新文件，已有生成输出需显式 `--overwrite`。数学和表格保护不等于全页布局验证，仍需实机打开、点击播放并检查PDF海报。PDF不会播放视频，静态画面应能独立解释任务；媒体来源与时间戳另写到该页备注和素材清单，不添加来源页脚。
