# md-report-docx

把 Markdown 报告转换成符合中文课程设计/论文排版要求的 Word 文档。

*Convert a Markdown report into a Word document with Chinese academic formatting.*

正文宋体、标题黑体、首行缩进 2 字符、A4 页边距 2.5cm、图表题注居中、代码块装箱加
语法配色、页脚页码、自动目录域。**不需要任何 Word 模板文件**，文档从零构建，换机器、
换项目结果一致。

## 特性

- **不依赖模板**：所有样式在代码里定义，由 `python-docx` 从空文档建起。避免了
  「拿别人的模板改改」带来的样式漂移和主题色泄漏。
- **中文排版默认就对**：中西文字体分开设置（`w:eastAsia`，只设 `font.name` 对中文
  无效）、标题去掉继承的自动编号、正文两端对齐加首行缩进 2 字符。
- **代码块**：转成 1×1 表格，浅灰底纹加细边框，内部按语法上色。pygments 自带的
  600 个词法器里，实测 68 个常用语言标记能真正着色；`diff` 和 `console` 单独处理。
- **表格列宽按内容自动分配**：汉字按 2、西文按 1.2 个半角宽估算，按比例分配到正文
  宽度，不用手工调。
- **真目录域和页码域**，不是手打的假目录。
- **可配置**：TOML / JSON / YAML 配置文件，只写要改的项，其余沿用默认值。
- **转换后自检**：正文逐段比对、代码块逐字比对、列宽合计、颜色白名单。

## 环境要求

Python **3.11+**（配置文件用标准库 `tomllib` 读取）。

```
pip install python-docx Pillow pygments
```

- `PyYAML` 可选，只有用 `.yaml` 配置时才需要。
- `pdfplumber` 可选，装了才能用 `--pdf` 做逐字符的溢出检查。

## 快速开始

```bash
python scripts/md_to_report_docx.py 报告.md
```

在同目录生成 `报告.docx`。

```bash
# 指定输出并自检
python scripts/md_to_report_docx.py 报告.md --out 输出.docx --check

# 打印全部默认排版参数，改哪项就把哪项写进配置文件
python scripts/md_to_report_docx.py --print-defaults
```

## 配置

在 Markdown 旁边放一个**同名**配置文件就会自动加载：

```
课程设计报告.md
课程设计报告.toml     ← 自动加载
```

```toml
[fonts]
cjk = "仿宋"

[sizes]
body = 10.5

[page]
margin_cm = 3.0

[tables]
width_mode = "rules"

[tables.rules]
"字段,类型,可空,说明" = [4.0, 3.6, 1.2, 7.2]
```

也可以 `--config 我的排版.toml` 显式指定。全部配置项、默认值和说明见
[references/config.md](references/config.md)。

## Markdown 语法支持

支持标题、段落、加粗、斜体、删除线、行内代码、链接、表格、图片与题注、代码围栏、
有序列表、无序列表（三层嵌套）、引用、分隔线、参考文献。

不支持的语法**不会报错**，会按普通正文排版。完整清单、代码围栏的语言标记、
以及容易踩的坑见 [references/syntax.md](references/syntax.md)。

## 自检

`--check` 检查这些内容，发现问题会打印 `错误：` 并以退出码 1 结束：

- **正文文本逐段不变** —— 加代码框前后比对非代码正文，防止转换丢字
- **代码块内容与 Markdown 围栏逐字一致**，数量也要对得上
- **每张数据表列宽合计等于正文宽度**，每张图片宽度不超版心
- **颜色白名单** —— 代码段落只能用配色表里的颜色，其余段落必须是黑色

最后一条能抓出主题色泄漏这类只在 Word 里才暴露的问题。

## 工作原理

| 文件 | 职责 |
| --- | --- |
| `scripts/md_to_report_docx.py` | 命令行入口与流程编排 |
| `scripts/report_style.py` | 配置加载与合并、页面设置、样式定义、表格列宽计算 |
| `scripts/report_builder.py` | 逐行解析 Markdown，生成 docx 元素 |
| `scripts/code_highlight.py` | 代码块装箱与语法配色 |
| `scripts/audit_report.py` | 自检 |

几个值得说明的实现选择：

**代码块整块分词、按行切分。** 逐行分词会丢跨行注释和字符串状态；把整个代码块合并
成一段又会让悬挂缩进失效（Word 会把 `<w:br/>` 之后的行都当成折行，一起缩进去）。
所以整块分词后按换行切开，再放回一行一段，两边的好处都留着。

**`diff` 和 `console` 不走 pygments。** 它的 diff 词法器会读 `---`/`+++` 里的文件名
后缀去挑子词法器，hunk 内容不是那个语言时就整行退化成 `Token.Error`，增删行反而
没有颜色；console 词法器会吞掉提示符之后的内容。这两个按行首符号自己上色。

**内容保底。** 任何词法器只要分词后的拼接结果与原文不符，自动退回纯文本并在转换时
打印提示——宁可不着色，也不能丢字。

## 已知限制

- 表格列宽是估算的，个别表格可能不如手工调的好看。要精确控制就把
  `tables.width_mode` 改成 `rules`，按表头写死列宽。
- `#` 一级标题会被当成封面（标题 + 下一个非空行作副标题 + 「目录」+ 目录域），
  **一份文档只应出现一个 `#`**，出现第二个会重复生成封面和目录。
- 目录和页码是 Word 域，打开文档时按提示更新，或用 `Ctrl+A` 后 `F9` 刷新。
- 题注只在紧跟图片之后时才识别，避免正文里的「图 3-1」被误判。
- 不支持脚注、数学公式、Setext 标题、自动链接、引用式链接、缩进代码块、原始 HTML、
  表格合并单元格、嵌套表格。

## 目录结构

```
md-report-docx/
├── SKILL.md                 技能说明（供 agent 运行时读取）
├── README.md
├── LICENSE
├── .gitignore
├── references/
│   ├── config.md            全部配置项
│   └── syntax.md            语法支持清单
├── scripts/
│   ├── md_to_report_docx.py
│   ├── report_style.py
│   ├── report_builder.py
│   ├── code_highlight.py
│   └── audit_report.py
└── _test/
    ├── 语法测试.md           语法回归样例
    ├── _check_fonts.py      查正文与代码块实际用的字体
    ├── _check_colors.py     查每个 run 的实际色值
    ├── _probe_languages.py  列出哪些语言标记能着色
    └── _probe_tokens.py     看词法器的原始输出
```

`SKILL.md` 是给 agent 运行时读的技能描述，内容与 README 有重叠但不冲突，
只当普通命令行工具用时可以忽略。

## 许可证

[MIT](LICENSE) © 2026 json0368
