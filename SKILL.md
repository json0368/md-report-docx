---
name: md-report-docx
description: 把 Markdown 报告转换成中文课程报告/论文排版的 Word 文档（.docx），含宋体正文、黑体标题、首行缩进、图表题注、代码框加语法配色、目录与页码。Convert a Markdown report into a Word document with Chinese academic report formatting.
---

# Markdown 报告转 Word

把一份 Markdown 报告转成符合中文课程设计/论文排版要求的 `.docx`：A4、2.5cm 边距、
宋体正文、黑体加粗、Times New Roman 西文、首行缩进 2 字符、两端对齐、图表题注居中、
代码块装箱加语法配色、页脚页码、自动目录域。

排版不依赖任何 Word 模板文件，文档是从零构建的，所以换机器、换项目都能得到一样的结果。

## 用法

```bash
python scripts/md_to_report_docx.py 报告.md
python scripts/md_to_report_docx.py 报告.md --out 输出.docx --check
python scripts/md_to_report_docx.py --print-defaults > 报告.json
```

`--check` 会在转换后自检，默认关闭。`--print-defaults` 打印全部默认排版参数，
改哪项就把哪项写进配置文件。

用 `load_workspace_dependencies` 返回的 Python 执行，它自带 `python-docx`、`Pillow`
和 `pygments`。不要另装系统 Python。

## 配置

在 Markdown 旁边放一个**同名**的 `.toml` / `.json` / `.yaml` 就会自动加载，
只写要改的项，其余沿用默认值：

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
```

也可以显式指定：`--config 我的排版.toml`。全部配置项见
[references/config.md](references/config.md)。

三种格式都可以用：`.toml` 由 Python 3.11+ 自带的 `tomllib` 读取，不需要额外依赖，
推荐优先使用；`.json` 用标准库；`.yaml` 需要环境里装了 PyYAML，没装时会报错并
提示改用另外两种。

## 支持的 Markdown 语法

标题、段落、加粗、斜体、删除线、行内代码、链接、表格、图片与题注、代码围栏、
有序列表、无序列表（支持三层嵌套）、引用、分隔线、参考文献。

不支持的语法不会报错，会按普通正文排版。完整清单和常见坑见
[references/syntax.md](references/syntax.md)。

## 自检

`--check` 检查这些内容，发现问题会打印 `错误：` 并以退出码 1 结束：

- **正文文本逐段不变** —— 加代码框前后，非代码正文必须完全一致，防止转换丢字
- **代码块内容与 Markdown 围栏逐字一致**，数量也要对得上
- **每张数据表列宽合计等于正文宽度**，每张图片宽度不超版心
- **颜色白名单** —— 代码段落只能用配色表里的颜色，其余段落必须是黑色
  （这条能抓出主题色泄漏，Word 里会莫名出现蓝色标题）

`--pdf 文件.pdf` 可以额外做逐字符的溢出检查（有没有文字跑出页面）。
这项需要环境里装了 `pdfplumber`；没装时自动跳过并说明原因。

## 代码块外观

代码块转成 1×1 表格：浅灰底纹 `F6F8FA`、`D9D9D9` 细边框、单元格留白，
内部用 pygments 分词上色（关键字 `1F4E79` 加粗、注释 `58724C`、字符串 `985C10` 等）。

代码块默认用**等宽字体**（`fonts.mono`，默认 `Consolas`），正文仍是 Times New Roman。
代码里的中文注释仍用宋体，因为等宽字体没有汉字字形。想跟正文保持一致：

```toml
[code]
font = "latin"
```

超过 30 行的代码块允许表格跨页断开，短块则整块保持在同一页。折行后的部分会缩进
0.6cm（`code.wrap_indent_cm`），不会跟下一行混在一起。

`diff` 和 `console` 按行首符号上色——`-` 红、`+` 绿、`@@` 蓝、`$` 提示符蓝——
不走 pygments：它的 diff 词法器会按文件名后缀挑子词法器，console 词法器会吞掉
内容。支持的语言标记见 [references/syntax.md](references/syntax.md)。

## 已知限制

- 表格列宽按内容自动分配（汉字算 2、西文算 1.2 个半角宽），个别表格可能不如手工调的好看。
  需要精确控制时把 `tables.width_mode` 改成 `rules`，再按表头写死列宽。
- `#` 一级标题会被当成封面：标题 + 下一个非空行作副标题 + 「目录」+ 目录域。
  **一份文档只应出现一个 `#`**，出现第二个会重复生成封面和目录。
- 目录和页码是 Word 域，打开文档时按提示更新，或用 `Ctrl+A` 后 `F9` 刷新。
- 题注必须贴着图或表才识别：`图 3-1` 写在图片下面，`表 3-1` 写在表格上面。
  这样正文里的「图 3-1」「表 3-1」才不会被误判。
  题注格式可用 `captions.figure_patterns` 和 `captions.table_patterns` 配置。
