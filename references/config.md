# 排版配置项

配置文件与 Markdown 同名（`.toml` / `.json` / `.yaml`）即自动加载，也可以用
`--config` 指定。只写要改的项即可，其余用下面的默认值。

想看当前版本的完整默认值，运行 `python scripts/md_to_report_docx.py --print-defaults`。

## page — 页面

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `width_cm` | `21.0` | 页宽 |
| `height_cm` | `29.7` | 页高 |
| `margin_cm` | `2.5` | 四边页边距，正文宽度 = 页宽 − 2 × 此值 |
| `header_distance_cm` | `1.2` | 页眉距顶边 |
| `footer_distance_cm` | `1.3` | 页脚距底边 |

## fonts — 字体

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `latin` | `Times New Roman` | 西文字体，用于正文和代码块 |
| `cjk` | `宋体` | 中文正文字体 |
| `cjk_bold` | `黑体` | 中文加粗（标题）字体 |
| `mono` | `Consolas` | 行内代码字体 |

中文必须靠 `w:eastAsia` 指定，只设 `font.name` 对中文无效——这一点脚本已经处理。

## sizes — 字号（磅）

| 键 | 默认值 | | 键 | 默认值 |
| --- | --- | --- | --- | --- |
| `title` | `20.0` | | `caption` | `10.5` |
| `subtitle` | `12.0` | | `code` | `10.5` |
| `body` | `12.0` | | `table` | `12.0` |
| `h1` | `16.0` | | `toc` | `12.0` |
| `h2` | `14.0` | | `footer` | `10.5` |
| `h3` | `12.0` | | | |

## body — 正文段落

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `first_line_indent_pt` | `24.0` | 首行缩进，24 磅 = 2 个 12 磅字符 |
| `line_spacing` | `1.5` | 行距倍数 |
| `align` | `justify` | `left` / `center` / `right` / `justify` |

## headings — 标题

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `line_spacing` | `1.35` | 行距倍数 |
| `space_before` | `[14.0, 10.0, 10.0]` | 一/二/三级标题的段前距 |
| `space_after` | `6.0` | 段后距 |
| `page_break_level1` | `true` | 正文第一个标题另起一页，让封面和目录单独占一页 |
| `page_break_prefixes` | `["附录 "]` | 一级标题以这些词开头时也另起一页 |

## cover — 封面与目录

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `break_before` | `课程设计报告` | 标题里出现这个词时前面换行；留空字符串表示不换行 |
| `toc_title` | `目录` | 目录标题文字；留空则不生成目录 |
| `subtitle_from_next_line` | `true` | 把 `#` 标题之后的第一个非空行当副标题 |

## captions — 图表题注

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `patterns` | `["^图 \\d+-\\d+ ", "^表 \\d+-\\d+ "]` | 题注匹配规则，按顺序尝试 |
| `require_preceding_image` | `true` | 只在紧跟图片之后才当题注，避免正文里的「图 3-1」被误判 |

改成 `图 1.1` 风格：`patterns = ["^图 \\d+\\.\\d+ "]`。

## images — 图片

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `max_width_cm` | `0.0` | 最大宽度；`0` 表示用正文宽度 |
| `max_height_cm` | `16.0` | 最大高度，宽度按高宽比反推 |

## tables — 表格

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `width_mode` | `auto` | `auto` 按内容分配；`equal` 均分；`rules` 用表头规则 |
| `minimum_column_cm` | `1.4` | 任何一列不低于此宽度 |
| `maximum_weight` | `40.0` | 单个单元格再长也不会把该列撑过这个权重 |
| `header_fill` | `F2F2F2` | 表头底纹 |
| `border_color` | `D9D9D9` | 边框颜色 |
| `center_values` | `["是","否","1:1","1:N","M:N","ref","index"]` | 这些值居中；纯数字也居中，其余左对齐 |
| `rules` | `{}` | `rules` 模式的表头规则 |

`rules` 模式的键是表头用逗号连接，值是各列宽度（厘米）：

```toml
[tables]
width_mode = "rules"

[tables.rules]
"字段,类型,可空,说明" = [4.0, 3.6, 1.2, 7.2]
```

匹配不到规则的表会退回均分，宽度数量写错会直接报错。

## lists — 列表与引用

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `bullet_chars` | `["•","◦","▪"]` | 一/二/三层无序列表的项目符号 |
| `bullet_indent_cm` | `0.74` | 每层缩进量 |
| `number_pattern` | `^\d+\. ` | 有序列表识别规则 |
| `number_hanging_indent_cm` | `0.6` | 有序列表悬挂缩进 |
| `quote_indent_cm` | `0.74` | 引用块缩进 |

## references — 参考文献

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `pattern` | `^\[\d+\] ` | 识别规则 |
| `hanging_indent_cm` | `0.65` | 悬挂缩进 |

## code — 代码块

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `frame` | `true` | 是否装箱；`false` 则只输出 Report Code 段落 |
| `shading` | `F6F8FA` | 单元格底纹 |
| `border_color` | `D9D9D9` | 边框颜色 |
| `split_after_lines` | `30` | 超过这个行数允许跨页断开 |
| `keep_with_next_for_text` | `true` | `text` 块里对齐的 ASCII 内容不要断开 |
| `font` | `mono` | `mono` 用 `fonts.mono` 等宽字体；`latin` 沿用正文字体 |
| `wrap_indent_cm` | `0.6` | 折行后的缩进量，让折下来的部分与上一行区分开；`0` 关闭 |
| `palette` | 见下 | 语法配色 |

`palette` 的键和默认色值：

| 键 | 色值 | | 键 | 色值 |
| --- | --- | --- | --- | --- |
| `text` | `24292F` | | `function` | `006D77` |
| `keyword` | `1F4E79` | | `decorator` | `6F42A0` |
| `string` | `985C10` | | `operator` | `455265` |
| `comment` | `58724C` | | `punctuation` | `334155` |
| `number` | `7C3F8C` | | `inserted` | `1A7F37` |
| `deleted` | `CF222E` | | | |

`inserted` 和 `deleted` 用于 `diff` 代码块的增删行。

## metadata — 文档属性

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `title` | `""` | 文档标题属性 |
| `subject` | `""` | 主题属性 |
| `keywords` | `""` | 关键词属性 |

留空则不写入。
