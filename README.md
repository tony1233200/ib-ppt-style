# 投行 PPT 风格规范与模板（IB PPT Style）

> 一套可直接执行的投行（investment banking）风格演示文稿规范：版式栅格、字号阶梯、深蓝配色、结论式标题、Source 脚注与保密页脚——附按规范生成的 `.pptx` 模板和一份行业分析样例。

面向董事会 / CFO / 投委会等专业读者的**阅读型**演示稿（不是路演大屏稿）：一页一个结论，图表与表格主导页面，文字只做证据与标注。

## 仓库内容

| 路径 | 说明 |
|---|---|
| [ib-style-guide.md](ib-style-guide.md) | **规范全文**：总则、栅格、字号阶梯、配色、表格样式、数据呈现、pitch book 章节顺序、页脚与保密、示意图规范、出稿检查清单 |
| `templates/ib-template.pptx` | 按规范生成的空白模板（7 页：封面 / 保密声明 / 章节页 / 内容页 / 表格页 / 图表页 / 结尾页） |
| [scripts/build_template.py](scripts/build_template.py) | 模板的生成脚本（python-pptx），改色值/字号后可重新生成模板 |
| `examples/sample-deck.pptx` | 按规范实际出稿的 9 页样例《中国轮胎行业概览》（数字为演示口径，页脚已标注） |

## 规范要点速览

- **画布**：16:9（1280×720），左右页边距 64px，标题区 / 正文区 / 页脚区像素级固定
- **配色锁定**：主色 navy `#0B2C5C`、辅助蓝 `#1F4E79`、浅蓝灰 `#D9E2EC`、墨色 `#14243D`；强调红 `#A6192E` 每页 ≤ 2 处，仅用于焦点数据；封面金线 `#B9975B`
- **字号阶梯**：封面标题 40pt → 章节页 28pt → 内容页结论标题 21pt → 正文 13.5pt → 表格 12pt → 脚注 8.25pt，同角色全稿恒定
- **结论式标题（action title）**：标题是完整结论句——✗「营收分析」 ✓「营收连续四季增长，四季度达 61 亿元」
- **表格**：表头深蓝底白字、隔行斑马纹、数字右对齐千分位、单位写进表头、合计行加主色线
- **图表**：同系深浅配色（禁彩虹色）、柱图零基线、关键点必标数据标签、口径写进脚注
- **页脚**：每页左下 `Source: <来源>, <日期>`，保密字样 `Confidential — For discussion purposes only`，右下页码
- **结构**：封面 → 保密声明 → 执行摘要 → 情景与目标 → 行业/市场 → 公司/业务 → 财务 → 估值 → 战略建议 → 下一步 → 附录

## 使用方式

1. **直接套模板**：打开 `templates/ib-template.pptx`，替换文字与数据即可；新增页面从模板页复制改。
2. **改规范再生成**：修改 `scripts/build_template.py` 顶部的色值常量，运行
   ```bash
   pip install python-pptx
   python scripts/build_template.py            # 输出 templates/ib-template.pptx
   python scripts/build_template.py out.pptx   # 或指定输出路径
   ```
3. **给 AI 工具当规范**：把 `ib-style-guide.md` 作为风格约束喂给 PPT 生成工具（本项目即以此方式配合 [ppt-master](https://github.com/hugohe3/ppt-master) 出稿，样例见 `examples/`）。

## 说明与来源

规范的章节结构、action title、来源标注等方法论整理自公开资料：Deckary《Investment Banking Pitch Book: Structure, Content & Best Practices》、Deckary《Consulting Slide Standards》；字号/边距的具体数值按投行通用模板惯例（Macabacus 系格式标准）结合 16:9 画布折算，属执行口径，不是任何一家银行的官方模板原文。示意图一节的令牌与分层规则源自 diagram-design 风格体系，并按本规范色值做了映射锁定。

样例与模板中的数字均为**示例数据（演示口径）**，不构成任何投资建议。

## License

MIT © 2026 tony1233200
