# blue-resume-pptx

把简历内容（一份 JSON）生成**单页 A4、可编辑的 PowerPoint 简历**。版式是蓝色顶栏，左栏放经历，右栏放技能卡片。仓库里同时附带一个 Claude 技能：把零散素材交给 Claude，它会先问清缺的信息，再写成 JSON，最后直接出 .pptx。

为什么用 PPT 做简历：Word 里图片乱跑、表格错位，PPT 可以自由排版，导出 PDF 后和 Word 做的没有区别。

| 中文示例 | 英文示例 |
|---|---|
| ![中文示例](examples/preview_zh_campus.png) | ![英文示例](examples/preview_en_marketing.png) |

每个示例都有三份文件：`.json`（输入数据）、`.pptx`（可编辑的 PPT）和同名 `.pdf`（导出效果，不用装 PowerPoint 也能直接看）。

| 示例 | 数据 | PPT | PDF |
|---|---|---|---|
| 中文 · 前端校招 | [zh_campus.json](examples/zh_campus.json) | [zh_campus.pptx](examples/zh_campus.pptx) | [zh_campus.pdf](examples/zh_campus.pdf) |
| 英文 · 市场营销 | [en_marketing.json](examples/en_marketing.json) | [en_marketing.pptx](examples/en_marketing.pptx) | [en_marketing.pdf](examples/en_marketing.pdf) |

> 示例的 PDF 和预览图是在 Linux 上用 LibreOffice 导出的，中文字体用思源黑体（Noto Sans CJK）代替黑体。在 Windows 版 PowerPoint 里打开 .pptx，会显示为黑体和微软雅黑。示例人物都是虚构的。

## 特点

- **全部可编辑**：每个板块都是真正的文本框，要点是带项目符号的列表，日期右对齐，联系方式胶囊可以点击（邮箱、GitHub、作品集）。
- **自动排版，保证一页**：放不下时按顺序处理：先把教育、证书这类短板块挪到右栏；再收紧行距和间距；最后把字号降到 8.5 号、8 号。实在放不下会提示超出几行。
- **数据自动加粗**：90 人、56%、$2.5M、3x 这类数字会自动加黑加粗；年份、月份、型号不加粗。重点短语用 `**...**` 手动标出。
- **中英文自动识别**：中文用黑体，英文用 Arial 并按单词换行。
- **可换主色**：一个 `theme` 字段，卡片底色会自动跟着变浅。

## 两个版本，任选其一

| | Python 版 `resume_pptx.py` | JS 版 `build_resume.js` |
|---|---|---|
| 适合 | ChatGPT 等网页 AI、只装了 Python 的电脑 | Claude 技能、豆包等有 Node.js 的环境 |
| 依赖 | python-pptx、Pillow（网页 AI 一般自带） | Node.js 18+，`npm install` |
| 文件 | **单个文件**，数据格式和给 AI 的说明都写在文件开头 | `skill/blue-resume-pptx/scripts/` |

两个版本用同一份 `data.json`，生成结果**逐个元素一致**（`tests/parity.py` 检查过 240 份英文简历和中文示例）。头像略有不同：Python 版用 PowerPoint 自带的圆形形状显示照片，在 PPT 里「更改图片」后仍然是圆形。

## 在网页 AI 里用（ChatGPT、豆包等）

已实测（2026-10）：

| 平台 | 结果 | 用的版本 |
|---|---|---|
| Claude | ✅ 可用 | JS 版（作为技能） |
| ChatGPT 网页版（免费版） | ✅ 可用：运行脚本并给出 .pptx 下载 | Python 版；环境为 Python 3.13、python-pptx 1.0.2、Pillow 12.3 |
| 豆包 | ✅ 可用：运行脚本并给出 .pptx 下载 | JS 版 |
| 只能聊天、不能运行代码的 AI | ⚠️ 只能输出 data.json，需要自己在电脑上运行脚本 | 任一版本 |

判断 AI 是否真的运行了脚本：看它是否贴出了脚本打印的那一行 JSON（含 `level`、`bottom`、`overflowPt`），以及 PPT 文件属性和文本框名称（「姓名」「列表_…」「核心数据卡片」等）是否来自脚本。

1. 下载 [`resume_pptx.py`](resume_pptx.py)。
2. 在 AI 对话里上传这个文件，再附上你的旧简历或零散信息（照片可选）。
3. 粘贴下面这段话：

```text
我上传了 resume_pptx.py（简历 PPT 生成脚本）和我的简历材料。请：
1. 先阅读 resume_pptx.py 开头的「给 AI 的使用说明」和 data.json 格式；
2. 按说明整理我的信息，缺的先问我，不要编造任何经历和数字；
3. 写好 data.json 后，在代码环境里运行：python resume_pptx.py data.json 简历.pptx
4. 把生成的 .pptx 给我下载，并告诉我脚本输出的 level 和 overflowPt。
注意：必须运行这个脚本生成，不要自己用 python-pptx 另画版式。
```

只能聊天、不能运行代码的 AI（比如 DeepSeek 网页版）也能用：让它只输出 `data.json`，你再在自己电脑上运行 `python resume_pptx.py data.json 简历.pptx`。

## 快速开始（本地）

Python 版：

```bash
pip install python-pptx Pillow
python resume_pptx.py examples/zh_campus.json 简历.pptx
```

JS 版（需要 Node.js 18 或以上）：

```bash
npm install
npm run example                                  # 重新生成 examples/zh_campus.pptx
npm run build -- 你的数据.json 输出文件名.pptx      # 用自己的数据生成
npm test                                         # 自动加粗规则单测
```

脚本会输出一行 JSON，说明生成结果：

```json
{"level":0,"overflowPt":0,"lang":"zh","movedToRight":[]}
```

- `level`：压缩档位，0 表示原始版式，4 表示 8 号字。
- `overflowPt`：大于 0 表示仍然超出一页。
- `movedToRight`：自动挪到右栏的板块。

生成后在 PowerPoint 或 WPS 里打开，检查一遍，再「另存为 PDF」投递，得到和 .pptx 对应的 PDF。如果装了 LibreOffice，也可以用命令行导出：

```bash
soffice --headless --convert-to pdf 输出文件名.pptx
```

## 数据格式

完整说明见 [resume_pptx.py](resume_pptx.py) 开头或 [skill/blue-resume-pptx/SKILL.md](skill/blue-resume-pptx/SKILL.md)，示例见 [examples/zh_campus.json](examples/zh_campus.json)。基本结构如下：

```json
{
  "name": "姓名",
  "title": "前端开发工程师",
  "contacts": [{ "type": "phone", "text": "138xxxx" }, { "type": "email", "text": "a@b.com" }],
  "summary": "一两句个人简介",
  "photo": "photo.jpg",
  "left": [
    { "title": "实习经历", "entries": [
      { "name": "公司 · 职位", "date": "2025.07 - 2025.12",
        "details": ["技术栈：React、TypeScript"],
        "bullets": ["要点，**重点**会加粗"] }
    ]}
  ],
  "right": [
    { "title": "技能", "tags": ["React", "TypeScript"] },
    { "title": "证书", "items": [{ "name": "CET-6", "date": "2024.06", "content": "562 分" }] }
  ]
}
```

## 作为 Claude 技能使用

把 `skill/blue-resume-pptx` 整个文件夹放进 Claude 的技能目录（例如 Claude Code 的 `~/.claude/skills/`），然后在仓库根目录执行一次 `npm install`。之后直接对 Claude 说「帮我做一份简历 PPT」，并附上旧简历或零散信息。技能会：

1. 提取信息，把缺的内容一次问清楚，并按岗位追问常见缺口：证书、量化成果、技术栈、成绩；
2. 如果你附上了 JD，先做一份匹配分析；
3. 内容超出一页时，先给出精简方案，你确认后再删；
4. 生成 .pptx 并检查渲染效果。

技能要求 AI 不编造经历和数字。凡是涉及事实的改写，都会先问你。

## 隐私提示

请不要把自己的真实简历数据和照片提交到公开仓库。可以把它们放在 `private/` 目录下，这个目录已经加入 `.gitignore`。

## 许可证

代码采用 MIT 许可证。联系方式图标来自 [Font Awesome Free](https://fontawesome.com)（通过 react-icons 渲染），按 [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/) 授权使用。
