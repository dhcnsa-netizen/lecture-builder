# 讲义制作工作流

<sub>技术标识 / 安装目录名：`lecture-builder`</sub>

> 一个 **Claude Code / Claude Agent Skill**：把老师手里的「知识点文档 + 题目文档（+ 解析版）」
> 一键合成为**学生版**和**教师版**两份排版精美、可直接打印的 Word/WPS 讲义。

把原始素材丢进来，自动得到：

| 版本 | 内容 |
|------|------|
| **学生版** | 知识点 + 例题（题干），每题后留空白「答案 / 答题区」 |
| **教师版** | 同样的版式，但每题后附**完整【答案】和【详解】** |

两份都是真正的 `.docx`，在 WPS 或 Microsoft Word 中均可打开、打印。

---

## 🎯 给已有素材的老师：30 秒出一套讲义

**适合谁**——手里**已经有「知识点文档」和「题目文档」**的中学老师（物理及任何含图含公式的学科）。
你不用再排版、不用重打图和公式，把素材丢进来即可。

**① 你需要提供的材料（都是 `.docx`）**

| 材料 | 是否必需 | 用途 |
|------|:---:|------|
| 知识点文档 | ✅ 必需 | 讲义的"核心知识"部分 |
| 题目文档 | ✅ 必需 | 例题题干（学生版） |
| 题目解析版 | ⭐ 可选 | 教师版的【答案】【详解】来自它；不提供则只出学生版 |

**② 怎么做（最快路径）**——把上面 3 个文件发给装了本 skill 的 Claude，说一句
“**用这份知识点和题目做学生版+教师版讲义**”，剩下的分节、归类、排版、编号全自动。
（也可手动写 `config.json` 跑脚本，见下文。）

**③ 你会得到什么格式**——**两份可直接打印的 Word/WPS `.docx`**：

- **学生版**：知识点 + 例题（题干），每题后留空白「答案 / 答题区」
- **教师版**：同样的版式 + 每题完整【答案】和【详解】

排版固定为 A4 纵向、1.5 倍行距、黑体标题 / 宋体正文、自动目录 + 页码，
**原图、原公式、原表格逐字节保真**，干净适合打印。

---

## ✨ 它解决了什么问题

老师做复习讲义时最痛的两点，这个工具都替你扛了：

1. **图和公式不能重打。** 物理题离开了图就没法做，选项常常本身就是公式或图片。本工具
   **直接复制源文件的原始 XML**（图片、Word 公式 OMML、表格逐字节搬运），绝不重新录入，
   所以**图、公式、表格 100% 保真**。
2. **重复的体力活全自动。** 封面、自动目录、页码、例题重新编号（例1/例2…）、学生版答题区 /
   教师版答案详解、A4 排版、字体字号行距——全部由引擎一次成型。

唯一需要"人脑"判断的，是**把内容映射成结构**：知识点分几节、每道题归到哪个知识点。
这一步由你（或调用本 skill 的 Claude）写成一个很小的 `config.json`，剩下交给确定性的引擎。

---

## 📐 产出的版式（固定、专业、适合打印）

- A4 纵向，页边距 上下 2cm / 左右 2.5cm，1.5 倍行距
- 标题黑体（一级三号居中），正文宋体（小四）
- 封面（标题 / 副标题 / 班级·姓名·日期 填写栏）
- **自动目录** + 居中**页码**页脚
- 每个知识点小节结构：`◆核心知识 → ◆重点公式 → ★笔记区 → ◆对应例题训练`
- 仅用 ■ ★ ◆ 和分隔线点缀，无花哨颜色，干净易印

---

## 🗂️ 仓库结构

```
.
├── SKILL.md                     # Skill 定义（触发条件 + 工作流说明）
├── scripts/
│   ├── lecture_engine.py        # 核心引擎：合并原始 XML → 学生版 + 教师版
│   ├── extract_outline.py       # 把 docx 逐段编号导出，便于规划分组
│   └── finalize.ps1             # 用 Word 刷新目录/页码（无 Word 则可跳过）
└── references/
    ├── config_schema.md         # 配置文件字段说明 + 公式语法
    └── example_config.json      # 一份完整的真实配置示例
```

---

## 🚀 安装为 Claude Skill

把整个文件夹放到 Claude 的 skills 目录即可被自动识别：

- **Claude Code（个人级）**：`~/.claude/skills/lecture-builder/`
  （Windows：`C:\Users\<你>\.claude\skills\lecture-builder\`）
- 安装后，当你说“**用这份知识点和题目做一份讲义，要学生版和教师版**”之类的话时，
  Claude 会自动触发本 skill 并完成整套流程。

> 也可以完全脱离 skill，作为普通命令行脚本使用（见下文）。

---

## 🧭 使用方法（命令行 / 引擎直跑）

### 依赖
- **Python 3**（标准库即可，无需第三方包；引擎纯标准库实现）
- 可选：**Microsoft Word**（仅用于把目录、页码"烤死"到文件里；只有 WPS 也能用，首次打开按 F9 / 接受"更新域"即可）

### 三步走

**① 给每个源文件导出大纲**，拿到每段的 `pt_index` 以及题目起始位置：

```bash
python scripts/extract_outline.py "知识点.docx"           > kp.txt
python scripts/extract_outline.py "题目.docx"             > q.txt
python scripts/extract_outline.py "题目（解析版）.docx"    > sol.txt
```

每行形如 `序号 标签 | 文本… 标记`，其中 `★START` 标出一道题的开头，
`〖IMG〗/〖MATH〗/〖TBL〗` 标出图 / 公式 / 表。这些 `pt_index` 就是写进配置的数字。

**② 写 `config.json`**（字段详见 [`references/config_schema.md`](references/config_schema.md)，
可直接照抄 [`references/example_config.json`](references/example_config.json) 改）：

- `sections`：知识点分节，每节给出知识点段落范围 `[start,end]`，以及归到本节的题目列表；
- `question_segments` / `solution_segments`：把题目/解析文件按"考点 + 题型"切成几大块，
  引擎会在块内自动找出每道题；
- `formulas`：给最核心的小节补 2–4 条 Word 公式（可选，原文里的公式已自动保留）。

**③ 生成两版讲义并打开**：

```bash
python scripts/lecture_engine.py --config config.json --open
```

会输出 `output_student` 和 `output_teacher` 指定的两个 `.docx`（默认
`lecture.docx` 和 `lecture_教师版.docx`），并用系统默认程序（WPS/Word）打开。

**④（可选）把目录页码烤死**——若装了 Word：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/finalize.ps1 -Path "lecture.docx"
```

---

## 🧩 配置示例（节选）

```json
{
  "title": "高二物理期末复习讲义",
  "subtitle": "专题：电磁感应",
  "knowledge_file": "知识点.docx",
  "question_file": "题目.docx",
  "solution_file": "题目（解析版）.docx",
  "output_student": "电磁感应_学生版.docx",
  "output_teacher": "电磁感应_教师版.docx",
  "question_segments": [
    {"kao": 1, "start": 8, "end": 52, "type": "choice"},
    {"kao": 1, "start": 55, "end": 62, "type": "solve"}
  ],
  "sections": [
    {"title": "一、楞次定律与感应电流方向的判断",
     "knowledge": [[0, 10]],
     "formulas": [["Φ=", {"it": "BS"}]],
     "problems": [{"kao": 1, "num": 3}, {"kao": 1, "num": 4}]}
  ]
}
```

> **公式语法**：每条公式是一个数组，元素是纯文本或 `{"frac":[分子,分母]}`、`{"sup":[底,指数]}`、
> `{"sub":[…]}`、`{"sqrt":…}`、`{"it":"斜体"}`。例如 `T = 2πm/qB` 写作
> `["T=", {"frac": [["2π", {"it":"m"}], ["q", {"it":"B"}]]}]`。详见配置文档。

---

## ⚠️ 常见问题

- **文件被占用 / 写入失败**：目标 `.docx` 正在 WPS/Word 中打开，先关闭再重新生成。
- **某题图或选项缺失**：该题的 `start/end` 把内容截断了，对照大纲把范围放宽。
- **目录显示占位文字**：域还没刷新，打开后按 F9 / 接受"更新域"，或运行 `finalize.ps1`。
- **只有两个文件（没有解析版）**：只会生成学生版；教师版需要含答案/详解的题目文件。
- **题号没变成"例N"**：本工具已能处理跨运行块、带前导空块的题号；若仍异常，多半是
  `★START` 没识别到，检查大纲里该题首段是否以"数字．"开头。

---

## 🛠️ 工作原理（一句话版）

`extract_outline.py` 和 `lecture_engine.py` **共用同一套解析器**，保证大纲里的段落序号与
引擎构建时完全一致；引擎按 `config.json` 把知识点段落和题目段落**原样拼接**，只在外层
注入封面、目录、标题、公式框、答题区/答案区，最终重新打包成合法的 `.docx`。

---

## 适用范围

为高中物理复习讲义打造，但对**任何含图、含公式的 Word 素材**都适用——
只要源文件是 `.docx`，知识点与题目可被切分归类，就能生成同样的两版讲义。

## License

[MIT](LICENSE)
