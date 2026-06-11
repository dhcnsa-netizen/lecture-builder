---
name: lecture-builder
description: >-
  Build printable two-version revision lectures (学生版 + 教师版 / student + teacher)
  as Word/WPS .docx from a teacher's own source files — a 知识点/knowledge-points doc
  plus a 题目/problem-set doc (and optionally a 解析版/answers-and-solutions doc). It
  preserves every original figure and equation exactly (by merging the source XML,
  not retyping), groups each problem under the matching knowledge point, renumbers
  examples, and adds a cover, auto table of contents, page numbers, formula boxes
  and answer/note regions. Use this whenever the user wants to turn lesson materials
  into a 讲义/lecture/复习讲义/handout/学案, make a student + teacher version of a
  worksheet, assemble knowledge + exercises into one formatted doc, or says things
  like “用知识点和题目生成讲义”, “做一份学生版和教师版”, “把解析粘到题目后面”. Works for
  physics and any subject whose source docs contain figures and equations.
---

# Lecture Builder（讲义生成器）

Turn a teacher's raw materials into a polished, print-ready revision lecture in two
flavours from the **same** content:

- **学生版 (student)** — knowledge + worked-example *statements*, each followed by a
  blank 答案/答题区 to fill in.
- **教师版 (teacher)** — identical layout but each example carries the full 【答案】 and
  【详解】 pasted from the solutions file.

Both are real `.docx` that open in WPS or Microsoft Word.

## Why this skill exists (read this first)

Source teaching docs are full of **figures and equations that cannot be retyped**
faithfully — physics problems are unsolvable without their diagrams, and option
lists are often images/formulas. So the engine **copies the original paragraph XML
verbatim** (images, OMML equations, tables) and only adds *new* structure (cover,
headings, formula boxes, answer lines) around it. Never transcribe a problem or a
formula by hand — let the engine carry the original bytes across. This is the whole
point and the source of the quality.

The only thing that needs human-level judgment each time is **mapping content to
structure**: which knowledge paragraphs form each section, where each problem
begins/ends, and which knowledge point each problem belongs to. You produce that
judgment as a small `config.json`; a deterministic engine does the heavy lifting.

## Inputs

Ask the user for (or locate) up to three `.docx` files:
1. **knowledge_file** — the 知识点 doc (required).
2. **question_file** — the 题目 doc with problem statements (required).
3. **solution_file** — the 解析版 / answers + 详解 doc (optional but needed for a
   meaningful 教师版). If the user only has two files, build just the student version
   and tell them a solutions doc is required for the teacher version.

## Workflow

### 1 — Outline every source file
Run the extractor on each file so you can see paragraph indices, figures, formulas,
and where problems start:

```bash
python <skill>/scripts/extract_outline.py "知识点.docx"          > kp_outline.txt
python <skill>/scripts/extract_outline.py "题目.docx"            > q_outline.txt
python <skill>/scripts/extract_outline.py "题目（解析版）.docx"   > sol_outline.txt
```

Read all three. Each line is `pt_index TAG | text… flags`, where `★START` marks a
problem's first paragraph and `〖IMG〗/〖MATH〗/〖TBL〗` flag figures/equations/tables.
Those `pt_index` numbers are exactly what go into the config.

### 2 — Design the knowledge sections
Follow the **knowledge file's own logical structure** (its 模型/类型/章节 order). Pick
4–6 sections that match the subject. For each, record the `[start,end]` pt ranges of
the knowledge paragraphs to include (skip pure decorative/watermark images at the very
top). Keep the original wording — you are only choosing ranges, not rewriting.

### 3 — Segment the problems
For the question file (and solution file), find the contiguous blocks of problems and
write them as segments `{kao,start,end,type}`:
- `kao` groups problems so identical numbers in different parts stay distinct (e.g.
  考点01 → kao 1, 考点02 → kao 2).
- `type` is `choice` (选择题 → short 答案 line) or `solve` (计算/论述 → ruled 答题区).
- `start,end` bound the block; the engine finds each problem inside it automatically.

The solution file usually has *different* indices than the question file (solutions
add paragraphs), so outline it separately and write `solution_segments` from its own
`★START` lines.

### 4 — Map each problem to a knowledge section (the judgment call)
For every problem, decide which knowledge point it best tests, by physics/subject
theme — not just by source order. A problem may combine topics; pick its **dominant**
model. Put the chosen `{kao,num}` into that section's `problems` list, in a sensible
order. Aim to place *all* problems somewhere; if a few don't fit any of the chosen
sections, widen the closest section's scope (and its title) rather than dropping them.
It's fine for some sections to hold many and others few — reflect reality.

### 5 — Write `config.json`
Fill in the schema. **Read `references/config_schema.md`** for every field, the
segment/section shapes, and the formula-token mini-language. Add 2–4 core `formulas`
to the most foundational section(s) only; the equations already inside the copied
knowledge are preserved automatically, so don't re-author the rest.

### 6 — Build
```bash
python <skill>/scripts/lecture_engine.py --config config.json --open
```
This writes the student file and (if `solution_file` is set) the teacher file, then
opens them. `--open` uses the OS default app (WPS/Word) on Windows.

### 7 — Finalize the table of contents & page numbers
The engine sets the doc to auto-update fields, so the TOC and page numbers fill in the
first time it opens. To **bake** them in (so a printed/handed-off file is correct
without a manual F9), run the finalizer if Microsoft Word is installed:

```powershell
powershell -ExecutionPolicy Bypass -File <skill>/scripts/finalize.ps1 -Path "lecture.docx"
powershell -ExecutionPolicy Bypass -File <skill>/scripts/finalize.ps1 -Path "lecture_教师版.docx"
```

If only WPS is installed (no Word COM), skip this — the TOC still updates on open;
just tell the user to accept “更新域” / press F9 once.

### 8 — Verify before declaring done
Spot-check the result (e.g. export a few pages to images with `pymupdf`/`fitz` or open
it) to confirm: figures render, equations look right, the TOC lists the sections,
examples are renumbered 例1/例2…, and student answer regions vs teacher solutions are
correct. Fix the config and rebuild if anything is off — rebuilding is cheap.

## Common pitfalls

- **File locked**: if a build fails with a permission error, the target `.docx` is open
  in WPS/Word — close it (or `Stop-Process -Name wps,WINWORD`) and rebuild.
- **Wrong segment bounds**: if a problem's figure or options go missing or bleed into
  the next example, your `start/end` clipped it — widen the range using the outline.
- **TOC shows placeholder text**: fields haven't updated yet; open once and update, or
  run the finalizer.
- **Don't** strip or rewrite the copied content; the engine already removes only
  bookmarks/numbering/heading-styles that would conflict, and neutralizes the green
  source-citation color for clean printing.

## Layout produced (fixed, professional, print-clean)

A4 portrait · margins 上下2cm 左右2.5cm · 1.5 line spacing · 黑体 headings (三号 H1) ·
宋体 body (小四) · centered page-number footer · cover with 班级/姓名/日期 fields ·
auto TOC · each section = ◆核心知识 → ◆重点公式 → ★笔记区 → ◆对应例题训练. Decoration
uses ■ ★ ◆ and rule lines only — no loud colors, suitable for printing.
