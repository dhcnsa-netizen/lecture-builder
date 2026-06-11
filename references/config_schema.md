# Lecture config schema

`lecture_engine.py` is driven by one JSON file. All `*_file` paths are relative to
the config file's own folder (or absolute). pt indices are the numbers printed by
`extract_outline.py` for that specific file.

## Top-level fields

| field | required | meaning |
|-------|----------|---------|
| `title` | yes | cover main title, e.g. `高二物理期末复习讲义` |
| `subtitle` | yes | cover subtitle, e.g. `专题：带电粒子在磁场中的运动` |
| `knowledge_file` | yes | source .docx with the knowledge points |
| `question_file` | yes | source .docx with the problems (statements only) → student version |
| `solution_file` | no | source .docx with answers + 详解 → teacher version. If omitted, only the student version is built. The teacher version copies whole problem blocks (statement + answer + solution) from this file. |
| `output_student` | no | default `lecture.docx` |
| `output_teacher` | no | default `lecture_教师版.docx` |
| `question_segments` | yes | how to slice the question file into problem groups (below) |
| `solution_segments` | no | same idea for the solution file; defaults to `question_segments` if the layouts match (they usually don't — set it explicitly) |
| `sections` | yes | the knowledge sections, each followed by its assigned problems |
| `answer_lines` | no | ruled lines in a student 答题区 (default 5) |
| `note_lines` | no | ruled lines in each 笔记区 (default 2) |

## Segments

A segment is one contiguous block of similarly-typed problems inside a source file:

```json
{"kao": 1, "start": 6, "end": 29, "type": "choice"}
```

- `kao` — a group id you choose (e.g. 1 = 考点01, 2 = 考点02). Problems are keyed by `(kao, num)` where `num` is the original printed number, so the same number in different kao groups stays distinct.
- `start`,`end` — inclusive pt index range (from `extract_outline.py`) covering that block.
- `type` — `"choice"` (selection question → short 【答案】 line) or `"solve"` (calculation → 答题区 with ruled lines). Only affects the **student** version.

The engine auto-detects each problem's first paragraph (a line beginning with `数字．`) inside the range, so you only give the block bounds, not every problem.

## Sections

```json
{
  "title": "一、洛伦兹力、安培力与圆周运动基础",
  "knowledge": [[2, 12]],
  "formulas": [ ["F=qvB"], ["r=", {"frac": ["mv", "qB"]}] ],
  "problems": [ {"kao": 1, "num": 1}, {"kao": 2, "num": 3} ]
}
```

- `knowledge` — list of `[start,end]` inclusive pt ranges **in the knowledge file** to copy verbatim (figures/tables/formulas preserved).
- `formulas` — optional; list of formulas rendered as real Word equations (OMML). Each formula is a token list (see below). Use for the 2–4 core formulas of a topic; the originals already inside the copied knowledge are kept too.
- `problems` — ordered list of `{kao,num}`; they appear as 例1, 例2 … (renumbered per section). Student gets statement + answer region; teacher gets statement + 答案 + 详解.

## Formula tokens

A formula is a list. Each item is either a plain string or one dict:

| token | renders |
|-------|---------|
| `"qvB"` or `{"text":"qvB"}` | upright text |
| `{"it":"v"}` | italic |
| `{"frac":[NUM, DEN]}` | fraction (NUM/DEN are each a string or token list) |
| `{"sup":[BASE, EXP]}` | superscript |
| `{"sub":[BASE, EXP]}` | subscript |
| `{"sqrt":X}` | square root |

Example — `T = 2πm / qB`:
```json
["T=", {"frac": [["2π", {"it":"m"}], ["q", {"it":"B"}]]}]
```
Keep formulas simple; for anything elaborate, rely on the originals copied from the knowledge file instead of re-authoring.

## Minimal example

```json
{
  "title": "高二物理期末复习讲义",
  "subtitle": "专题：带电粒子在磁场中的运动",
  "knowledge_file": "知识点.docx",
  "question_file": "题目.docx",
  "solution_file": "题目（解析版）.docx",
  "question_segments": [
    {"kao": 1, "start": 6,  "end": 29,  "type": "choice"},
    {"kao": 1, "start": 32, "end": 48,  "type": "solve"},
    {"kao": 2, "start": 53, "end": 95,  "type": "choice"},
    {"kao": 2, "start": 98, "end": 145, "type": "solve"}
  ],
  "solution_segments": [
    {"kao": 1, "start": 6,   "end": 70,  "type": "choice"},
    {"kao": 1, "start": 72,  "end": 118, "type": "solve"},
    {"kao": 2, "start": 121, "end": 244, "type": "choice"},
    {"kao": 2, "start": 246, "end": 499, "type": "solve"}
  ],
  "sections": [
    {"title": "一、洛伦兹力、安培力与圆周运动基础",
     "knowledge": [[2, 12]],
     "formulas": [["F=qvB"], ["r=", {"frac": ["mv", "qB"]}]],
     "problems": [{"kao":1,"num":1},{"kao":1,"num":2},{"kao":2,"num":3}]},
    {"title": "二、带电粒子在有界磁场中的运动",
     "knowledge": [[13, 60]],
     "problems": [{"kao":2,"num":1},{"kao":2,"num":2}]}
  ]
}
```
