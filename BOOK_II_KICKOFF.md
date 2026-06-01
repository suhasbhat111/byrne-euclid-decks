# Book II Kickoff — Byrne Euclid Decks

## Status when Book II begins
- Book I: **fully complete** — 48 proposition PDFs + combined PDF + 48 PPTXs +
  combined PPTX + Basic Principles PPTX, all in `build/`.
- Pipeline: `make all` rebuilds PDFs; `python3 packager/packager.py --all`
  rebuilds all PPTXs; `make pptx` does both.
- Parser gate: `cd packager && python3 test_parser.py` must print `ALL PASS`
  after any generator change before rebuilding PPTXs.

---

## What Book II covers
14 propositions on rectangles and squares (areas of right-angled triangles,
completing the square, etc.). Byrne's 1847 edition covers II.1–II.14.
The jemmybutton reference PDF has the Book II pages — check them before
writing any generator.

---

## Starting a new generator: the checklist

Copy these conventions exactly from Book I (they are the conventions in
`project_byrne_setup.md` memory — re-read it at the start of a new session):

1. `\documentclass[12pt]{article}` + `\geometry{a5paper, margin=1.5cm}`
2. `\usepackage[lining]{ebgaramond}` always; add `\usepackage{amsmath}` when needed
3. `\def\mpPre{u := 1cm; textLabels := false;}`
4. Title format: `\textbf{Book~II.\enspace Prop.\ <N>. <Prob.|Theor.>}`
   — **"Book II." not "Book I."** — the packager's parser validates this.
5. `% STATEMENT` / `% PROOF` tags — only needed if you follow the PROB-tagged
   layout (I01–I10 style). All other layouts are detected automatically.
6. Layer-1 geometric assertions (`errmessage "ASSERT..."`) inside
   `\defineNewPicture` — mandatory, same as Book I.
7. Build: `lualatex generator_II<NN>.tex` from the `generators/` directory.
8. After build: `cp generator_II<NN>.pdf build/II<NN>.pdf` and update
   `STATUS.md`.

---

## Key differences between Book I and Book II layouts

Book II propositions tend to be **heavier on algebraic-style equations**
(rectangle notation, products of lines) and lighter on pure geometry. Expect:
- More `\drawPolygon` / `byPolygon` colored rectangle fills
- Equation lines like `rect(AB, BC) + sq(BD) = sq(AD)` expressed as
  colored symbols — check jemmybutton carefully for the color choices
- Same PROB / THEOR layout families as Book I

---

## Packager: what needs updating for Book II

The packager (`packager/packager.py`) needs **one change** before it can
handle Book II generators:

- The parser gate regex currently validates:
  `^Book [IVXLCDM]+\. Prop\. [IVXLCDM]+\. (Prob|Theor)\.$`
  This already matches "Book II. Prop. I. Prob." — **no change needed**.

- The `--all` flag in `packager.py main` globs `generator_I*.tex`. For Book II,
  update the glob to `generator_I*.tex generator_II*.tex` or add a `--book`
  argument, or just run Book II separately with `--all` after a glob update.

- The `basic_principles_pptx.py` script and the combined PPTX merge script
  (used to build `book_I_complete.pptx`) will need a Book II equivalent —
  same pattern, different source files.

---

## Recommended first prompt for the next session

Paste this to start a Book II session:

```
We're starting Book II of the Byrne Euclid Decks project. Book I is 100%
complete (48 PDFs + PPTXs, combined decks, parser gate passing).

Read the project memory (project_byrne_setup.md) and BOOK_II_KICKOFF.md
before doing anything. The jemmybutton reference PDF is the authoritative
source for wording, layout, color choices, and figure scale — check it
before writing each generator.

The pipeline (lualatex, packager.py, test_parser.py gate) is the same as
Book I. Book II title format is "Book II. Prop. N. Prob/Theor." — the
packager parser already handles this.

Start with generator_II01.tex (II.1 is a Theor. — a rectangle contained by
two straight lines). Follow the jemmybutton reference for the exact statement
wording and figure. Apply all Book I generator conventions from the memory.
Build it, verify the PDF visually, then we will proceed proposition by
proposition.
```

---

## Files to read at the start of a Book II session

1. `memory/project_byrne_setup.md` — conventions, parser rules, PPTX scheme
2. `memory/feedback_jemmybutton_reference.md` — the "always check jemmybutton" rule
3. `STATUS.md` — current build status (add II.xx rows as you go)
4. `packager/PPTX_PLAN.md` — PPTX packager architecture (unchanged for Book II)
5. jemmybutton PDF pages for Book II — visual reference before each generator
