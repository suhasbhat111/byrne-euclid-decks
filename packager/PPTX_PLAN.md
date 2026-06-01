# PPTX Generation Plan — implementation spec for packager.py v2

**Author:** planned by Opus 4.8, to be implemented by a cheaper model.
**Status:** approved scheme = refined Approach A (phrase-coalesced editable text + inline symbol images).

---

## 0. Goal & editability contract

Produce one editable `.pptx` per Book I proposition (I01–I48) plus the Basic
Principles, matching the look of the existing `build/I<NN>.pdf`.

"Editable" means, concretely:
- **Prose is editable as phrases.** Each maximal run of consecutive words (and
  the spaces between them) on a flowed line is ONE text box, not one box per
  word. Click a sentence fragment, retype it.
- **Inline Byrne symbols** (colored lines / circles / triangles / angles that
  appear *between* words in the proof) remain **images** — PPTX cannot embed an
  image inline in flowing text, so this is unavoidable and accepted.
- **Title** is an editable text box.
- Known limitation (accepted): editing a phrase box to be longer than its
  measured width will not reflow or push the trailing symbol — true reflow is
  impossible in PPTX. Good enough for fixing wording/typos.

Existing `packager.py` has solid *rendering* machinery (symbol rendering, font
metrics, line-flow, placement) — KEEP it. But the **parser is broken on all 48**,
not just the 38 untagged ones. It was eyeball-validated on the rendered I01/I04
output, which hid dirty extraction underneath. **Measured ground truth (run the
parser yourself to confirm — see §5 gate):**

| Defect | Count | Generators |
|---|---|---|
| Title leaks raw LaTeX (`~`, `\enspace`) | **48/48** | all |
| Statement contains unstripped structure (`\begin{minipage}`, `\vspace`, …) | **46/48** | all but I44/I45 (which are worse — see below) |
| Proof comes out EMPTY | **9** | I11, I22, I23, I26, I31, I42, I44, I45, I46 |
| Figure body EMPTY (two-arg `\defineNewPicture`) | **2** | I44, I45 |

So Task A is effectively a **parser rewrite**, not a tweak. Symbol rendering /
layout / placement stay as-is.

---

## 1. Task A — Rewrite `parse_generator` (fix all defects below)

The split rule, verified against the source:

1. **Title** = contents of the first `\textbf{...}` (already done).
2. **Statement** = text starting at the first `\itshape` (all 48 have exactly
   one), up to the proof-start marker.
3. **Proof-start marker**, in priority order (verified against all 48):
   - `% PROOF` comment tag (I01–I10), else
   - `\upshape` (35 generators — the THEOR single-minipage layout), else
   - the `\begin{center}` that follows the statement/figure minipages.
     This is the **PROB-without-tag** layout used by exactly these 8:
     **I11, I22, I23, I31, I42, I44, I45, I46** — they have NO `%PROOF`, NO
     `\upshape`, and NO `\bigskip`, so `\begin{center}` is the only boundary.
   - (Do NOT rely on `\bigskip` as the boundary — those 8 have zero of them.)
4. **Proof** = from proof-start to the first of `\begin{flushright}` /
   `\end{minipage}` / `\end{document}`.
   - Two structural families, for reference:
     - **PROB** (I01–10, I11, I22, I23, I31, I42, I44, I45, I46): statement in
       an `\itshape` minipage, proof in a `\begin{center}` block → proof is
       CENTERED (`proof_centered=True`, already keyed off `\begin{center}`).
     - **THEOR** (the `\upshape` set): figure-left + one text minipage holding
       title, `\itshape` statement, `\upshape` proof → proof is LEFT-aligned.
5. **Two-page special case:** only `generator_I26.tex` contains `\newpage`.
   Split the proof on `\newpage` into two case-blocks → emit TWO proof slides
   ("Proof — Case 1", "Proof — Case 2"). All other props: one statement slide +
   one proof slide.

### Concrete defects to fix (each verified present today)

- **B1 — Title cleanup (48/48).** Today titles come out as
  `'Book~I.\\enspace Prop. I. Prob.'`. Strip to plain text: `~`→space,
  `\enspace`/`\,`/`\;`/`\quad`→space, `\ `→space, `\.`→`.`, then remove any
  residual `\cmd`. Target exactly: `Book I. Prop. I. Prob.` (or `… Theor.`).
- **B2 — Statement isolation (46/48).** Statement currently includes the title
  `\begin{center}`, the figure minipage, `\vspace`, `\noindent`,
  `\begin/\end{minipage}`, `\centering`, `\hfill`, `\drawCurrentPicture`. Fix by
  extracting ONLY the text inside the `\itshape` region (from `\itshape` to its
  matching `\end{minipage}` for the two-minipage PROB layout, or to the
  proof-start marker for THEOR), then strip the structural commands listed below.
  The figure minipage (with `\drawCurrentPicture`) must NOT leak into statement.
- **B3 — Empty proof (9: I11,I22,I23,I26,I31,I42,I44,I45,I46).** Caused by the
  missing `\begin{center}` proof-marker (8 PROB-without-tag) and `\newpage`
  (I26). Implement the marker priority above. **Search for the proof marker
  strictly AFTER the `\itshape` position** so the *title* `\begin{center}` (which
  precedes `\itshape`) is never mistaken for the proof. (PROB-without-tag props
  have two `\begin{center}` — title then proof; this rule selects the right one.)
- **B4 — Two-arg figure → empty body (I44, I45).** Source uses
  `\defineNewPicture[sfA][sfB]{…}`; the regex `\[([^\]]*)\]?\s*\{` matches only
  ONE bracket, fails on the second, returns `None` → body `''`, scale wrong, and
  the statement falls back to the file's comment header. Fix the regex to consume
  TWO optional `[..]` groups. **Semantics matter:** `sfA` = inline-symbol scale,
  `sfB` = main-figure scale. Thread BOTH: pass `sfB` to `_figure_tex` (main
  figure) and `sfA` to `cached_render` (inline symbols). Single-arg props keep
  `sfB = defaultScale = 1`.
- **B5 — `\upshape`/`\itshape` glued to next word (e.g. I32 `\upshapeThrough`).**
  When a font command abuts a letter, the tokenizer reads `\upshapeThrough` as
  one unknown command and DROPS the word "Through". Strip `\itshape`/`\upshape`
  (and any bare `\noindent`/`\centering`) as whole commands BEFORE tokenizing,
  inserting a space at the boundary.
- **B6 — Overflow truncation.** `build_slide` silently stops placing lines once
  `y > bottom` → long proofs are cut with no warning. At-risk (many `\\` breaks):
  I47, I48, I21, I40, I35, I24. Mitigate: measure total proof height; if it
  exceeds the content box, step `PROOF_PT` down (14→12→11) until it fits, OR at
  minimum **print a loud `OVERFLOW` warning** naming the prop so QA catches it.
  Do NOT ship silent truncation.

**Structural commands to strip** before tokenizing (extend `clean_*`):
`\begin{minipage}[..]{..}` / `\end{minipage}`, `\noindent`, `\centering`,
`\vspace{..}`, `\hfill`, `\itshape`, `\upshape`, `\medskip`, `\bigskip`,
`\linewidth` math, and the leading `\textbf{title}`. Keep `\drawUnitLine`,
`\drawAngle`, `\drawLine`, `$...$`, the named back-refs (`\triangleABC`),
`\therefore`, `\perp`, citation `(...)`.

> Note: `proof_centered` already keys off `\begin{center}` in the proof, which
> is the correct test — PROB proofs are centered, and THEOR props that wrap their
> proof in `\begin{center}` (e.g. I32, I47) are correctly centered too. Do NOT
> hard-code "THEOR = left".

---

## 2. Task B — Phrase coalescing (the requested refinement)

Today `place_line` emits one zero-margin textbox **per word**. Change it so that,
**within each flowed line**, consecutive non-pic LToks (`text` + `space`) are
merged into a **single text box**, broken only by `pic` tokens.

Algorithm in `place_line` (or a small pre-pass per line):
1. Walk the line's LToks left to right, tracking running `x`.
2. Accumulate a buffer of consecutive `text`/`space` LToks.
3. When a `pic` is hit (or line ends): flush the buffer as ONE textbox spanning
   from the buffer's start-x to its end-x; then place the pic; continue.
4. Flush trailing buffer at line end.

**Mixed bold/italic inside one buffer:** do NOT split the box. Build the textbox
with multiple runs in the single paragraph — one `run` per LTok with its own
`font.bold`/`font.italic` (python-pptx supports many runs per paragraph). Width =
sum of the LTok widths; `word_wrap=False`, zero margins (as today).

Result per line: `[phrasebox] [pic] [phrasebox] [pic] ...`. Pure-prose lines
(statement body, concluding sentences) collapse to a single editable box.

Acceptance: re-render I01; slide 2 should have far fewer than 31 text shapes
(expect ~one box per prose run between symbols), and visually match the PDF.

---

## 3. Task C — Basic Principles (5 decks or 1 deck)

The Basic Principles sources (definitions / postulates / axioms / elucidations /
symbols) are list-structured, NOT statement/proof. Two acceptable options —
**pick the image route unless the text route is quick**:
- **Preferred (simple, robust):** one slide per page of each
  `build/basic_principles_*.pdf`, placed as a full-slide image (use
  `pdftocairo -png -r 200`). Reference material; fidelity > editability here.
- Optional later: a dedicated text builder if editable basic-principles text is
  wanted.

Deliver as a single `basic_principles.pptx` (slides in canonical order:
definitions, postulates, axioms, elucidations, symbols).

---

## 4. Task D — Batch driver + Makefile

- Add a loop (in `packager.py main` via `--all`, or a small `build_all.py`) that
  runs the packager over `generators/generator_I*.tex` → `build/I<NN>.pptx`.
- Add `make pptx` target that builds all 48 + basic_principles.pptx.
- Keep the `pic_cache` so re-runs are fast.

---

## 5. Acceptance — automated gate (must PASS before any visual QA)

This exact script encodes the defects found in §1. It must print `ALL PASS`
with no listed generators. Run it after the parser rewrite, before building
PPTX. (It is the same probe used to find the defects — do not weaken it.)

```python
# packager/test_parser.py  — run: python3 test_parser.py
import packager as P, re
bad = {'empty_proof':[], 'empty_body':[], 'dirty_stmt':[], 'bad_title':[]}
TITLE_RE = re.compile(r'^Book [IVXLCDM]+\. Prop\. [IVXLCDM]+\. (Prob|Theor)\.$')
for n in range(1,49):
    t=f'I{n:02d}'; d=P.parse_generator(f'../generators/generator_{t}.tex')
    if not d['proof'].strip():                                   bad['empty_proof'].append(t)
    if len(d['body'])==0:                                        bad['empty_body'].append(t)
    if re.search(r'\\(begin|end|vspace|noindent|centering|hfill|minipage|drawCurrentPicture)', d['statement']): bad['dirty_stmt'].append(t)
    if not TITLE_RE.match(d['title']):                           bad['bad_title'].append(t)
ok=all(not v for v in bad.values())
print('ALL PASS' if ok else 'FAIL')
[print(f'  {k}: {v}') for k,v in bad.items() if v]
```

Today this prints: empty_proof=9, empty_body=2, dirty_stmt=46, bad_title=48.
**Target: all empty.**

Then, build + visual gate:
- [ ] All 48 + Basic Principles `.pptx` build without exceptions; no `OVERFLOW`
      warning printed (see B6).
- [ ] I26 yields a 3-slide deck (statement + 2 proof cases).
- [ ] I44 & I45 show a real main figure (regression test for B4).
- [ ] Phrase coalescing: I01 proof slide has materially fewer than 31 text shapes.
- [ ] Visual diff vs source PDF (via LibreOffice, §5b-A) for at least:
      I01 (prob), I13 (theor, left), I32/I47 (theor, centered), I26 (two-case),
      I44 (two-arg figure).
- [ ] `basic_principles.pptx` present, 5 categories in canonical order.

## 5b. Recommended improvements (Opus second-pass)

These raise quality/robustness for little cost. Do A first — without it the
coder is working blind.

- **A. Install a render-to-image path for QA (DO THIS FIRST).** There is no
  installed tool to view a `.pptx` locally (no LibreOffice/soffice; pdftocairo
  and gs only do PDF). Install LibreOffice for a headless QA loop:
  `brew install --cask libreoffice`, then verify each deck with
  `soffice --headless --convert-to pdf <deck>.pptx` → `pdftocairo -png` →
  visual-diff against the source `build/I<NN>.pdf`. Without this the coder
  cannot see its own output and will iterate blindly.
- **B. Paragraph-level coalescing for symbol-free runs (editability win).**
  Beyond per-line phrase coalescing (Task B): when a *contiguous block of
  flowed lines contains NO pic tokens* (e.g. the whole statement, or a plain
  concluding sentence), emit it as ONE `word_wrap=True` multi-line text box
  spanning those lines, instead of one fixed box per line. That paragraph then
  reflows naturally when edited in PowerPoint — true editability for the prose
  that has no inline symbols. Lines that DO contain symbols stay as fixed
  per-line `[phrase][pic][phrase]` boxes.
- **C. Use real font metrics for vertical alignment.** Replace the magic
  `cap_h ≈ 0.72×pt` and `baseline = 0.78×line_h` constants with PIL
  `ImageFont.getmetrics()` (returns ascent, descent). Align symbol baselines to
  the true text baseline. Removes the main source of per-proposition tuning.
- **D. Bump symbol render DPI.** `RENDER_DPI` 150 → ~300 for inline symbols
  (they are small; cost is negligible, cache absorbs re-runs) so they stay crisp
  on retina/projector. Keep the figure at ~200+.
- **E. Normalize slide layout (note, not a change):** the packager always puts
  figure-left / text-right regardless of PROB-vs-THEOR. That is fine and
  intended for slides — do NOT try to mirror each PDF's left/right placement.

## 6. Out of scope / do NOT do
- Do not re-author or re-typeset proof prose from upstream (licensing — see
  project memory). Parse only the existing generators.
- Do not rebuild PDFs. PPTX work reads generators (for text) + renders symbols.
- No combined Book I PPTX (per project decision — individual decks only).
