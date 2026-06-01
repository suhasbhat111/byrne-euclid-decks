# Project Status

Format: `ID | status | notes`
Status values: `not-started` | `built` | `verified`

## Basic Principles (Book I)

|ID|Status|Notes|
|---|---|---|
|BP.Definitions|verified|35 defs, jemmybutton figures, fig-left layout, 5 pages|
|BP.Postulates|verified|3 postulates, text-only, 1 page|
|BP.Axioms|verified|12 axioms, Axiom XII with parallel-postulate figure, fig-left, 3 pages|
|BP.Elucidations|verified|jemmybutton prose + 2 figures (angle vertex, multi-line vertex), 4 pages|
|BP.Symbols|verified|math symbols + colored angle/point pictograms + abbreviations (def./post./ax./hyp./const./Q.E.D.), 3 pages|

## Book I Propositions

|ID|Status|Notes|
|---|---|---|
|I.1|verified|jemmybutton-matched: stmt left / fig right, Book I title, EB Garamond, citations (post. III / ax. I / def. 15)|
|I.2|verified|jemmybutton-matched: stmt left / fig right, construction circles, citations (post. I / pr. I.1 / def. 15 / const.)|
|I.3|verified|jemmybutton-matched: stmt left / fig right, circle cut-off, citations (pr. I.2 / post. III / def. 15 / ax. I)|
|I.4|verified|jemmybutton-matched: fig left 45% / text right 52%, scale [1/5], one-para proof, no letter refs, (ax. X)|
|I.5|verified|isosceles base angles, inline sub-triangle figures|
|I.6|verified|equal angles => equal sides, proof by contradiction|
|I.7|verified|unique triangle on base, aligned brace proof|
|I.8|verified|SSS congruence, rotated inline angle symbols|
|I.9|verified|bisect angle, kite figure|
|I.10|verified|bisect line, equilateral triangle construction|
|I.11|verified|jemmybutton-matched: stmt left / fig right, equilateral triangle construction, perpendicular from point on line, (pr.~I.3 / pr.~I.1 / pr.~I.8 / def.~10)|
|I.12|verified||
|I.13|verified||
|I.14|verified||
|I.15|verified||
|I.16|verified||
|I.17|verified||
|I.18|verified||
|I.19|verified||
|I.20|verified||
|I.21|built|interior point lines < two sides, angle greater; fig-left layout|
|I.22|built|construct triangle from three lines; PROB layout, two circles|
|I.23|built|make angle equal to given angle; PROB layout|
|I.24|built|greater included angle => greater base; two-triangle fig, custom angle macro|
|I.25|built|greater base => greater angle (converse I.24); two-triangle fig|
|I.26|built|AAS congruence, two cases; two-page PDF|
|I.27|built|alternate angles equal => parallel|
|I.28|built|external = internal opposite => parallel|
|I.29|built|parallel => alternate angles equal; uses ax. XII|
|I.30|built|lines parallel to same line => parallel to each other|
|I.31|built|draw line through point parallel to given line; PROB layout|
|I.32|built|exterior angle = sum of remote interior; angle sum = 2 right angles|
|I.33|built|lines joining equal parallel lines are equal and parallel|
|I.34|built|opposite sides/angles of parallelogram equal; diagonal bisects|
|I.35|built|parallelograms on same base, same parallels are equal; colored polygon areas|
|I.36|built|parallelograms on equal bases, same parallels are equal|
|I.37|built|triangles on same base, same parallels are equal; colored polygon areas|
|I.38|built|triangles on equal bases, same parallels are equal|
|I.39|built|equal triangles on same base, same side => same parallels; proof by contradiction|
|I.40|built|equal triangles on equal bases, same side => same parallels; proof by contradiction|
|I.41|built|parallelogram = twice triangle on same base between same parallels|
|I.42|built|construct parallelogram = triangle with given angle; PROB layout|
|I.43|built|complements of parallelograms about diagonal are equal|
|I.44|built|apply parallelogram to line = triangle with given angle; scale 1/5 inline, 1/2 main|
|I.45|built|parallelogram = rectilinear figure with given angle; scale 1/6 inline, 1/3 main|
|I.46|built|construct square on given line; PROB layout|
|I.47|verified|Pythagorean theorem; complex figure with 3 squares, byLineStylize; proof set in \footnotesize to fit one A5 page (longest proof in Book I) — fixed 2026-05-31 blank-page-1 overflow|
|I.48|verified|converse of Pythagorean theorem (visually QA'd in full sweep)|

## Deliverables (Book I) — ALL COMPLETE

|Deliverable|Status|Location|
|---|---|---|
|48 proposition PDFs|✅ verified|`build/I01.pdf … I48.pdf` (visual sweep of all 48 done)|
|Combined Book I PDF|✅ done|`build/book_I_complete.pdf` (65 pages)|
|48 proposition PPTXs|✅ done|`build/I01.pptx … I48.pptx` (fit-to-fill design, all QA'd vs PDFs)|
|Basic Principles PPTX|✅ done|`build/basic_principles.pptx` (16 slides, image-per-page)|
|Combined Book I PPTX|✅ done|`build/book_I_complete.pptx` (113 slides)|

## How to rebuild
- PDFs: `make all` (lualatex per generator) — do NOT edit build/ directly
- PPTX: `make pptx` (runs packager.py --all + basic_principles_pptx.py + merge_pptx.py)
- Parser gate before PPTX: `make pptx-check` (must print ALL PASS)
- Visual QA loop: LibreOffice `soffice --headless --convert-to pdf X.pptx` → `pdftocairo -png` → compare to `build/I<NN>.pdf`

## Next steps / TODO
- [ ] Book II (props II.1–II.14) — not started. See `BOOK_II_KICKOFF.md` for the
      starting prompt and conventions. Packager change needed: extend `--all`
      glob to include `generator_II*.tex`; add Book II merge/basic-principles.
- [ ] (Optional polish) Multi-line `$\left.\begin{aligned}..\right\}$` equation
      groups (I07,I22,I35,I37,I38,I42) render as cramped inline images — could
      be laid out natively if desired.
- [ ] (Optional) Embed EB Garamond in PPTX for 100% portability (SIL OFL allows).
