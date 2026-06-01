# Byrne's Euclid — Book I Decks

A recreation of **Oliver Byrne's 1847 colour edition of Euclid's *Elements*** as
modern, self-contained slide decks (PDF + PowerPoint), one per proposition.

Byrne's famous innovation was to replace letter-labels in geometric proofs with
**coloured figures referenced inline in the prose** — "the square on ▬ equals
the sum of the squares on ▬ and ▬". This project reproduces that style for all
of **Book I** (48 propositions + the Basic Principles: definitions, postulates,
axioms, elucidations, symbols).

## What's here

| Output | File(s) |
|---|---|
| 48 proposition PDFs | `build/I01.pdf … I48.pdf` |
| Combined Book I PDF | `build/book_I_complete.pdf` (65 pp) |
| 48 proposition PPTX decks | `build/I01.pptx … I48.pptx` |
| Basic Principles PPTX | `build/basic_principles.pptx` |
| Combined Book I PPTX | `build/book_I_complete.pptx` (113 slides) |

The PPTX decks use a **fit-to-fill layout**: editable prose as real text runs,
Byrne's coloured symbols as inline images, dynamically sized to fill each slide,
with a tricolour title rule as a nod to Byrne's colour method.

## Repository layout

```
generators/      LaTeX source — one generator_I<NN>.tex per proposition (the ONLY files to edit)
packager/        packager.py (tex → editable PPTX), merge/basic-principles/test scripts, PPTX_PLAN.md
build/           compiled PDFs + PPTX decks (committed deliverables)
lib/             shared snippets
Makefile         `make all` (PDFs), `make pptx` (decks), `make pptx-check` (parser gate)
STATUS.md        per-proposition status + next steps
BOOK_II_KICKOFF.md   notes for extending to Book II
```

## Building from source

**Requirements**
- TeX Live / MacTeX **2026** with LuaHBTeX + MetaPost, and the
  [`byrne`](https://ctan.org/pkg/byrne) package (ships in TeX Live 2026)
- [EB Garamond](https://github.com/octaviopardo/EBGaramond12) font
- Ghostscript, Python 3 with `python-pptx` and `Pillow`
- (Optional, for visual QA) LibreOffice — headless `soffice` to render PPTX

**Commands**
```bash
make all          # build all proposition PDFs (lualatex per generator)
make pptx-check   # parser sanity gate — must print ALL PASS
make pptx         # build all PPTX decks + combined deck
```

**Configuration** — `packager/packager.py` has a few hardcoded tool paths near
the top (TeX Live 2026 binary dir, Ghostscript, the EB Garamond font directory).
Adjust them for your system if they differ.

## Author

Created by **Suhas Bhat** ([@suhasbhat111](https://github.com/suhasbhat111)), 2026.

Developed with **Claude Code** (v2.1.145, Anthropic), using Claude Opus 4.8 and
Claude Sonnet 4.6.

## Credits & acknowledgements

- **Oliver Byrne** — *The First Six Books of the Elements of Euclid* (London,
  1847), the public-domain original this recreation is based on.
- **Sergey Slyusarev** (a.k.a. *jemmybutton*) — author of the
  [`byrne`](https://ctan.org/pkg/byrne) LaTeX + MetaPost package
  ([github.com/jemmybutton/byrne-latex](https://github.com/jemmybutton/byrne-latex))
  and of the full LaTeX recreation of Byrne's Euclid
  ([github.com/jemmybutton/byrne-euclid](https://github.com/jemmybutton/byrne-euclid)).
  **This project depends entirely on his `byrne` package** to draw every coloured
  figure and inline symbol, and used his recreation as the typographic and figure
  reference throughout. His work is licensed **GPL-3.0**; it is *not* redistributed
  here — install the package from CTAN/TeX Live. Enormous thanks to him: these
  decks would not exist without his package. 🙏

## License

The author's own contributions — the per-proposition LaTeX **generators** and the
Python **packaging code** — are dedicated to the public domain under
[**CC0 1.0**](LICENSE), and come **with no warranty of any kind**. The underlying
text and geometry are from Byrne's 1847 *Elements*, which is itself public domain.

The coloured figures and inline symbols are **rendered by the `byrne` package
(GPL-3.0) by Sergey Slyusarev (jemmybutton)**, which is **not included or
redistributed** here — you install it yourself from CTAN/TeX Live.

## Notice & good-faith policy

This is a **personal, non-commercial project, created purely for learning**. It is
shared freely and in good faith. The author **does not seek, accept, or pay any
money** in connection with it, and claims no ownership over the public-domain
source material or over the third-party tools it relies on.

**If you are a rights holder — or anyone with a genuine concern about any part of
this repository — please open a GitHub issue. The material in question will be
amended or removed promptly and without argument.** Nothing here is intended to
infringe anyone's rights or to compete with the works it gratefully builds upon.
