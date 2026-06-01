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

## Credits

- **Oliver Byrne** — *The First Six Books of the Elements of Euclid* (London,
  1847), the public-domain original this recreation is based on.
- The [`byrne`](https://ctan.org/pkg/byrne) LaTeX/MetaPost package and the
  **jemmybutton** vector recreation of Byrne's Euclid, used as the typographic
  and figure reference throughout. (No upstream files are redistributed here.)

## License

Dedicated to the public domain under [**CC0 1.0**](LICENSE). The source material
(Byrne 1847) is itself public domain; this recreation adds no restrictions — use
it freely, no attribution required (though credit is always appreciated).
