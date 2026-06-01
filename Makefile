LUALATEX = lualatex
FLAGS    = -interaction=nonstopmode -output-directory=../build

GENERATORS = $(wildcard generators/generator_*.tex)
PDFS       = $(patsubst generators/generator_%.tex,build/%.pdf,$(GENERATORS))

.PHONY: all pptx clean

all: $(PDFS)

build/%.pdf: generators/generator_%.tex
	cd generators && $(LUALATEX) $(FLAGS) $(<F)

pptx:
	cd packager && python3 packager.py --all
	cd packager && python3 basic_principles_pptx.py
	cd packager && python3 merge_pptx.py

pptx-check:
	cd packager && python3 test_parser.py

clean:
	rm -f build/*.pdf build/*.aux build/*.log build/*.out
