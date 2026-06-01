"""Parser gate — must print ALL PASS before building any PPTX."""
import packager as P, re, sys

TITLE_RE = re.compile(r'^Book [IVXLCDM]+\. Prop\. [IVXLCDM]+\. (Prob|Theor)\.$')
bad = {'empty_proof': [], 'empty_body': [], 'dirty_stmt': [], 'bad_title': []}

for n in range(1, 49):
    t = f'I{n:02d}'
    d = P.parse_generator(f'../generators/generator_{t}.tex')
    if not d['proof'].strip():
        bad['empty_proof'].append(t)
    if len(d['body']) == 0:
        bad['empty_body'].append(t)
    if re.search(r'\\(begin|end|vspace|noindent|centering|hfill|minipage|drawCurrentPicture)',
                 d['statement']):
        bad['dirty_stmt'].append(t)
    if not TITLE_RE.match(d['title']):
        bad['bad_title'].append(t)

ok = all(not v for v in bad.values())
print('ALL PASS' if ok else 'FAIL')
for k, v in bad.items():
    if v:
        print(f'  {k}: {v}')

if not ok:
    sys.exit(1)
