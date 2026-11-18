import sys, docx
from docx.shared import RGBColor
from docx.text.paragraph import Paragraph
from docx.table import Table
src, out = sys.argv[1], sys.argv[2]
d = docx.Document(src); RED = RGBColor(0xFF, 0, 0)
lines = [f'# 저자 작성·확인 목록 ({src}의 붉은 글씨 자동 추출)', '', '붉은 글씨를 채우거나 확인한 뒤 검정으로 바꾸거나 삭제하세요. 위치는 직전 제목 기준입니다.', '']
def red(p): return ''.join(r.text for r in p.runs if r.font.color is not None and r.font.color.type is not None and r.font.color.rgb == RED).strip()
sec, n = '표지', 0
for el in d.element.body.iterchildren():
    tag = el.tag.split('}')[1]
    if tag == 'p':
        p = Paragraph(el, d)
        if p.style.name.startswith('Heading'): sec = p.text.strip(); continue
        t = red(p)
        if t: n += 1; lines.append(f'{n}. **[{sec}]** {t}')
    elif tag == 'tbl':
        seen = set()
        for row in Table(el, d).rows:
            for c in row.cells:
                for p in c.paragraphs:
                    t = red(p).replace('\n', ' ')
                    if t and t not in seen:
                        seen.add(t); n += 1; lines.append(f'{n}. **[{sec} · 표/그림 칸]** {t}')
open(out, 'w', encoding='utf-8').write('\n'.join(lines) + '\n'); print(n)
