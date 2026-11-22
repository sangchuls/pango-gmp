import docx, glob, copy, re
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
SRC=glob.glob('/root/.claude/uploads/b5e6f573-7b81-5e78-b347-81f5624c7faa/be7335e7-*.docx')[0]
d=docx.Document(SRC)
def find(prefix, nth=0):
    hits=[p for p in d.paragraphs if p.text.strip().startswith(prefix)]
    assert hits, prefix
    return hits[nth]
def fmt(p, indent=0.5, left=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    pf=p.paragraph_format; p.alignment=align; pf.first_line_indent=Cm(indent) if indent else None
    pf.left_indent=Cm(left) if left else None; pf.line_spacing=1.5; pf.space_after=Pt(2); pf.space_before=Pt(0)
def settext(p, text, label=None, **kw):
    for r in list(p.runs): r._r.getparent().remove(r._r)
    for h in p._p.findall(qn('w:hyperlink')): p._p.remove(h)
    if label:
        r=p.add_run(label+': '); r.font.bold=True
    r=p.add_run(text)
    for r in p.runs: r.font.color.rgb=RGBColor(0,0,0); r.font.size=None
    fmt(p, **kw); return p
def after(p, text, label=None, **kw):
    np_=OxmlElement('w:p'); p._p.addnext(np_)
    q=docx.text.paragraph.Paragraph(np_, p._parent); return settext(q, text, label, **kw)
def drop(p): p._p.getparent().remove(p._p)
def block(first_prefix, last_prefix, items):
    """first~last 문단을 items(텍스트 또는 (라벨,텍스트,kw))로 교체"""
    ps=d.paragraphs; i=[k for k,p in enumerate(ps) if p.text.strip().startswith(first_prefix)][0]
    j=[k for k,p in enumerate(ps) if k>=i and p.text.strip().startswith(last_prefix)][0]
    olds=ps[i:j+1]; cur=None
    for n,it in enumerate(items):
        lab,txt,kw=(it if isinstance(it,tuple) else (None,it,{}))
        cur=settext(olds[0],txt,lab,**kw) if n==0 else after(cur,txt,lab,**kw)
    for p in olds[1:]: drop(p)
# ---------- 1.2
block('· 목적:','· RQ3',[
 '본 연구의 목적은 초기 판단 오류로 인한 설계 변경 리스크를 줄이기 위해, 바이브 코딩으로 개발한 검토 앱을 초기 계획설계에 도입하는 방식과 원리를 분석하는 것이다. 이를 위해 다음 세 가지 연구 질문을 설정한다.',
 ('RQ1','앱은 전문 지식을 어떻게 로직화하고, 지식을 명시화·검증·공유하는가?',{'indent':0,'left':0.5}),
 ('RQ2','기존 검토 방식(유추·직관·업체의 비공식 도움)에 대비하여 어떤 점이 개선되었는가?',{'indent':0,'left':0.5}),
 ('RQ3','이후 개발 단계에서 앱들을 통합하여 지식 확장의 완성도를 높이려면 어떤 방향으로 나아가야 하는가?',{'indent':0,'left':0.5})])
# ---------- 1.3 (범위 문장 정합: 효과 검증·통합 구현은 이후 단계)
settext(find('연구의 방법은 사례조사'),'연구의 방법은 사례조사와 문헌조사 방법을 기반으로 한다. 연구 대상은 저자가 개발해 웹으로 배포하여 실무에 적용해 보고 있는 시범 적용(pilot) 앱 3가지이다. 본 연구는 구체적인 앱의 구성 원리를 분석하고, 앱의 효과 검증과 이를 통합한 프레임워크 구현은 이후 단계의 목표로 한다.')
# ---------- 2.2 그림 1 언급
p=find('데이터센터의 기본 공간 구성'); settext(p,p.text.replace('운영·휴게 공간이다.','운영·휴게 공간이다(그림 1).'))
# ---------- 2.3 표 1 언급
p=find('선행연구에서는 기획설계'); settext(p,p.text.replace('확인되지 않았다.','확인되지 않았다(표 1).'))
# ---------- 4.1
block('명시화 : 전력 이중화','검증 : 산식과 코드',[
 ('지식 명시화','전력 이중화 종류와 용량, 단위 및 장비 사양 등을 입력한다.',{'indent':0}),
 ('도구 구현','브라우저에 정보값을 입력하면 실시간으로 전력 분배 시스템을 다이어그램으로 구현하고(그림 3 (a)), 이를 통해 기준층 평면 다이어그램 및 전기실 구성을 협의한다.',{'indent':0}),
 ('검증','산식과 코드를 LLM 등을 통해 크로스체크하여 검증한다.',{'indent':0})])
# ---------- 4.2
block('명시화 : 서버수량','(항온항습기의 높이',[
 ('지식 명시화','서버 수량, 유닛 배치, 기둥 간격, 랙 부하 등의 값을 입력하고, 실무에서 전문 분야가 사용하는 산식을 로직에 포함한다(예: 필요 풍량 8,000kW × 170CFM/kW × 1.699 × 1.1 = 2,541,704CMH). 결과는 입면·단면·평면으로 표현된다(그림 3 (b)).',{'indent':0}),
 ('검증','유사 사례의 검토 결과(항온항습기의 높이, 개수 등)와 비교하여 결과값의 오차 범위를 확인한다.',{'indent':0})])
# ---------- 4.3 (표 4와 정합: 저장소 내 별도 검증 절차 없음)
block('명시화 : 발전기 용량','다만 전문업체',[
 ('지식 명시화','발전기 용량, 제조사, 풍량·풍속값, 개구율을 입력한다.',{'indent':0}),
 ('도구 구현','앱을 통해 대수와 풍량에 따른 건축 개구 입면적의 변화를 확인하고, 그 결과로 요구되는 입면적의 규모와 건축물의 층고·단면을 확인할 수 있다(그림 3 (c)).',{'indent':0}),
 ('검증','전문업체의 카탈로그값을 기반으로 하며, 앱 안에 별도의 검증 절차는 없다. 다만 전문업체에서 제공한 정보의 업데이트는 분기별로 확인이 필요하다.',{'indent':0})])
# ---------- 4.4
block('세가지 개발 사례를','기존 방식 (유사사례',[
 '세 가지 개발 사례를 분석틀의 세부 요소별로 비교한 결과는 표 4와 같으며, 근거 표시는 사례 C ●, 사례 A ◐, 사례 B ○로 차이가 있었다.',
 '기존 방식(유사 사례를 통한 유추, 직관)에 대비해, 앱을 통해 정량적이고 일관된 검토를 신속하게 하여 기획설계에 대응할 수 있는 시스템과 지속적으로 업데이트할 수 있는 파이프라인을 구축한 것을 기존 대비 개선점으로 볼 수 있다. 선언적·절차적 지식은 명시화되었지만, 상황적(앱의 연결)·전략적(온톨로지 기반 종합 검토) 지식까지 시스템에 추가할 수 있도록 하는 것이 목표이며, 이후 연구에서는 프레임워크와 샘플 버전 개발까지를 목표로 한다.'])
# ---------- 5.1
settext(find('건축 공간에 가장 영향을'),'건축 공간에 가장 영향을 미치는 전문 분야 검토 요소를 찾아내고, 이를 앱으로 구현하여 기존의 직관·유추 의존도를 낮추었다. 근거와 이론에 기반한 정량적인 검토, 그리고 변화에 신속하게 대응할 수 있는 시스템을 앱으로 생성하였고, 이를 지속적으로 개발해 업데이트할 수 있도록 파이프라인을 구축하여 체계화하였다. 나아가 다음 단계는 이러한 전문 분야의 앱을 총괄하여 관리하고 통찰력 있는 의견을 제시할 수 있는 상위 단계의 지능형 앱을 개발하는 것을 목표로 하며, 이를 통해 초기 기획설계 오류로 인한 설계 변경 리스크를 감소시키고 정확도를 높이려고 한다.')
# ---------- 5.2
block('▶ 작성 메모','또한 분석틀은',[
 '향후 연구에서는 앱을 통합하고 제어할 수 있도록 온톨로지로 지식 구조화를 시도하고, LLM 기반 AI 에이전트가 앱을 호출해 산식을 기반으로 한 종합 결론과 검토 의견서를 도출할 수 있도록 한다(Chen 외, 2024; Dong 외, 2025). 또한 분석틀은 시설 유형에 종속되지 않아 산업시설·방송통신시설 등 다른 특수 시설로 확장이 가능하다.'])
# ---------- 5.3
block('데이터센터의 한정된','시범 버전이라는',[
 '데이터센터라는 한정된 영역의 사례를 기반으로 하여 다른 분야에 대한 연구도 필요하다는 점, 그리고 실제 공식 배포를 통한 다수의 피드백과 전문 분야 기관의 공식적인 확인 절차에 이르지 못한 시범 버전이라는 점이 본 연구의 한계이다.'])
# ---------- 참고문헌: 본문 미인용 IEA 삭제
drop(find('International Energy Agency'))
# ---------- 생성형 AI 활용 표기
block('▶ 작성 메모','· 실제 사용 범위',[
 '사례 앱 3개는 저자가 Claude Code(Anthropic)를 활용해 개발하였다. 본 보고서의 구성 정리, 그림·표 작성, 맞춤법과 문장 순서 교정에 생성형 AI(Claude)를 보조 도구로 사용하였으며, 본문은 저자가 작성하였다.'])
# ---------- 빈 문단 정리 + 본문 서식 통일
body_start=[k for k,p in enumerate(d.paragraphs) if p.text.strip()=='1. 서론'][0]
refs=[k for k,p in enumerate(d.paragraphs) if p.text.strip()=='참고문헌'][0]
for k,p in enumerate(list(d.paragraphs)):
    t=p.text.strip()
    if k<body_start: continue
    has_img=bool(p._p.xpath('.//w:drawing'))
    if not t and not has_img and p.style.name=='Normal':
        prev=p._p.getprevious()
        if prev is not None and prev.tag==qn('w:p'): drop(p); continue
for k,p in enumerate(d.paragraphs):
    t=p.text.strip()
    if p.style.name!='Normal' or not t or t.startswith(('자료:','표 ','그림 ')): continue
    for r in p.runs:
        if r.font.color.type is not None and str(r.font.color.rgb)=='8C8C8C': r.font.color.rgb=RGBColor(0,0,0); r.font.size=None
d.save('건축계획연구_과제_제출본.docx'); print('saved')
# ===== Pass 2: 표 명사 통일·판정 빈도 재계산
def cell_set(c, text):
    p=c.paragraphs[0]
    rs=p.runs
    if rs:
        rs[0].text=text
        for r in rs[1:]: r._r.getparent().remove(r._r)
    else: p.add_run(text)
    for extra in c.paragraphs[1:]: extra._p.getparent().remove(extra._p)
t2=[t for t in d.tables if t.rows[0].cells[0].text.strip()=='분석 관점' and t.rows[0].cells[1].text.strip()=='이론적 근거'][0]
cell_set(t2.rows[1].cells[2],'① 전문 검토 영역 ② 적정 변수·단위 입력 ③ 산식·규칙 근거 ④ 근거·가정 표시')
t4=[t for t in d.tables if t.rows[0].cells[0].text.strip()=='분석 관점' and t.rows[0].cells[2].text.strip()=='사례 A'][0]
fix={'● (타LLM기반)':'●(타 LLM 기반)','● (기존사례와 비교)':'●(기존 사례와 비교)','◐(개인,팀원일부)':'◐(개인, 팀원 일부)','◐(깃허브,엑셀)':'◐(깃허브, 엑셀)','● (엑셀,카달로그)':'●(엑셀, 카탈로그)','해당없음':'해당 없음','⑤ 구현 형태':'⑤ 구현 형태'}
for row in t4.rows:
    for c in row.cells:
        v=c.text.strip()
        if v in fix: cell_set(c,fix[v])
cnt={}
for ci in (2,3,4):
    vals=[t4.rows[r].cells[ci].text.strip() for r in range(1,len(t4.rows)-1)]
    vals=[v for k,v in enumerate(vals) if k!=4]  # ⑤ 제외
    f=sum(v.startswith('●') for v in vals); h=sum(v.startswith('◐') for v in vals); o=sum(v.startswith('○') for v in vals); na=sum(v.startswith('해당') for v in vals)
    cnt[ci]=(f,h,o,na)
last=t4.rows[-1]
cell_set(last.cells[0],'판정 빈도'); cell_set(last.cells[1],'● / ◐ / ○')
for ci in (2,3,4):
    f,h,o,na=cnt[ci]; cell_set(last.cells[ci],'%d / %d / %d'%(f,h,o)+(' (해당 없음 %d)'%na if na else ''))
print('판정 빈도',cnt)
src=[p for p in d.paragraphs if p.text.startswith('자료: 저자 작성 (● 충족')][0]
settext(src,'자료: 저자 작성 (● 충족, ◐ 부분 충족, ○ 미흡; ⑤와 해당 없음은 판정에서 제외; 2026.11 버전 기준)',indent=0,align=WD_ALIGN_PARAGRAPH.CENTER)
for r in src.runs: r.font.size=Pt(8.5)
# 4.4 본문을 표 4와 일치
p=find('세 가지 개발 사례를'); settext(p,p.text.replace('근거 표시는 사례 C ●, 사례 A ◐, 사례 B ○로 차이가 있었다.','근거·가정 표시는 사례 C가 충족(●), 사례 A와 B가 부분 충족(◐)으로 차이가 있었다.'),indent=0.5)
d.save('건축계획연구_과제_제출본.docx'); print('saved2')
# 표 3 첫 열 폭
t3=[t for t in d.tables if t.rows[0].cells[0].text.strip()=='사례'][0]
from docx.shared import Cm
for row in t3.rows:
    for c,w in zip(row.cells,[1.3,3.0,4.3,2.6,5.4]): c.width=Cm(w)
grid=t3._tbl.tblGrid
for gc,w in zip(grid.findall(qn('w:gridCol')),[1.3,3.0,4.3,2.6,5.4]): gc.set(qn('w:w'),str(int(w*567)))
d.save('건축계획연구_과제_제출본.docx'); print('saved3')
