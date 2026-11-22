import docx, copy, re, subprocess, sys, json
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
SRC='과제보고서_최종본.docx'; OUT=sys.argv[1]; PAGES=json.load(open(sys.argv[2])) if len(sys.argv)>2 else {}
SHORT='건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석'
EN='A Study on Micro-App Analysis for Expanding Knowledge in Early-Stage Architectural Planning: Focusing on Three Data Center Review Apps Developed with Vibe Coding'
d=docx.Document(SRC)
TW=Cm(16.6)
def fonts(rpr_owner, ascii_, east):
    rpr=rpr_owner.get_or_add_rPr() if hasattr(rpr_owner,'get_or_add_rPr') else rpr_owner
    rf=rpr.find(qn('w:rFonts'))
    if rf is None:
        rf=OxmlElement('w:rFonts'); rpr.insert(0,rf)
    for k in list(rf.attrib): del rf.attrib[k]
    for k,v in (('ascii',ascii_),('hAnsi',ascii_),('eastAsia',east),('cs',ascii_)): rf.set(qn('w:'+k),v)
# ---- 그림 1-1 교체 (장 이름 가운데 정렬)
_bl=d.element.body.xpath('.//a:blip')[0]
_part=d.part.related_parts[_bl.get(qn('r:embed'))]
_part._blob=open('../figures/fig_1_1_flow.png','rb').read()
# ---- 페이지 설정 (A4)
sec=d.sections[0]
sec.page_width=Cm(21.0); sec.page_height=Cm(29.7)
sec.left_margin=sec.right_margin=Cm(2.2); sec.top_margin=Cm(2.8); sec.bottom_margin=Cm(2.5)
sec.header_distance=Cm(1.4); sec.footer_distance=Cm(1.2)
# ---- 스타일
n=d.styles['Normal']; n.font.size=Pt(10.5); fonts(n.element.get_or_add_rPr(),'Times New Roman','바탕')
n.paragraph_format.line_spacing=1.5
for name,size,before,after in (('Heading 1',20,0,18),('Heading 2',13,16,6),('Heading 3',11.5,10,4)):
    st=d.styles[name]; st.font.size=Pt(size); st.font.bold=True; st.font.color.rgb=RGBColor(0,0,0)
    fonts(st.element.get_or_add_rPr(),'맑은 고딕','맑은 고딕')
    pf=st.paragraph_format; pf.space_before=Pt(before); pf.space_after=Pt(after); pf.line_spacing=1.3; pf.keep_with_next=True
    pf.alignment=WD_ALIGN_PARAGRAPH.LEFT
d.styles['Heading 1'].paragraph_format.page_break_before=True
# 기존 런 글꼴(eastAsia 맑은 고딕) → 바탕 (제목 제외)
for p in d.element.body.iter(qn('w:p')):
    ppr=p.find(qn('w:pPr')); ps=ppr.find(qn('w:pStyle')) if ppr is not None else None
    heading = ps is not None and ps.get(qn('w:val')) in ('Heading1','Heading2','Heading3')
    for r in p.findall(qn('w:r')):
        rpr=r.find(qn('w:rPr'))
        if rpr is None: continue
        rf=rpr.find(qn('w:rFonts'))
        if rf is not None:
            if heading: fonts(rpr,'맑은 고딕','맑은 고딕')
            else: fonts(rpr,'Times New Roman','바탕')
# ---- 본문 단락 서식 (들여쓰기·양쪽 정렬)
for p in d.paragraphs:
    if p.style.name!='Normal': continue
    t=p.text.strip()
    if len(t)>60 and p.alignment is None and p.paragraph_format.left_indent is None and not t.startswith(('자료:','주제어')):
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.first_line_indent=Cm(0.5)
        p.paragraph_format.space_after=Pt(2)
for p in d.paragraphs:
    m=re.match(r'^<표 (\d+-\d+)>\s*(.*)$',p.text.strip())
    if m and p.runs:
        p.runs[0].text='표 %s. %s'%(m.group(1),m.group(2)); p.runs[0].font.bold=False
        for r in p.runs[1:]: r._r.getparent().remove(r._r)
# 표: 9pt, 줄간격 1.15, 글자·선 검정, 음영 연한 회색
def shade(tc,fill):
    tcPr=tc.get_or_add_tcPr()
    for e in tcPr.findall(qn('w:shd')): tcPr.remove(e)
    if fill:
        sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),fill); tcPr.append(sh)
for tb in d.tables:
    tblPr=tb._tbl.tblPr
    if tb._tbl.xpath('.//w:drawing'):
        # 그림 배치용 표: 선·음영 없음
        for tag in ('w:tblStyle','w:tblLook','w:tblBorders'):
            for e in tblPr.findall(qn(tag)): tblPr.remove(e)
        bd=OxmlElement('w:tblBorders')
        for side in ('top','left','bottom','right','insideH','insideV'):
            e=OxmlElement('w:'+side); e.set(qn('w:val'),'nil'); bd.append(e)
        tblPr.append(bd)
        for row in tb.rows:
            for c in row.cells:
                shade(c._tc,None)
                tcPr=c._tc.get_or_add_tcPr()
                for e in tcPr.findall(qn('w:tcBorders')): tcPr.remove(e)
                for pp in c.paragraphs: pp.paragraph_format.line_spacing=1.0; pp.paragraph_format.first_line_indent=None
        continue
    for tag in ('w:tblStyle','w:tblLook'):
        for e in tblPr.findall(qn(tag)): tblPr.remove(e)
    for e in tblPr.findall(qn('w:tblBorders')): tblPr.remove(e)
    bd=OxmlElement('w:tblBorders')
    for side,val,sz,col in (('top','single','12','000000'),('left','nil','0','auto'),('bottom','single','12','000000'),('right','nil','0','auto'),('insideH','single','4','000000'),('insideV','single','4','000000')):
        e=OxmlElement('w:'+side); e.set(qn('w:val'),val); e.set(qn('w:sz'),sz); e.set(qn('w:space'),'0'); e.set(qn('w:color'),col); bd.append(e)
    tblPr.append(bd)
    for ri,row in enumerate(tb.rows):
        for c in row.cells:
            shade(c._tc,'D9D9D9' if ri==0 else None)
            tcPr=c._tc.get_or_add_tcPr()
            for e in tcPr.findall(qn('w:tcBorders')): tcPr.remove(e)
            for e in tcPr.findall(qn('w:vAlign')): tcPr.remove(e)
            va=OxmlElement('w:vAlign'); va.set(qn('w:val'),'center'); tcPr.append(va)
            if ri==0:
                for pp in c.paragraphs:
                    pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
                    for rr in pp.runs: rr.font.bold=True
            for p in c.paragraphs:
                p.paragraph_format.line_spacing=1.15; p.paragraph_format.first_line_indent=None
                p.paragraph_format.space_before=Pt(1); p.paragraph_format.space_after=Pt(1)
            for r in c._tc.iter(qn('w:r')):
                rpr=r.find(qn('w:rPr'))
                if rpr is None:
                    rpr=OxmlElement('w:rPr'); r.insert(0,rpr)
                for tag in ('w:color','w:sz','w:szCs','w:u'):
                    for e in rpr.findall(qn(tag)): rpr.remove(e)
                fonts(rpr,'Times New Roman','바탕')
                col=OxmlElement('w:color'); col.set(qn('w:val'),'000000'); rpr.append(col)
                sz=OxmlElement('w:sz'); sz.set(qn('w:val'),'18'); rpr.append(sz)
                szc=OxmlElement('w:szCs'); szc.set(qn('w:val'),'18'); rpr.append(szc)
# ---- 표지·차례 제거 대상 (앞 4개 단락 = 구 표지)
body=d.element.body
paras=d.paragraphs
abs_head=[p for p in paras if p.style.name=='Heading 1' and p.text.strip()=='국문초록'][0]
for p in paras[:4]: p._p.getparent().remove(p._p)
def newp(before, text='', size=11, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=0, font=('Times New Roman','바탕'), ls=1.3, italic=False):
    p=before.insert_paragraph_before('')
    p.alignment=align; pf=p.paragraph_format; pf.space_before=Pt(sb); pf.space_after=Pt(sa); pf.line_spacing=ls; pf.first_line_indent=None
    if text:
        lines=text.split('\n')
        for i,l in enumerate(lines):
            r=p.add_run(l); r.font.size=Pt(size); r.font.bold=bold; r.font.italic=italic; fonts(r._r.get_or_add_rPr(),*font)
            if i<len(lines)-1: r.add_break()
    return p
# 표지
cov=[]
cov.append(newp(abs_head,'연 구 보 고 서',14,False,sb=0,sa=0))
cov.append(newp(abs_head,'건축 초기 계획설계의 지식 확장을 위한\nMicro-App 분석',22,True,sb=110,sa=10,ls=1.5))
cov.append(newp(abs_head,'— 바이브 코딩으로 개발한 데이터센터 검토 앱 3개 사례를 중심 연구 —',13,False,sb=8,sa=24))
cov.append(newp(abs_head,EN.replace(': ',':\n',1),11.5,False,sb=0,sa=40,ls=1.4))
cov.append(newp(abs_head,'담당교수  김영주',13,False,sb=60,sa=36))
cov.append(newp(abs_head,'이 보고서를 건축계획연구 과제 보고서로 제출합니다.',12,False,sb=0,sa=30))
cov.append(newp(abs_head,'2026년 11월 23일',13,False,sb=0,sa=50))
cov.append(newp(abs_head,'한양대학교 공학대학원',15,True,sb=0,sa=6))
cov.append(newp(abs_head,'건축학전공',13,True,sb=0,sa=6))
cov.append(newp(abs_head,'임상철 (학번 2026160492)',13,False,sb=0,sa=0))
def sectpr(kind):
    sp=copy.deepcopy(d.element.body.find(qn('w:sectPr')))
    for tag in ('w:headerReference','w:footerReference','w:pgNumType','w:titlePg'):
        for e in sp.findall(qn(tag)): sp.remove(e)
    if kind:
        pg=OxmlElement('w:pgNumType'); pg.set(qn('w:fmt'),kind); pg.set(qn('w:start'),'1'); 
        # insert after pgMar
        mar=sp.find(qn('w:pgMar')); mar.addnext(pg)
    return sp
last_cov=cov[-1]._p; ppr=last_cov.get_or_add_pPr(); ppr.append(sectpr(None))
# 차례 (구역2 시작)
toc_title=newp(abs_head,'차      례',20,True,sb=0,sa=24,font=('맑은 고딕','맑은 고딕'))
# TOC 스타일
def tocstyle(name,sid,indent,bold,sb):
    st=d.styles.add_style(name,WD_STYLE_TYPE.PARAGRAPH); st.element.name_val=name.lower() if False else name
    st.base_style=d.styles['Normal']; st.paragraph_format.left_indent=Cm(indent); st.paragraph_format.space_before=Pt(sb); st.paragraph_format.space_after=Pt(0); st.paragraph_format.line_spacing=1.25
    st.paragraph_format.tab_stops.add_tab_stop(TW,WD_TAB_ALIGNMENT.RIGHT,WD_TAB_LEADER.DOTS)
    st.font.size=Pt(10.5); st.font.bold=bold
    return st
s1=tocstyle('TOC 1','TOC1',0,True,8); s2=tocstyle('TOC 2','TOC2',0.9,False,0)
for st,nm in ((s1,'toc 1'),(s2,'toc 2')):
    st.element.find(qn('w:name')).set(qn('w:val'),nm)
heads=[(p.style.name,p.text.strip()) for p in d.paragraphs if p.style.name in ('Heading 1','Heading 2')]
def norm(s): return re.sub(r'\s+','',s)
toc_ps=[]
for i,(sn,tx) in enumerate(heads):
    p=newp(abs_head,'',10.5,False,align=WD_ALIGN_PARAGRAPH.LEFT,sb=0,sa=0)
    p.style=(s1 if sn=='Heading 1' else s2)
    pf=p.paragraph_format; pf.line_spacing=None; pf.space_before=None; pf.space_after=None; pf.left_indent=None; pf.first_line_indent=None
    if i==0:
        for kind,txt in (('begin',None),('instr',' TOC \\o "1-2" \\h \\z \\u '),('separate',None)):
            r=p.add_run()
            if kind=='instr':
                it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text=txt; r._r.append(it)
            else:
                fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),kind); r._r.append(fc)
    p.add_run(tx); r=p.add_run(); r.add_tab(); p.add_run(str(PAGES.get(norm(tx),'0')))
    toc_ps.append(p)
r=toc_ps[-1].add_run(); fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),'end'); r._r.append(fc)
# 구역2 끝: 주제어 단락
kw=[p for p in d.paragraphs if p.text.startswith('주제어')][0]
kw._p.get_or_add_pPr().append(sectpr('lowerRoman'))
# 마지막 구역: 본문
last=body.find(qn('w:sectPr'))
for e in last.findall(qn('w:pgNumType')): last.remove(e)
pg=OxmlElement('w:pgNumType'); pg.set(qn('w:fmt'),'decimal'); pg.set(qn('w:start'),'1'); last.find(qn('w:pgMar')).addnext(pg)
# ---- 머리글/바닥글
def page_footer(footer):
    footer.is_linked_to_previous=False
    for p in footer.paragraphs[1:]: p._p.getparent().remove(p._p)
    p=footer.paragraphs[0]
    for r in list(p.runs): r._r.getparent().remove(r._r)
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    def run(txt=None,fld=None):
        r=p.add_run(txt or ''); r.font.size=Pt(10); fonts(r._r.get_or_add_rPr(),'Times New Roman','바탕')
        if fld:
            for kind in ('begin','instr','separate','text','end'):
                rr=p.add_run(); rr.font.size=Pt(10)
                if kind=='instr':
                    it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text=' PAGE '; rr._r.append(it)
                elif kind=='text': rr.text='1'
                else:
                    fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),kind); rr._r.append(fc)
    run('- '); run(fld=True); run(' -')
def title_header(header):
    header.is_linked_to_previous=False
    p=header.paragraphs[0]
    for r in list(p.runs): r._r.getparent().remove(r._r)
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(SHORT); r.font.size=Pt(9.5); fonts(r._r.get_or_add_rPr(),'Times New Roman','맑은 고딕')
    ppr=p._p.get_or_add_pPr(); b=OxmlElement('w:pBdr'); bt=OxmlElement('w:bottom')
    for k,v in (('val','single'),('sz','6'),('space','4'),('color','000000')): bt.set(qn('w:'+k),v)
    b.append(bt); ppr.append(b)
def empty(hf):
    hf.is_linked_to_previous=False
    for p in hf.paragraphs:
        for r in list(p.runs): r._r.getparent().remove(r._r)
S=d.sections
assert len(S)==3,len(S)
empty(S[0].header); empty(S[0].footer)
empty(S[1].header); page_footer(S[1].footer)
title_header(S[2].header); page_footer(S[2].footer)
d.save(OUT)
print('saved',OUT,len(heads))
