import docx, copy, re, io
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from PIL import Image
TPL='과제양식_원본.docx'; SRC='사용자_최종수정본.docx'; OUT='건축계획연구_기말과제_2026160492_임상철_최종.docx'
FIG='/home/user/pango-gmp/research/figures/'
d=docx.Document(TPL); body=d.element.body; els=list(body)
from docx.enum.text import WD_LINE_SPACING
bt=d.styles['Body Text'].paragraph_format; bt.line_spacing_rule=WD_LINE_SPACING.SINGLE
d.styles['절제목'].paragraph_format.space_before=Pt(6)
T={'cap_fig':els[22],'cap_tab':els[29],'tbl':els[30],'ref':els[41],'body':els[6],'empty':els[14]}
T={k:copy.deepcopy(v) for k,v in T.items()}
for el in els[4:-1]: body.remove(el)          # 샘플 본문 삭제 (제목표·1단 구역 유지)
sect=body[-1]
COLW=Cm(8.3)
def put(el): sect.addprevious(el); return el
def set_runs(p_el, parts):
    """parts: [(text, {'b':bool,'i':bool,'sz':pt})]  첫 run의 rPr 복제"""
    rs=p_el.findall(qn('w:r')); base=rs[0].find(qn('w:rPr')) if rs else None
    for r in p_el.findall(qn('w:r'))+p_el.findall(qn('w:hyperlink')): p_el.remove(r)
    for t,o in parts:
        r=OxmlElement('w:r')
        rpr=copy.deepcopy(base) if base is not None else OxmlElement('w:rPr')
        for tag in ('w:b','w:i','w:bCs','w:iCs'):
            for e in rpr.findall(qn(tag)): rpr.remove(e)
        bb=OxmlElement('w:b')
        if not o.get('b'): bb.set(qn('w:val'),'0')
        rpr.append(bb)
        if o.get('i'): rpr.append(OxmlElement('w:i'))
        if o.get('sz'):
            for tag in ('w:sz','w:szCs'):
                for e in rpr.findall(qn(tag)): rpr.remove(e)
                e=OxmlElement(tag); e.set(qn('w:val'),str(int(o['sz']*2))); rpr.append(e)
        r.append(rpr); te=OxmlElement('w:t'); te.set(qn('xml:space'),'preserve'); te.text=t; r.append(te); p_el.append(r)
    return p_el
def P(style, text, parts=None, align=None, first=None):
    p=d.add_paragraph(style=style); 
    if parts: 
        for t,o in parts:
            r=p.add_run(t); r.bold=o.get('b'); r.italic=o.get('i')
    else: p.add_run(text)
    if align is not None: p.alignment=align
    if first is not None: p.paragraph_format.first_line_indent=first
    for r in p.runs: r.font.name='맑은 고딕'; r._r.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'맑은 고딕')
    return p
def blank(): put(copy.deepcopy(T['empty']))
def caption(kind, text):
    el=copy.deepcopy(T['cap_fig'] if kind=='fig' else T['cap_tab']); set_runs(el,[(text,{})]); put(el)
def source(text):
    el=copy.deepcopy(T['cap_tab']); set_runs(el,[(text,{})])
    ppr=el.find(qn('w:pPr'))
    jc=ppr.find(qn('w:jc')) if ppr is not None else None
    if jc is not None: ppr.remove(jc)
    put(el)
def picture(blob_or_path, width=COLW):
    p=d.add_paragraph(style='Body Text'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent=Cm(0)
    p.paragraph_format.line_spacing=1.0
    src=io.BytesIO(blob_or_path) if isinstance(blob_or_path,bytes) else blob_or_path
    p.add_run().add_picture(src, width=width)
def table(rows, widths, left_cols=()):
    tb=copy.deepcopy(T['tbl']); trs=tb.findall(qn('w:tr')); hdr, bodyr=trs[0], trs[1]
    for tr in trs: tb.remove(tr)
    grid=tb.find(qn('w:tblGrid'))
    for g in list(grid): grid.remove(g)
    tw=sum(widths)
    for w in widths:
        g=OxmlElement('w:gridCol'); g.set(qn('w:w'),str(int(4710*w/tw))); grid.append(g)
    n=len(widths)
    for ri,row in enumerate(rows):
        src=copy.deepcopy(hdr if ri==0 else bodyr); tcs=src.findall(qn('w:tc'))
        first,mid,last=tcs[0],tcs[1],tcs[-1]
        for tc in tcs: src.remove(tc)
        for ci,val in enumerate(row):
            tc=copy.deepcopy(first if ci==0 else (last if ci==n-1 else mid))
            tc.find(qn('w:tcPr')).find(qn('w:tcW')).set(qn('w:w'),str(int(4710*widths[ci]/tw)))
            ps=tc.findall(qn('w:p'))
            for extra in ps[1:]: tc.remove(extra)
            set_runs(ps[0],[(val,{'b':ri==0})])
            if ri>0 and ci in left_cols:
                jc=ps[0].find(qn('w:pPr')).find(qn('w:jc'))
                if jc is not None: jc.set(qn('w:val'),'left')
            src.append(tc)
        tb.append(src)
    put(tb)
def ref(text):
    el=copy.deepcopy(T['ref']); parts=[]
    for k,seg in enumerate(text.split('*')):
        if seg: parts.append((seg,{'i':k%2==1}))
    set_runs(el,parts); put(el)

# ---------- 제목 표
t0=d.tables[0]
cell=t0.rows[1].cells[0]
def cellp(p,text,size=None):
    for r in p.runs[1:]: r._r.getparent().remove(r._r)
    p.runs[0].text=text
    if size: p.runs[0].font.size=Pt(size)
ps=cell.paragraphs
cellp(ps[0],'건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석')
cellp(ps[1],'- 바이브 코딩으로 개발한 데이터센터 검토 앱 3개 사례를 중심으로 -')
cellp(ps[2],'A Study on Micro-App Analysis for Knowledge Expansion in Early-Stage Architectural Planning')
cellp(ps[3],'- Focused on Three Data Center Review Apps Developed through Vibe Coding -')
a=t0.rows[2].cells[1].paragraphs[0]; cellp(a,'임상철 Lim, Sang-Chul / 한양대학교 공학대학원 건축학전공 석사과정')
k=t0.rows[4].cells[1].paragraphs
cellp(k[0],'데이터센터, 초기 계획설계, 지식 확장, 바이브 코딩, Micro-App')
cellp(k[1],'Data Center, Early-Stage Architectural Planning, Knowledge Expansion, Vibe Coding, Micro-App')

# ---------- 본문 (제출본에서 순서대로)
FIX=[('p. 15) 분류에 의해 선언적','p. 15)에 따른 선언적'),('앱 에 담긴 지식을 구분 하고자','앱에 담긴 지식을 구분하고자'),('운영·휴게 공간이다 (그림 1)','운영·휴게 공간이다(그림 1)'),('만든3가지','만든 3가지'),('ClaudeCod로','Claude Code로'),('기존대비 개선된 점으로 볼수있다','기존 대비 개선된 점으로 볼 수 있다'),('명시화 되었지만','명시화되었지만'),('(앱의 연결)· 전략적','(앱의 연결)·전략적'),('샘플 버전 개발 까지를','샘플 버전 개발까지를')]
def fix(t):
    for a,b in FIX: t=t.replace(a,b)
    return t
S=docx.Document(SRC); sb=S.element.body
pending=''
started=False; in_refs=False; fig_no=0
def ptext(el): return ''.join(x.text or '' for x in el.iter(qn('w:t'))).strip()
for el in sb:
    if el.tag==qn('w:p'):
        p=Paragraph(el,S); t=ptext(el); st=p.style.name
        if not started:
            if st=='Heading 1' and t.startswith('1.'): started=True
            else: continue
        if el.find('.//'+qn('w:drawing')) is not None:
            rid=el.find('.//'+'{http://schemas.openxmlformats.org/drawingml/2006/main}blip').get(qn('r:embed'))
            picture(S.part.related_parts[rid].blob); continue
        if not t: continue
        t=fix(t)
        if st=='Heading 1':
            if t=='참고문헌':
                in_refs=True; blank(); P('장제목','References'); continue
            blank(); P('장제목',t); continue
        if st=='Heading 2':
            t=re.sub(r'^(\d+)\.(\d+)\s',r'\1.\2. ',t); P('절제목',t); continue
        if in_refs:
            if t.startswith('※'):
                blank(); P('Body Text',t,first=Cm(0)).runs[0].font.size=Pt(7.5); continue
            # 원문 이탤릭 보존
            parts=[]
            for r in p.runs:
                if r.text: parts.append(('*' if r.italic else '')+r.text+('*' if r.italic else ''))
            ref(''.join(parts).replace('**','')); continue
        m=re.match(r'^(그림|표) (\d+)\.\s*(.*)$',t)
        if m:
            caption('fig' if m.group(1)=='그림' else 'tab', '<%s %s> %s'%(m.group(1),m.group(2),m.group(3))); continue
        if t.startswith('자료:'): source(t.replace(' (주소를 클릭하면 앱으로 연결)','')); continue
        if t.startswith('연구의 방법은 사례조사'):
            t='본 연구에서 Micro-App은 지식 확장을 위해 도메인 지식을 기반으로 개발한 소규모 검토 앱을 말한다. '+t
        t=(pending+' '+t).strip() if pending else t; pending=''
        if not re.search(r'[.?!)]$',t): pending=t; continue
        lm=re.match(r'^(RQ\d|지식 명시화|도구 구현|검증):\s*(.*)$',t)
        if lm:
            P('Body Text',None,parts=[(lm.group(1)+': ',{'b':True}),(lm.group(2),{})]); continue
        P('Body Text',t)
    elif el.tag==qn('w:tbl') and started:
        if el.find('.//'+qn('w:drawing')) is not None:
            # 사례 화면 3개 → 합성 이미지
            A=Image.open(FIG+'fig_app1_b_diagram_crop.png').convert('RGB'); B=Image.open(FIG+'fig_caseB_crop.png').convert('RGB'); C=Image.open(FIG+'fig_caseC_crop.png').convert('RGB')
            W=1600; A=A.resize((W,int(A.height*W/A.width))); h=520
            B=B.resize((int(B.width*h/B.height),h)); C=C.resize((int(C.width*h/C.height),h))
            sc=(W-20)/(B.width+C.width); B=B.resize((int(B.width*sc),int(h*sc))); C=C.resize((int(C.width*sc),int(h*sc)))
            canvas=Image.new('RGB',(W,A.height+20+B.height),'white'); canvas.paste(A,(0,0)); canvas.paste(B,(0,A.height+20)); canvas.paste(C,(B.width+20,A.height+20))
            from PIL import ImageDraw, ImageFont
            dr=ImageDraw.Draw(canvas)
            try: fnt=ImageFont.truetype('/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc',40)
            except: fnt=None
            for lab,xy in (('(a) 사례 A',(12,10)),('(b) 사례 B',(12,A.height+30)),('(c) 사례 C',(B.width+32,A.height+30))):
                dr.rectangle([xy[0]-6,xy[1]-4,xy[0]+190,xy[1]+50],fill='white'); dr.text(xy,lab,fill='black',font=fnt)
            bio=io.BytesIO(); canvas.save(bio,'PNG'); picture(bio.getvalue())
            continue
        rows=[[fix(ptext(tc)).replace('개발 기간(커밋)','개발 기간') for tc in tr.findall(qn('w:tc'))] for tr in el.findall(qn('w:tr'))]
        hdr=rows[0]
        if hdr[0]=='분야': widths=[2.2,2.6,3.6]; left=(1,2)
        elif hdr[1]=='이론적 근거': widths=[1.5,2.2,2.6,2.2]; left=(1,2,3)
        elif hdr[0]=='사례': widths=[0.6,1.6,2.1,1.5,2.8]; left=(1,2,4)
        else: widths=[1.4,2.2,1.3,1.6,1.3]; left=()
        table(rows,widths,left)
d.save(OUT); print('saved',OUT)
