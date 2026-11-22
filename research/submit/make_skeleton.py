import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
FIG='/home/user/pango-gmp/research/figures/'
d=docx.Document()
sec=d.sections[0]; sec.page_width=Cm(21); sec.page_height=Cm(29.7)
sec.left_margin=sec.right_margin=Cm(2.2); sec.top_margin=Cm(2.5); sec.bottom_margin=Cm(2.2); sec.footer_distance=Cm(1.0)
def fonts(rpr,a,e):
    rf=rpr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.insert(0,rf)
    for k in list(rf.attrib): del rf.attrib[k]
    for k,v in (('ascii',a),('hAnsi',a),('eastAsia',e),('cs',a)): rf.set(qn('w:'+k),v)
n=d.styles['Normal']; n.font.size=Pt(10.5); fonts(n.element.get_or_add_rPr(),'Times New Roman','바탕')
n.paragraph_format.line_spacing=1.5; n.paragraph_format.space_after=Pt(2)
for nm,sz,b,a in (('Heading 1',14,12,6),('Heading 2',11.5,8,3)):
    st=d.styles[nm]; st.font.size=Pt(sz); st.font.bold=True; st.font.color.rgb=RGBColor(0,0,0)
    fonts(st.element.get_or_add_rPr(),'맑은 고딕','맑은 고딕'); pf=st.paragraph_format
    pf.space_before=Pt(b); pf.space_after=Pt(a); pf.keep_with_next=True; pf.line_spacing=1.2
# 쪽 번호
fp=sec.footer.paragraphs[0]; fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
def fld(p,code):
    for kind in ('begin','instr','separate','t','end'):
        r=p.add_run(); r.font.size=Pt(9.5)
        if kind=='instr': it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text=code; r._r.append(it)
        elif kind=='t': r.text='1'
        else: fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),kind); r._r.append(fc)
fp.add_run('- ').font.size=Pt(9.5); fld(fp,' PAGE '); fp.add_run(' -').font.size=Pt(9.5)
def P(text='',size=None,bold=False,align=None,sa=None,sb=None,color=None,italic=False):
    p=d.add_paragraph(); 
    if align is not None: p.alignment=align
    if sa is not None: p.paragraph_format.space_after=Pt(sa)
    if sb is not None: p.paragraph_format.space_before=Pt(sb)
    if text:
        r=p.add_run(text); r.font.bold=bold; r.font.italic=italic
        if size: r.font.size=Pt(size)
        if color: r.font.color.rgb=RGBColor.from_string(color)
    return p
def MEMO(lines,n):
    p=P('▶ 작성 메모 (이 회색 글을 지우고 본문을 직접 작성, 약 %s줄)'%n,9,True,color='8C8C8C',sa=0)
    for l in lines:
        q=P('· '+l,9,color='8C8C8C',sa=0); q.paragraph_format.left_indent=Cm(0.4); q.paragraph_format.line_spacing=1.25
    P('',sa=2)
H=lambda t,l=1: d.add_heading(t,level=l)
def BODY(t):
    p=d.add_paragraph(t); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.first_line_indent=Cm(0.5); return p
def cap(t,above=True):
    p=P(t,9.5,False,WD_ALIGN_PARAGRAPH.CENTER,sa=3 if above else 1,sb=4 if above else 2); p.paragraph_format.keep_with_next=above; return p
def SRC(t): P('자료: '+t,8.5,align=WD_ALIGN_PARAGRAPH.LEFT if False else WD_ALIGN_PARAGRAPH.CENTER,sa=6)
def link(par,url,text):
    rid=d.part.relate_to(url,RT.HYPERLINK,is_external=True)
    h=OxmlElement('w:hyperlink'); h.set(qn('r:id'),rid); r=OxmlElement('w:r'); rpr=OxmlElement('w:rPr')
    c=OxmlElement('w:color'); c.set(qn('w:val'),'000000'); rpr.append(c); u=OxmlElement('w:u'); u.set(qn('w:val'),'single'); rpr.append(u)
    sz=OxmlElement('w:sz'); sz.set(qn('w:val'),'16'); rpr.append(sz)
    r.append(rpr); t=OxmlElement('w:t'); t.text=text; r.append(t); h.append(r); par._p.append(h)
def TABLE(header,rows,widths,links=None):
    t=d.add_table(rows=1+len(rows),cols=len(header)); t.alignment=WD_TABLE_ALIGNMENT.CENTER
    tblPr=t._tbl.tblPr; bd=OxmlElement('w:tblBorders')
    for side,val,sz in (('top','single','12'),('left','nil','0'),('bottom','single','12'),('right','nil','0'),('insideH','single','4'),('insideV','single','4')):
        e=OxmlElement('w:'+side); e.set(qn('w:val'),val); e.set(qn('w:sz'),sz); e.set(qn('w:space'),'0'); e.set(qn('w:color'),'000000'); bd.append(e)
    tblPr.append(bd)
    for ri,row in enumerate([header]+rows):
        for ci,val in enumerate(row):
            c=t.rows[ri].cells[ci]; c.width=Cm(widths[ci])
            tcPr=c._tc.get_or_add_tcPr(); va=OxmlElement('w:vAlign'); va.set(qn('w:val'),'center'); tcPr.append(va)
            if ri==0:
                sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),'D9D9D9'); tcPr.append(sh)
            p=c.paragraphs[0]; p.paragraph_format.line_spacing=1.15; p.paragraph_format.space_after=Pt(1); p.paragraph_format.space_before=Pt(1)
            if ri==0: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            if links and ri>0 and ci==links[0]:
                link(p,val,val.replace('https://','').rstrip('/'))
            else:
                r=p.add_run(val); r.font.size=Pt(9); r.font.bold=(ri==0)
            tr=t.rows[ri]._tr; trPr=tr.get_or_add_trPr(); cs=OxmlElement('w:cantSplit'); trPr.append(cs)
            if ri==0: th=OxmlElement('w:tblHeader'); th.set(qn('w:val'),'true'); trPr.append(th)
    return t
def FIG1(f,w):
    p=P(align=WD_ALIGN_PARAGRAPH.CENTER,sa=0,sb=4); p.paragraph_format.keep_with_next=True; p.paragraph_format.line_spacing=1.0
    p.add_run().add_picture(FIG+f,width=Cm(w))

# ===== 제목
P('건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석',16,True,WD_ALIGN_PARAGRAPH.CENTER,sa=2)
P('— 바이브 코딩으로 개발한 데이터센터 검토 앱 3개 사례를 중심으로 —',11,False,WD_ALIGN_PARAGRAPH.CENTER,sa=2)
P('한양대학교 공학대학원 건축학전공  임상철 (2026160492)  |  건축계획연구 (담당교수 김영주)',9.5,False,WD_ALIGN_PARAGRAPH.CENTER,sa=10)

H('1. 서론')
H('1.1 연구의 배경',2)
BODY('건축 설계의 단계 중 초기 기획·계획설계 단계는 사업의 규모와 설계의 밑그림을 그리는 중요한 과정이다. 초기 단계의 결정이 건물 성능·생애주기 비용에 미치는 영향이 가장 크다(Bragança 외, 2014).')
BODY('일반 시설(상주 공간 중심)은 발주처의 요구 사항, 건축물 용도에 따른 설계 각론, 유사 사례, 건축가의 경험으로 정합성 있는 계획안 도출이 가능하다. 하지만 특수 시설(산업시설, 데이터센터 등)의 경우는 설비·전기 등의 공간 비중이 커서 검토의 한계가 있다. 전문 분야(구조·기계·전기 분야 등) 프로젝트 팀이 구성되기 어려운 초기 기획 단계에서 건축가가 정확한 기획설계안을 도출하기 위해서는 전문 분야의 지식 확장이 필요하며, 이를 앱(App)으로 생성하고자 한다. 본문에서 언급하는 지식 확장의 정의는 ① 전문 분야 검토 항목의 식별, ② 근거의 이해, ③ 조건 변경 시 정량적이고 신속한 대응이다.')
BODY('데이터센터에서는 전기 수전 용량에 따른 서버실과 항온항습실, 전원 공급 시스템 안정성을 위한 전기실, 정전 시 전원 공급을 유지할 수 있도록 해 주는 비상 발전기실을 계획하고, 이를 이중화와 여유율 공간까지 고려하는 것이 건축계획의 중대한 요소이다. 이 분야는 전기·설비 전문 설계 분야와 협업을 해야 하지만, 프로젝트 초기에 계약도 안 된 상황에서 기획설계를 위해 정식으로 검토하는 것은 현실적으로 어려운 실정이다. 이에 건축 외 분야의 전문 지식이 필요한 특수 시설 중 데이터센터를 사례로 구체적인 연구를 진행해 보고자 한다.')
H('1.2 연구의 목적과 연구 질문',2)
MEMO(['목적: 초기 판단 오류로 인한 설계 변경 리스크를 줄이기 위해, 바이브 코딩으로 개발한 검토 앱을 초기 계획설계에 도입하는 방식과 원리를 분석',
 'RQ1 앱은 전문 지식을 어떻게 반영하고 명시화·검증·공유하는가',
 'RQ2 기존 검토 방식(유추·직관·업체 비공식 도움)에 비해 무엇을 개선하는가',
 'RQ3 앱들을 통합하여 지식 확장의 완성도를 높이려면 어떤 방향이 필요한가'],5)
H('1.3 연구의 방법 및 범위',2)
MEMO(['방법: 문헌 조사(분석틀 도출) + 사례분석 (설문·인터뷰·실험 없음)',
 '대상: 저자가 개발해 웹으로 배포, 동료 실무자가 사용하고 협력업체가 결과를 보며 의견을 주는 시범 적용(pilot) 앱 3개',
 '시점: 2026년 11월 버전 (개발이 계속되는 단계, 공식 인증 전)',
 '제외: 앱의 효과(정확도·소요 시간) 검증, 통합 프레임워크 구현'],4)

H('2. 이론적 배경')
H('2.1 설계 지식의 명시화와 지식 유형',2)
MEMO(['암묵지 → 형식지로의 표출화(externalization)가 지식 창조의 핵심 (Nonaka & Takeuchi, 1995)',
 '본 연구의 앱: 협력사와의 협의·문의로 흩어져 있던 지식과 전문가의 검토 원리를 로직으로 옮긴 표출화의 한 경로',
 '지식 유형 (de Jong & Ferguson-Hessler, 1996; 이병수, 2016, p. 15): 선언적(입력값·기준), 절차적(산식·계산 순서), 상황적(조건과 연결 관계), 전략적(검토 순서와 판단)'],5)
H('2.2 데이터센터의 공간 구성',2)
MEMO(['공간: 서버룸(본 연구의 데이터홀), 서포팅 인프라(기계·전기), 운영 공간 → 그림 1 (이명진, 2021)',
 '메인 층(서버실·항온항습실) 연면적 : 기계·전기실 연면적 ≈ 1 : 1 (김재성·이원재, 2018)',
 '서버실 층고 4.5m 이상, 수변전실·발전기실 적정 층고 7m, 발전기실 급·배기구 면적 확보 (김재성·이원재, 2018)',
 '냉각은 이중바닥·천장 높이·랙 배치의 영향을 받음 (Cai & Gou, 2026)'],5)
FIG1('fig_2_1_space.png',15.5)
cap('그림 1. 데이터센터의 공간구성 프로세스와 본 연구 사례의 위치',False)
SRC('저자 작성 (이명진, 2021, 그림 1을 참고하여 재구성; 비율은 김재성·이원재, 2018)')
H('2.3 선행연구 검토',2)
MEMO(['표 1의 네 흐름을 각 1~2문장으로 서술 (한정되었다는 표현 대신 "주로 ~방향으로 연구되었다")',
 '연구 공백: 설비 용량 산정을 공간 영향 검토로 잇는 도구, 실무자가 생성형 AI로 직접 만든 검토 도구의 지식을 명시화·검증·공유 관점에서 분석한 연구는 확인되지 않음'],5)
cap('표 1. 선행연구 분석')
TABLE(['분야','선행연구','주된 연구 방향'],[
 ['설계 지식 구조화·AI 설계지원','이병수 외(2016), 김혜정 외(2019), 이용주(2020)','매스 규모산정 지식의 BIM 구조화, AI 디자인 대안 생성, 전문가 시스템 기반 친환경 설계지원'],
 ['LLM·온톨로지·에이전트','Chen 외(2024), Dong 외(2025)','LLM·온톨로지 기반 BIM 규정 검토, 다중 에이전트 설계 조정'],
 ['데이터센터 공간 구성','김재성·이원재(2018), 이명진(2021)','국내 사례의 실 프로그램, 공간구성 프로세스'],
 ['데이터센터 냉각','Cai & Gou(2026)','패시브·액티브 냉각 전략의 종합 리뷰']],[3.6,4.6,8.4])
SRC('저자 작성')

H('3. 연구방법')
H('3.1 연구 모형과 분석틀',2)
MEMO(['그림 2: 이론적 근거 → 분석틀(4관점 12요소) → 사례 3개 판정 → 결과와 제언의 진행 모형 (인과 검증 모형 아님)',
 '표 2: 관점별 이론적 근거와 세부 요소',
 '판정: ● 충족(화면·공개 문서에서 확인) ◐ 부분 충족(일부 또는 앱 밖에서만 확인) ○ 미흡, ⑤는 판정 없이 기술',
 '자료: 소스코드, 커밋 이력, 설명서, 실행 화면 / 계산은 연구자의 독립 계산과 대조'],5)
FIG1('fig_3_1_model.png',15.5)
cap('그림 2. 연구 모형',False)
SRC('저자 작성')
cap('표 2. 분석틀')
TABLE(['분석 관점','이론적 근거','세부 요소','핵심 질문'],[
 ['지식 명시화','Nonaka & Takeuchi(1995), 이병수 외(2016)','① 검토 영역 ② 입력 변수·단위 ③ 산식·규칙 ④ 근거·가정 표시','앱이 아는 것을 화면에 드러내는가'],
 ['도구 구현','이병수 외(2016)','⑤ 구현 형태 ⑥ 출력·공간 표현','결과를 건축 공간으로 보여 주는가'],
 ['검증','Peffers 외(2007) 평가 단계','⑦ 문서-코드 일치 ⑧ 기준 사례 시험 ⑨ 자체 점검','계산이 맞는지 확인할 수 있는가'],
 ['공유·개선','Peffers 외(2007) 공유 단계, Nonaka & Takeuchi(1995)','⑩ 배포·접근 ⑪ 의견 수집·처리 ⑫ 지식 원본 관리','다른 사람이 쓰고 고칠 수 있는가']],[2.4,4.4,5.6,4.2])
SRC('저자 작성 (Nonaka & Takeuchi, 1995; 이병수 외, 2016; Peffers 외, 2007 참고)')
H('3.2 분석 대상 사례',2)
MEMO(['선정 기준: 저자 개발·배포·시범 적용 / 서로 다른 검토 영역 / 소스와 이력 확인 가능',
 '개발 과정: 협력업체 협의·장비 사양 정리 → Claude Code로 로직 구현 → GitHub 커밋 → Vercel 배포 → 의견 반영',
 '세 앱 모두 HTML·CSS·JavaScript 분리, 사례 C는 화면을 React로 구성하고 계산 함수를 별도 파일로 둠'],4)
cap('표 3. 분석 대상 사례')
TABLE(['사례','앱','검토 영역 → 공간 영향','개발 기간(커밋)','접속 주소'],[
 ['A','전력 이중화 검토','이중화·변압기 → 전기실(평면)','2026.08~11 (73)','https://power-redundancy-calculator.vercel.app/'],
 ['B','데이터홀 단면 검토','랙·냉각 장비 → 층고·건물 규모','2026.05~10 (17)','https://datahall-section-check.vercel.app/'],
 ['C','발전기실 급·배기 계산','급·배기 풍량 → 입면·층고','2026.06~09 (5)','https://generator-calc.vercel.app/']],[1.0,3.0,4.4,2.7,5.5],links=(4,))
SRC('각 사례의 GitHub 저장소 커밋 이력, 2026.11 기준 (주소를 클릭하면 앱으로 연결)')

H('4. 사례분석')
t=d.add_table(rows=2,cols=3); t.alignment=WD_TABLE_ALIGNMENT.CENTER
for i,(f,lab) in enumerate((('fig_app1_b_diagram_crop.png','(a) 사례 A'),('fig_caseB_crop.png','(b) 사례 B'),('fig_caseC_crop.png','(c) 사례 C'))):
    c=t.rows[0].cells[i]; p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run().add_picture(FIG+f,height=Cm(4.0))
    q=t.rows[1].cells[i].paragraphs[0]; q.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=q.add_run(lab); r.font.size=Pt(9)
cap('그림 3. 사례 앱의 주요 화면',False)
SRC('저자 개발 앱 실행 화면 캡처, 2026.11.18~20')
H('4.1 사례 A: 전력 이중화 검토 앱',2)
MEMO(['명시화: 이중화 7종(2oo3, 3oo4, 4oo5, N+1, N+2C, 2N, 2N+1)+직접 입력, 단위 입력, 유효용량 산식 화면 표시(4,000 × 0.99 × 0.9 × 0.96 = 3,421.4kW) / 기준값 출처는 일부만(랙 밀도별 냉각 기준은 코드 메모에만) ◐',
 '구현: 브라우저 실행, 전기실 연결도·층 평면·데이터홀 평면·단면·계통도로 표현',
 '검증: 설명서 산식과 코드 일치, 시험 코드 없음, 33개 조건 363개 결과 독립 계산과 불일치 없음, 부하율 100% 초과·환기 공간 부족 경고',
 '공유: 공개 배포, 의견 분류·상태 관리, 의견이 브라우저에만 저장, 73 커밋, 냉각 산식이 사례 B와 중복'],5)
H('4.2 사례 B: 데이터홀 단면 검토 앱',2)
MEMO(['명시화: 층고·보 깊이·기둥 간격·랙·IT 부하·냉각 장비 입력, 단계별 계산 표시(필요 풍량 8,000kW × 170CFM/kW × 1.699 × 1.1 = 2,541,704CMH, 보 아래 1.97m = 8.3 − 6.33), 계수 출처 미표시 ○',
 '구현: 입면·단면·평면 자동 작성, 3D·OBJ·DXF 내보내기',
 '검증: 설명서는 기능 위주(산식은 화면·코드), 시험 코드 없음, 기둥 간섭·풍속 초과(5.85m/s) 경고',
 '공유: 공개 배포(저장소와 동일), 의견 반영·보류·반려 분류, 배포본에서 서버 수집 미동작, 냉각 산식 사례 A와 중복'],5)
H('4.3 사례 C: 발전기실 급·배기 계산 앱',2)
MEMO(['명시화: 제조사 4곳 모델 데이터 자동 입력, 카탈로그 출처 쪽수·산식 메모 화면 표시 ●',
 '구현: React 화면, 계산 함수 1파일을 화면과 API가 공유, 급·배기 계통 다이어그램',
 '검증: 설명서에 기준(급기 7m/s, 배기 8m/s, 여유율 1.2)만, 시험 코드 없음, Cummins 3,000kW 8대 급기 69.12㎡·배기 54.72㎡ 일치, 카탈로그 7.20㎡ vs 역산 7.17㎡ 점검, 풍속 초과 경고',
 '공유: 공개 배포, API 동작, 의견 수집 기능 없음 ○, 지식 원본 관리 가장 명확'],5)
H('4.4 종합 비교와 기존 방식 대비 개선',2)
MEMO(['표 4: 충족 A 7·B 6·C 8, 미흡 B(④)·C(⑪) 각 1, ⑧ 기준 사례 시험은 세 사례 모두 ◐',
 '세 사례 모두 충족: ①②③⑥⑨⑩ / 사례별 차이: 근거 표시(C ●, A ◐, B ○), 지식 원본(C 1파일, A·B 냉각 산식 중복)',
 '기존 방식(유추·직관·업체 비공식 도움) → 앱: 정량적이고 일관된 검토 규칙, 같은 입력 같은 결과, 조건 변경 시 즉시 재계산 (개략 검토용, 업체 상세 검토 대체 아님)',
 '지식 유형: 선언적·절차적 지식은 명시화, 상황적(앱 사이의 연결)·전략적(검토 순서) 지식은 설계자에게 남음'],6)
cap('표 4. 분석틀에 따른 사례 비교')
J=[['지식 명시화','① 검토 영역','●','●','●'],['','② 입력 변수·단위','●','●','●'],['','③ 산식·규칙','●','●','●'],['','④ 근거·가정 표시','◐','○','●'],
 ['도구 구현','⑤ 구현 형태','정적 웹앱','정적 웹앱 + 선택형 서버','웹앱 + API'],['','⑥ 출력·공간 표현','●','●','●'],
 ['검증','⑦ 문서-코드 일치','●','◐','◐'],['','⑧ 기준 사례 시험','◐','◐','◐'],['','⑨ 자체 점검','●','●','●'],
 ['공유·개선','⑩ 배포·접근','●','●','●'],['','⑪ 의견 수집·처리','◐','◐','○'],['','⑫ 지식 원본 관리','◐','◐','●'],
 ['판정 빈도 (11개)','● / ◐ / ○','7 / 4 / 0','6 / 4 / 1','8 / 2 / 1']]
TABLE(['분석 관점','세부 요소','사례 A','사례 B','사례 C'],J,[2.8,3.8,3.0,4.0,3.0])
SRC('저자 작성 (● 충족, ◐ 부분 충족, ○ 미흡; ⑤는 판정 대상에서 제외; 2026.11 버전 기준)')

H('5. 결론 및 제언')
H('5.1 결론',2)
MEMO(['RQ1: 협의·장비 사양을 로직으로 옮겨 개발·배포, 산식·결과는 명시되나 근거 표시와 기준 사례 시험은 사례별 차이',
 'RQ2: 유추·직관에 의존하던 검토를 정량적이고 일관된 규칙으로 바꾸어 검토 방식을 체계화',
 'RQ3: 앱 사이의 연결(상황적 지식)과 검토 순서(전략적 지식)를 다루기 위한 통합 필요',
 '기대: 검토 속도와 정량적 근거 증가 → 초기 계획의 정합성 향상, 설계 변경 리스크 감소 (효과 검증은 후속 연구)'],6)
H('5.2 제언 및 향후 연구',2)
MEMO(['제언 ① 지식 원본 단일화 ② 근거 수준(확인·가정) 표시 ③ 독립 계산 대조를 기준 사례 시험으로 저장소에 포함 ④ 공통 포털과 의견 이력 관리',
 '향후 연구: 앱 통합 → 온톨로지로 지식 구조화 → LLM 기반 AI 에이전트가 앱을 호출해 결론·근거·산식을 담은 검토의견서 도출 (Chen 외, 2024; Dong 외, 2025 참고)',
 '분석틀은 시설 유형에 종속되지 않아 산업시설·방송통신시설 등 다른 특수 시설로 확장 가능'],5)
H('5.3 연구의 한계',2)
MEMO(['저자 개발 앱 3개로 일반화 한계, 연구자=개발자(판정 근거를 공개 자료로 한정해 보완)',
 '시범 적용 단계의 2026.11 버전, 효과(정확도·소요 시간)는 미검증'],3)

H('참고문헌')
from docx.enum.text import WD_BREAK
REFS=['김재성, 이원재. (2018). 2000년대 이후 국내 데이터센터의 실(room) 프로그램과 공간구성에 관한 연구. *한국문화공간건축학회논문집, 61*, 149–157.',
'김혜정, 김미경, 전한종. (2019). 인공지능 기술을 활용한 BIM 설계의 디자인 대안 생성 프로세스에 관한 기초적 연구. *대한건축학회 학술발표대회 논문집, 39*(2), 136–139.',
'이명진. (2021). 데이터센터의 공간구성. *대한설비공학회 동계학술발표대회 논문집*, 160–163.',
'이병수. (2016). *건축기획단계 매스 규모산정을 위한 지식기반 BIM 시스템 구현* [석사학위논문, 한양대학교 대학원].',
'이병수, 지승열, 전한종. (2016). 건축기획 매스 규모산정의 설계지식 재현을 위한 지식기반 BIM 시스템 구현. *한국CDE학회 논문집, 21*(3), 252–266. https://doi.org/10.7315/CDE.2016.252',
'이용주. (2020). *인공지능 기술의 전문가 시스템을 활용한 BIM 기반 친환경 건축 설계지원에 관한 연구* [박사학위논문, 한양대학교].',
'Bragança, L., Vieira, S. M., & Andrade, J. B. (2014). Early stage design decisions: The way to achieve sustainable buildings at lower costs. *The Scientific World Journal, 2014*, Article 365364. https://doi.org/10.1155/2014/365364',
'Cai, S., & Gou, Z. (2026). Towards energy-efficient data centers: A comprehensive review of passive and active cooling strategies. *Energy and Built Environment, 7*, 206–226. https://doi.org/10.1016/j.enbenv.2024.08.009',
'Chen, N., Lin, X., Jiang, H., & An, Y. (2024). Automated building information modeling compliance check through a large language model combined with deep learning and ontology. *Buildings, 14*(7), Article 1983. https://doi.org/10.3390/buildings14071983',
'de Jong, T., & Ferguson-Hessler, M. G. M. (1996). Types and qualities of knowledge. *Educational Psychologist, 31*(2), 105–113.',
'Dong, Y., Zhan, Z., Hu, Y., Doe, D. M., & Han, Z. (2025). AI BIM coordinator for non-expert interaction in building design using LLM-driven multi-agent systems. *Automation in Construction, 180*, Article 106563. https://doi.org/10.1016/j.autcon.2025.106563',
'International Energy Agency. (2025). *Energy and AI*. https://www.iea.org/reports/energy-and-ai',
'Nonaka, I., & Takeuchi, H. (1995). *The knowledge-creating company: How Japanese companies create the dynamics of innovation*. Oxford University Press.',
'Peffers, K., Tuunanen, T., Rothenberger, M. A., & Chatterjee, S. (2007). A design science research methodology for information systems research. *Journal of Management Information Systems, 24*(3), 45–77. https://doi.org/10.2753/MIS0742-1222240302']
for ref in REFS:
    p=d.add_paragraph(); pf=p.paragraph_format; pf.left_indent=Cm(1.0); pf.first_line_indent=Cm(-1.0); pf.line_spacing=1.2; pf.space_after=Pt(2)
    for k,part in enumerate(ref.split('*')):
        if part: r=p.add_run(part); r.font.size=Pt(9); r.font.italic=(k%2==1)
H('생성형 AI 활용 표기',2)
MEMO(['실제 사용 범위를 본인이 직접 기재 (예: 사례 앱 개발에 Claude Code 사용, 보고서 구성·자료 정리·그림 작성·검토에 생성형 AI 활용, 본문은 저자 작성)'],2)
d.save('건축계획연구_과제_뼈대.docx'); print('saved')
