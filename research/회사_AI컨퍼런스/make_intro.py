import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
d=docx.Document()
s=d.sections[0]; s.page_width=Cm(21); s.page_height=Cm(29.7); s.left_margin=s.right_margin=Cm(2.0); s.top_margin=Cm(2.0); s.bottom_margin=Cm(1.8)
def fonts(rpr,f):
    rf=rpr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.insert(0,rf)
    for k in ('ascii','hAnsi','eastAsia','cs'): rf.set(qn('w:'+k),f)
n=d.styles['Normal']; n.font.size=Pt(10); fonts(n.element.get_or_add_rPr(),'맑은 고딕'); n.paragraph_format.line_spacing=1.35; n.paragraph_format.space_after=Pt(2)
def P(t='',size=None,bold=False,align=None,sb=0,sa=2,color=None,indent=None):
    p=d.add_paragraph(); p.paragraph_format.space_before=Pt(sb); p.paragraph_format.space_after=Pt(sa)
    if align is not None: p.alignment=align
    if indent: p.paragraph_format.left_indent=Cm(indent)
    if t:
        r=p.add_run(t); r.font.bold=bold
        if size: r.font.size=Pt(size)
        if color: r.font.color.rgb=RGBColor.from_string(color)
    return p
def H(t): 
    p=P(t,12.5,True,sb=10,sa=4)
    pb=OxmlElement('w:pBdr'); b=OxmlElement('w:bottom')
    for k,v in (('val','single'),('sz','6'),('space','2'),('color','8E8A84')): b.set(qn('w:'+k),v)
    pb.append(b); p._p.get_or_add_pPr().append(pb)
def SUB(t): P(t,10.5,True,sb=5,sa=1)
def B(label,t=None):
    p=P(indent=0.4); p.paragraph_format.first_line_indent=Cm(-0.4)
    r=p.add_run('• '); 
    if t is None: p.add_run(label)
    else:
        rb=p.add_run(label+': '); rb.font.bold=True; p.add_run(t)
def TXT(t): P(t,indent=0.2)

P('AI 활용사례 소개',10,False,color='8C8C8C',sa=0)
P('데이터센터 초기 설계 검토 앱 (Micro-App 3종)',18,True,sa=2)
P('— 전력 이중화 · 데이터홀 단면 · 발전기실 급·배기 —',11,False,color='555555',sa=8)
p=P(sb=0,sa=6); r=p.add_run('개발 목적: '); r.font.bold=True
p.add_run('협력업체 검토 전 초기 설계 단계에서 설비·전기 공간의 규모를 근거 있게 산정하여, 설계 변경 리스크를 줄이고 설계자의 검토 역량을 확장')

H('1. 기존 업무의 문제점 및 개발 배경')
TXT('데이터센터는 전기실, 발전기실, 데이터홀, 냉동기 등 설비·전기 공간의 비중이 커서, 이들 공간의 규모가 건물 전체의 평면, 입면, 층고를 좌우함.')
TXT('그러나 초기 기획·계획설계 단계에서는 전기·설비 협력업체와의 계약이 이루어지지 않은 경우가 많아, 건축가가 유사 사례와 경험으로 규모를 추정하거나 협력업체에 비공식적으로 검토를 요청해야 했음.')
TXT('이후 정식 검토 결과가 초기 가정과 크게 다르면 전기실 규모, 발전기실 그릴창 면적, 데이터홀 층고 등이 바뀌어 설계 변경과 재작업이 발생함.')
TXT('이에 따라 협력업체와 협의하며 축적한 검토 방법과 장비 사양을 앱으로 정리하여, 설계자가 조건만 입력하면 근거와 함께 결과를 즉시 확인할 수 있는 도구를 직접 개발함.')

H('2. 프로그램 개요 및 주요 기능')
TXT('본 프로그램은 데이터센터 초기 설계에서 건축 규모에 가장 큰 영향을 주는 세 가지 검토를 각각 수행하는 웹 기반 검토 앱(3종)임. 별도 설치 없이 링크로 접속하여 사용함.')
t=d.add_table(rows=4,cols=4); t.alignment=WD_TABLE_ALIGNMENT.CENTER
rows=[['구분','앱','입력','결과 및 공간 영향'],
 ['A','전력 이중화 검토','이중화 방식(7종), IT 용량, 변압기 용량, 장비 사양','변압기 대수·부하율, 전기실 구성, 급전 다이어그램 → 평면'],
 ['B','데이터홀 단면 검토','서버 수량, 랙 배치, 기둥 간격, 랙 부하, 냉각 장비','필요 풍량, 보 아래 환기 공간, 입면·단면·평면 → 층고·규모'],
 ['C','발전기실 급·배기 계산','발전기 제조사·모델, 대수, 풍속, 개구율','급·배기 개구 면적, 풍속, 계통도 → 입면·층고']]
tb=t._tbl.tblPr; bd=OxmlElement('w:tblBorders')
for side,val,sz in (('top','single','12'),('left','nil','0'),('bottom','single','12'),('right','nil','0'),('insideH','single','4'),('insideV','single','4')):
    e=OxmlElement('w:'+side); e.set(qn('w:val'),val); e.set(qn('w:sz'),sz); e.set(qn('w:space'),'0'); e.set(qn('w:color'),'000000'); bd.append(e)
tb.append(bd)
for ri,row in enumerate(rows):
    for ci,v in enumerate(row):
        c=t.rows[ri].cells[ci]; c.width=Cm([1.2,3.3,6.0,6.5][ci])
        if ri==0:
            sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),'D9D9D9'); c._tc.get_or_add_tcPr().append(sh)
        p=c.paragraphs[0]; p.paragraph_format.line_spacing=1.15
        if ri==0 or ci==0: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r=p.add_run(v); r.font.size=Pt(9); r.font.bold=(ri==0)
SUB('주요 기능')
B('즉시 계산','조건을 입력하면 변압기 대수, 필요 풍량, 개구 면적 등을 바로 산출')
B('근거 표시','계산식과 값을 화면에 함께 표시 (예: 4,000 × 0.99 × 0.9 × 0.96 = 3,421.4kW), 제조사 카탈로그 출처 표시')
B('도면 연동','결과를 급전 다이어그램, 평면·단면·입면, 급·배기 계통도로 표현하여 공간 영향을 바로 확인')
B('자동 경고','부하율 초과, 풍속 초과, 기둥 간섭, 환기 공간 부족 등을 자동 표시')
B('대안 비교','이중화 방식, 랙 밀도, 발전기 대수 등 조건을 바꾸면 즉시 재계산')
B('공유·출력','웹 링크로 팀원·협력업체와 결과 공유, 도면 DXF·OBJ 내보내기(사례 B)')

H('3. 프로그램 제작 과정')
SUB('① 실무 지식 정리')
TXT('프로젝트 진행 중 협력업체와 협의하고 문의하며 메모해 둔 검토 방법, 장비 사양, 계산 근거와 전문가의 검토 순서를 모아 입력·규칙·산식 형태로 정리함.')
SUB('② 검토 항목 선정 및 화면 기획')
TXT('건축 규모에 영향이 큰 전력 이중화, 데이터홀 단면, 발전기실 급·배기를 우선 대상으로 선정하고, 입력 → 계산 → 도면 표현의 화면 구조를 설계함.')
SUB('③ AI 기반 개발 (바이브 코딩)')
TXT('Claude Code를 활용하여 자연어로 필요한 기능과 계산 조건을 설명하고, AI가 작성한 코드를 실행 화면으로 확인하며 수정을 반복함.')
SUB('④ 배포 및 지속 개선')
TXT('GitHub로 버전을 관리하고 Vercel로 웹에 배포하는 파이프라인을 구축함. 팀원과 협력업체의 의견을 받아 다시 AI로 반영하는 방식으로 개선 중임(개발 기간 2026.05~11, 사례 A 기준 73회 업데이트).')

H('4. 프로그램의 주요 장점 및 기대효과')
SUB('① 협력업체 회신 대기 시간 단축')
TXT('계약 전 초기 단계에서도 설계자가 직접 전기실·발전기실·데이터홀 규모를 개략 검토할 수 있어, 비공식 문의와 회신 대기에 소요되던 시간을 줄일 수 있음.')
SUB('② 근거 있는 초기 계획')
TXT('유사 사례 추정과 직관에 의존하던 판단을 산식과 근거가 드러나는 정량적 검토로 바꾸어, 발주처와 협의 시 설명력을 높일 수 있음.')
SUB('③ 설계 변경 리스크 감소')
TXT('조건 변경 시 즉시 재계산하여 대안을 비교할 수 있으므로, 이후 정식 검토 결과와의 차이로 인한 규모 변경과 재작업을 줄일 수 있음.')
SUB('④ 실무 지식의 자산화')
TXT('개인의 경험과 협의 메모로 흩어져 있던 전문 지식을 누구나 쓸 수 있는 도구로 정리하여, 담당자가 바뀌어도 같은 기준으로 검토할 수 있음.')
SUB('⑤ AI를 활용한 자체 업무 도구 개발')
TXT('전문 개발 인력이나 상용 프로그램 없이 설계 실무자가 AI 코딩 도구로 필요한 프로그램을 직접 기획·개발할 수 있다는 점에서 의미가 있음.')

H('5. 개발 과정의 시행착오 및 개선 과제')
B('기준값의 근거 표시','일부 계수와 기준값의 출처가 화면에 표시되지 않아, 확인된 값과 가정한 값을 구분하는 표시가 필요함.')
B('계산 검증 체계','산식 검증을 다른 LLM 크로스체크와 유사 사례 비교로 수행하였으나, 기준 사례 시험을 앱에 포함하여 수정 시마다 자동 확인하는 체계가 필요함.')
B('지식의 중복 관리','냉각 관련 산식이 두 앱(A·B)에 각각 구현되어, 한쪽을 수정하면 다른 쪽과 어긋날 수 있음. 공통 산식을 한곳에서 관리하는 구조가 필요함.')
B('의견 수집','사용자 의견이 개인 브라우저에 저장되어 이력 관리가 어려움. 의견을 한곳에 모으고 처리 상태를 관리하는 기능이 필요함.')
B('외부 자료 최신화','제조사 카탈로그 기반 데이터(사례 C)는 정보 변경 여부를 분기별로 확인할 필요가 있음.')
B('보안 및 배포 환경','현재 개인 저장소와 외부 웹 서비스로 배포되어 있어, 사내 확산 시 보안 검토와 사내 환경 이전이 필요함.')

H('6. 향후 보완 및 발전 방향')
SUB('단기 — 앱별 완성도 향상')
TXT('근거·가정 표시, 기준 사례 시험, 의견 이력 관리를 보완하고, 냉동기 설치 공간 검토 앱을 추가하여 주요 설비 공간 검토를 완성하고자 함.')
SUB('중기 — 앱 통합 및 사내 공유')
TXT('공통 산식과 기준값을 한곳에서 관리하고 앱들을 하나의 포털로 묶어, 사내 팀원이 함께 쓰고 의견을 모을 수 있는 환경을 구축하고자 함.')
SUB('장기 — AI 기반 종합 검토 시스템으로 확장')
TXT('온톨로지로 설비와 공간의 관계를 구조화하고, AI 에이전트가 각 앱을 호출하여 수전용량·냉각 방식 변경이 전기실·발전기실·데이터홀에 미치는 영향을 함께 분석하고, 근거와 산식이 담긴 종합 검토의견서를 작성하는 시스템을 목표로 함.')
TXT('나아가 산업시설, 방송통신시설 등 설비 비중이 큰 다른 시설로 적용 범위를 넓히고자 함.')

H('7. 종합 의견')
TXT('본 프로그램은 초기 설계에서 협력업체의 검토를 기다려야 했던 실무의 불편에서 출발하여, 설계자가 AI를 활용해 직접 검토 도구를 만든 사례임.')
TXT('핵심은 계산을 자동화하는 기능이 아니라, 개인의 경험과 협의 과정에 흩어져 있던 전문 지식을 근거가 드러나는 도구로 정리하여 조직이 함께 쓸 수 있는 자산으로 전환하는 것에 있음.')
TXT('현재 팀원과 협력업체가 참여하는 시범 적용 단계로, 향후 실제 프로젝트에서 효과를 검증하고 사용자 의견을 반영해 개선한다면 데이터센터를 넘어 설비 비중이 큰 건축 프로젝트 전반의 초기 설계 품질을 높이는 데 기여할 수 있을 것으로 기대됨.')
P('※ 앱 접속: power-redundancy-calculator.vercel.app · datahall-section-check.vercel.app · generator-calc.vercel.app',8.5,False,color='8C8C8C',sb=8)
d.save('AI활용사례_데이터센터검토앱_소개.docx'); print('ok')
