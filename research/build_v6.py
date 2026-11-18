# -*- coding: utf-8 -*-
"""과제 제출본(v6): 핵심만 담은 간결판. 원본 서식(plan_v2_original.docx)의 표지·스타일·바닥글을 유지하고 본문을 새로 작성.
확인되지 않은 내용과 붉은 표기는 넣지 않는다. 참고문헌 3편."""
import copy, math
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

SRC, OUT, FIG = 'plan_v2_original.docx', '연구계획서_과제제출본.docx', 'figures/'
d = docx.Document(SRC)
body = d.element.body
TBL_TPL = copy.deepcopy(d.tables[0]._tbl)
# 표지(문단 0~3)만 남기고 본문 삭제
keep = {id(p._p) for p in d.paragraphs[:4]}
for el in list(body):
    if el.tag == qn('w:sectPr') or id(el) in keep:
        continue
    body.remove(el)
for sec in d.sections:
    for fp in sec.footer.paragraphs:
        for r in fp.runs:
            r.text = r.text.replace('연구계획서 수정 초안', '연구계획서')

def H1(t): d.add_paragraph(t, style='Heading 1')
def H2(t): d.add_paragraph(t, style='Heading 2')
def P(t, align=None, size=None, bold=None):
    p = d.add_paragraph()
    r = p.add_run(t)
    if size: r.font.size = Pt(size)
    if bold is not None: r.bold = bold
    if align is not None: p.alignment = align
    return p
def caption(t): P(t, align=WD_ALIGN_PARAGRAPH.CENTER, size=9)

def table(header, rows, widths=None):
    tbl = copy.deepcopy(TBL_TPL)
    sect = body[-1]; sect.addprevious(tbl)
    trs = tbl.tr_lst; hdr, tpl = trs[0], trs[1]
    for tr in trs[2:]: tbl.remove(tr)
    ncol = len(header)
    def fill(tr, vals):
        tcs = tr.tc_lst
        for tc in tcs[ncol:]: tr.remove(tc)          # 열 수 맞추기
        while len(tr.tc_lst) < ncol: tr.append(copy.deepcopy(tr.tc_lst[-1]))
        for i, (tc, v) in enumerate(zip(tr.tc_lst, vals)):
            if widths:
                w = tc.get_or_add_tcPr().find(qn('w:tcW'))
                if w is not None: w.set(qn('w:w'), str(widths[i]))
            ps = tc.findall(qn('w:p'))
            for extra in ps[1:]: tc.remove(extra)
            pp = Paragraph(ps[0], d._body)
            for r in list(pp.runs): r._r.getparent().remove(r._r)
            pp.add_run(v)
    fill(hdr, header)
    for v in rows:
        tr = copy.deepcopy(tpl); tbl.append(tr); fill(tr, v)
    tbl.remove(tpl)
    # 표 격자 열 정의 맞추기
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        for c in cols[ncol:]: grid.remove(c)
        while len(grid.findall(qn('w:gridCol'))) < ncol: grid.append(copy.deepcopy(cols[0]))
        if widths:
            for c, w in zip(grid.findall(qn('w:gridCol')), widths): c.set(qn('w:w'), str(w))
    tr0 = tbl.tr_lst[0].get_or_add_trPr(); h = OxmlElement('w:tblHeader'); h.set(qn('w:val'), 'true'); tr0.append(h)
    d.add_paragraph()  # 표 뒤 간격

def images(items, cols=2, width=7.9):
    rows = math.ceil(len(items) / cols)
    t = d.add_table(rows=rows, cols=cols); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (path, sub) in enumerate(items):
        c = t.cell(i // cols, i % cols); c.vertical_alignment = 1
        p0 = c.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.add_run().add_picture(path, width=Cm(width))
        if sub:
            p1 = c.add_paragraph(); p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p1.add_run(sub); r.font.size = Pt(8.5)
    # 그림 표가 페이지에서 갈리지 않게: 행 분할 금지 + 다음 행·캡션과 함께
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr(); cs = OxmlElement('w:cantSplit'); cs.set(qn('w:val'), 'true'); trPr.append(cs)
        for c in row.cells:
            for pp in c.paragraphs: pp.paragraph_format.keep_with_next = True

# 시나리오 기준값 (앱 ① 산식, 가정 조건)
EFF = 4000 * 0.99 * 0.9 * 0.96
def tx(load, n, c, full=False):
    per = load * 1000 if full else load * 1000 / n
    u = max(1, math.ceil(per / EFF - 1e-9)); return u, u * (n + c)
S1a, S1b, S2a, S2b = tx(10, 3, 1), tx(20, 3, 1), tx(20, 3, 1), tx(20, 1, 1, True)
assert (S1a, S1b, S2a, S2b) == ((1, 4), (2, 8), (2, 8), (6, 12))
S3a, S3b = math.ceil(6000 / 30), math.ceil(6000 / 130)

# ===================== 본문 =====================
H1('국문초록')
P('데이터센터와 같은 설비 집약적 건축물은 전력 공급, 냉각, 이중화 방식이 설비실 규모와 공간 구성을 좌우한다. 그러나 초기 계획설계 단계에는 전문 분야의 검토 결과가 확정되지 않은 경우가 많아, 건축설계자가 경험에 기대어 기술 조건을 가정해야 한다. 본 연구는 건축설계자의 실무 지식을 Markdown(MD) 문서로 명시화하고, 생성형 AI를 활용한 바이브 코딩으로 검토용 Micro-App을 개발한 뒤, 이를 하나의 포털에서 공유하고 사용자 의견을 받아 개선하는 설계 지원 프레임워크를 제안한다.')
P('연구는 설계과학연구 방법론(DSRM)에 따라 수행하며, 저자가 선행 개발한 두 앱(Power Redundancy Calculator, AI Factory Study)을 출발점으로 삼는다. 계획서 작성 단계에서 앱 ①의 변압기 대수 산정 기능을 33개 입력 조건, 363개 항목에 대해 독립적으로 작성한 산식과 비교한 예비 시험에서 불일치는 나타나지 않았다. 향후 데이터센터 초기 계획 시나리오에서 기존 수작업, 개별 앱 사용, 통합 포털 사용의 세 조건을 비교하여 검토 정확도, 소요 시간, 사용성을 평가한다. 본 연구는 설계자를 AI로 대체하는 것이 아니라, 설계자가 다루는 전문 지식의 범위를 넓히고 그 근거를 관리하는 방법을 제시하는 데 의의가 있다.')
P('주제어: 초기 계획설계, 지식 명시화, Markdown, 바이브 코딩, Micro-App, 사용자 피드백, 데이터센터')

H1('제1장 서론')
H2('1.1 연구의 배경')
P('건축 기획과 초기 계획설계 단계의 결정은 사업 추진 여부와 이후 기본설계의 방향에 큰 영향을 준다. 데이터센터는 운영·관리 인력이 상주하는 사무·하역 공간과 함께, 사람이 상주하지 않는 데이터홀, 전기실, 기계실, 발전기실, 통신실 등 무중단 운영을 위한 기술 공간이 중첩되는 시설이다. 수전용량, IT 부하, 냉각 방식, 전력 이중화 방식이 바뀌면 장비 수량과 배치, 전기실·기계실·샤프트 규모와 동선을 다시 검토해야 한다.')
P('초기 단계에는 전문 용역의 참여 시점이 서로 다르고 기술 조건도 확정되지 않은 경우가 많다. 이때 건축설계자가 경험과 사례에 기댄 유추로 기술 조건을 가정하면, 이후 계획과의 차이가 커질 위험이 있다. 본 연구는 이 문제를 설계자의 지식 부족이 아니라, 분야별 지식을 정리하고 검증 가능한 형태로 관리하는 문제로 다룬다.')
P('생성형 AI는 실무자가 자연어로 요구사항을 설명하여 소규모 검토 도구를 직접 만들 수 있게 한다. 그러나 이렇게 만든 도구의 계산 정확도는 자동으로 보장되지 않으며, 개별 앱이 흩어져 있으면 재사용과 개선 이력 관리도 어렵다. 따라서 지식을 명시적으로 기록하고, 검증된 앱으로 구현하며, 공유와 개선을 반복할 수 있는 절차가 필요하다.')

H2('1.2 연구 목적과 연구 질문')
P('본 연구의 목적은 건축설계자의 전문 지식을 MD 문서로 명시화하고, 바이브 코딩으로 개발한 Micro-App을 통합 포털에서 공유하며, 사용자 의견을 통해 지속적으로 개선하는 설계 지원 프레임워크를 개발하고 평가하는 것이다.')
P('RQ1. 초기 계획설계에 필요한 전문 분야 지식(산식, 기준, 적용 조건)을 어떤 MD 구조와 검증 절차로 명시화할 수 있는가?')
P('RQ2. MD를 바탕으로 개발한 Micro-App을 어떻게 한곳에서 공유하고, 사용자 의견을 받아 지식과 앱을 갱신하는 절차로 연결할 수 있는가?')
P('RQ3. 제안한 방식은 기존 수작업 및 개별 앱 사용과 비교하여 검토 정확도, 소요 시간, 사용성에서 어떤 차이를 보이는가?')

H2('1.3 연구 범위')
P('적용 대상은 데이터센터 초기 계획설계의 전력·IT 인프라와 그에 따른 공간 검토이다. 개발 범위는 저자가 선행 개발한 두 앱과 이를 묶는 통합 포털이며, 포털은 앱 접근, 공통 설명, MD 열람, 개선 의견 전달을 제공하는 접근성 중심의 통합으로 한정한다. 앱 간 계산 결과의 자동 교환, 온톨로지 기반 의미 연결, BIM 연동은 후속 연구로 둔다. 사용자 의견은 담당자가 검토하여 반영하며 자동으로 지식을 수정하지 않는다. 앱의 결과는 개략 검토용이며 전문 엔지니어의 상세 설계를 대체하지 않는다.')

H1('제2장 이론적 배경 및 선행연구')
H2('2.1 암묵지의 명시화')
P('Nonaka와 Takeuchi(1995)는 개인의 경험에 축적된 암묵지가 표출화(externalization)를 거쳐 조직이 공유할 수 있는 형식지로 전환되는 과정을 지식 창조의 핵심으로 설명하였다. 본 연구에서 MD 문서화는 설계자의 경험 지식을 입력 조건, 산식, 적용 범위, 출처로 표출하는 단계이고, 앱을 통한 공유와 사용자 의견의 반영은 표출된 지식이 다시 실무에서 쓰이고 보완되는 순환에 해당한다.')
H2('2.2 선행연구와 본 연구의 차이')
P('이병수·지승열·전한종(2016)은 건축기획 단계의 매스 규모산정에 쓰이는 설계지식을 지식기반 BIM 템플릿과 플러그인, 규칙 기반 추론으로 재현하여 건축가에게 개략 결과를 빠르게 제공하고 축적된 지식을 재사용하는 방법을 제시하였다. 이 연구는 설계자의 경험 지식을 도구로 구조화할 수 있음을 보여 준다. 본 연구는 대상을 건축 외 전문 분야인 데이터센터의 전력·IT 인프라로 넓히고, 실무자가 생성형 AI로 직접 도구를 만들어 공유하고 개선하는 운영 절차까지 다룬다는 점에서 차이가 있다.')
H2('2.3 바이브 코딩과 Micro-App')
P('본 연구에서 Micro-App은 특정 설계 검토 문제를 위해 입력값, 계산 규칙, 출력 결과를 갖춘 소규모 실행 도구로 정의한다. 바이브 코딩은 사용자가 자연어 요구사항과 지식 문서를 제시하고 생성형 AI의 코드 생성과 수정을 활용하는 개발 방식으로 정의한다. 생성형 AI는 코드 작성을 돕지만 계산 결과의 타당성을 보증하지 않으므로, 계산 로직은 재현 가능한 함수로 구현하고 별도의 시험과 전문가 검토로 확인한다.')

H1('제3장 연구 방법')
H2('3.1 연구 절차')
P('본 연구는 Peffers 외(2007)의 설계과학연구 방법론(DSRM)에 따라 다음 절차로 수행한다. 전체 프레임워크는 그림 3-1과 같다.')
table(['DSRM 단계', '본 연구의 활동'], [
    ['문제 식별', '데이터센터 초기 계획에서 기술 가정 변경이 공간 재검토로 이어지는 문제 정리, 실무자 인터뷰'],
    ['목표 설정', '지식 명시화, 검증된 앱, 공유·개선 절차의 요구 조건 정의'],
    ['설계·개발', 'MD 지식 문서 작성, 기존 앱 보완, 통합 포털과 의견 전달 기능 구현'],
    ['시연', '데이터센터 초기 계획 시나리오에 적용'],
    ['평가', '세 조건 비교 실험, 사용자 의견 처리 전후 비교'],
    ['결과 공유', '논문 작성, MD·앱 변경 이력 공개'],
], widths=[2400, 7000])
images([(FIG + 'fig_3_1_framework.png', '')], cols=1, width=16)
caption('그림 3-1  연구 프레임워크: 지식 문서화–앱 개발–공유–피드백 개선 순환')

H2('3.2 MD 지식 문서')
P('각 앱에 대응하는 MD 문서에는 다음 항목을 기록하고, MD와 소스코드의 변경 이력을 연결하여 관리한다. 출처가 확인되지 않은 값은 가정으로 표시하고, 출처 간 수치가 다르면 하나를 고르지 않고 함께 적는다.')
table(['구성 항목', '기록 내용 (앱 ① 예시)'], [
    ['목적·적용 범위', '전력 이중화 방식에 따른 변압기 대수와 전기실 구성의 개략 검토'],
    ['입력·단위', 'IT 부하(MW), 이중화 방식(N+C), 변압기 정격(kVA), 역률, 안전율, UPS 효율'],
    ['산식·근거', '유효용량 = 정격 × 역률 × 안전율 × UPS 효율, 모듈당 대수 = ⌈설계부하 ÷ 유효용량⌉'],
    ['출력', '모듈당·총 변압기 대수, 평상시·상실 시 부하율'],
    ['검증', '기준 사례, 기대값, 앱 결과, 불일치와 수정 이력'],
], widths=[2400, 7000])

H2('3.3 검증 절차')
P('검증은 세 단계로 나누어 수행하고 각 단계에서 확인한 범위를 구분하여 보고한다. 첫째, MD의 변수와 산식을 소스코드의 함수와 대조한다. 둘째, 독립적으로 작성한 산식으로 기준 사례를 만들고 실제 앱 화면의 값과 비교한다. 셋째, 분야별 전문가가 산식의 적용 조건과 한계를 검토한다.')
P('예비 시험으로 앱 ①을 브라우저에서 실행하여 변압기 정격 2,000~4,000kVA, 안전율 0.75~0.9, POD 부하 3,000·6,000kW 등 33개 조건에서 유효용량과 다섯 가지 이중화 방식의 변압기 대수(363개 항목)를 비교하였고 불일치는 없었다. 이 결과는 문서의 산식과 구현이 일치함을 보인 것이며, 산식 자체의 공학적 타당성은 셋째 단계에서 확인한다.')

H2('3.4 통합 포털과 피드백 체계')
P('통합 포털은 앱별 목적, 적용 범위, 검증 상태, 링크, MD 문서와 변경 이력을 한 화면에서 제공한다(그림 3-2). 사용자는 대상 앱, 문제 유형(오류·계산·기능·UI·기타), 재현 조건, 개선 제안을 적어 의견을 보내고, 의견은 이메일로 담당자에게 전달된다. 담당자는 의견을 분류하고 전문 검토를 거쳐 MD와 앱을 수정한 뒤 재검증하여 배포한다. 의견 접수 이력은 담당자 측에 보존하여 처리율과 수정 내역을 평가에 활용한다.')
images([(FIG + 'fig_4_3_wireframe.png', '')], cols=1, width=16)
caption('그림 3-2  통합 포털 화면 구상안과 개선 의견 입력 양식 (구현 전 계획안)')

H2('3.5 평가 설계')
P('건축 실무자를 대상으로 같은 과제를 세 조건에서 수행하게 한다. 조건 1은 기존 수작업과 자료 검색, 조건 2는 개별 앱 사용, 조건 3은 통합 포털을 통한 앱과 MD 문서 활용이다. 참가자는 세 조건을 모두 수행하며 학습 효과를 줄이기 위해 조건 순서를 균형화하고, 난이도가 같은 과제 세트를 조건마다 다르게 배정한다. 과제는 한 번에 하나의 조건만 바꾸는 시나리오로 구성하며, 표의 수치는 실제 프로젝트가 아닌 가정 조건이다.')
table(['시나리오', '변경 조건 (가정)', '계산 항목 기준값 (앱 ① 산식)'], [
    ['S1 전력 확대', 'IT 부하 10MW → 20MW, 3oo4, 변압기 4,000kVA', f'총 {S1a[1]}대 → {S1b[1]}대 (모듈당 {S1a[0]}대 → {S1b[0]}대)'],
    ['S2 이중화 변경', 'IT 20MW, N+1(N=3) → 2N', f'총 {S2a[1]}대 → {S2b[1]}대'],
    ['S3 랙 밀도 변경', 'POD 6,000kW 고정, 랙당 30kW → 130kW', f'{S3a}랙 → {S3b}랙'],
], widths=[2000, 3800, 3600])
P('계산 항목은 위 기준값으로, 공간·배치 판단 항목은 전문가가 사전에 작성한 기준 답안으로 채점한다. 측정 지표는 다음과 같다. 표본이 작으므로 기술통계와 사례 해석을 중심으로 결과를 보고한다.')
table(['평가 영역', '측정 방법'], [
    ['검토 정확도', '기준 답안의 필수 항목 중 올바르게 식별한 비율, 계산 결과 일치 여부'],
    ['근거 제시', '출처와 적용 조건을 함께 제시한 판단의 비율'],
    ['소요 시간', '과제 시작부터 제출까지의 시간'],
    ['사용성', '사용성 설문과 사후 인터뷰'],
    ['개선 효과', '의견 처리 전후 버전의 오류·누락 수 비교'],
], widths=[2400, 7000])

H1('제4장 선행 개발 앱과 예비 검증')
H2('4.1 선행 개발 앱')
P('연구의 출발점은 저자가 Claude Code를 사용해 바이브 코딩으로 개발한 두 웹앱이다. 아래 내용은 2026년 11월 18일 기준 소스코드와 배포본을 확인한 결과이며, 두 앱의 배포본은 저장소 소스와 동일하다.')
table(['구분', '앱 ① Power Redundancy Calculator', '앱 ② AI Factory Study'], [
    ['목적', '이중화 방식별 변압기 대수, 전기실·POD 급전 관계, 데이터홀 평면·냉각 개략 검토', 'AI 서버 랙·POD 구성과 건축 공간 요구를 공개 자료의 출처와 함께 정리'],
    ['성격', '입력값 기반 계산·시각화 도구', '출처와 근거 수준(확인·해석·가정·미검증)을 표기한 지식 문서형 앱'],
    ['규모·이력', 'HTML/CSS/JS 약 10,800줄, 73커밋', 'HTML/CSS/JS 약 1,500줄, 3커밋'],
    ['검증', '산식·코드 대조, 기준 사례 예비 시험(363항목 불일치 없음)', '자체 회귀 시험 38개 통과'],
    ['의견 수집', '앱 내 입력·분류 후 이메일 전송', '없음'],
], widths=[1800, 3800, 3800])
images([
    (FIG + 'fig_app1_a_result.png', '(a) 계산 결과'),
    (FIG + 'fig_app1_b_diagram_crop.png', '(b) 전기실 세트와 POD 급전 관계'),
    (FIG + 'fig_app1_c_datahall_crop.png', '(c) 데이터홀 평면·냉각 검토'),
    (FIG + 'fig_app1_d_feedback.png', '(d) 개선 의견 입력 창'),
])
caption('그림 4-1  앱 ① Power Redundancy Calculator 주요 화면 (기본 입력값)')
images([
    (FIG + 'fig_app2_a_home.png', '(a) 홈'),
    (FIG + 'fig_app2_b_pod.png', '(b) POD 랙 구성 편집'),
    (FIG + 'fig_app2_c_glossary.png', '(c) 용어와 출처 각주'),
    (FIG + 'fig_app2_d_sources.png', '(d) 근거 목록'),
])
caption('그림 4-2  앱 ② AI Factory Study 주요 화면')

H2('4.2 예비 검증에서 확인한 점')
P('첫째, 앱 ①의 변압기 대수 산정은 문서의 산식과 일치하였다. 둘째, 앱 ②는 수치마다 출처를 등록하고 공개 출처가 없는 값은 가정으로 표시하며, 자체 시험 기록에서 랙 외형(0.8×1.4m → 0.6×1.07m)과 바닥 하중(약 1.6 → 2.8tf/m²)을 스스로 정정한 이력이 있다. 셋째, 두 앱 사이의 불일치가 발견되었다. 앱 ②는 랙당 전력 227kW를 공개 출처가 없는 가정값으로 낮추었으나 앱 ①은 이 값을 기본 랙 구성에 그대로 쓰고 있다. 이는 한 앱에서 수정된 지식이 다른 앱으로 전파되지 않는 문제를 보여 주며, MD를 지식 원본으로 두고 변경 이력을 관리해야 하는 이유이자 본 연구의 개선 절차가 다룰 첫 사례이다. 넷째, 앱 ①의 의견은 사용자 브라우저에만 저장되어 담당자 측 이력이 남지 않으므로, 통합 포털에서 접수 이력을 보존하도록 보완한다.')

H2('4.3 연구 일정')
table(['단계', '수행 내용', '산출물'], [
    ['1단계', '실무자 인터뷰, 검토 항목 정리', '검토 항목 목록'],
    ['2단계', '앱별 MD 작성, 검증 1·2단계 수행', 'MD 문서, 검증 기록'],
    ['3단계', '통합 포털과 의견 전달 기능 구현', '포털 시제품'],
    ['4단계', '시나리오 비교 실험', '정확도·시간·사용성 자료'],
    ['5단계', '의견 반영과 재검증', '수정 이력, 전후 비교'],
    ['6단계', '결과 분석과 논문 작성', '연구 결과와 한계'],
], widths=[1400, 4600, 3400])

H1('제5장 기대 효과와 한계')
P('본 연구는 건축설계자의 실무 지식을 문서와 실행 도구로 바꾸고, 공유와 개선까지 하나의 절차로 운영하는 방법을 제시한다. 실무적으로는 초기 계획 단계에서 전문 분야의 개략 검토 도구에 쉽게 접근하고, 계산의 전제와 근거를 함께 확인하며, 전문 엔지니어와 협의할 쟁점을 미리 정리하는 데 도움이 될 것으로 기대한다.')
P('한계로, 통합은 앱 접근과 지식·의견 관리에 한정되며 앱 간 계산 결과의 자동 연결은 다루지 않는다. 데이터센터 사례에 집중하므로 다른 건축 유형으로 일반화하는 데 제약이 있다. 후속 연구에서는 앱 간 데이터 연동, 온톨로지를 이용한 분야 간 관계 표현, BIM 연계를 단계적으로 검토한다.')

H1('참고문헌')
P('Nonaka, I., & Takeuchi, H. (1995). The Knowledge-Creating Company: How Japanese Companies Create the Dynamics of Innovation. Oxford University Press.')
P('Peffers, K., Tuunanen, T., Rothenberger, M. A., & Chatterjee, S. (2007). A design science research methodology for information systems research. Journal of Management Information Systems, 24(3), 45–77.')
P('이병수, 지승열, 전한종 (2016). 건축기획 매스 규모산정의 설계지식 재현을 위한 지식기반 BIM 시스템 구현. 한국CDE학회 논문집, 21(3), 252–266. https://doi.org/10.7315/CDE.2016.252')

H1('부록. 생성형 AI 활용 표기')
P('선행 개발 앱 두 개는 저자가 Claude Code(Anthropic)를 활용해 개발하였다. 본 계획서 작성 과정의 앱 소스 분석, 기준 사례 시험 스크립트 작성, 화면 캡처와 그림 작성, 문장 초안 작성에 생성형 AI(Claude)를 보조 도구로 사용하였으며, 모든 내용은 저자가 검토하여 확정하였다.')
P('그림 출처: 그림 4-1은 power-redundancy-calculator.vercel.app, 그림 4-2는 ai-factory-study.vercel.app의 배포본을 2026년 11월 18일 캡처한 것이다. 그림 3-1과 3-2는 저자가 작성하였다.', size=9)

d.save(OUT); print('saved', OUT)
