# -*- coding: utf-8 -*-
"""연구계획서 v4(제출용 풀세트) 생성.
원본: plan_v2_original.docx  →  출력: 연구계획서_v4_제출본.docx
표기 규칙: «...» 안의 글은 붉은색(저자가 작성·확인할 부분). [작성 필요] 등 꺾쇠 표기도 자동으로 붉게 칠한다.
"""
import copy, math, re
import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

SRC = 'plan_v2_original.docx'
OUT = '연구계획서_v4_제출본.docx'
FIG = 'figures/'
RED = RGBColor(0xFF, 0x00, 0x00)

d = docx.Document(SRC)
P = list(d.paragraphs)   # 원본 문단 인덱스(고정)
T = list(d.tables)       # 0:2.7 1:3.3 2:3.5 3:3.6 4:4.5
PARENT = P[12]._parent
BODY_TPL = P[12]._p
H1_TPL = P[109]._p
DONE = set()             # 이미 서식을 정한 문단(_p id)

# ---------------- 기본 도구 ----------------
def _clear(p):
    for r in list(p.runs):
        r._r.getparent().remove(r._r)

def rich(p, text, base_bold=None, size=None):
    """«» 구간은 붉은색·굵게. 나머지는 기본 서식."""
    _clear(p)
    for i, seg in enumerate(re.split(r'«|»', text)):
        if not seg:
            continue
        r = p.add_run(seg)
        red = (i % 2 == 1)
        if red:
            r.font.color.rgb = RED
            r.bold = True
        elif base_bold is not None:
            r.bold = base_bold
        if size:
            r.font.size = Pt(size)
    DONE.add(id(p._p))
    return p

def par_after(anchor_el, text, tpl=None, align=None, size=None, bold=None):
    el = copy.deepcopy(tpl if tpl is not None else BODY_TPL)
    anchor_el.addnext(el)
    p = Paragraph(el, PARENT)
    rich(p, text, base_bold=bold, size=size)
    if align is not None:
        p.alignment = align
    return el

def append(p, text):
    rich(p, p.text + text)

def replace(p, old, new):
    assert old in p.text, (old, p.text[:60])
    rich(p, p.text.replace(old, new))

def delete(p):
    p._p.getparent().remove(p._p)

def styled_table(anchor_el, header, rows):
    """원본 표(Light Shading Accent 1, 3열) 서식을 복제해 새 표를 만든다."""
    tbl = copy.deepcopy(T[0]._tbl)
    anchor_el.addnext(tbl)
    trs = tbl.tr_lst
    hdr, tpl = trs[0], trs[1]
    for tr in trs[2:]:
        tbl.remove(tr)
    def fill(tr, vals):
        for tc, v in zip(tr.tc_lst, vals):
            pp = Paragraph(tc.findall(qn('w:p'))[0], PARENT)
            rich(pp, v)
    fill(hdr, header)
    for v in rows:
        tr = copy.deepcopy(tpl); tbl.append(tr); fill(tr, v)
    tbl.remove(tpl)
    return tbl

def _cell_border(cell, color='FF0000', style='dashed', sz='12'):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement(f'w:{side}')
        e.set(qn('w:val'), style); e.set(qn('w:sz'), sz); e.set(qn('w:space'), '0'); e.set(qn('w:color'), color)
        b.append(e)
    tcPr.append(b)

def _move(tbl, anchor_el):
    anchor_el.addnext(tbl._tbl)
    return tbl._tbl

def caption(anchor_el, num, text):
    return par_after(anchor_el, f'그림 {num}  {text}', align=WD_ALIGN_PARAGRAPH.CENTER, size=9)

def image_grid(anchor_el, items, cols=2, width_cm=7.9):
    """items: [(이미지 경로, '(a) 설명'), ...]"""
    rows = math.ceil(len(items) / cols)
    t = d.add_table(rows=rows, cols=cols)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (path, sub) in enumerate(items):
        c = t.cell(i // cols, i % cols)
        c.vertical_alignment = 1  # center
        p0 = c.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.add_run().add_picture(path, width=Cm(width_cm))
        if sub:
            p1 = c.add_paragraph(); p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rich(p1, sub, size=8.5)
    return _move(t, anchor_el)

def placeholder_grid(anchor_el, labels, height_cm=5.5):
    """붉은 점선 사진 자리. labels: ['설명', ...] (한 줄에 나란히)"""
    t = d.add_table(rows=1, cols=len(labels))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.rows[0].height = Cm(height_cm); t.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    for i, lab in enumerate(labels):
        c = t.cell(0, i); _cell_border(c); c.vertical_alignment = 1
        p0 = c.paragraphs[0]; p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rich(p0, '«[사진 삽입]»', size=10)
        p1 = c.add_paragraph(); p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rich(p1, f'«{lab}»', size=9)
    return _move(t, anchor_el)

# ---------------- 시나리오 기준값 (앱 ① 산식, 가정 조건) ----------------
EFF = 4000 * 0.99 * 0.9 * 0.96                     # 3,421.44 kW
def units(load_mw, n, c, full=False):
    per = load_mw * 1000 if full else load_mw * 1000 / n
    u = max(1, math.ceil(per / EFF - 1e-9))
    return u, u * (n + c), (load_mw * 1000 / n) / (u * EFF) * 100
s1a, s1b = units(10, 3, 1), units(20, 3, 1)
s2a, s2b = units(20, 3, 1), units(20, 1, 1, full=True)
s3a, s3b = math.ceil(6000 / 30), math.ceil(6000 / 130)
assert (s1a[:2], s1b[:2], s2a[:2], s2b[:2], s3a, s3b) == ((1, 4), (2, 8), (2, 8), (6, 12), 200, 47)

# =====================================================================
# 표지
rich(P[4], '«[확인 필요] 제출 전 제목·부제 길이, 과목 제출 양식, 표지 정보(소속·담당교수·날짜) 확인»')
par_after(P[4]._p, '«이 문서의 붉은 글씨와 붉은 점선 칸은 저자가 직접 작성하거나 확인할 부분이다. 제출 전 내용을 채우고 검정으로 바꾸거나 삭제한다.»', size=9.5)

# 국문초록
append(P[8], ' 계획서 작성 단계에서는 선행 개발한 앱 2개를 대상으로 문서-코드 대조와 기준 사례 예비 시험을 수행하여 제안한 검증 절차의 실행 가능성을 확인하였다.')

# 1.1
rich(P[15], '«[작성 필요] 프로젝트 A·B(익명): 초기 기술 가정(IT 부하, 이중화 방식, 랙 밀도 등)이 바뀌어 전기실·기계실·데이터홀 공간을 재검토한 사례 1~2개를 3~5문장으로 기술. 포함 항목: ① 변경 전·후 조건 ② 영향을 받은 공간과 면적 변화 ③ 재검토 기간·회차 ④ 당시 판단 근거. 수치마다 실제/가정 여부 표기»')
g = placeholder_grid(P[15]._p, ['프로젝트 A 변경 전 평면(또는 단면)\n권장: 가로형, 폭 8cm, 실명·발주처 가림', '프로젝트 A 변경 후 평면(또는 단면)\n변경으로 영향받은 실을 표시'])
caption(g, '1-1', '«프로젝트 A의 초기 기술 가정 변경 전·후 공간 비교(익명)»')

# 1.4
append(P[26], ' 아울러 저자가 선행 개발한 데이터홀 평면·냉각 검토 앱과 발전기실 급·배기 검토 앱은 앱 ①·②의 보완 사례로 4.1절에 간략히 기술하며, 온톨로지 기반 검토의견서 에이전트는 임시 형태의 시제품이므로 평가 범위에 포함하지 않는다.')
rich(P[28], '통합 플랫폼의 구현 범위는 «앱 목록과 링크, 앱별 목적·적용 범위·검증 상태의 공통 설명, MD 지식 문서 열람, 변경 이력, 개선 의견 양식과 이메일 전달»로 한다(4.3절, 그림 4-3). «[확인 필요] 위 범위와 MD 열람 방식(포털 내 표시 또는 저장소 링크) 확정»')

# 2.3
replace(P[36], '바이브 코딩은 사용자가', '바이브 코딩(Karpathy, 2025)은 사용자가')
append(P[37], ' 본 연구의 선행 앱은 저자가 Claude Code를 사용해 이러한 방식으로 개발한 것이며, 그 과정은 4.1절에서 회고적 사례로 기술한다.')

# 2.5
rich(P[42], '지식 명시화의 이론적 근거로, Nonaka와 Takeuchi(1995)의 지식 창조 모형은 개인의 암묵지가 표출화(externalization)를 거쳐 조직이 공유하는 형식지로 전환되는 과정을 설명한다. 본 연구의 MD 문서화는 표출화 단계에, 앱을 통한 공유와 피드백 반영은 연결화(combination)와 내면화(internalization) 단계에 대응시킬 수 있다. 실무자가 직접 도구를 만드는 측면은 최종 사용자 개발(end-user development) 연구와 관련된다. Ko 외(2011)는 최종 사용자가 만든 소프트웨어의 오류와 검증 문제를 정리하였고, Fischer 외(2004)의 메타 디자인 개념은 사용자가 사용 중에 도구를 수정·확장할 수 있도록 설계하는 방향을 제시한다. 이 연구들은 본 연구의 피드백 순환 설계와 생성 코드 검증 절차의 이론적 근거로 활용한다. «[검증 필요] 위 3편의 원문을 확인하고 대응 관계 서술 보완»')

# 2.6
append(P[44], ' 저자가 별도로 시험 제작한 온톨로지 기반 검토 에이전트는 임시 형태의 시제품으로서 본 연구의 개발·평가 범위에 포함하지 않으며, 후속 연구의 출발점으로만 다룬다.')

# 3.1 + 그림 3-1
replace(P[51], 'Peffers 외(2007)의 설계과학연구 방법론(DSRM)을 참조하여', 'Hevner 외(2004)의 설계과학 연구 지침과 Peffers 외(2007)의 설계과학연구 방법론(DSRM)을 참조하여')
append(P[51], ' 연구 프레임워크는 그림 3-1과 같다.')
g = image_grid(P[51]._p, [(FIG + 'fig_3_1_framework.png', '')], cols=1, width_cm=16)
caption(g, '3-1', '연구 프레임워크: 지식 문서화–앱 개발–공유–피드백 개선 순환')

# 3.2
rich(P[54], '인터뷰 대상은 «데이터센터 설계·시공 실무 경력 5년 이상»의 건축·전기·기계 분야 실무자로 하고, 분야별 «2~4명(전체 8~10명 내외)»을 목표로 유의 표집과 눈덩이 표집을 병행한다. 질문지는 ① 초기 설계에서 반복하는 기술 검토 항목 ② 입력 조건과 산출 근거 ③ 검토 실패·재작업 경험 ④ 도구 사용과 개선 의견 전달 방식의 4개 영역으로 구성한다. 인터뷰는 서면 동의를 받아 녹음·전사하고, 프로젝트명과 발주처는 익명 처리한다. «[작성 필요] 대상자 익명 코드 목록, 수행 일정, 질문지 원문(부록 추가), 학교 IRB 심의 대상 여부»')

# 3.3 검증 절차 (MD 표 뒤)
e = par_after(T[1]._tbl, '검증은 다음 3단계로 구분하여 수행하고, 각 단계에서 확인한 범위를 분리하여 보고한다. 첫째, 문서-코드 대조 단계에서는 MD의 변수·산식과 소스코드의 함수를 일대일로 대조하여 불일치를 기록한다. 둘째, 기준 사례 시험 단계에서는 독립적으로 작성한 계산식으로 기준 사례를 만들고 실제 앱의 화면 값과 비교한다. 셋째, 전문가 검토 단계에서는 분야별 검토자가 기준 답안과 적용 한계를 확인한다.')
par_after(e, '예비 시험으로 앱 ①에서 변압기 정격 2,000~4,000kVA, 안전율 0.75~0.9, POD 부하 3,000·6,000kW 등 33개 조건, 363개 항목을 비교하여 불일치를 발견하지 못했다. 이 시험은 설명서의 산식과 구현의 일치를 확인한 것으로서 산식의 공학적 타당성이나 세트 분할, 급전 방식, 냉각 계산을 포함하지 않는다. 앱 ②는 자체 회귀 시험(38개 항목)을 갖추고 있으며 출처 등록과 근거 수준(해석·가정·미검증) 표기를 사실 관리 규칙으로 운영한다. 전문가 검토 단계는 아직 수행하지 않았다. 출처가 확인되지 않은 입력값은 MD와 화면에 가정으로 표시하고, 출처 간 수치가 다르면 하나를 선택하지 않고 병기한다.')

# 3.4
replace(P[59], '오류·기능 개선·지식 추가 등으로 분류한다', '오류·계산·기능·UI·기타(4.3절)로 분류하고, 지식 추가가 필요한 의견은 별도로 표시한다')
par_after(P[60]._p, '저자가 개발한 앱에서는 두 가지 전송 방식이 확인된다. 앱 ①은 사용자 브라우저에 의견을 저장한 뒤 주간 취합본을 이메일(mailto) 본문으로 보내며, 접수 이력이 사용자 브라우저에만 남는다. 데이터홀 검토 앱은 서버가 모든 사용자의 의견을 한곳에 수집하고 관리자가 반영·보류·반려로 분류하며, 첨부된 설계값으로 오류 조건을 재현할 수 있다. 본 연구는 두 방식의 접수 이력 보존성과 구현 부담을 비교하여 통합 플랫폼의 전송 방식을 선택한다.')
rich(P[61], '«[확인 필요] 최종 전송 방식 선택, 이메일 수신자, 개인정보 수집 범위, 접수 이력 보존 기간, 피드백 처리 책임자»')

# 3.5 기준값
e = par_after(T[2]._tbl, '표의 수치는 실제 프로젝트가 아닌 가정 조건이다. 평가 과제 중 계산 항목의 기준값은 앱 ①의 검증된 산식 범위(변압기 대수 산정)에서 다음과 같이 산출되며, 공간·배치 판단 항목은 전문가 기준 답안으로 별도로 정한다.')
t = styled_table(e, ['시나리오', '조건 (가정)', '계산 항목 기준값'], [
    ['S1 전력 확대', '3oo4, 변압기 4,000kVA, 역률 0.99, 안전율 0.9, UPS 효율 0.96 (유효용량 3,421.4kW)', f'10MW: 모듈당 {s1a[0]}대, 총 {s1a[1]}대 → 20MW: 모듈당 {s1b[0]}대, 총 {s1b[1]}대 (1개 모듈 상실 시 부하율 {s1b[2]:.1f}%)'],
    ['S2 이중화 변경', 'IT 20MW, S1과 같은 변압기 조건', f'N+1(N=3): 모듈당 {s2a[0]}대, 총 {s2a[1]}대 → 2N: 모듈당 {s2b[0]}대, 총 {s2b[1]}대'],
    ['S3 랙 밀도 변경', 'POD 6,000kW 고정', f'랙당 30kW: {s3a}랙 → 랙당 130kW: {s3b}랙 (올림)'],
    ['S4 사용자 개선', '앱 오류·누락 의견 1건 접수', '계산 기준값 없음. 처리 단계별 소요 시간과 수정 범위 기록'],
])
rich(P[65], '«[작성 필요] 프로젝트 A·B의 실제 조건(익명)으로 시나리오 수치를 바꿀지 결정, 시나리오별 전문가 기준 답안(정답 항목·허용 오차) 작성, 관련 평면·단면 이미지 준비»')
P[65]._p.getparent().remove(P[65]._p); t.addnext(P[65]._p)

# 3.6
p67 = P[67].text
for a, b in [('A(기존', '조건 1(기존'), ('B(개별', '조건 2(개별'), ('C(통합', '조건 3(통합'), ('C의 효과', '조건 3의 효과')]:
    assert a in p67, a; p67 = p67.replace(a, b)
rich(P[67], p67)
replace(P[68], 'A/B/C 단면 비교와 별도로', '조건 1~3의 비교와 별도로')
rich(P[71], '참가자는 «건축 실무자 12명 내외(경력 5년 이상과 미만 각 6명)»로 하고, 참가자 내 설계로 세 조건을 모두 수행하되 조건 순서는 라틴 방격으로 균형화한다. 난이도를 맞춘 동형 과제 세트 «3개(세트당 계산 항목 3개, 판단 항목 3개)»를 조건별로 하나씩 사용한다. 채점은 «2인»이 독립적으로 수행하고 일치도는 Cohen의 카파(Cohen, 1960)로 보고하며, 사용성은 SUS(Brooke, 1996)로 측정한다. 조건 간 차이는 Friedman 검정과 사후 Wilcoxon 부호순위 검정(Bonferroni 보정)으로 탐색적으로 분석하고 효과 크기와 사례 해석을 함께 제시한다. 단기 과제 수행 결과만으로 전문 지식의 확장을 입증한다고 주장하지 않으며, 지식 활용 지표를 함께 해석한다. «[확인 필요] 참가자 수·과제 수·평가자 수 확정»')

# 3.7
append(P[73], ' 실제 프로젝트 사례는 프로젝트 A, B 등 익명 기호로 표기하며 사업명, 발주처, 대지 위치 등 식별 정보는 기재하지 않는다. 앱 분석 대상은 저자가 직접 개발한 앱으로 한정한다. 인간 참여자를 대상으로 하는 인터뷰와 사용자 실험은 «[확인 필요] 한양대학교 IRB 심의 또는 면제 여부»를 확인한 후 수행한다.')

# 4.1
rich(P[76], '연구 착수 시점의 선행 개발 자산은 저자가 Claude Code로 개발한 정적 웹앱 두 개이다. 본 절의 서술은 2026년 11월 18일 기준 소스코드, 커밋 이력, 시험 기록과 배포본을 확인한 내용에 한하며, 개발 과정은 회고적 사례로 기술한다. 앱 ①은 Power Redundancy Calculator(https://power-redundancy-calculator.vercel.app/), 앱 ②는 AI Factory Study(https://ai-factory-study.vercel.app/)이다. 두 앱의 배포본은 저장소 소스와 동일함을 확인하였다.')
t = styled_table(P[76]._p, ['구분', '앱 ① Power Redundancy Calculator', '앱 ② AI Factory Study'], [
    ['목적', '수전·IT 용량에서 데이터홀별 이중화 방식, 변압기 대수, 전기실·POD 급전 관계, 기준층 평면, 랙·냉각 개략 검토', 'AI Factory가 요구하는 건축 공간을 공개 자료의 출처와 함께 정리한 참고 앱'],
    ['성격', '입력값 기반 계산·시각화 도구', '출처 등록과 근거 수준 표기를 갖춘 지식 문서형 앱'],
    ['주요 입력', '수전·IT 용량, 데이터홀 수, 이중화 방식(2oo3, 3oo4, N+1, N+2C, 2N 등), 변압기 정격(상한 4,000kVA), 역률, UPS 효율, 안전율, POD 부하·수량, 랙 구성', '랙 종류별 수량과 랙당 전력(kW), 열 수, 용도·용량 한도'],
    ['주요 출력', '모듈당 설계부하, 모듈당·총 변압기 대수, 평상시·비상시 부하율, 구성 다이어그램, 자동 검토 의견, FWU·CDU 대수', '랙 구성 합계 전력, 열 배치안, 하중·외형 참고값, 출처 각주'],
    ['계산 근거', '유효용량 = 정격 × 역률 × 안전율 × UPS 효율, 모듈당 변압기 = ⌈설계부하 ÷ 유효용량⌉, 총 대수 = 모듈당 × (N+C)', '랙 수 × 랙당 kW 합계, 중량 ÷ 접지면적 등 단순 산술. 입력값은 출처 또는 가정·해석·미검증 태그로 구분'],
    ['구현 규모', 'HTML/CSS/JS 약 10,800줄, 73커밋 (2026.08~11)', 'HTML/CSS/JS 약 1,500줄, 3커밋 (2026.10)'],
    ['검증 현황', '산식·코드 대조 및 기준 사례 예비 시험(363항목 불일치 없음, 변압기 대수 산정 범위)', '자체 회귀 시험 38개 통과, 원문 대조일이 등록된 출처 19건/22건'],
    ['개선 의견', '앱 내 입력·분류·상태 관리, 이메일(mailto) 전송', '없음'],
])
e = par_after(t, '두 앱은 앱 ①에서 앱 ②로 이동하는 링크와 POD 구성 JSON의 수동 붙여넣기로만 연결되어 있으며 계산 결과가 자동으로 교환되지는 않는다. 이는 1.4절에서 정한 접근성 중심의 통합 수준과 일치한다. 그림 4-1과 그림 4-2는 두 앱의 주요 화면이다. 앱 ①은 기본 입력값에서 보 밑 환기(RA) 공간이 부족하면(그림 4-1 (c), −0.05m) 붉게 경고하는 등 입력 조건의 문제를 화면에서 바로 드러낸다.')
g = image_grid(e, [
    (FIG + 'fig_app1_a_result.png', '(a) ③ 계산 결과: 모듈당 설계부하·유효용량·총 변압기 대수'),
    (FIG + 'fig_app1_b_diagram_crop.png', '(b) ④ 구성 다이어그램: 전기실 세트와 POD 급전 관계'),
    (FIG + 'fig_app1_c_datahall_crop.png', '(c) ⑥ 데이터홀 평면 검토: 냉각 분담·POD 평면'),
    (FIG + 'fig_app1_d_feedback.png', '(d) 개선 의견 입력 창'),
])
e = caption(g, '4-1', '앱 ① Power Redundancy Calculator 주요 화면 (기본 입력값, 가정 조건)')
g = image_grid(e, [
    (FIG + 'fig_app2_a_home.png', '(a) 홈: 주제별 진입 화면'),
    (FIG + 'fig_app2_b_pod.png', '(b) Server POD Study: 랙 구성 편집'),
    (FIG + 'fig_app2_c_glossary.png', '(c) 용어: 출처 각주와 근거 수준 표기'),
    (FIG + 'fig_app2_d_sources.png', '(d) 근거: 출처 간 수치 불일치 목록'),
])
e = caption(g, '4-2', '앱 ② AI Factory Study 주요 화면')
e = par_after(e, '현재 한계는 다음과 같다. 첫째, 앱 ①의 공랭 분담 곡선, 일부 층고 상수, 랙 밀도 기준은 코드상 설계 관행에서 도출한 값이며 제조사 확정값이 아니다. 둘째, 앱 ①은 개선 의견을 사용자 브라우저에만 저장하므로 개발자 측 접수 이력이 자동으로 보존되지 않는다. 셋째, NVL72 랙당 전력 227kW에 대해 앱 ②는 공개 출처가 없어 가정으로 표시한 반면 앱 ①은 이를 기본 랙 구성에 사용한다. 이는 앱 간 지식 갱신이 자동으로 전파되지 않는다는 점을 보여 주며, 본 연구의 개선 절차가 다룰 첫 사례로 활용한다. 넷째, 실무 프로젝트 적용 결과와 전문 엔지니어 검토는 아직 확보하지 못했다.')
e = par_after(e, '개발 이력 측면에서 앱 ①의 저장소 이력은 2026년 8월 16일의 파일 분리 커밋부터 남아 있어 그 이전 이력은 확인할 수 없고, 앱 ②는 v0.10.3 첫 커밋(2026년 10월 28일)이며 이전 버전은 시험 기록으로만 확인된다.')
e = par_after(e, '보완 사례로 두 개의 앱을 추가로 참조한다. 데이터홀 검토 앱은 랙·냉각 설비(FWU, CDU)·기둥·보 간의 간섭을 평면·단면으로 검토하며 앱 ①의 데이터홀 검토 기능과 풍량·CDU 산식이 같다. 이는 같은 지식이 여러 저장소에 중복 구현될 때 생길 수 있는 불일치 위험과 MD 단일 원본의 필요성을 보여 준다. 이 앱은 서버가 사용자 의견을 수집하고 관리자가 반영·보류·반려로 분류하는 피드백 체계를 갖추고 있다. 발전기실 급·배기 계산 앱은 제조사 카탈로그를 코드와 분리된 데이터로 두고 계산 로직을 한 파일에 두어 화면과 API가 함께 사용하며, 풍량에서 역산한 개구면적과 카탈로그 값을 함께 출력하는 자체 점검을 포함한다. 이 앱의 결과는 수기 계산과 일치하였다. 두 보완 사례 앱은 자동 시험을 갖추지 않았다.')
delete(P[77])
rich(P[78], '«[확인 필요] 그림 4-1·4-2는 2026-11-18 배포본을 기본 입력값으로 캡처한 것이다. 논문용으로 특정 가정 조건(예: 3.5절 S1)을 넣은 화면이 필요하면 교체»')
P[78]._p.getparent().remove(P[78]._p); e.addnext(P[78]._p)

# 4.2
rich(P[80], '각 앱에 대응하는 MD 문서를 작성하고, 실제 코드의 변수·산식·결과와 대조한다. 기술 기준의 출처와 적용 시점을 기록하며, 기준 사례 시험을 통과한 기능만 실무 검토에 활용한다. 현재 앱 ①의 변압기 대수 산정 범위는 예비 시험을 마쳤고, 그 밖의 기능(세트 분할·급전 방식·냉각)과 전문가 검토는 계획 단계이다. MD 수정이 코드 수정으로 자동 반영되는 구조는 현재 계획에 포함하지 않으며, 수정·검증·배포 절차를 명시적으로 운영한다.')

# 4.3
e = par_after(P[82]._p, '피드백 입력 항목은 대상 앱, 문제 유형(앱 ①의 분류 체계인 오류·계산·기능·UI·기타), 재현 조건(설계값 파일 첨부), 개선 제안, 선택 연락처로 구성하는 것을 기본안으로 한다. 그림 4-3은 구현 전 화면 구상안이다.')
g = image_grid(e, [(FIG + 'fig_4_3_wireframe.png', '')], cols=1, width_cm=16)
e = caption(g, '4-3', '통합 플랫폼 화면 구상안(와이어프레임)과 개선 의견 입력 양식 — 구현 전 계획안')
g = placeholder_grid(e, ['개선 의견 이메일 수신 화면\n(플랫폼 구현 후 캡처, 개인정보 가림)'], height_cm=4.5)
g = caption(g, '4-4', '«개선 의견 이메일 수신 예시(구현 후 삽입)»')
rich(P[83], '«[작성 필요] 통합 플랫폼 구현 후 그림 4-4 캡처 삽입. 구현 전 제출이면 그림 4-4를 삭제하고 그림 4-3만 유지»')
P[83]._p.getparent().remove(P[83]._p); g.addnext(P[83]._p)  # 캡션 뒤로

# 4.4
rich(P[86], '«[작성 필요] 데이터센터 사례조사 1~2건(프로젝트 A, B 등 익명 표기): 프로젝트 조건, 앱 적용 전후 검토 절차, 화면·도면, 판단 근거, 전문가 검토 의견. 아래 그림 4-5 자리에 이미지 삽입»')
g = placeholder_grid(P[86]._p, ['프로젝트 A 조건을 넣은 앱 ① 화면\n(가정/실제 수치 구분 표기)', '프로젝트 A 관련 평면·단면\n(익명, 식별 정보 가림)'])
caption(g, '4-5', '«프로젝트 A 적용 사례: 앱 검토 화면과 관련 도면(익명)»')
rich(P[87], '개발자 재검토에 의한 개선 사례는 앱 ②의 시험 기록(2026년 10월 28일)에서 확인된다. 랙 외형 가정을 0.8×1.4m에서 언론 보도 기준 0.6×1.07m로, 랙 직하부 하중을 약 1.6에서 약 2.8tf/m²로 정정하였고, 공개 출처를 찾을 수 없는 랙당 전력(227/188kW)을 가정·미검증으로 격하하였다. 이는 사용자 피드백이 아닌 개발자의 자체 재검토 사례이며, 앞서 기술한 앱 간 227kW 불일치와 함께 오류 확인, 지식 수정, 재검증, 반영 범위 확인의 개선 순환을 추적하는 소재로 활용한다.')
par_after(P[87]._p, '«[작성 필요] 사용자 피드백 사례: 실제 접수 의견이 있으면 추가하고, 없으면 가상 시나리오로 표시하여 향후 실증 계획으로 기술»')

# 4.5 일정(제안안)
months = ['2026.12~2027.01', '2027.01~02', '2027.03~04', '2027.05~07', '2027.07~08', '2027.09~11']
for i, m in enumerate(months, start=1):
    c = T[4].rows[i].cells[0]
    rich(c.paragraphs[0], f'{c.text.strip()}\n«{m}»')
rich(P[90], '2단계 중 앱 ①·② 분석과 기준 사례 예비 시험은 계획서 작성 단계에서 선행 수행하였다. «[확인 필요] 표의 기간은 제안안이다. 학기·논문 심사 일정에 맞춰 확정»')

# 5.1, 5.4
append(P[93], ' 이 중 앱 ①·②의 분석 기록과 기준 사례 시험 스크립트는 계획서 작성 단계에서 이미 작성되었다.')
append(P[99], ' 발전기실 급·배기 계산 앱처럼 계산 로직을 단일 파일로 두고 화면과 API가 함께 쓰는 구조는 후속 연구의 API 기반 계산 연동을 위한 선행 구현으로 참고할 수 있다.')

# 참고문헌 추가 (가나다·알파벳 순서는 원본 형식을 따라 끝에 추가)
NEW_REFS = [
    '• Brooke, J. (1996). SUS: A "quick and dirty" usability scale. In P. W. Jordan, B. Thomas, B. A. Weerdmeester, & I. L. McClelland (Eds.), Usability Evaluation in Industry (pp. 189–194). Taylor & Francis. «[원문 확인]»',
    '• Cohen, J. (1960). A coefficient of agreement for nominal scales. Educational and Psychological Measurement, 20(1), 37–46. «[원문 확인]»',
    '• Fischer, G., Giaccardi, E., Ye, Y., Sutcliffe, A. G., & Mehandjiev, N. (2004). Meta-design: A manifesto for end-user development. Communications of the ACM, 47(9), 33–37. «[원문 확인]»',
    '• Hevner, A. R., March, S. T., Park, J., & Ram, S. (2004). Design science in information systems research. MIS Quarterly, 28(1), 75–105. «[원문 확인]»',
    '• Karpathy, A. (2025, February 2). "There\'s a new kind of coding I call \'vibe coding\'..." [Post]. X. «[원문 확인: URL·게시일]»',
    '• Ko, A. J., Abraham, R., Beckwith, L., Blackwell, A., Burnett, M., Erwig, M., Scaffidi, C., Lawrance, J., Lieberman, H., Myers, B., Rosson, M. B., Rothermel, G., Shaw, M., & Wiedenbeck, S. (2011). The state of the art in end-user software engineering. ACM Computing Surveys, 43(3), Article 21. «[원문 확인]»',
    '• Nonaka, I., & Takeuchi, H. (1995). The Knowledge-Creating Company. Oxford University Press. «[원문 확인]»',
]
e = P[107]._p
for r in NEW_REFS:
    e = par_after(e, r)
rich(P[108], '«[검증 필요] 추가한 참고문헌 7편의 서지 사항 원문 대조, 학교 인용 양식에 맞춰 정렬. 기존 초안의 Gruber(1993), Noy & McGuinness(2001)는 본문 인용이 없어 제외함»')

# 부록 A: 상태표로 교체
for i in range(110, 118):
    delete(P[i])
e = par_after(P[109]._p, '붉은 글씨는 저자가 작성하거나 확인할 항목이다.')
styled_table(e, ['항목', '상태', '위치'], [
    ['앱 ①·②의 기능·입출력·산식·검증 자료', '완료 (2026-11-18, 소스·배포본 대조)', '4.1, 4.2'],
    ['앱 화면 캡처 (앱별 4장)', '완료 (기본 입력값). «교체 여부 확인»', '그림 4-1, 4-2'],
    ['연구 프레임워크 그림, 플랫폼 와이어프레임', '완료 (구상안)', '그림 3-1, 4-3'],
    ['실무 사례(프로젝트 A·B) 조건·도면', '«작성 필요»', '1.1, 3.5, 4.4, 그림 1-1, 4-5'],
    ['개선 의견 이메일 수신 예시', '«작성 필요 (구현 후)»', '그림 4-4'],
    ['인터뷰 대상·질문지·일정, IRB', '«작성 필요»', '3.2, 3.7'],
    ['실험 참가자·과제·평가자 수', '«확인 필요 (제안안 기재)»', '3.6'],
    ['시나리오 전문가 기준 답안(판단 항목)', '«작성 필요»', '3.5'],
    ['연구 일정', '«확인 필요 (제안안 기재)»', '4.5'],
    ['선행연구 원문·서지 정보', '«검증 필요»', '2.2, 2.5, 2.7, 참고문헌'],
    ['통합 플랫폼 구현 범위', '«확인 필요»', '1.4, 3.4'],
    ['제목·제출 양식·생성형 AI 표기 양식', '«확인 필요»', '표지, 부록 B'],
])

# 부록 B, C
h = par_after(P[109]._p, '부록 B. 생성형 AI 활용 표기', tpl=H1_TPL)
# 부록 A 표 뒤로 옮기기 위해 문서 끝(마지막 요소 앞)에 배치
body = d.element.body
sect = body[-1] if body[-1].tag == qn('w:sectPr') else None
def to_end(el):
    el.getparent().remove(el)
    if sect is not None: sect.addprevious(el)
    else: body.append(el)
to_end(h)
b1 = par_after(h, '본 연구의 선행 개발 앱과 시제품은 저자가 Claude Code(Anthropic)를 활용하여 개발하였다. 본 계획서의 앱 분석(소스코드·커밋 이력 확인, 산식 대조, 기준 사례 시험 스크립트 작성), 화면 캡처, 그림 작성과 초안 문장 작성에도 생성형 AI(Claude)를 보조 도구로 활용하였으며, 내용은 저자의 검토를 거쳐 확정한다. 생성형 AI가 작성한 코드와 문장의 정확성은 AI가 보증하지 않으며 3.3절의 검증 절차를 따른다. «[확인 필요] 학교의 생성형 AI 활용 표기 양식»')
h2 = par_after(b1, '부록 C. 그림 출처와 재현 정보', tpl=H1_TPL)
c1 = par_after(h2, '그림 4-1: https://power-redundancy-calculator.vercel.app/ (저장소 커밋 babefae), 2026-11-18 캡처, 기본 입력값(가정 조건). 그림 4-2: https://ai-factory-study.vercel.app/ (커밋 12122d7), 2026-11-18 캡처. 그림 3-1, 4-3: 저자 작성(생성형 AI 보조), 그림 4-3은 구현 전 구상안. 그림 1-1, 4-4, 4-5: 저자 제공 예정.')
par_after(c1, '앱 ① 기준 사례 예비 시험은 실제 앱을 브라우저로 실행해 입력값을 바꾸고 화면 값을 독립 산식과 비교하는 스크립트로 수행하였으며(33개 조건, 363개 항목), 스크립트와 앱별 분석 기록은 연구 자료로 보관한다.')

# ---------------- 마지막: 남은 꺾쇠 표기 붉게 ----------------
MARK = re.compile(r'\[(?:작성 필요|검증 필요|보완 권장|확인 필요|원문[^\]]*|사진 삽입)[^\]]*\]')
def recolor(p):
    if id(p._p) in DONE or not p.runs:
        return
    txt = p.text
    if not MARK.search(txt):
        return
    if txt.strip().startswith('['):
        rich(p, '«' + txt + '»')
    else:
        rich(p, MARK.sub(lambda m: '«' + m.group(0) + '»', txt))
for p in d.paragraphs:
    recolor(p)
for tb in d.tables:
    for row in tb.rows:
        for c in row.cells:
            for p in c.paragraphs:
                recolor(p)

# 바닥글: 제출용 표기
for sec in d.sections:
    for fp in sec.footer.paragraphs:
        if '수정 초안' in fp.text:
            rich(fp, fp.text.replace('연구계획서 수정 초안', '연구계획서'))
# 표 머리행 반복
for tb in d.tables:
    trPr = tb.rows[0]._tr.get_or_add_trPr()
    if trPr.find(qn('w:tblHeader')) is None and tb.style is not None and tb.style.name.startswith('Light'):
        h = OxmlElement('w:tblHeader'); h.set(qn('w:val'), 'true'); trPr.append(h)
d.save(OUT)
print('saved', OUT, 'S1', s1a, s1b, 'S2', s2a, s2b, 'S3', s3a, s3b)
