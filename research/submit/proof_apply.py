import docx, glob, copy
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
SRC=glob.glob('/root/.claude/uploads/b5e6f573-7b81-5e78-b347-81f5624c7faa/a4fb27a4-*.docx')[0]
d=docx.Document(SRC)
def find(prefix):
    for p in d.paragraphs:
        if p.text.strip().startswith(prefix): return p
    raise KeyError(prefix)
def body(p,text):
    for r in p.runs[1:]: r._r.getparent().remove(r._r)
    if not p.runs: p.add_run('')
    r=p.runs[0]; r.text=text; r.font.size=None; r.font.color.rgb=RGBColor(0,0,0); r.font.bold=False; r.font.italic=False
    pf=p.paragraph_format; p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; pf.first_line_indent=Cm(0.5); pf.left_indent=None; pf.line_spacing=1.5; pf.space_after=Pt(2)
def drop(p): p._p.getparent().remove(p._p)
def drop_memo(head_prefix):
    # 해당 절 제목 다음의 ▶메모 + '· ' 줄 삭제
    hs=[p for p in d.paragraphs if p.text.strip().startswith(head_prefix)][0]
    el=hs._p.getnext()
    while el is not None:
        t=''.join(x.text or '' for x in el.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')).strip()
        if t.startswith('▶ 작성 메모') or t.startswith('· '):
            nxt=el.getnext(); el.getparent().remove(el); el=nxt; continue
        if t=='' and el.tag.endswith('}p'):
            el=el.getnext(); continue
        if t.startswith('·')==False and not t.startswith('▶ 작성'):
            # 본문 단락을 지나쳐 다음 메모도 찾아 삭제(3.2처럼 본문 뒤에 메모가 남은 경우)
            nxt=el.getnext()
            if nxt is not None:
                tt=''.join(x.text or '' for x in nxt.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')).strip()
                if tt.startswith('▶ 작성 메모'):
                    el=nxt; continue
            break
# ---- 1.1
p1=find('건축 설계의 단계 중'); body(p1,'건축 설계의 단계 중 초기 기획·계획설계 단계는 사업의 규모와 설계의 밑그림을 그리는 중요한 과정이다. 초기 단계의 결정이 건물 성능·생애주기 비용에 미치는 영향이 가장 크다(Bragança 외, 2014).')
drop(find('초기 단계 의 결정'))
body(find('일반시설(상주 공간 중심)'),'일반 시설(상주 공간 중심)은 발주처의 요구 사항, 건축물 용도에 따른 설계 각론, 유사 사례, 건축가의 경험으로 정합성 있는 계획안 도출이 가능하다. 하지만 특수 시설(산업시설, 데이터센터 등)의 경우는 설비·전기 등의 공간 비중이 커서 검토의 한계가 있다. 전문 분야(구조·기계·전기 분야 등) 프로젝트 팀이 구성되기 어려운 초기 기획 단계에서 건축가가 정확한 기획설계안을 도출하기 위해서는 전문 분야의 지식 확장이 필요하며, 이를 앱(App)으로 생성하고자 한다. 본문에서 언급하는 지식 확장의 정의는 ① 전문 분야 검토 항목의 식별, ② 근거의 이해, ③ 조건 변경 시 정량적이고 신속한 대응이다.')
drop(find('건축 외 분야 전문지식이 필요한'))
body(find('데이터센터는 전기 수전 용량'),'데이터센터에서는 전기 수전 용량에 따른 서버실과 항온항습실, 전원 공급 시스템 안정성을 위한 전기실, 정전 시 전원 공급을 유지할 수 있도록 해 주는 비상 발전기실을 계획하고, 이를 이중화와 여유율 공간까지 고려하는 것이 건축계획의 중대한 요소이다. 이 분야는 전기·설비 전문 설계 분야와 협업을 해야 하지만, 프로젝트 초기에 계약도 안 된 상황에서 기획설계를 위해 정식으로 검토하는 것은 현실적으로 어려운 실정이다. 이에 건축 외 분야의 전문 지식이 필요한 특수 시설 중 데이터센터를 사례로 구체적인 연구를 진행해 보고자 한다.')
# ---- 1.3
drop_memo('1.3 연구의 방법')
p=find('연구의 방법은 사례조사'); body(p,'연구의 방법은 사례조사와 문헌조사 방법을 기반으로 한다. 연구 대상은 저자가 개발해 웹으로 배포하여 실무에 적용해 보고 있는 시범 적용(pilot) 앱 3가지로 연구한다. 구체적인 앱의 구성 원리와 앱의 효과를 검증하고, 이를 통합한 프레임워크 구현까지를 목표로 한다.')
drop(find('연구대상은 저자가')); drop(find('구체적인 앱의 구성 원리'))
# ---- 2.1
p=find('앱의 개발 방법은'); body(p,'앱은 실무 진행 시 전문 협력업체와의 협의·문의로 흩어져 있던 지식과 전문가의 검토 원리를 로직으로 하여 개발하였고, 지식 유형(de Jong & Ferguson-Hessler, 1996; 이병수, 2016, p. 15) 분류에 의해 선언적(입력값·기준), 절차적(산식·계산 순서), 상황적(설계 조건과 연결 관계), 전략적(검토 순서와 종합 판단) 지식에 접근하고자 한다.')
drop(find('상황적(설계 조건과 연결관계)'))
# ---- 2.2
p=find('▶데이터센터의 기본 공간 구성'); body(p,'데이터센터의 기본 공간 구성은 서버가 위치하는 데이터홀(서버룸), 이를 지원하기 위한 기계·전기실, 그리고 사람이 상주하는 운영 사무실·보안·하역 영역 등으로 구성되며, 기준층(서버실·항온항습실) 연면적과 기계·전기실 연면적의 비율은 일반적으로 1:1이다(김재성·이원재, 2018). 사람이 중점적으로 상주하는 공간은 운영·휴게 공간이다.')
p=find('서버실 층고 4.5m'); body(p,'김재성·이원재(2018)는 서버실 층고 4.5m 이상, 수변전실·발전기실 적정 층고 7m, 발전기실 급·배기구 면적 확보를 언급하였지만, 이는 통일된 수치가 아니며 프로젝트마다 시스템 및 요건에 따라 차이가 있을 수 있다.')
for q in d.paragraphs:
    if q.text.strip()=='ㅁ':
        for r in q.runs:
            if r.text.strip()=='ㅁ': r.text=''
# ---- 2.3
p=find('선행연구에서 기획 설계'); body(p,'선행연구에서는 기획설계 단계의 자동화, AI 기술 도입 등을 통해 BIM 템플릿 생성(이병수 외, 2016), 친환경 검토 플러그인(이용주, 2020), Robot Structural Analysis 및 Fusion 360을 이용한 AI 솔루션(김혜정 외, 2019) 등 다양한 연구가 선행되었다. 하지만 AI 바이브 코딩을 통해 앱을 생성하고 이를 검증하며, 추후 이를 통합하여 종합 솔루션까지 도출해 줄 수 있는 프레임워크에 대한 연구 분석은 확인되지 않았다.')
# ---- 3.1
drop_memo('3.1 연구 모형과 분석틀')
p=find('이론적 근거, 분석틀, 사례 분석'); body(p,'이론적 근거, 분석틀, 사례 분석, 결과 도출 단계로 연구의 흐름을 정하였다(그림 2). 분석틀의 경우 지식 명시화, 도구 구현, 검증, 공유 및 개선의 4가지 관점, 12개 요소로 구성하였다(표 2). 검토 분야, 입력 변수, 근거 및 산식을 명시화하고 이를 바이브 코딩으로 구현하는 형태를 표현하며, 결과값을 기반으로 기준 사례와 비교하여 정확도를 점검하고, 이를 보고 배포하여 지속적으로 개선하고 의견을 수렴하여 업데이트할 수 있도록 한다. 사례는 데이터센터의 중요한 검토 요소 중 3가지 요소를 기반으로 만든 앱을 통해 분석하고, 결과값을 조합하여 종합적인 결론에 이를 수 있는 프레임워크를 연구한다. 이를 증빙하기 위한 자료는 배포되어 있는 앱의 화면을 첨부하여 사용한다.')
# ---- 3.2
p=find('분석 대상은 실무에서 직접협의'); body(p,'분석 대상은 실무에서 직접 협의했던 지식으로 만든 3가지 앱이며, 시범 적용 중인 단계이다. 소스코드와 업데이트 이력 확인이 가능하다. 개발 과정은 협력업체 협의, 장비 사양, 계산 근거 등의 지식을 취합하고 Claude Code로 바이브 코딩을 진행하여 파이프라인을 구축하고, HTML·CSS·JavaScript 파일 분리 구조로 구성하여 GitHub로 커밋하고 웹 배포(Vercel)를 하여 지속적으로 사용해 보고 의견을 반영하여 업데이트하는 구조이다(표 3).')
drop_memo('3.2 분석 대상 사례')
d.save('건축계획연구_과제_작성중.docx'); print('saved')
