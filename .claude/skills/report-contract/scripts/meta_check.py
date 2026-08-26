# -*- coding: utf-8 -*-
"""메타 표현 점검기 — 보고서가 자기 자신을 설명하는 문장을 찾는다.
사용: python3 meta_check.py <파일...>     (.pptx .docx .md .txt .js 지원)
계약 C11 · 자기검증 G7. 0건이어야 제출한다."""
import sys, re, zipfile, os

RULES = [
 ('M1 제작 용어', r'프레임장|적용장|소결장|상향\s*링크|3부\s*구조|판정\s*기준으로\s*선언|같은\s*절\s*안에서\s*닫|배치\s*단위'),
 ('M2 독자 안내', r'읽는\s*경로|독자별|만으로\s*판단할\s*수\s*있|이\s*장에서는|앞서\s*살펴본|다음\s*장에서|본\s*장은'),
 ('M3 보고서 자칭', r'본\s*보고서는.{0,12}구성|본문\s*\d+장|수록\s*장|장에\s*귀속|장\s*단위로\s*확인|이\s*절은'),
 ('M4 장번호 인용', r'\(\s*\d{1,3}\s*장\s*(참조|인용)?\s*\)|\d{1,3}장에서\s*(확인|제시|서술)|근거\s*[①-⑨].{0,6}\d{1,3}\s*장'),
 ('M5 작성 자기지시', r'표준\s*서식으로\s*(기술|작성)|템플릿에\s*따라|원칙에\s*따라\s*작성|같은\s*서식으로\s*제시'),
 ('M6 설계 메타', r'\[미확정\]\s*배지|대체\s*원고|상태\s*:\s*(신규|승계)|계약\s*C\d+'),
 ('M7 자기평가', r'체계적으로\s*구성|논리적으로\s*구성|일관되게\s*구성|짜임새\s*있게'),
 ('M8 결정 요청', r'결정\s*요청|의사결정\s*안건|결정\s*:\s*_|기관\s*결정\s*:'),
 ('S1 학술 용어', r'소결|접는\s*규칙|기능\s*접기|접어\s|접음'),
 ('S2 구어·비유', r'갈래|밀림|얹어|얹은|닿지\s*않|에\s*닿음|(?<!엇)갈림|(?<!엇)갈려|굳어지|굳어짐|뼈대'),
 ('S3 수사 표기', r'석\s*달|한\s*해\s*(뒤|만에)|세\s*해|두\s*해|잇달아|마흔|서른|쉰두|여든|스무\s|열두\s|열세\s|열다섯|여섯\s*(개|영역|가지)|여덟\s*(개|기능|가지)|첫\s해'),
]

def text_of(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.pptx', '.docx', '.xlsx'):
        z = zipfile.ZipFile(path)
        parts = [n for n in z.namelist()
                 if re.match(r'(ppt/slides/slide|word/document|xl/sharedStrings)', n) and n.endswith('.xml')]
        out = []
        for n in sorted(parts):
            s = z.read(n).decode('utf-8', 'ignore')
            s = re.sub(r'</a:t>\s*</a:r>\s*<a:r>\s*<a:rPr[^>]*/?>(?:</a:rPr>)?\s*<a:t>', '', s)
            out.append((n, ' '.join(re.findall(r'<(?:a:t|w:t|t)[^>]*>([^<]*)<', s))))
        return out
    return [(path, open(path, encoding='utf-8', errors='ignore').read())]

HEAD_RULES = [
 ('H1 머리글 어미', r'(것|는가|으면|이유|곳|부분|지점|모습|쓰임)$'),
]
hits = 0
for f in sys.argv[1:]:
    for name, txt in text_of(f):
        for label, pat in RULES:
            for m in re.finditer(pat, txt):
                s = max(0, m.start()-30); e = min(len(txt), m.end()+30)
                print(f'HIT [{label}] {os.path.basename(f)} :: {name.split("/")[-1]}')
                print(f'      …{txt[s:e].strip()}…')
                hits += 1
print(f'\n{hits} hit / {len(sys.argv)-1} file — 0이어야 제출')
sys.exit(1 if hits else 0)
