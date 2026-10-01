# 🚜 장비주행성 검토 프로그램

연약지반 복토 위 시공장비의 주행 가능 여부를 검토하고, **실무 구조계산서(엑셀·PDF·Word)** 를 바로 출력하는 Streamlit 웹 프로그램입니다.
계산식은 실무 계산서 `2.1 장비주행성`(Meyerhof and Hanna) 및 `3.1 장비주행성`(원지반 작용응력)과 동일하며, 원본 엑셀 값과 소수점까지 일치하도록 검증되었습니다.

## 검토 절차
1. **접지압** `P = W/(2·b·a)` (덤프트럭 `P = 0.4W/(b·a)`)
2. **원지반 작용응력(σ)** — 하중분산 응력법 `σ = P·b·a(1+ε)/[(b+2H·tanθ)(a+2H·tanθ)] + γ1·H`
3. **원지반 허용지지력(qa)** — Meyerhof and Hanna(1978) 층상지반 지지력식 (토목섬유 인장력 T 고려)
4. **판정** — `qa > σ` 이면 O.K. 장비별 검토 복토두께는 qa > σ 를 만족하는 최소 두께를 자동 판정(직접 지정도 가능)

## 출력물
| 파일 | 구성 |
|---|---|
| **엑셀(.xlsx)** | 시트1 `1.원지반 작용응력`(3.1 형식) · 시트2 `2.허용지지력 검토`(2.1 형식) · 시트3 그래프. **모든 계산 셀이 엑셀 수식으로 연결** — 노란색 입력칸(지반정수·복토두께)을 고치면 전체가 자동 재계산 |
| **PDF** | A4 세로, 이론삽도·산정식·결과표·그래프·결론. 장비가 많으면 표가 다음 쪽으로 이어지고 머리글 반복 |
| **Word(.docx)** | PDF와 동일 구성, 편집용 |

## 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

## GitHub → Streamlit Community Cloud 배포
1. 이 폴더의 파일을 **구조 그대로** GitHub 저장소에 올립니다.
   ```
   app.py  core.py  excel_report.py  reports.py
   requirements.txt  packages.txt  README.md  장비목록_예시.csv
   assets/theory_load_dispersion.png
   assets/theory_meyerhof_hanna.png
   .streamlit/config.toml
   ```
2. https://share.streamlit.io → **Create app** → 저장소/브랜치 선택, Main file path = `app.py` → Deploy
3. `packages.txt`(fonts-nanum)는 서버에 한글 폰트를 설치해 PDF/그래프의 한글이 깨지지 않게 합니다. **반드시 함께 올려주세요.**
   (Python 버전은 배포 화면 Advanced settings 에서 3.11 또는 3.12 권장)

## 장비목록 불러오기/저장
- 화면의 **⬇ 현재 장비목록 저장(CSV)** 로 저장 → 다음에 **장비목록 불러오기** 로 그대로 사용
- 열 구성: `검토, 장비명, 규격, 중량W(kN), 폭b(m), 길이a(m), 덤프식, 지정복토두께(m)` (예시: `장비목록_예시.csv`)

## 참고
- Ks(펀칭전단계수)는 원본과 같이 `LOOKUP(Φ1, 20~50°)` (구간 하한값) 적용, Φ1 < 20° 는 0(관입저항 무시).
- 원본 2.1의 Fcd 는 `1+0.4(0/b)`(=1) 로 고정되어 있었으나, 본 프로그램은 근입심도 Df 입력값을 반영합니다(Df=0이면 동일).
- 원본 3.1 시트의 허용지지력(평판재하시험값 직접입력) 대신 Meyerhof and Hanna 식으로 산정한 qa 를 연결했습니다.
