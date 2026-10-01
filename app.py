# -*- coding: utf-8 -*-
"""
장비주행성 검토 프로그램 (Streamlit)
------------------------------------------------------------
· 원지반 작용응력(σ)  : 하중분산 응력법          (3.1 장비주행성 계산서 수식)
· 원지반 허용지지력(qa): Meyerhof and Hanna(1978) (2.1 장비주행성 계산서 수식)
· 판정                : qa > σ → O.K
· 출력                : ① 수식 연결 엑셀 계산서(.xlsx)  ② PDF 계산서  ③ Word 계산서

실행 :  streamlit run app.py
"""

import io
import json

import pandas as pd
import streamlit as st

import core
import excel_report
import reports

st.set_page_config(page_title="장비주행성 검토", page_icon="🚜", layout="wide")

EQ_COLS = ["검토", "장비명", "규격", "중량W(kN)", "폭b(m)", "길이a(m)", "덤프식", "지정복토두께(m)"]
H_OPTIONS = [round(i * 0.1, 1) for i in range(0, 31)]       # 0.0 ~ 3.0 m
H_DEFAULT = [0.1, 0.2, 0.3, 0.4, 0.5, 0.6]


def _coerce(df):
    for c in ("중량W(kN)", "폭b(m)", "길이a(m)", "지정복토두께(m)"):
        df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    for c in ("장비명", "규격"):
        df[c] = df[c].astype(object).where(df[c].notna(), None)
    return df


def default_equipment_df():
    return _coerce(pd.DataFrame(
        [[True, n, s, W, b, a, d, float("nan")] for n, s, W, b, a, d in core.DEFAULT_EQUIPMENT], columns=EQ_COLS))


def normalize_uploaded(df):
    """업로드한 장비목록(CSV/XLSX)을 표준 열 구성으로 정리"""
    df = df.copy()
    alias = {"장비": "장비명", "중량": "중량W(kN)", "W": "중량W(kN)", "폭": "폭b(m)", "b": "폭b(m)",
             "길이": "길이a(m)", "a": "길이a(m)", "L": "길이a(m)", "덤프": "덤프식", "복토두께": "지정복토두께(m)"}
    df = df.rename(columns={c: alias.get(str(c).strip(), str(c).strip()) for c in df.columns})
    for c in EQ_COLS:
        if c not in df.columns:
            df[c] = True if c == "검토" else (False if c == "덤프식" else None)
    df = _coerce(df[EQ_COLS].copy())
    for c in ("검토", "덤프식"):
        df[c] = df[c].map(lambda v: v if isinstance(v, bool) else
                          str(v).strip().lower() in ("true", "1", "1.0", "o", "y", "yes", "예")).astype(bool)
    return df.reset_index(drop=True)


# ------------------------------------------------------------------
# 사이드바 : 계산서 정보 · 지반정수 · 복토두께
# ------------------------------------------------------------------
with st.sidebar:
    st.header("📋 계산서 정보")
    project = st.text_input("PROJECT (프로젝트명)", value="", placeholder="계산서에 표기될 프로젝트명")

    st.header("🌍 지반 및 설계정수")
    d = core.DEFAULT_PARAMS
    st.caption("원지반 작용응력(σ) — 하중분산 응력법")
    gamma1 = st.number_input("복토층 단위중량 γ1 (kN/m³)", 0.0, 30.0, d["gamma1"], 0.5)
    theta = st.number_input("하중분산각 θ (°)  ※ 토목섬유-수평면 각 겸용", 0.0, 60.0, d["theta"], 1.0)
    eps = st.number_input("충격계수 ε (i)", 0.0, 1.0, d["eps"], 0.05)
    st.caption("허용지지력(qa) — Meyerhof and Hanna(1978)")
    c2 = st.number_input("원지반 점착력 c2 (kPa)", 0.0, 500.0, d["c2"], 0.5)
    phi1 = st.number_input("복토층 내부마찰각 Φ1 (°)", 0.0, 50.0, d["phi1"], 1.0,
                           help="펀칭전단계수 Ks는 원본과 같이 LOOKUP(20~50°) 적용, 20° 미만은 0")
    phi2 = st.number_input("원지반 내부마찰각 Φ2 (°)", 0.0, 50.0, d["phi2"], 1.0)
    Fs = st.number_input("안전율 Fs", 1.0, 5.0, d["Fs"], 0.1)
    T = st.number_input("토목섬유 허용인장력 T (kN/m)", 0.0, 2000.0, d["T"], 1.0,
                        help="0이면 토목섬유 보강효과 미반영. 예) 30 tonf/m ≈ 294.2 kN/m")
    Df = st.number_input("근입심도 Df (m)", 0.0, 5.0, d["Df"], 0.1)
    gamma2 = st.number_input("원지반 단위중량 γ2 (kN/m³) — 표기용", 0.0, 30.0, d["gamma2"], 0.5)

    st.header("📏 복토두께")
    thicknesses = st.multiselect("작용응력 표에 표시할 복토두께 (m) — 다중 선택", H_OPTIONS, default=H_DEFAULT,
                                 format_func=lambda v: f"{v:.1f} m")
    thicknesses = sorted(set(thicknesses)) or [0.1]
    include_zero = st.checkbox("최소 복토두께 판정 시 0.0 m(복토 없음)부터 검토", value=True,
                               help="장비별 검토 복토두께 = qa > σ 를 만족하는 최소 두께 (원본 3.1 검토결과표 방식)")

params = dict(gamma1=gamma1, gamma2=gamma2, theta=theta, eps=eps, c2=c2, phi1=phi1, phi2=phi2, Df=Df, T=T, Fs=Fs)

# ------------------------------------------------------------------
# 본문
# ------------------------------------------------------------------
st.title("🚜 장비주행성 검토 프로그램")
st.caption("하중분산 응력법(원지반 작용응력 σ) + Meyerhof and Hanna(1978) 층상지반 지지력식(허용지지력 qa) · "
           "실무 계산서(2.1 / 3.1 장비주행성) 수식과 동일")

with st.expander("📖 적용 이론 및 산정식", expanded=False):
    c_l, c_r = st.columns(2)
    with c_l:
        st.markdown("**① 원지반상 작용응력(σ) — 하중분산 응력법**")
        st.image(core.IMG_DISPERSION, use_container_width=True)
        st.latex(r"P=\frac{W}{2\,b\,a}\qquad(\text{Dump}:\ P=\frac{0.4\,W}{b\,a})")
        st.latex(r"\sigma=\frac{P\,b\,a\,(1+\varepsilon)}{(b+2H\tan\theta)(a+2H\tan\theta)}+\gamma_1H")
    with c_r:
        st.markdown("**② 원지반 허용지지력(qa) — Meyerhof and Hanna(1978)**")
        st.image(core.IMG_MEYERHOF, use_container_width=True)
        st.latex(r"q_a=\frac{1}{F_s}\Big[(1+0.2\tfrac{b}{L})c_2N_{c(2)}F_{cs(2)}F_{cd(2)}"
                 r"+\gamma_1H^2(1+\tfrac{b}{L})\frac{K_s\tan\phi_1}{b}+\frac{2T\sin\theta}{b+H}\Big]")
        st.latex(r"N_{q(2)}=e^{\pi\tan\phi_2}\tan^2(45+\tfrac{\phi_2}{2}),\quad "
                 r"N_{c(2)}=\frac{N_{q(2)}-1}{\tan\phi_2},\quad F_{cs(2)}=1+\tfrac{b}{L}\tfrac{N_{q(2)}}{N_{c(2)}}")
    st.info(f"현재 지반정수 → Ks = {core.ks_lookup(phi1):.2f},  Nc(2) = {core.nc2(phi2):.3f},  "
            f"Nq(2) = {core.nq2(phi2):.3f},  Nr(2) = {core.nr2(phi2):.3f}")

# ---------------- 장비 입력 ----------------
st.subheader("🚛 검토 장비 및 제원")
st.caption("· 표 아래 ➕ 로 장비 추가, 행 선택 후 🗑 로 삭제 · 접지압(P)은 자동 계산 "
           "(덤프식 체크 시 P=0.4W/(b·a)) · '지정복토두께'를 비워두면 최소 복토두께를 자동 판정")

if "eq_df" not in st.session_state:
    st.session_state.eq_df = default_equipment_df()
    st.session_state.eq_ver = 0

c1, c2_, c3 = st.columns([2, 1, 1])
with c1:
    up = st.file_uploader("장비목록 불러오기 (CSV/XLSX, 열: 장비명·규격·중량W(kN)·폭b(m)·길이a(m)·덤프식·지정복토두께(m))",
                          type=["csv", "xlsx"], key="eq_upload")
    if up is not None and st.session_state.get("eq_upload_name") != up.name + str(up.size):
        try:
            raw = pd.read_csv(up, encoding="utf-8-sig") if up.name.lower().endswith(".csv") else pd.read_excel(up)
            st.session_state.eq_df = normalize_uploaded(raw)
            st.session_state.eq_ver += 1
            st.session_state.eq_upload_name = up.name + str(up.size)
            st.rerun()
        except Exception as ex:  # noqa: BLE001
            st.error(f"장비목록을 읽지 못했습니다: {ex}")
with c3:
    if st.button("↺ 기본 장비목록으로 초기화", use_container_width=True):
        st.session_state.eq_df = default_equipment_df()
        st.session_state.eq_ver += 1
        st.session_state.pop("eq_upload_name", None)
        st.rerun()

edited = st.data_editor(
    st.session_state.eq_df,
    key=f"eq_editor_{st.session_state.eq_ver}",
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "검토": st.column_config.CheckboxColumn("검토", default=True, width="small"),
        "장비명": st.column_config.TextColumn("장비명", required=True),
        "규격": st.column_config.TextColumn("규격", default="-"),
        "중량W(kN)": st.column_config.NumberColumn("중량 W (kN)", min_value=0.0, step=0.1, format="%.2f"),
        "폭b(m)": st.column_config.NumberColumn("접지폭 b (m)", min_value=0.0, step=0.01, format="%.3f"),
        "길이a(m)": st.column_config.NumberColumn("접지길이 a (m)", min_value=0.0, step=0.01, format="%.3f"),
        "덤프식": st.column_config.CheckboxColumn("덤프식", default=False, width="small"),
        "지정복토두께(m)": st.column_config.NumberColumn("지정 복토두께 (m)", min_value=0.0, step=0.1, format="%.2f",
                                                    help="비워두면 자동판정"),
    },
)
with c2_:
    st.download_button("⬇ 현재 장비목록 저장(CSV)", edited.to_csv(index=False).encode("utf-8-sig"),
                       "장비목록.csv", "text/csv", use_container_width=True)

# 입력 검증
equipment, skipped = [], []
for idx, r in edited.iterrows():
    if not bool(r.get("검토", True)):
        continue
    raw_name = r.get("장비명")
    name = "" if raw_name is None or pd.isna(raw_name) else str(raw_name).strip()
    nums = [r.get("중량W(kN)"), r.get("폭b(m)"), r.get("길이a(m)")]
    if not name and all(v is None or pd.isna(v) for v in nums):
        continue                                   # 완전히 빈 행(➕로 추가만 한 행)은 조용히 무시
    try:
        W_, b_, a_ = (float(v) for v in nums)
        if not name or any(pd.isna([W_, b_, a_])) or W_ <= 0 or b_ <= 0 or a_ <= 0:
            raise ValueError
    except (TypeError, ValueError):
        skipped.append(name or f"{idx + 1}번째 행")
        continue
    hfix = r.get("지정복토두께(m)")
    spec = r.get("규격")
    equipment.append(dict(name=name, spec=("-" if spec is None or pd.isna(spec) or str(spec).strip() == "" else str(spec)),
                          W=W_, b=b_, a=a_, dump=bool(r.get("덤프식", False)),
                          H_fix=None if hfix is None or pd.isna(hfix) else float(hfix)))
if skipped:
    st.warning("입력값(장비명·중량·폭·길이)이 비어있거나 0 이하여서 검토에서 제외된 장비: " + ", ".join(skipped))
if not equipment:
    st.error("검토할 장비가 없습니다. 장비를 1개 이상 입력/선택하세요.")
    st.stop()

rows = core.evaluate(equipment, thicknesses, params, include_zero=include_zero)

# ---------------- 결과 ----------------
st.subheader("📊 검토 결과")
n_ng = sum(r["judge"] == "N.G" for r in rows)
m1, m2, m3 = st.columns(3)
m1.metric("검토 장비", f"{len(rows)} 종")
m2.metric("적합 (O.K)", f"{len(rows) - n_ng} 종")
m3.metric("부적합 (N.G)", f"{n_ng} 종")

res_df = pd.DataFrame([{
    "장비명": r["name"], "규격": r["spec"], "접지압 P (kPa)": round(r["P"], 2),
    "검토 복토두께 H (m)": r["H"], "두께결정": "자동" if r["auto"] else "지정",
    "작용응력 σ (kPa)": round(r["sigma"], 2), "허용지지력 qa (kPa)": round(r["qa"], 2),
    "qa/σ": round(r["FS"], 2), "판정": r["judge"]} for r in rows])


def _judge_color(v):
    if v == "O.K":
        return "background-color:#E6F4EA;color:#1B5E20;font-weight:bold"
    if v == "N.G":
        return "background-color:#FDE9E9;color:#B71C1C;font-weight:bold"
    return ""


sty = res_df.style.format({"검토 복토두께 H (m)": "{:.1f}", "접지압 P (kPa)": "{:.2f}", "작용응력 σ (kPa)": "{:.2f}",
                           "허용지지력 qa (kPa)": "{:.2f}", "qa/σ": "{:.2f}"})
sty = (sty.map if hasattr(sty, "map") else sty.applymap)(_judge_color, subset=["판정"])
st.dataframe(sty, use_container_width=True, hide_index=True)
if n_ng:
    st.error("⚠ 부적합 장비: " + ", ".join(f"{r['name']}({r['spec']})" for r in rows if r["judge"] == "N.G")
             + " — 복토두께 상향, 토목섬유 보강(T) 또는 장비 변경을 검토하세요.")
else:
    st.success("✅ 모든 장비가 주행 가능(O.K)한 것으로 검토되었습니다.")

tab1, tab2 = st.tabs(["복토두께별 작용응력 σ (kPa)", "그래프"])
with tab1:
    mat = pd.DataFrame([{"장비명": r["name"], "규격": r["spec"], "P (kPa)": r["P"],
                         **{f"{h:.1f} m": round(r["sig_tbl"][h], 2) for h in thicknesses}} for r in rows])
    st.dataframe(mat, use_container_width=True, hide_index=True)
with tab2:
    st.image(core.chart_png(rows), use_container_width=True)

# ---------------- 계산서 출력 ----------------
st.subheader("🖨️ 구조계산서 출력")
st.caption("· **엑셀** : 2.1/3.1 장비주행성 계산서와 같은 배치로, 모든 계산 셀이 엑셀 수식으로 서로 연결됨 "
           "(노란색 입력칸 수정 시 자동 재계산) · **PDF** : A4 바로 인쇄용 · **Word** : 편집용")

key = json.dumps({"p": project, "params": params, "h": thicknesses, "z": include_zero, "eq": equipment},
                 ensure_ascii=False, sort_keys=True, default=str)


@st.cache_data(show_spinner=False, max_entries=20)
def make_files(cache_key):
    k = json.loads(cache_key)
    rr = core.evaluate(k["eq"], k["h"], k["params"], include_zero=k["z"])
    return (excel_report.build_excel(k["p"], rr, k["h"], k["params"]),
            reports.build_pdf(k["p"], rr, k["h"], k["params"]),
            reports.build_docx(k["p"], rr, k["h"], k["params"]))


with st.spinner("계산서 생성 중..."):
    try:
        xlsx_b, pdf_b, docx_b = make_files(key)
    except Exception as ex:  # noqa: BLE001
        st.exception(ex)
        st.stop()

fname = (project.strip() or "장비주행성") + "_장비주행성검토"
for ch in '\\/:*?"<>|':
    fname = fname.replace(ch, "_")
d1, d2, d3 = st.columns(3)
d1.download_button("📊 엑셀 계산서 (수식 연결) .xlsx", xlsx_b, f"{fname}.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True, type="primary")
d2.download_button("📄 PDF 계산서 (A4 인쇄용)", pdf_b, f"{fname}.pdf", "application/pdf", use_container_width=True)
d3.download_button("📝 Word 계산서 .docx", docx_b, f"{fname}.docx",
                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                   use_container_width=True)

if core.find_kr_font()[0] is None:
    st.warning("서버에 한글 폰트가 없어 PDF에 대체 폰트를 사용했습니다. GitHub 저장소에 packages.txt(fonts-nanum)를 "
               "함께 올리면 해결됩니다.")
