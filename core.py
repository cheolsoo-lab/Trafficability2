# -*- coding: utf-8 -*-
"""
core.py — 장비주행성 검토 계산 엔진
------------------------------------------------------------
원본 실무 계산서와 동일한 수식으로 계산합니다.
  · 3.1 장비주행성(1-1.원지반 작용응력) : 접지압 P, 하중분산 응력 σ
  · 2.1 장비주행성(Meyerhof and Hanna)  : 허용지지력 qa, 형상계수, Ks

엑셀 출력물의 셀 수식과 1:1로 대응되도록 작성했으며(ROUND, LOOKUP 동작 포함)
엑셀에서 재계산한 값과 소수점까지 동일하게 나오는 것을 검증했습니다.
"""

import io
import math
import os

import matplotlib

matplotlib.use("Agg")
import matplotlib.font_manager as fm  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSET_DIR = os.path.join(BASE_DIR, "assets")
IMG_DISPERSION = os.path.join(ASSET_DIR, "theory_load_dispersion.png")   # 3.1 삽도
IMG_MEYERHOF = os.path.join(ASSET_DIR, "theory_meyerhof_hanna.png")      # 2.1 삽도


def read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


# ------------------------------------------------------------------
# 한글 폰트 탐색 (Windows / macOS / Streamlit Cloud(fonts-nanum) / Linux)
# ------------------------------------------------------------------
KR_FONT_CANDIDATES = [
    # (regular, bold)
    (r"C:\Windows\Fonts\malgun.ttf", r"C:\Windows\Fonts\malgunbd.ttf"),
    ("/usr/share/fonts/truetype/nanum/NanumGothic.ttf", "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf"),
    ("/Library/Fonts/NanumGothic.ttf", "/Library/Fonts/NanumGothicBold.ttf"),
    ("/System/Library/Fonts/Supplemental/AppleGothic.ttf", None),
    ("/Library/Fonts/AppleGothic.ttf", None),
    ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", None),
]


def find_kr_font():
    """(regular_path, bold_path) — 없으면 (None, None)"""
    for reg, bold in KR_FONT_CANDIDATES:
        if os.path.exists(reg):
            return reg, (bold if bold and os.path.exists(bold) else reg)
    return None, None


_KR_REG, _KR_BOLD = find_kr_font()
_mpl_family = None
if _KR_REG:
    try:
        fm.fontManager.addfont(_KR_REG)
        _mpl_family = fm.FontProperties(fname=_KR_REG).get_name()
    except Exception:
        _mpl_family = None
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = [f for f in [_mpl_family, "Malgun Gothic", "NanumGothic",
                                               "AppleGothic", "DejaVu Sans"] if f]
plt.rcParams["axes.unicode_minus"] = False


# ------------------------------------------------------------------
# 엑셀 함수와 동일 동작 헬퍼
# ------------------------------------------------------------------
def xl_round(x, n=2):
    """엑셀 ROUND (0.5 → 올림, 은행가 반올림 아님)"""
    m = 10 ** n
    return math.floor(abs(x) * m + 0.5) / m * (1 if x >= 0 else -1)


KS_PHI = [20, 25, 30, 35, 40, 45, 50]
KS_VAL = [1.89, 2.22, 3.06, 4.45, 6.95, 11.12, 19.15]


def ks_lookup(phi1):
    """펀칭전단계수 Ks — 엑셀 LOOKUP(φ1, 20~50, Ks) 와 동일(구간 하한값 채택).
    φ1 < 20° 는 조견표 범위 밖이므로 0(관입저항 무시, 안전측)으로 처리."""
    if phi1 < KS_PHI[0]:
        return 0.0
    val = KS_VAL[0]
    for p, k in zip(KS_PHI, KS_VAL):
        if phi1 >= p:
            val = k
    return val


def nq2(phi2):
    r = math.radians(phi2)
    return math.exp(math.pi * math.tan(r)) * math.tan(math.radians(45 + phi2 / 2)) ** 2


def nc2(phi2):
    if phi2 == 0:
        return 5.14
    return (nq2(phi2) - 1) / math.tan(math.radians(phi2))


def nr2(phi2):
    return (nq2(phi2) - 1) * math.tan(math.radians(1.4 * phi2))


# ------------------------------------------------------------------
# 3.1 장비주행성 수식
# ------------------------------------------------------------------
def contact_pressure(W, b, a, dump=False):
    """H열: =ROUND(E/(2*F*G),2)   (덤프: =ROUND(E*0.4/(F*G),2))"""
    if b <= 0 or a <= 0:
        return 0.0
    return xl_round(W * 0.4 / (b * a)) if dump else xl_round(W / (2 * b * a))


def sigma(P, b, a, H, gamma1, theta, eps):
    """=$H*$F*$G*(1+ε)/(($F+2*H*TAN(RADIANS(θ)))*($G+2*H*TAN(RADIANS(θ))))+H*γ1"""
    t = math.tan(math.radians(theta))
    return P * b * a * (1 + eps) / ((b + 2 * H * t) * (a + 2 * H * t)) + H * gamma1


# ------------------------------------------------------------------
# 2.1 장비주행성 수식 (Meyerhof and Hanna, 1978)
# ------------------------------------------------------------------
def shape_factors(b, L, phi2, Df):
    Fcs = 1 + (b / L) * (nq2(phi2) / nc2(phi2))                                  # Q열
    Fqs = xl_round(1 if phi2 == 0 else 1 + math.tan(math.radians(phi2)) * b / L)  # R열
    Frs = xl_round(1 - 0.4 * b / L)                                              # S열
    Fcd = 1 + 0.4 * (Df / b)                                                     # T열
    return Fcs, Fqs, Frs, Fcd


def qa_meyerhof(b, L, H, p):
    """M열: =(1/Fs)*((1+0.2*(b/L))*c2*Nc*Fcs*Fcd + γ1*H^2*(1+b/L)*(Ks*TAN(φ1)/b)
                      + (2*T*SIN(θ))/(b+H))"""
    Fcs, _, _, Fcd = shape_factors(b, L, p["phi2"], p["Df"])
    term1 = (1 + 0.2 * (b / L)) * p["c2"] * nc2(p["phi2"]) * Fcs * Fcd
    term2 = p["gamma1"] * H ** 2 * (1 + b / L) * (ks_lookup(p["phi1"]) * math.tan(math.radians(p["phi1"])) / b)
    term3 = (2 * p["T"] * math.sin(math.radians(p["theta"]))) / (b + H)
    return (1 / p["Fs"]) * (term1 + term2 + term3)


DEFAULT_PARAMS = dict(
    gamma1=18.0,   # 복토층 단위중량 γ1 (kN/m³)
    gamma2=18.0,   # 원지반 단위중량 γ2 (kN/m³) — 표기용(2.1 원본과 동일)
    theta=30.0,    # 하중분산각 / 토목섬유-수평면 각 θ (°)
    eps=0.0,       # 충격계수 ε
    c2=7.0,        # 원지반 점착력 (kPa)
    phi1=25.0,     # 복토층 내부마찰각 (°)
    phi2=30.0,     # 원지반 내부마찰각 (°)
    Df=0.0,        # 근입심도 (m)
    T=0.0,         # 토목섬유 허용인장력 (kN/m)
    Fs=1.5,        # 안전율
)

# 3.1 장비주행성.xlsm 장비 제원 (원본과 동일)
DEFAULT_EQUIPMENT = [
    # 장비명, 규격, 중량W(kN), 폭b(m), 길이a(m), 덤프식
    ("Dozer", "6t", 64.1, 0.71, 2.03, False),
    ("Dozer", "13t", 131.0, 0.77, 2.61, False),
    ("Dozer", "19t", 190.0, 0.76, 2.80, False),
    ("Dozer", "24t", 240.0, 0.55, 2.84, False),
    ("Dozer", "30t", 348.9, 0.63, 3.21, False),
    ("back hoe", "25t", 268.0, 0.60, 3.61, False),
    ("PBD", "50t", 855.9, 1.20, 7.67, False),
    ("PBD", "106t", 1056.0, 1.40, 9.09, False),
    ("DCM", "120t", 1177.2, 0.86, 5.23, False),
    ("TCM", "50t", 490.5, 0.60, 4.47, False),
    ("GCP", "-", 879.0, 0.76, 5.11, False),
    ("Pile driver", "105t", 1030.05, 0.80, 5.52, False),
    ("크레인", "230t", 2254.0, 1.20, 7.52, False),
    ("Dump", "15t", 260.0, 0.20, 0.50, True),
    ("Dump", "24t", 432.0, 0.23, 0.575, True),
]


def evaluate(equipment, thicknesses, p, include_zero=True):
    """
    equipment : list of dict(name, spec, W, b, a, dump, H_fix(None|float))
    thicknesses : 매트릭스(복토두께별 σ) 표에 표시할 두께 목록
    반환 : 장비별 결과 dict 목록
      - 검토 복토두께 H 는 지정값(H_fix)이 있으면 그 값,
        없으면 후보 두께 중 qa > σ 를 만족하는 최소 두께(없으면 최대 두께, N.G)
    """
    cands = sorted(set(([0.0] if include_zero else []) + list(thicknesses)))
    rows = []
    for e in equipment:
        P = contact_pressure(e["W"], e["b"], e["a"], e["dump"])
        sig_tbl = {h: sigma(P, e["b"], e["a"], h, p["gamma1"], p["theta"], p["eps"]) for h in thicknesses}

        if e.get("H_fix") is not None:
            H_sel, auto = float(e["H_fix"]), False
        else:
            H_sel, auto = cands[-1], True
            for h in cands:
                if qa_meyerhof(e["b"], e["a"], h, p) > sigma(P, e["b"], e["a"], h, p["gamma1"], p["theta"], p["eps"]):
                    H_sel = h
                    break

        s = sigma(P, e["b"], e["a"], H_sel, p["gamma1"], p["theta"], p["eps"])
        q = qa_meyerhof(e["b"], e["a"], H_sel, p)
        Fcs, Fqs, Frs, Fcd = shape_factors(e["b"], e["a"], p["phi2"], p["Df"])
        rows.append(dict(
            name=e["name"], spec=e["spec"], W=e["W"], b=e["b"], a=e["a"], dump=e["dump"],
            P=P, sig_tbl=sig_tbl, H=H_sel, auto=auto, sigma=s, qa=q,
            FS=(q / s if s > 0 else float("inf")), judge=("O.K" if q > s else "N.G"),
            Fcs=Fcs, Fqs=Fqs, Frs=Frs, Fcd=Fcd,
        ))
    return rows


def group_spans(names):
    """연속된 동일 장비명 → [(start_idx, end_idx, name)] (엑셀 셀 병합용)"""
    spans, i = [], 0
    while i < len(names):
        j = i
        while j + 1 < len(names) and names[j + 1] == names[i]:
            j += 1
        spans.append((i, j, names[i]))
        i = j + 1
    return spans


# ------------------------------------------------------------------
# 산정식 이미지 (matplotlib mathtext — LaTeX 설치 불필요)
# ------------------------------------------------------------------
FORMULA_TEX = {
    "P": r"$P=\frac{W}{2\,b\,a}\qquad \left(\mathrm{Dump}:\ P=\frac{0.4\,W}{b\,a}\right)$",
    "sigma": r"$\sigma=\frac{P\cdot b\cdot a\,(1+\varepsilon)}{(b+2H\tan\theta)\,(a+2H\tan\theta)}+\gamma_1 H$",
    "qa": r"$q_a=\frac{1}{F_s}\left[\left(1+0.2\frac{b}{L}\right)c_2 N_{c(2)}F_{cs(2)}F_{cd(2)}"
          r"+\gamma_1H^2\left(1+\frac{b}{L}\right)\frac{K_s\tan\phi_1}{b}+\frac{2T\sin\theta}{b+H}\right]$",
    "Nq": r"$N_{q(2)}=e^{\pi\tan\phi_2}\,\tan^2\left(45+\frac{\phi_2}{2}\right)$",
    "Nc": r"$N_{c(2)}=\frac{N_{q(2)}-1}{\tan\phi_2}\quad(\phi_2=0:\ 5.14)$",
    "Fcs": r"$F_{cs(2)}=1+\frac{b}{L}\cdot\frac{N_{q(2)}}{N_{c(2)}}\qquad F_{cd(2)}=1+0.4\frac{D_f}{b}$",
}

_formula_cache = {}


def formula_png(key, fontsize=18, dpi=200):
    if key in _formula_cache:
        return _formula_cache[key]
    tex = FORMULA_TEX[key]
    fig = plt.figure(figsize=(0.1, 0.1))
    fig.text(0, 0, tex, fontsize=fontsize)
    buf = io.BytesIO()
    fig.savefig(buf, dpi=dpi, transparent=False, facecolor="white",
                bbox_inches="tight", pad_inches=0.04, format="png")
    plt.close(fig)
    _formula_cache[key] = buf.getvalue()
    return _formula_cache[key]


def chart_png(rows, dpi=150):
    """장비별 작용응력(σ) vs 허용지지력(qa) 비교 그래프"""
    labels = [f"{r['name']}({r['spec']})\nH={r['H']:.1f}m" for r in rows]
    sig = [r["sigma"] for r in rows]
    qa = [r["qa"] for r in rows]
    x = list(range(len(rows)))
    w = max(7.0, 0.62 * len(rows) + 2.5)
    fig, ax = plt.subplots(figsize=(w, 4.4), dpi=dpi)
    ax.bar([i - 0.2 for i in x], sig, width=0.4, color="#90a4ae", label="작용응력 σ (kPa)")
    ax.bar([i + 0.2 for i in x], qa, width=0.4, color="#4f81bd", label="허용지지력 qa (kPa)")
    for i, r in enumerate(rows):
        ax.text(i, max(r["sigma"], r["qa"]) * 1.02, r["judge"], ha="center", va="bottom", fontsize=8,
                color=("#1b5e20" if r["judge"] == "O.K" else "#b71c1c"), fontweight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=7.5, rotation=0 if len(rows) <= 8 else 45,
                       ha="center" if len(rows) <= 8 else "right")
    ax.set_ylabel("kPa")
    ax.set_title("장비별 작용응력(σ) vs 허용지지력(qa)")
    ax.grid(axis="y", linestyle="--", alpha=0.4)
    ax.legend(fontsize=8, loc="upper left", bbox_to_anchor=(1.0, 1.0))
    ax.set_ylim(0, max(sig + qa + [1]) * 1.15)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi)
    plt.close(fig)
    return buf.getvalue()
