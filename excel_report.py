# -*- coding: utf-8 -*-
"""
excel_report.py — 수식이 살아있는 엑셀 구조계산서 생성
------------------------------------------------------------
시트1 '1.원지반 작용응력'  : 3.1 장비주행성.xlsm(1-1.원지반 작용응력)과 동일 배치
시트2 '2.허용지지력 검토'  : 2.1 장비_주행성.xlsx(Meyerhof and Hanna)와 동일 배치
시트3 '3.검토결과 그래프' : 셀에 연결된 엑셀 네이티브 차트

모든 계산 셀은 원본과 같은 엑셀 수식으로 작성되며 서로 연결됩니다.
  시트1 '허용지지력' 칸  ← 시트2 qa(M열) 수식
  시트2 P·b·L·H 칸      ← 시트1 장비표/검토결과 셀
  시트2 γ1·θ·i         ← 시트1 입력셀(S13·S14·S15)
따라서 엑셀에서 지반정수, 장비제원, 복토두께를 바꾸면 전체 결과가 자동 갱신됩니다.
(각 수식에는 프로그램이 계산한 결과값도 함께 저장되어 재계산 전에도 값이 표시됩니다.)
"""

import io

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell
from PIL import Image

import core

S1 = "1.원지반 작용응력"
S2 = "2.허용지지력 검토"
S3 = "3.검토결과 그래프"
R1 = f"'{S1}'!"
R2 = f"'{S2}'!"

FONT = "맑은 고딕"


def _c(r, c, row_abs=False, col_abs=False):
    """0-based (row, col) → 'A1' 표기"""
    return xl_rowcol_to_cell(r, c, row_abs, col_abs)


def _col_px(width_chars):
    return int(width_chars * 7 + 5)


def _fit_image(path, max_w, max_h):
    w, h = Image.open(path).size
    s = min(max_w / w, max_h / h)
    return s, w * s, h * s


def build_excel(project, rows, thicknesses, p):
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    wb.set_properties({"title": "장비주행성 검토 구조계산서"})
    wb.set_calc_mode("auto")
    wb.set_size(1600, 1000)

    def F(**kw):
        base = {"font_name": FONT, "font_size": 9, "valign": "vcenter"}
        base.update(kw)
        return wb.add_format(base)

    f_txt = F()
    f_bold = F(bold=True)
    f_title = F(bold=True, font_size=11)
    f_proj = F(bold=True, font_size=10)
    f_center = F(align="center")
    f_hdr = F(bold=True, align="center", text_wrap=True, border=1, bg_color="#F2F2F2")
    f_cell = F(align="center", border=1)
    f_cell_wrap = F(align="center", border=1, text_wrap=True, font_size=8)
    f_n1 = F(align="center", border=1, num_format="0.0")
    f_n2 = F(align="center", border=1, num_format="0.00")
    f_n3 = F(align="center", border=1, num_format="0.000")
    f_in1 = F(align="center", border=1, num_format="0.0", bg_color="#FFF9DB")   # 입력셀(복토두께)
    f_in2 = F(align="center", border=1, num_format="0.00", bg_color="#FFF9DB")
    f_judge = F(align="center", border=1, bold=True)
    f_eq = F(align="center")
    f_par1 = F(align="center", num_format="0.0", bg_color="#FFF9DB")
    f_par2 = F(align="center", num_format="0.00", bg_color="#FFF9DB")
    f_link = F(align="center", num_format="0.00")
    f_ng = wb.add_format({"font_color": "#C00000", "bg_color": "#FDE9E9"})
    f_ok = wb.add_format({"font_color": "#006100"})

    n_eq = len(rows)
    n_h = len(thicknesses)
    gamma1, theta, eps = p["gamma1"], p["theta"], p["eps"]
    Ks = core.ks_lookup(p["phi1"])
    Nc, Nq, Nr = core.nc2(p["phi2"]), core.nq2(p["phi2"]), core.nr2(p["phi2"])

    # =================================================================
    # 시트1 : 1.원지반 작용응력  (3.1 장비주행성 형식)
    # =================================================================
    ws = wb.add_worksheet(S1)
    ws.hide_gridlines(2)
    widths = {0: 2.8, 1: 8.5, 2: 4.7, 3: 1.6, 4: 6.4, 5: 5.8, 6: 5.8, 7: 7.2, 8: 6.1, 9: 6.1}
    last_h_col = 10 + n_h - 1
    last_col = max(15, last_h_col)                       # 최소 P열까지(원본 인쇄폭)
    for c in range(0, last_col + 1):
        ws.set_column(c, c, widths.get(c, 6.3))
    # 입력 매개변수 열 (원본: R·S열 / 두께 열이 많으면 그 오른쪽으로 자동 이동, 인쇄영역 밖)
    pc = max(17, last_col + 2)
    PL = xl_col_to_name(pc + 1)
    G1, TH, EP = f"${PL}$13", f"${PL}$14", f"${PL}$15"
    ws.set_column(pc, pc, 17)
    ws.set_column(pc + 1, pc + 1, 7)

    ws.write(0, 0, "PROJECT :", f_proj)
    ws.merge_range(0, 2, 0, last_col, project or "", F(bold=True, font_size=10, align="left"))
    ws.write(3, 1, "1. 복토 두께에 따른 원지반에 작용하는 응력(σ)산정", f_title)
    ws.merge_range(4, 1, 4, 8, "장비 접지압에 대한 Model", F(bold=True, align="center"))
    ws.merge_range(4, 9, 4, 14, "원지반상 작용응력(σ)", F(bold=True, align="center"))
    for r in range(5, 16):
        ws.set_row(r, 15)

    # 이론 삽도 (B6 ~ I16)
    img_w = sum(_col_px(widths.get(c, 6.3)) for c in range(1, 9))
    s, iw, ih = _fit_image(core.IMG_DISPERSION, img_w - 6, 11 * 20 - 4)
    ws.insert_image(5, 1, "dispersion.png", {"image_data": io.BytesIO(core.read_bytes(core.IMG_DISPERSION)), "x_scale": s, "y_scale": s,
                                             "x_offset": int((img_w - iw) / 2), "y_offset": 2,
                                             "object_position": 3})
    # 산정식 (J6~)
    eq_w = sum(_col_px(widths.get(c, 6.3)) for c in range(9, 16))
    for key, row, max_h in (("sigma", 5, 54), ("P", 8, 38)):
        png = core.formula_png(key)
        w0, h0 = Image.open(io.BytesIO(png)).size
        sc = min((eq_w - 10) / w0, max_h / h0)
        ws.insert_image(row, 9, f"f_{key}.png", {"image_data": io.BytesIO(png), "x_scale": sc, "y_scale": sc,
                                                 "x_offset": 4, "y_offset": 3, "object_position": 3})
    legend = [("σ", " : 원지반상 작용응력(kPa)"), ("P", " : 장비접지압(kPa),  W : 장비중량(kN)"),
              ("H", " : 복토층의 두께(m)"), ("θ", " : 하중분산각(°)"), ("a,b", " : 접지길이, 접지폭(m)")]
    for i, (sym, txt) in enumerate(legend):
        ws.write(10 + i, 9, sym, f_center)
        ws.write(10 + i, 10, txt, f_txt)
    ws.write(15, 9, "ε", f_center)
    ws.write_formula(15, 10, '=" : 충격계수,  γ1 : 복토 단위중량 = "&TEXT(' + G1 + ',"0.0")&" kN/m³"', f_txt,
                     f" : 충격계수,  γ1 : 복토 단위중량 = {gamma1:.1f} kN/m³")

    # 입력 매개변수 (원본과 같이 R·S열, 인쇄영역 밖)
    ws.write(11, pc, "▼ 입력값(수정 가능)", f_bold)
    ws.write(12, pc, "복토단위중량(kN/m3)", f_txt)
    ws.write_number(12, pc + 1, gamma1, f_par1)
    ws.write(13, pc, "하중분산각 θ(°)", f_txt)
    ws.write_number(13, pc + 1, theta, f_par1)
    ws.write(14, pc, "충격계수 ε", f_txt)
    ws.write_number(14, pc + 1, eps, f_par2)

    # 표 머리글 (17~18행)
    h0 = 16
    ws.set_row(h0, 22)
    ws.set_row(h0 + 1, 34)
    ws.merge_range(h0, 1, h0 + 1, 1, "장비명", f_hdr)
    ws.merge_range(h0, 2, h0 + 1, 3, "규격", f_hdr)
    ws.merge_range(h0, 4, h0 + 1, 4, "중량\n(W, kN)", f_hdr)
    ws.merge_range(h0, 5, h0, 6, "장비", f_hdr)
    ws.write(h0 + 1, 5, "폭(b)", f_hdr)
    ws.write(h0 + 1, 6, "길이(a)", f_hdr)
    ws.merge_range(h0, 7, h0 + 1, 7, "접지압\n(P, kPa)", f_hdr)
    ws.merge_range(h0, 8, h0 + 1, 8, "하중\n분산각\n(θ)", f_hdr)
    ws.merge_range(h0, 9, h0 + 1, 9, "충격\n계수\n(ε)", f_hdr)
    hdr_txt = "복토 두께에 따른 원지반상의 작용응력(σ, kN/m²)"
    if n_h > 1:
        ws.merge_range(h0, 10, h0, last_h_col, hdr_txt, f_hdr)
    else:
        ws.write(h0, 10, "σ (kN/m²)", f_hdr)
    for j, h in enumerate(thicknesses):
        ws.write_number(h0 + 1, 10 + j, h, F(bold=True, align="center", border=1, bg_color="#F2F2F2",
                                                num_format="0.0"))

    # 장비 행 (19행~)
    d0 = h0 + 2
    names = [r["name"] for r in rows]
    for a, b_, nm in core.group_spans(names):
        if b_ > a:
            ws.merge_range(d0 + a, 1, d0 + b_, 1, nm, f_cell_wrap)
        else:
            ws.write(d0 + a, 1, nm, f_cell_wrap)
    for i, r in enumerate(rows):
        rr = d0 + i
        R = rr + 1  # 엑셀 행번호
        ws.merge_range(rr, 2, rr, 3, r["spec"], f_cell)
        ws.write_number(rr, 4, r["W"], f_n2)
        ws.write_number(rr, 5, r["b"], f_n2)
        ws.write_number(rr, 6, r["a"], F(align="center", border=1, num_format="0.00#"))
        fP = f"=ROUND(E{R}*0.4/(F{R}*G{R}),2)" if r["dump"] else f"=ROUND(E{R}/(2*F{R}*G{R}),2)"
        ws.write_formula(rr, 7, fP, f_n2, r["P"])
        if i == 0:
            ws.write_formula(rr, 8, "=" + TH, f_n1, theta)
            ws.write_formula(rr, 9, "=" + EP, f_n2, eps)
        else:
            ws.write_formula(rr, 8, f"=$I${d0 + 1}", f_n1, theta)
            ws.write_formula(rr, 9, f"=$J${d0 + 1}", f_n2, eps)
        for j, h in enumerate(thicknesses):
            col = xl_col_to_name(10 + j)
            hr = f"{col}${h0 + 2}"
            fs = (f"=$H{R}*$F{R}*$G{R}*(1+$J{R})/(($F{R}+2*{hr}*TAN(RADIANS($I{R})))"
                  f"*($G{R}+2*{hr}*TAN(RADIANS($I{R}))))+{hr}*" + G1)
            ws.write_formula(rr, 10 + j, fs, f_n2, r["sig_tbl"][h])
    end1 = d0 + n_eq - 1

    # ---------------- 2. 장비주행성 검토결과 ----------------
    t2 = end1 + 2
    ws.write(t2, 1, "2. 장비주행성 검토결과", f_title)
    hh = t2 + 2
    ws.set_row(hh, 16)
    ws.set_row(hh + 1, 16)
    heads = [((1, 1), "장비명"), ((2, 3), "규격"), ((4, 5), "접지압\n(P, kPa)"), ((6, 7), "복토 두께\n(m)"),
             ((8, 9), "작용응력\n(kPa)"), ((10, 11), "허용지지력\n(kPa)"), ((12, 13), "판정"), ((14, 15), "비고")]
    for (c1, c2), t in heads:
        ws.merge_range(hh, c1, hh + 1, c2, t, f_hdr)
    s0 = hh + 2
    s2_first = 33  # 시트2 장비 첫 행(0-based) = 엑셀 34행
    for a, b_, nm in core.group_spans(names):
        if b_ > a:
            ws.merge_range(s0 + a, 1, s0 + b_, 1, nm, f_cell_wrap)
        else:
            ws.write(s0 + a, 1, nm, f_cell_wrap)
    for i, r in enumerate(rows):
        rr = s0 + i
        R = rr + 1
        Rt = d0 + i + 1        # 시트1 장비표 행
        R2row = s2_first + i + 1
        ws.merge_range(rr, 2, rr, 3, r["spec"], f_cell)
        ws.merge_range(rr, 4, rr, 5, "", f_n2)
        ws.write_formula(rr, 4, f"=H{Rt}", f_n2, r["P"])
        ws.merge_range(rr, 6, rr, 7, "", f_in1)
        ws.write_number(rr, 6, r["H"], f_in1)
        ws.merge_range(rr, 8, rr, 9, "", f_n2)
        fs = (f"=$H{Rt}*$F{Rt}*$G{Rt}*(1+$J{Rt})/(($F{Rt}+2*G{R}*TAN(RADIANS($I{Rt})))"
              f"*($G{Rt}+2*G{R}*TAN(RADIANS($I{Rt}))))+G{R}*" + G1)
        ws.write_formula(rr, 8, fs, f_n2, r["sigma"])
        ws.merge_range(rr, 10, rr, 11, "", f_n2)
        ws.write_formula(rr, 10, f"={R2}M{R2row}", f_n2, r["qa"])
        ws.merge_range(rr, 12, rr, 13, "", f_judge)
        ws.write_formula(rr, 12, f'=IF(K{R}>I{R},"O.K","N.G")', f_judge, r["judge"])
    end2 = s0 + n_eq - 1
    note = ("허용지지력은\nMeyerhof and\nHanna(1978)\n층상지반\n지지력식에 의한\n산정값\n"
            "(2.허용지지력\n검토 참조)\n\n복토두께(노란칸)\n수정 시 자동\n재계산")
    if n_eq > 1:
        ws.merge_range(s0, 14, end2, 15, note, f_cell_wrap)
    else:
        ws.merge_range(s0, 14, s0, 15, "2.허용지지력 검토 참조", f_cell_wrap)
    ws.conditional_format(s0, 12, end2, 13, {"type": "cell", "criteria": "==", "value": '"N.G"', "format": f_ng})
    ws.conditional_format(s0, 12, end2, 13, {"type": "cell", "criteria": "==", "value": '"O.K"', "format": f_ok})

    ws.set_paper(9)
    ws.set_portrait()
    ws.center_horizontally()
    ws.set_margins(left=0.5, right=0.4, top=0.6, bottom=0.6)
    ws.print_area(0, 0, end2 + 1, last_col)
    ws.fit_to_pages(1, 0)
    if end2 > 58:                         # 장비가 많으면 '2. 검토결과'를 다음 페이지로
        ws.set_h_pagebreaks([t2])
    ws.repeat_rows(0)
    ws.set_footer("&C&9- &P / &N -")
    ws.freeze_panes(0, 0)

    # =================================================================
    # 시트2 : 2.허용지지력 검토  (2.1 장비_주행성 형식, Meyerhof and Hanna)
    # =================================================================
    w2 = wb.add_worksheet(S2)
    w2.hide_gridlines(2)
    w2.set_column(0, 0, 2.8)
    w2.set_column(1, 1, 6.2)
    w2.set_column(2, 14, 6.3)
    w2.set_column(3, 3, 7.4)   # 접지압(P) — 1000 kPa 이상 표시
    w2.set_column(15, 15, 2.8)
    w2.set_column(16, 19, 6.2)

    w2.write(0, 0, "PROJECT :", f_proj)
    w2.merge_range(0, 2, 0, 15, project or "", F(bold=True, font_size=10, align="left"))
    w2.write(3, 1, "▶ Meyerhof and Hanna(1978)모델", f_title)
    w2.write(4, 1, ": 층상지반에 대한 지지력 검토식으로 토목섬유의 인장력을 고려할 수 있도록 변형된 공식", f_txt)
    w2.write(5, 1, ": 연약한 점토층 위에 단단한 모래층이 있는 경우(c1=0, Φ2=0) 설계적용", f_txt)
    for r in range(6, 16):
        w2.set_row(r, 15)
    # 삽도 (C7 ~ N16)
    area_w = sum(_col_px(6.3) for _ in range(2, 14))
    s, iw, ih = _fit_image(core.IMG_MEYERHOF, area_w, 10 * 20 - 4)
    w2.insert_image(6, 2, "meyerhof.png", {"image_data": io.BytesIO(core.read_bytes(core.IMG_MEYERHOF)), "x_scale": s, "y_scale": s,
                                           "x_offset": int((area_w - iw) / 2), "y_offset": 2, "object_position": 3})
    # 산정식 (17행)
    w2.set_row(16, 48)
    png = core.formula_png("qa")
    w0, h0p = Image.open(io.BytesIO(png)).size
    full_w = _col_px(6.2) + 13 * _col_px(6.3)
    sc = min((full_w - 10) / w0, 62 / h0p)
    w2.insert_image(16, 1, "f_qa.png", {"image_data": io.BytesIO(png), "x_scale": sc, "y_scale": sc,
                                        "x_offset": 4, "y_offset": 2, "object_position": 3})
    w2.write(17, 9, "*원지반의 지지력+관입저항력+토목섬유 인장력", F(font_size=8, italic=True))

    # 범례/입력 (20~27행) — 원본 2.1과 동일 셀 위치
    w2.write(19, 1, "여기서,", f_txt)
    w2.write(19, 2, "a, b", f_txt)
    w2.write(19, 3, ": 장비의 접지 길이, 접지폭", f_txt)
    w2.write(19, 9, "B", f_txt)
    w2.write(19, 10, ": 원지반상 하중 작용 폭", f_txt)
    w2.write(19, 17, "원지반 지지력계수", f_bold)

    def leg(r, c_sym, sym, c_txt, txt):
        w2.write(r, c_sym, sym, f_txt)
        w2.write(r, c_txt, txt, f_txt)

    leg(20, 2, "γ1", 3, ": 복토층 단위중량(kN/m3)")
    w2.write_string(20, 7, "=", f_eq)
    w2.write_formula(20, 8, "=" + R1 + G1, f_link, gamma1)
    leg(20, 9, "γ2", 10, ": 원지반 단위중량(kN/m3)")
    w2.write_string(20, 13, "=", f_eq)
    w2.write_number(20, 14, p["gamma2"], f_par1)

    leg(21, 2, "Df", 3, ": 근입심도(m)")
    w2.write_string(21, 7, "=", f_eq)
    w2.write_number(21, 8, p["Df"], f_par2)
    leg(21, 9, "c2", 10, ": 원지반 점착력(kPa)")
    w2.write_string(21, 13, "=", f_eq)
    w2.write_number(21, 14, p["c2"], f_par1)

    w2.write(22, 2, "Nq(2),Nr(2),Nc(2)", f_txt)
    w2.write(22, 4, ": 원지반의 지지력계수", f_txt)
    w2.write(22, 9, "Fqs(2),Frs(2),Fcs(2)", f_txt)
    w2.write(22, 11, ": 형상계수", f_txt)

    leg(23, 2, "θ", 3, ": 토목섬유가 수평면과 이루는 각")
    w2.write_string(23, 7, "=", f_eq)
    w2.write_formula(23, 8, "=" + R1 + TH, F(align="center", num_format="0.0"), theta)
    leg(23, 9, "H", 10, ": 복토층 두께")

    leg(24, 2, "Ks", 3, ": 펀칭전단계수")
    w2.write_string(24, 7, "=", f_eq)
    w2.write_formula(24, 8, "=IF(O25<G29,0,LOOKUP(O25,G29:M29,G30:M30))", f_link, Ks)
    leg(24, 9, "Φ1", 10, ": 복토층 내부마찰각")
    w2.write_string(24, 13, "=", f_eq)
    w2.write_number(24, 14, p["phi1"], f_par1)

    leg(25, 2, "i", 3, ": 충격계수")
    w2.write_string(25, 7, "=", f_eq)
    w2.write_formula(25, 8, "=" + R1 + EP, f_link, eps)
    leg(25, 9, "Φ2", 10, ": 원지반 내부마찰각")
    w2.write_string(25, 13, "=", f_eq)
    w2.write_number(25, 14, p["phi2"], f_par1)

    leg(26, 2, "T", 3, ": 토목섬유의 허용인장력(kN/m)")
    w2.write_string(26, 7, "=", f_eq)
    w2.write_number(26, 8, p["T"], f_par1)
    leg(26, 9, "Fs", 10, ": 안전율")
    w2.write_string(26, 13, "=", f_eq)
    w2.write_number(26, 14, p["Fs"], f_par1)

    # 지지력계수 (R21~S23)
    w2.write(20, 17, "Nc(2)", f_txt)
    w2.write_formula(20, 18, "=IF(O26=0,5.14,(S22-1)/TAN(O26*PI()/180))", F(num_format="0.000"), Nc)
    w2.write(21, 17, "Nq(2)", f_txt)
    w2.write_formula(21, 18, "=EXP(PI()*TAN(O26*PI()/180))*(TAN((45+O26/2)*PI()/180))^2",
                     F(num_format="0.000"), Nq)
    w2.write(22, 17, "Nr(2)", f_txt)
    w2.write_formula(22, 18, "=(S22-1)*TAN(1.4*O26*PI()/180)", F(num_format="0.000"), Nr)

    # Ks 조견표 (29~30행)
    w2.merge_range(28, 3, 28, 5, "복토층 마찰각(Φ1˚)", f_hdr)
    w2.merge_range(29, 3, 29, 5, "Ks", f_hdr)
    for j, (ph, kv) in enumerate(zip(core.KS_PHI, core.KS_VAL)):
        w2.write_number(28, 6 + j, ph, f_cell)
        w2.write_number(29, 6 + j, kv, f_n2)

    # 검토표 (31~)
    w2.write_formula(30, 1, '="■ 검토구간 (토목섬유 허용인장력 T = "&TEXT(I27,"0.0")&" kN/m)"', f_title,
                     f"■ 검토구간 (토목섬유 허용인장력 T = {p['T']:.1f} kN/m)")
    w2.set_row(31, 18)
    w2.set_row(32, 26)
    w2.merge_range(31, 1, 32, 2, "사용장비", f_hdr)
    w2.merge_range(31, 3, 31, 5, "장비 제원", f_hdr)
    w2.write(32, 3, "접지압(P)", f_hdr)
    w2.write(32, 4, "폭(b)", f_hdr)
    w2.write(32, 5, "길이(L)", f_hdr)
    w2.merge_range(31, 6, 31, 9, "원지반상 작용 응력", f_hdr)
    w2.merge_range(32, 6, 32, 7, "복토 두께(H)", f_hdr)
    w2.merge_range(32, 8, 32, 9, "작용 응력(σ)", f_hdr)
    w2.merge_range(31, 10, 32, 11, "Geotextile\n허용강도(T)", f_hdr)
    w2.merge_range(31, 12, 31, 13, "허용 지지력(qa)", f_hdr)
    w2.merge_range(32, 12, 32, 13, "Meyerhof and Hanna", F(bold=True, align="center", text_wrap=True, border=1,
                                                             bg_color="#F2F2F2", font_size=7.5))
    w2.merge_range(31, 14, 32, 14, "적합\n여부", f_hdr)
    w2.merge_range(31, 16, 31, 19, "원지반 형상계수(hansen)", f_hdr)
    for j, t in enumerate(["Fcs(2)", "Fqs(2)", "Frs(2)", "Fcd(2)"]):
        w2.write(32, 16 + j, t, f_hdr)

    for i, r in enumerate(rows):
        rr = s2_first + i
        R = rr + 1
        Rt = d0 + i + 1      # 시트1 장비표 행
        Rs = s0 + i + 1      # 시트1 검토결과 행
        w2.merge_range(rr, 1, rr, 2, f"{r['name']}({r['spec']})", f_cell_wrap)
        w2.write_formula(rr, 3, f"={R1}H{Rt}", f_n2, r["P"])
        w2.write_formula(rr, 4, f"={R1}F{Rt}", f_n2, r["b"])
        w2.write_formula(rr, 5, f"={R1}G{Rt}", F(align="center", border=1, num_format="0.00#"), r["a"])
        w2.merge_range(rr, 6, rr, 7, "", f_n1)
        w2.write_formula(rr, 6, f"={R1}G{Rs}", f_n1, r["H"])
        w2.merge_range(rr, 8, rr, 9, "", f_n2)
        w2.write_formula(rr, 8, (f"=(D{R}*E{R}*F{R})*(1+$I$26)/((E{R}+2*(G{R}*TAN(RADIANS($I$24))))"
                                 f"*(F{R}+(2*G{R}*TAN(RADIANS($I$24)))))+$I$21*G{R}"), f_n2, r["sigma"])
        w2.merge_range(rr, 10, rr, 11, "", f_n1)
        w2.write_formula(rr, 10, "=$I$27", f_n1, p["T"])
        w2.merge_range(rr, 12, rr, 13, "", f_n2)
        w2.write_formula(rr, 12, (f"=(1/$O$27)*((1+0.2*(E{R}/F{R}))*$O$22*$S$21*Q{R}*T{R}"
                                  f"+$I$21*G{R}^2*(1+(E{R}/F{R}))*($I$25*TAN(RADIANS($O$25))/E{R})"
                                  f"+(2*K{R}*SIN(RADIANS($I$24)))/(E{R}+G{R}))"), f_n2, r["qa"])
        w2.write_formula(rr, 14, f'=IF(M{R}>I{R},"O.K","N.G")', f_judge, r["judge"])
        w2.write_formula(rr, 16, f"=1+(E{R}/F{R})*($S$22/$S$21)", f_n3, r["Fcs"])
        w2.write_formula(rr, 17, f"=ROUND(IF($O$26=0,1,1+TAN($O$26*PI()/180)*E{R}/F{R}),2)", f_n2, r["Fqs"])
        w2.write_formula(rr, 18, f"=ROUND(1-0.4*E{R}/F{R},2)", f_n2, r["Frs"])
        w2.write_formula(rr, 19, f"=1+0.4*($I$22/E{R})", f_n3, r["Fcd"])
    e2 = s2_first + n_eq - 1
    w2.conditional_format(s2_first, 14, e2, 14, {"type": "cell", "criteria": "==", "value": '"N.G"', "format": f_ng})
    w2.conditional_format(s2_first, 14, e2, 14, {"type": "cell", "criteria": "==", "value": '"O.K"', "format": f_ok})
    w2.write(e2 + 2, 1, "※ 노란색 셀은 입력값이며, 수정하면 시트 전체가 자동 재계산됩니다. "
                        "(P·b·L·H는 '1.원지반 작용응력' 시트와 연결)", F(font_size=8, italic=True))

    w2.set_paper(9)
    w2.set_portrait()
    w2.center_horizontally()
    w2.set_margins(left=0.5, right=0.4, top=0.6, bottom=0.6)
    w2.print_area(0, 0, e2 + 2, 15)          # 원본과 동일하게 A~P열 인쇄 (형상계수 Q~T는 보조열)
    w2.fit_to_pages(1, 0)
    w2.set_footer("&C&9- &P / &N -")

    # =================================================================
    # 시트3 : 검토결과 그래프 (셀 연결 네이티브 차트)
    # =================================================================
    w3 = wb.add_worksheet(S3)
    w3.hide_gridlines(2)
    w3.set_column(0, 0, 11)
    w3.write(0, 0, "PROJECT :", f_proj)
    w3.write(0, 1, project or "", f_proj)
    w3.write(2, 0, "3. 장비별 작용응력(σ) vs 허용지지력(qa) 비교", f_title)
    ch = wb.add_chart({"type": "column"})
    cats = f"={R2}$B${s2_first + 1}:$B${e2 + 1}"
    ch.add_series({"name": "작용응력 σ (kPa)", "categories": cats,
                   "values": f"={R2}$I${s2_first + 1}:$I${e2 + 1}", "fill": {"color": "#90A4AE"}})
    ch.add_series({"name": "허용지지력 qa (kPa)", "categories": cats,
                   "values": f"={R2}$M${s2_first + 1}:$M${e2 + 1}", "fill": {"color": "#4F81BD"}})
    ch.set_title({"name": "장비별 작용응력(σ) vs 허용지지력(qa)", "name_font": {"name": FONT, "size": 12}})
    ch.set_y_axis({"name": "kPa", "major_gridlines": {"visible": True, "line": {"dash_type": "dash"}},
                   "num_font": {"name": FONT}})
    ch.set_x_axis({"num_font": {"name": FONT, "size": 8, "rotation": -45 if n_eq > 8 else 0}})
    ch.set_legend({"position": "bottom", "font": {"name": FONT}})
    ch.set_size({"width": max(720, 48 * n_eq + 200), "height": 420})
    w3.insert_chart(4, 0, ch)
    w3.set_paper(9)
    w3.set_landscape()
    w3.fit_to_pages(1, 1)

    wb.close()
    return buf.getvalue()
