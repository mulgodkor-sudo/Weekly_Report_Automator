"""
plant_mh.py
Plant M/H 입력용 집계 + Excel 시트 작성 (화면/Excel 공통)

회사 M/H 입력 시스템 화면과 최대한 같은 배치:
  No. / 구분(General·Project) / Project Code / Func. Code / Description /
  날짜마다 ST·OT 두 칸 (주말 포함 모든 날짜) / ST TOTAL / OT TOTAL
  - 머리글 4단: 년월 / 요일 / 일 / ST·OT
  - OT 칸 연두색, 토·일 머리글 빨간 글씨
"""
from __future__ import annotations
from collections import OrderedDict
from datetime import date, datetime, timedelta

GENERAL_PC = "000000"
DAY_KO     = ["월", "화", "수", "목", "금", "토", "일"]

# 색상 (화면/Excel 공통)
C_HDR_BG   = "#F2F2F2"   # 머리글 배경
C_GRID     = "#BFBFBF"   # 테두리
C_OT_BG    = "#CCEB9E"   # OT 칸 연두색
C_WEEKEND  = "#E00000"   # 토·일 머리글 글씨
C_TOTAL_BG = "#FFE699"   # 합계 행
NOTE = "※ ST: 평일 08:30~17:30 (하루 최대 8H)  |  OT: 그 외 시간 및 토·일"


def to_date(d) -> date:
    return d.date() if isinstance(d, datetime) else d


def week_sunday(d: date) -> date:
    """d가 속한 주(월~일)의 일요일"""
    return d + timedelta(days=6 - d.weekday())


def date_list(start: date, end: date) -> list[date]:
    """start ~ end 모든 날짜 (주말 포함)"""
    return [start + timedelta(days=i) for i in range((end - start).days + 1)]


def mh_gubun(pc: str) -> str:
    """M/H 시스템 구분: 000000 → General, 그 외 → Project"""
    if not pc:
        return ""
    return "General" if pc == GENERAL_PC else "Project"


def fmt_h(v: float) -> str:
    """3.0 → '3', 3.5 → '3.5' (시스템 표기와 동일)"""
    return f"{v:g}"


def aggregate(rows: list[dict]) -> list[dict]:
    """
    [실적] 행 → (Project Code, Func. Code) 단위 날짜별 [ST, OT] 집계
    정렬: General 먼저 → Project Code → Func. Code
    """
    agg: dict[tuple, dict] = OrderedDict()
    for r in rows:
        if r.get("source") != "this_week" or r.get("this_week_h", 0) <= 0:
            continue
        pc = r.get("project_code") or ""
        fc = r.get("func_code")    or ""
        item = agg.setdefault((pc, fc), dict(
            gubun=mh_gubun(pc), pc=pc, fc=fc,
            desc=r.get("func_name") or "", cells={}))
        cell = item["cells"].setdefault(to_date(r["date"]), [0.0, 0.0])
        cell[0] = round(cell[0] + r.get("this_week_st", r.get("this_week_h", 0.0)), 1)
        cell[1] = round(cell[1] + r.get("this_week_ot", 0.0), 1)

    items = list(agg.values())
    items.sort(key=lambda it: (it["gubun"] != "General", it["pc"], it["fc"]))
    for it in items:
        it["st_total"] = round(sum(c[0] for c in it["cells"].values()), 1)
        it["ot_total"] = round(sum(c[1] for c in it["cells"].values()), 1)
    return items


def day_totals(items: list[dict], dates: list[date]) -> dict:
    tot = {d: [0.0, 0.0] for d in dates}
    for it in items:
        for d, (st, ot) in it["cells"].items():
            if d in tot:
                tot[d][0] = round(tot[d][0] + st, 1)
                tot[d][1] = round(tot[d][1] + ot, 1)
    return tot


def month_groups(dates: list[date]) -> list[tuple[str, int, int]]:
    """[(라벨 '2026년 09월', 시작 index, 끝 index)] — 날짜 기준 index"""
    groups = []
    for i, d in enumerate(dates):
        lbl = f"{d.year}년 {d.month:02d}월"
        if groups and groups[-1][0] == lbl:
            groups[-1] = (lbl, groups[-1][1], i)
        else:
            groups.append((lbl, i, i))
    return groups


# ════════════════════════════════════════════════════════════════════
# Excel 시트
# ════════════════════════════════════════════════════════════════════

FIXED_HDRS   = ["No.", "구분", "Project\nCode", "Func. Code", "Description"]
FIXED_WIDTHS = [5, 9, 10, 11, 32]


def write_sheet(wb, ws, items: list[dict], dates: list[date]) -> None:
    base = {"font_name": "맑은 고딕", "font_size": 9, "border": 1,
            "border_color": C_GRID, "align": "center", "valign": "vcenter"}

    def F(**kw):
        return wb.add_format({**base, **kw})

    hdr     = F(bold=True, bg_color=C_HDR_BG, text_wrap=True)
    hdr_we  = F(bold=True, bg_color=C_HDR_BG, font_color=C_WEEKEND)
    txt     = F()
    txt_l   = F(align="left")
    num_st  = F(num_format="0.##;-0.##;0")
    num_ot  = F(num_format="0.##;-0.##;0", bg_color=C_OT_BG)
    tot_lbl = F(bold=True, bg_color=C_TOTAL_BG)
    tot_st  = F(bold=True, bg_color=C_TOTAL_BG, num_format="0.##;-0.##;0")
    tot_ot  = F(bold=True, bg_color=C_TOTAL_BG, num_format="0.##;-0.##;0")
    note    = wb.add_format({"font_name": "맑은 고딕", "font_size": 9,
                             "font_color": "#7F7F7F"})

    nf    = len(FIXED_HDRS)
    c_st  = nf + 2 * len(dates)       # ST TOTAL 열
    c_ot  = c_st + 1                  # OT TOTAL 열
    H0    = 1                         # 머리글 시작 행 (0행은 안내문)
    D0    = H0 + 4                    # 데이터 시작 행

    ws.write(0, 0, NOTE, note)

    # 열 너비
    for ci, w in enumerate(FIXED_WIDTHS):
        ws.set_column(ci, ci, w)
    if dates:
        ws.set_column(nf, c_st - 1, 4.5)
    ws.set_column(c_st, c_ot, 7)

    # ── 머리글 ──
    for r in range(H0, D0):
        ws.set_row(r, 17)
    for ci, h in enumerate(FIXED_HDRS):
        ws.merge_range(H0, ci, D0 - 1, ci, h, hdr)
    for lbl, i0, i1 in month_groups(dates):
        ws.merge_range(H0, nf + 2 * i0, H0, nf + 2 * i1 + 1, lbl, hdr)
    for i, d in enumerate(dates):
        c  = nf + 2 * i
        fm = hdr_we if d.weekday() >= 5 else hdr
        ws.merge_range(H0 + 1, c, H0 + 1, c + 1, DAY_KO[d.weekday()], fm)
        ws.merge_range(H0 + 2, c, H0 + 2, c + 1, d.day, fm)
        ws.write(H0 + 3, c,     "ST", fm)
        ws.write(H0 + 3, c + 1, "OT", fm)
    ws.merge_range(H0, c_st, D0 - 1, c_st, "ST\nTOTAL", hdr)
    ws.merge_range(H0, c_ot, D0 - 1, c_ot, "OT\nTOTAL", hdr)

    # ── 데이터 ──
    def put(r, c, v, fmt, show_zero):
        if v or show_zero:
            ws.write_number(r, c, v, fmt)
        else:
            ws.write_blank(r, c, None, fmt)

    for ri, it in enumerate(items):
        r = D0 + ri
        ws.set_row(r, 18)
        ws.write(r, 0, ri + 1,   txt)
        ws.write(r, 1, it["gubun"], txt)
        ws.write_string(r, 2, it["pc"], txt)
        ws.write(r, 3, it["fc"], txt)
        ws.write(r, 4, it["desc"], txt_l)
        for i, d in enumerate(dates):
            st, ot = it["cells"].get(d, (0.0, 0.0))
            has = d in it["cells"]          # 입력한 날은 0도 표시 (시스템과 동일)
            put(r, nf + 2 * i,     st, num_st, has)
            put(r, nf + 2 * i + 1, ot, num_ot, has)
        put(r, c_st, it["st_total"], num_st, True)
        put(r, c_ot, it["ot_total"], num_st, True)

    # ── 합계 ──
    tr  = D0 + len(items)
    tot = day_totals(items, dates)
    ws.set_row(tr, 20)
    ws.merge_range(tr, 0, tr, nf - 1, "합  계", tot_lbl)
    for i, d in enumerate(dates):
        put(tr, nf + 2 * i,     tot[d][0], tot_st, False)
        put(tr, nf + 2 * i + 1, tot[d][1], tot_ot, False)
    put(tr, c_st, round(sum(it["st_total"] for it in items), 1), tot_st, True)
    put(tr, c_ot, round(sum(it["ot_total"] for it in items), 1), tot_ot, True)

    ws.freeze_panes(D0, nf)


def save_workbook(path: str, items: list[dict], dates: list[date]) -> None:
    import xlsxwriter
    wb = xlsxwriter.Workbook(path)
    write_sheet(wb, wb.add_worksheet("Plant MH"), items, dates)
    wb.close()
