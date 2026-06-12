"""
monthly_dialog.py  ─  월간 업무 정리 다이얼로그 v2
"""
from __future__ import annotations
import os, threading
from calendar import monthrange
from datetime import datetime, timedelta
from tkinter import ttk, filedialog, messagebox
import tkinter as tk

from config            import get_fc, get_func_name
from outlook_reader    import get_events
from monthly_processor import process_monthly, HOLIDAY_FC

# ── Excel 행 높이 추정 ──────────────────────────────────────────────
def _est_h(body: str, dates: list[str], col_body: int = 44) -> float:
    """본문 줄 수 + 날짜 수 기반 행 높이(pt) 추정"""
    body_lines = 0
    for ln in (body or "").split("\n"):
        ln_len = sum(2 if ord(c) > 127 else 1 for c in ln)
        body_lines += max(1, -(-ln_len // col_body))
    date_lines = len(dates)
    lines = max(body_lines, date_lines, 1)
    return min(200.0, max(20.0, lines * 14.5))


# ══════════════════════════════════════════════════════════════════
class MonthlyDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("월간 업무 정리")
        self.geometry("1080x760")
        self.minsize(860, 480)
        self.resizable(True, True)
        self.configure(bg="#F0F2F5")
        self._result    = None
        self._month_str = ""
        self._build()

    # ── UI ────────────────────────────────────────────────────────
    def _build(self):
        bar = tk.Frame(self, bg="#1F4E79", height=42)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="  📋  월간 업무 정리",
                 font=("맑은 고딕", 12, "bold"),
                 fg="white", bg="#1F4E79").pack(side="left", padx=12, pady=8)

        body = tk.Frame(self, bg="#F0F2F5", padx=12, pady=8)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        # ── 기간 설정 ──
        fr_d = ttk.LabelFrame(body, text="  📅  기간 설정", padding=(10,6))
        fr_d.grid(row=0, column=0, sticky="ew", pady=(0,8))

        P = dict(padx=6, pady=4)
        ttk.Label(fr_d, text="시작일 :").grid(row=0, column=0, sticky="e", **P)
        self.v_start = tk.StringVar()
        ttk.Entry(fr_d, textvariable=self.v_start, width=14,
                  font=("맑은 고딕",10)).grid(row=0, column=1, sticky="w", **P)
        ttk.Label(fr_d, text=" ~ 종료일 :").grid(row=0, column=2, **P)
        self.v_end = tk.StringVar()
        ttk.Entry(fr_d, textvariable=self.v_end, width=14,
                  font=("맑은 고딕",10)).grid(row=0, column=3, sticky="w", **P)

        self._set_prev_month()

        tk.Button(fr_d, text="🔍  불러오기",
                  command=self._load,
                  font=("맑은 고딕",10,"bold"),
                  bg="#1F4E79", fg="white",
                  activebackground="#2E75B6", activeforeground="white",
                  relief="flat", padx=14, pady=5, cursor="hand2"
                  ).grid(row=0, column=4, padx=(14,4), pady=4)

        self.progress = ttk.Progressbar(fr_d, mode="indeterminate", length=110)
        self.progress.grid(row=0, column=5, padx=4)

        ttk.Label(fr_d,
                  text="* 전월 날짜 자동 입력 / YYYY-MM-DD 형식으로 수정 가능 "
                       "/ Non-KPI 순서는 FC 엑셀 파일 기준",
                  foreground="gray", font=("맑은 고딕",8)
                  ).grid(row=1, column=0, columnspan=6, sticky="w", padx=6)

        # ── 결과 Treeview ──
        fr_tv = ttk.LabelFrame(body, text="  📊  결과  (같은 내용은 날짜 병합 / 별표 이후 내용 제거)", padding=(6,4))
        fr_tv.grid(row=1, column=0, sticky="nsew")
        fr_tv.rowconfigure(0, weight=1)
        fr_tv.columnconfigure(0, weight=1)

        COLS = ("gubun","pc","fc","subject","dates","body","hours")
        HDRS = ("구분","Project Code","FC","업무명","날짜","수행내용","시간")
        WIDS = (88,80,80,180,115,355,58)

        self.tv = ttk.Treeview(fr_tv, columns=COLS, show="headings", height=20)
        ysb = ttk.Scrollbar(fr_tv, orient="vertical",   command=self.tv.yview)
        xsb = ttk.Scrollbar(fr_tv, orient="horizontal", command=self.tv.xview)
        self.tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        for cid, hdr, w in zip(COLS, HDRS, WIDS):
            self.tv.heading(cid, text=hdr, anchor="center")
            anc = "w" if cid in ("subject","body","dates") else "center"
            self.tv.column(cid, width=w, minwidth=40, anchor=anc)

        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        self.tv.grid(row=0, column=0, sticky="nsew")

        self.tv.tag_configure("odd",      background="#FFFFFF")
        self.tv.tag_configure("even",     background="#F0F4FF")
        self.tv.tag_configure("subtotal", background="#FFE699",
                               font=("맑은 고딕",9,"bold"))
        self.tv.tag_configure("holiday",    background="#FCE4EC")
        self.tv.tag_configure("grand_total", background="#1F4E79",
                               foreground="white",
                               font=("맑은 고딕", 10, "bold"))
        self.tv.tag_configure("sum_hdr",    background="#37474F",
                               foreground="white",
                               font=("맑은 고딕", 9, "bold"))
        self.tv.tag_configure("sum_odd",    background="#FAFAFA")
        self.tv.tag_configure("sum_even",   background="#ECEFF1")

        # ── 하단 버튼 ──
        fr_btn = tk.Frame(self, bg="#F0F2F5")
        fr_btn.pack(fill="x", padx=12, pady=(4,10))

        self.lbl_stat = tk.Label(fr_btn, text="",
                                 bg="#F0F2F5", fg="gray",
                                 font=("맑은 고딕",8), anchor="w")
        self.lbl_stat.pack(side="left", padx=4)

        tk.Button(fr_btn, text="✅  닫기",
                  command=self.destroy,
                  font=("맑은 고딕", 11, "bold"),
                  bg="#455A64", fg="white",
                  activebackground="#546E7A", activeforeground="white",
                  relief="flat", padx=20, pady=8, cursor="hand2",
                  width=10,
                  ).pack(side="right", padx=(8, 0))

        self.btn_excel = tk.Button(
            fr_btn, text="📄  Excel 저장",
            command=self._save_excel,
            font=("맑은 고딕", 11, "bold"),
            bg="#1F5E20", fg="white",
            activebackground="#2E7D32", activeforeground="white",
            relief="flat", padx=20, pady=8, cursor="hand2",
            width=14,
            state="disabled",
        )
        self.btn_excel.pack(side="right")

    def _set_prev_month(self):
        today = datetime.today()
        last  = (today.replace(day=1) - timedelta(days=1))
        y, m  = last.year, last.month
        _, ld = monthrange(y, m)
        self.v_start.set(f"{y}-{m:02d}-01")
        self.v_end.set(f"{y}-{m:02d}-{ld:02d}")
        self._month_str = f"{y}년 {m}월"

    # ── 불러오기 ──────────────────────────────────────────────────
    def _load(self):
        try:
            start = datetime.strptime(self.v_start.get().strip(), "%Y-%m-%d")
            end   = datetime.strptime(self.v_end.get().strip(),   "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("날짜 오류", "YYYY-MM-DD 형식으로 입력하세요.", parent=self)
            return

        self._month_str = (
            f"{start.year}년 {start.month}월"
            if start.month == end.month else
            f"{start.year}-{start.month:02d} ~ {end.year}-{end.month:02d}"
        )
        self.progress.start()
        self.lbl_stat.config(text="아웃룩에서 일정을 불러오는 중...")

        def task():
            _com_ok = False
            try:
                import pythoncom; pythoncom.CoInitialize(); _com_ok = True
            except Exception: pass
            try:
                events  = get_events(start, end)
                fc_data = get_fc()
                result  = process_monthly(events, fc_data)
                self.after(0, lambda: self._on_ok(result))
            except Exception as e:
                msg = str(e)
                self.after(0, lambda: self._on_err(msg))
            finally:
                if _com_ok:
                    try: import pythoncom; pythoncom.CoUninitialize()
                    except Exception: pass

        threading.Thread(target=task, daemon=True).start()

    def _on_ok(self, result):
        self.progress.stop()
        self._result = result
        self._fill_tv(result)
        gc   = len(result["groups"])
        hh   = result["holiday"]["total_h"]
        self.lbl_stat.config(text=f"업무 그룹 {gc}건  |  휴가/공가 {hh:.1f}H")
        self.btn_excel.config(state="normal")

    def _on_err(self, msg):
        self.progress.stop()
        self.lbl_stat.config(text=f"오류: {msg[:60]}")
        messagebox.showerror("불러오기 오류", msg, parent=self)

    # ── Treeview ──────────────────────────────────────────────────
    def _fill_tv(self, result):
        self.tv.delete(*self.tv.get_children())
        ci = 0
        for g in result["groups"]:
            tag   = "even" if ci % 2 else "odd"
            ci   += 1
            first = True
            for row in g["rows"]:
                dates_str = ", ".join(row["dates"])
                vals = (
                    g["gubun"]        if first else "",
                    g["project_code"] if first else "",
                    g["func_code"]    if first else "",
                    g["subject"]      if first else "",
                    dates_str,
                    row["body"],
                    f"{row['total_h']:.1f}H",
                )
                self.tv.insert("", "end", values=vals, tags=(tag,))
                first = False
            # 소계
            self.tv.insert("", "end", values=(
                "","","","", "소  계","", f"{g['total_h']:.1f}H"
            ), tags=("subtotal",))

        hol = result["holiday"]
        if hol["dates"] or hol["total_h"] > 0:
            self.tv.insert("", "end", values=(
                "휴가/공가","—", HOLIDAY_FC,"휴가/공가",
                ", ".join(hol["dates"]) if hol["dates"] else "—",
                "—", f"{hol['total_h']:.1f}H"
            ), tags=("holiday",))

        # ── 월간 총 합계 ──────────────────────────────────────────
        grand = round(
            sum(g["total_h"] for g in result["groups"]) + hol["total_h"], 1
        )
        self.tv.insert("", "end", values=(
            "", "", "", "", "★  월간 총 합계", "", f"{grand:.1f}H"
        ), tags=("grand_total",))

        # ── PC / FC 별 소요시간 요약 ──────────────────────────────
        self.tv.insert("", "end", values=(
            "", "", "", "", "", "", ""
        ), tags=("odd",))  # 빈 구분 행

        self.tv.insert("", "end", values=(
            "[ PC / FC 별 소요시간 ]", "", "", "", "", "", ""
        ), tags=("sum_hdr",))

        summary = _build_summary(result)
        for si, s in enumerate(summary):
            tag = "sum_even" if si % 2 else "sum_odd"
            fn = s["func_name"][:28] if len(s["func_name"]) > 28 else s["func_name"]
            self.tv.insert("", "end", values=(
                s["gubun"], s["project_code"], s["func_code"],
                fn, "", "", f"{s['total_h']:.1f}H"
            ), tags=(tag,))

        # 휴가 요약 행
        if hol["total_h"] > 0:
            self.tv.insert("", "end", values=(
                "휴가/공가", "—", HOLIDAY_FC, "휴가/공가",
                "", "", f"{hol['total_h']:.1f}H"
            ), tags=("holiday",))

        # 요약 합계
        self.tv.insert("", "end", values=(
            "", "", "", "합  계", "", "", f"{grand:.1f}H"
        ), tags=("grand_total",))

    # ── Excel 저장 ────────────────────────────────────────────────
    def _save_excel(self):
        if not self._result: return
        path = filedialog.asksaveasfilename(
            title="월간 업무 정리 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일","*.xlsx")],
            initialfile=f"월간업무정리_{self._month_str.replace(' ','')}.xlsx",
            parent=self,
        )
        if not path: return
        try:
            _write_excel(self._result, path, self._month_str)
            if messagebox.askyesno("저장 완료",
                                   f"저장되었습니다.\n{path}\n\n바로 여시겠습니까?",
                                   parent=self):
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("저장 오류", str(e), parent=self)


# ══════════════════════════════════════════════════════════════════
# PC / FC 별 소요시간 집계
# ══════════════════════════════════════════════════════════════════

def _build_summary(result: dict) -> list[dict]:
    """
    groups에서 (gubun, project_code, func_code) 별 시간 합산.
    main groups와 동일한 정렬 순서 유지.
    """
    from collections import OrderedDict
    agg: dict[tuple, dict] = OrderedDict()
    for g in result["groups"]:
        key = (g["gubun"], g["project_code"], g["func_code"])
        if key not in agg:
            agg[key] = {
                "gubun":        g["gubun"],
                "project_code": g["project_code"],
                "func_code":    g["func_code"],
                "func_name":    get_func_name(g["func_code"]) or g["subject"],
                "total_h":      0.0,
            }
        agg[key]["total_h"] = round(agg[key]["total_h"] + g["total_h"], 1)
    return list(agg.values())


# ══════════════════════════════════════════════════════════════════
# Excel 생성
# # ══════════════════════════════════════════════════════════════════

def _write_excel(result: dict, path: str, month_str: str):
    import xlsxwriter
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("월간업무정리")

    def F(**kw):
        base = {"font_name":"맑은 고딕","font_size":10,
                "border":1,"valign":"vcenter"}
        return wb.add_format({**base, **kw})

    TITLE = wb.add_format({"font_name":"맑은 고딕","font_size":14,
                            "bold":True,"valign":"vcenter"})
    HDR   = F(bold=True, align="center", bg_color="#595959",
              font_color="#FFFFFF", text_wrap=True)
    ODD   = F(text_wrap=True, align="left")
    EVEN  = F(text_wrap=True, align="left",   bg_color="#F0F4FF")
    OC    = F(align="center")
    EC    = F(align="center", bg_color="#F0F4FF")
    OH    = F(align="center", num_format="0.0")
    EH    = F(align="center", num_format="0.0", bg_color="#F0F4FF")
    SUB   = F(bold=True, align="center", bg_color="#FFE699", num_format="0.0")
    SUBL  = F(bold=True, align="center", bg_color="#FFE699")
    HOL   = F(bold=True, align="center", bg_color="#FCE4EC")
    HOLT  = F(align="left", text_wrap=True, bg_color="#FCE4EC")
    HOLH  = F(bold=True, align="center", bg_color="#FCE4EC", num_format="0.0")

    ws.set_column(0, 0, 12)   # 구분
    ws.set_column(1, 1, 10)   # PC
    ws.set_column(2, 2, 10)   # FC
    ws.set_column(3, 3, 24)   # 업무명
    ws.set_column(4, 4, 16)   # 날짜
    ws.set_column(5, 5, 50)   # 수행내용
    ws.set_column(6, 6, 8)    # 시간

    ws.set_row(0, 28)
    ws.merge_range(0, 0, 0, 6, f"월간 업무 정리  —  {month_str}", TITLE)

    ws.set_row(1, 28)
    for ci, h in enumerate(["구분","Project Code","FC","업무명","날짜","수행내용","시간(H)"]):
        ws.write(1, ci, h, HDR)

    r = 2
    ci = 0
    for g in result["groups"]:
        fe  = EVEN  if ci % 2 else ODD
        fec = EC    if ci % 2 else OC
        feh = EH    if ci % 2 else OH
        ci += 1
        first = True
        for row in g["rows"]:
            # 날짜 줄바꿈 표시 (3개 이상이면 줄바꿈)
            dates_txt = "\n".join(row["dates"]) if len(row["dates"]) > 2 else ", ".join(row["dates"])
            h = _est_h(row["body"], row["dates"])
            ws.set_row(r, h)
            ws.write(r, 0, g["gubun"]        if first else "", fec)
            ws.write(r, 1, g["project_code"] if first else "", fec)
            ws.write(r, 2, g["func_code"]    if first else "", fec)
            ws.write(r, 3, g["subject"]      if first else "", fe)
            ws.write(r, 4, dates_txt, fe)
            ws.write(r, 5, row["body"],   fe)
            ws.write(r, 6, row["total_h"], feh)
            first = False
            r += 1

        ws.set_row(r, 18)
        for c in range(5): ws.write(r, c, "", SUBL)
        ws.write(r, 4, "소  계", SUBL)
        ws.write(r, 5, "", SUBL)
        ws.write(r, 6, g["total_h"], SUB)
        r += 1

    hol = result["holiday"]
    if hol["dates"] or hol["total_h"] > 0:
        dates_txt = "\n".join(hol["dates"]) if len(hol["dates"]) > 2 else ", ".join(hol["dates"])
        ws.set_row(r, max(20.0, len(hol["dates"]) * 14))
        ws.write(r, 0, "휴가/공가", HOL)
        ws.write(r, 1, "—", HOL)
        ws.write(r, 2, HOLIDAY_FC, HOL)
        ws.write(r, 3, "휴가/공가", HOL)
        ws.write(r, 4, dates_txt if hol["dates"] else "—", HOLT)
        ws.write(r, 5, "—", HOL)
        ws.write(r, 6, hol["total_h"], HOLH)
        r += 1

    # 월간 총 합계 행
    grand = round(sum(g["total_h"] for g in result["groups"]) + hol["total_h"], 1)
    GRAND  = wb.add_format({"font_name":"맑은 고딕","font_size":12,"bold":True,
                             "border":2,"align":"center","valign":"vcenter",
                             "bg_color":"#1F4E79","font_color":"#FFFFFF",
                             "num_format":"0.0"})
    GRANDL = wb.add_format({"font_name":"맑은 고딕","font_size":12,"bold":True,
                             "border":2,"align":"center","valign":"vcenter",
                             "bg_color":"#1F4E79","font_color":"#FFFFFF"})
    ws.set_row(r, 26)
    ws.merge_range(r, 0, r, 5, f"★  {month_str}  월간 총 업무 시간", GRANDL)
    ws.write(r, 6, grand, GRAND)
    r += 2   # 빈 행

    # ── PC / FC 별 소요시간 요약 시트 ──────────────────────────────
    ws2 = wb.add_worksheet("PC_FC별 소요시간")
    ws2.set_column(0, 0, 12)
    ws2.set_column(1, 1, 12)
    ws2.set_column(2, 2, 12)
    ws2.set_column(3, 3, 34)
    ws2.set_column(4, 4, 10)

    # 타이틀
    T2 = wb.add_format({"font_name":"맑은 고딕","font_size":13,"bold":True,
                         "valign":"vcenter"})
    ws2.set_row(0, 26)
    ws2.merge_range(0, 0, 0, 4, f"PC / FC 별 소요시간 요약  —  {month_str}", T2)

    # 헤더
    H2 = wb.add_format({"font_name":"맑은 고딕","font_size":10,"bold":True,
                         "border":1,"align":"center","valign":"vcenter",
                         "bg_color":"#37474F","font_color":"#FFFFFF"})
    ws2.set_row(1, 24)
    for ci, h in enumerate(["구분","Project Code","FC","업무명(FC 기준)","시간(H)"]):
        ws2.write(1, ci, h, H2)

    summary = _build_summary(result)
    S_ODD  = wb.add_format({"font_name":"맑은 고딕","font_size":10,
                             "border":1,"valign":"vcenter"})
    S_EVEN = wb.add_format({"font_name":"맑은 고딕","font_size":10,
                             "border":1,"valign":"vcenter","bg_color":"#ECEFF1"})
    S_NUM  = wb.add_format({"font_name":"맑은 고딕","font_size":10,
                             "border":1,"align":"center","valign":"vcenter",
                             "num_format":"0.0"})
    S_ENUM = wb.add_format({"font_name":"맑은 고딕","font_size":10,
                             "border":1,"align":"center","valign":"vcenter",
                             "num_format":"0.0","bg_color":"#ECEFF1"})
    S_C    = wb.add_format({"font_name":"맑은 고딕","font_size":10,
                             "border":1,"align":"center","valign":"vcenter"})
    S_EC   = wb.add_format({"font_name":"맑은 고딕","font_size":10,
                             "border":1,"align":"center","valign":"vcenter",
                             "bg_color":"#ECEFF1"})

    sr = 2
    for si, s in enumerate(summary):
        fe  = S_EVEN if si % 2 else S_ODD
        fec = S_EC   if si % 2 else S_C
        fen = S_ENUM if si % 2 else S_NUM
        ws2.set_row(sr, 20)
        ws2.write(sr, 0, s["gubun"],        fec)
        ws2.write(sr, 1, s["project_code"], fec)
        ws2.write(sr, 2, s["func_code"],    fec)
        ws2.write(sr, 3, s["func_name"],    fe)
        ws2.write(sr, 4, s["total_h"],      fen)
        sr += 1

    # 휴가 요약
    if hol["total_h"] > 0:
        ws2.set_row(sr, 20)
        fe = S_EVEN if sr % 2 else S_ODD
        fec = S_EC if sr % 2 else S_C
        fen = S_ENUM if sr % 2 else S_NUM
        ws2.write(sr, 0, "휴가/공가", fec)
        ws2.write(sr, 1, "—", fec)
        ws2.write(sr, 2, HOLIDAY_FC, fec)
        ws2.write(sr, 3, "휴가/공가", fe)
        ws2.write(sr, 4, hol["total_h"], fen)
        sr += 1

    # 합계 행
    STOT  = wb.add_format({"font_name":"맑은 고딕","font_size":11,"bold":True,
                            "border":2,"align":"center","valign":"vcenter",
                            "bg_color":"#1F4E79","font_color":"#FFFFFF"})
    STOTN = wb.add_format({"font_name":"맑은 고딕","font_size":11,"bold":True,
                            "border":2,"align":"center","valign":"vcenter",
                            "bg_color":"#1F4E79","font_color":"#FFFFFF",
                            "num_format":"0.0"})
    ws2.set_row(sr, 24)
    ws2.merge_range(sr, 0, sr, 3, f"월간 총 업무 시간", STOT)
    ws2.write(sr, 4, grand, STOTN)

    wb.close()
