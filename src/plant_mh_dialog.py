"""
plant_mh_dialog.py
Plant M/H 입력 확인 다이얼로그

회사 M/H 입력 시스템과 같은 배치로 [실적]을 표시 (plant_mh.py 공통 레이아웃)
  - 왼쪽 고정: No. / 구분 / Project Code / Func. Code / Description
  - 오른쪽 스크롤: 날짜마다 ST·OT (주말 포함 모든 날짜) / ST TOTAL / OT TOTAL
  - OT 칸 연두색, 토·일 머리글 빨간 글씨

조회 기간:
  - 기본값은 메인 화면의 Weekly Report 조회 기간
  - 시작일/종료일을 직접 조정해 [불러오기] (시작일을 바꿔도 종료일은 유지)
  - 종료일이 속한 주의 일요일까지 표시
"""
from __future__ import annotations
from datetime import datetime
import os
import threading

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import plant_mh as pm

FONT      = ("맑은 고딕", 9)
FONT_B    = ("맑은 고딕", 9, "bold")
FIXED_W   = [4, 8, 9, 9, 28]                   # 고정 열 너비(문자)
DAY_W     = 4                                  # ST/OT 칸 너비(문자)
TOTAL_W   = 7


class PlantMHDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, start: str = "", end: str = ""):
        super().__init__(parent)
        self.title("Plant M/H 입력 확인")
        self.geometry("1280x520")
        self.minsize(760, 340)
        self.resizable(True, True)
        self.configure(bg="#F0F2F5")

        self.v_start = tk.StringVar(value=start)
        self.v_end   = tk.StringVar(value=end)
        self._items: list[dict] = []
        self._dates: list = []

        self._build()
        self._load()

    # ════════════════════════════════════════════════════════════════
    # UI 빌드
    # ════════════════════════════════════════════════════════════════
    def _build(self):
        # ── 타이틀 바 ──────────────────────────────────────────────
        bar = tk.Frame(self, bg="#1F4E79", height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="  📊  Plant M/H 입력 확인  ([실적] 기준)",
                 font=("맑은 고딕", 11, "bold"),
                 fg="white", bg="#1F4E79").pack(side="left", padx=12, pady=6)

        # ── 조회 기간 ──────────────────────────────────────────────
        fr_d = ttk.LabelFrame(self, text="  📅  조회 기간", padding=(10, 5))
        fr_d.pack(fill="x", padx=12, pady=(8, 0))

        P = dict(padx=6, pady=3)
        ttk.Label(fr_d, text="시작일 :").grid(row=0, column=0, sticky="e", **P)
        ttk.Entry(fr_d, textvariable=self.v_start, width=14,
                  font=("맑은 고딕", 10)).grid(row=0, column=1, sticky="w", **P)
        ttk.Label(fr_d, text=" ~ 종료일 :").grid(row=0, column=2, **P)
        ttk.Entry(fr_d, textvariable=self.v_end, width=14,
                  font=("맑은 고딕", 10)).grid(row=0, column=3, sticky="w", **P)

        self.btn_load = tk.Button(
            fr_d, text="🔍  불러오기", command=self._load,
            font=("맑은 고딕", 10, "bold"),
            bg="#1F4E79", fg="white",
            activebackground="#2E75B6", activeforeground="white",
            relief="flat", padx=14, pady=4, cursor="hand2")
        self.btn_load.grid(row=0, column=4, padx=(14, 4), pady=3)

        self.progress = ttk.Progressbar(fr_d, mode="indeterminate", length=110)
        self.progress.grid(row=0, column=5, padx=4)

        ttk.Label(fr_d,
                  text="* 기본값은 Weekly Report 조회 기간 / YYYY-MM-DD 형식으로 수정 가능 "
                       "/ 시작일을 바꿔도 종료일은 유지 / 종료일이 속한 주의 일요일까지 표시",
                  foreground="gray", font=("맑은 고딕", 8)
                  ).grid(row=1, column=0, columnspan=6, sticky="w", padx=6)

        # ── 표 영역: 왼쪽 고정 + 오른쪽 가로 스크롤 ─────────────────
        area = tk.Frame(self, bg="#F0F2F5")
        area.pack(fill="both", expand=True, padx=12, pady=(6, 0))
        area.rowconfigure(0, weight=1)
        area.columnconfigure(1, weight=1)

        self.cv_left  = tk.Canvas(area, bg="white", highlightthickness=0)
        self.cv_right = tk.Canvas(area, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(area, orient="vertical",   command=self._yview)
        hsb = ttk.Scrollbar(area, orient="horizontal", command=self.cv_right.xview)
        self.cv_right.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.cv_left.grid(row=0, column=0, sticky="ns")
        self.cv_right.grid(row=0, column=1, sticky="nsew")
        vsb.grid(row=0, column=2, sticky="ns")
        hsb.grid(row=1, column=1, sticky="ew")

        # 격자선: 프레임 배경을 테두리색으로 두고 셀 사이 1px 간격
        self.fr_left  = tk.Frame(self.cv_left,  bg=pm.C_GRID)
        self.fr_right = tk.Frame(self.cv_right, bg=pm.C_GRID)
        self.cv_left.create_window((0, 0), window=self.fr_left, anchor="nw")
        self.cv_right.create_window((0, 0), window=self.fr_right, anchor="nw")
        self.fr_left.bind("<Configure>", lambda e: self._on_left_resize())
        self.fr_right.bind("<Configure>", lambda e: self.cv_right.configure(
            scrollregion=self.cv_right.bbox("all")))

        for w in (self.cv_left, self.cv_right):
            w.bind("<MouseWheel>", self._on_wheel)
            w.bind("<Shift-MouseWheel>", self._on_shift_wheel)

        # ── 하단 ────────────────────────────────────────────────────
        fr_btn = tk.Frame(self, bg="#F0F2F5")
        fr_btn.pack(fill="x", padx=12, pady=(4, 8))
        self.lbl_stat = tk.Label(fr_btn, text="",
                                 fg="#1F4E79", bg="#F0F2F5", font=FONT_B)
        self.lbl_stat.pack(side="left")
        tk.Label(fr_btn, text=pm.NOTE,
                 fg="gray", bg="#F0F2F5", font=("맑은 고딕", 8)
                 ).pack(side="left", padx=(16, 0))

        ttk.Button(fr_btn, text="✅  닫기",
                   command=self.destroy, width=12).pack(side="right")
        self.btn_excel = tk.Button(
            fr_btn, text="📄  Excel 저장", command=self._save_excel,
            font=FONT_B, bg="#1F5E20", fg="white",
            activebackground="#2E7D32", activeforeground="white",
            disabledforeground="#AAAAAA",
            relief="flat", padx=14, pady=4, cursor="hand2", state="disabled")
        self.btn_excel.pack(side="right", padx=(0, 8))

    # ── 스크롤 ───────────────────────────────────────────────────────
    def _yview(self, *args):
        self.cv_left.yview(*args)
        self.cv_right.yview(*args)

    def _on_wheel(self, e):
        self._yview("scroll", int(-e.delta / 120), "units")

    def _on_shift_wheel(self, e):
        self.cv_right.xview("scroll", int(-e.delta / 120), "units")

    def _on_left_resize(self):
        self.cv_left.configure(scrollregion=self.cv_left.bbox("all"),
                               width=self.fr_left.winfo_reqwidth())

    # ════════════════════════════════════════════════════════════════
    # 조회
    # ════════════════════════════════════════════════════════════════
    def _load(self):
        try:
            start = datetime.strptime(self.v_start.get().strip(), "%Y-%m-%d")
            end   = datetime.strptime(self.v_end.get().strip(),   "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("날짜 오류", "YYYY-MM-DD 형식으로 입력하세요.", parent=self)
            return
        if start > end:
            messagebox.showerror("날짜 오류", "시작일이 종료일보다 늦습니다.", parent=self)
            return

        # 종료일이 속한 주의 일요일까지 (주말 OT 포함)
        show_end = datetime.combine(pm.week_sunday(end.date()), datetime.min.time())
        dates = pm.date_list(start.date(), show_end.date())

        self.btn_load.config(state="disabled")
        self.btn_excel.config(state="disabled")
        self.progress.start()
        self.lbl_stat.config(text="아웃룩에서 일정을 불러오는 중...")

        def task():
            _com_ok = False
            try:
                import pythoncom; pythoncom.CoInitialize(); _com_ok = True
            except Exception:
                pass
            try:
                from outlook_reader  import get_events
                from event_processor import process_this_week, build_rows
                rows = build_rows(process_this_week(get_events(start, show_end)), [])
                self.after(0, lambda: self._on_ok(pm.aggregate(rows), dates))
            except Exception as e:
                msg = str(e)
                self.after(0, lambda: self._on_err(msg))
            finally:
                if _com_ok:
                    try: import pythoncom; pythoncom.CoUninitialize()
                    except Exception: pass

        threading.Thread(target=task, daemon=True).start()

    def _on_ok(self, items, dates):
        if not self.winfo_exists():
            return
        self.progress.stop()
        self.btn_load.config(state="normal")
        self._items, self._dates = items, dates
        self._render()
        st = round(sum(it["st_total"] for it in items), 1)
        ot = round(sum(it["ot_total"] for it in items), 1)
        d0, d1 = dates[0], dates[-1]
        self.lbl_stat.config(
            text=f"{d0.month}/{d0.day} ~ {d1.month}/{d1.day}   |   "
                 f"ST {pm.fmt_h(st)}H  /  OT {pm.fmt_h(ot)}H"
                 + ("" if items else "   (해당 기간 [실적] 없음)"))
        self.btn_excel.config(state="normal" if items else "disabled")

    def _on_err(self, msg):
        if not self.winfo_exists():
            return
        self.progress.stop()
        self.btn_load.config(state="normal")
        self.lbl_stat.config(text=f"오류: {msg[:60]}")
        messagebox.showerror("불러오기 오류", msg, parent=self)

    # ════════════════════════════════════════════════════════════════
    # 표 그리기
    # ════════════════════════════════════════════════════════════════
    def _cell(self, parent, r, c, text="", w=DAY_W, bg="white", fg="black",
              bold=False, anchor="center", rs=1, cs=1):
        lbl = tk.Label(parent, text=text, width=w, bg=bg, fg=fg,
                       font=FONT_B if bold else FONT, anchor=anchor,
                       padx=2, pady=0, bd=0)
        lbl.grid(row=r, column=c, rowspan=rs, columnspan=cs, sticky="nsew",
                 padx=(0, 1), pady=(0, 1))
        lbl.bind("<MouseWheel>", self._on_wheel)
        lbl.bind("<Shift-MouseWheel>", self._on_shift_wheel)
        return lbl

    def _render(self):
        items, dates = self._items, self._dates
        L, R = self.fr_left, self.fr_right
        for fr in (L, R):
            for w in fr.winfo_children():
                w.destroy()

        # 모든 행 높이를 고정해 왼쪽/오른쪽 표 행을 맞춤 (화면 배율 대응: 글꼴 기준)
        from tkinter import font as tkfont
        row_h  = tkfont.Font(self, font=FONT_B).metrics("linespace") + 6
        n_rows = 4 + len(items) + 1
        for fr in (L, R):
            for r in range(n_rows + 1):
                fr.rowconfigure(r, minsize=row_h + 1)

        H = pm.C_HDR_BG
        # ── 왼쪽 고정 머리글 (4단 병합) ──
        for ci, (hd, w) in enumerate(zip(["No.", "구분", "Project\nCode",
                                          "Func. Code", "Description"], FIXED_W)):
            self._cell(L, 0, ci, hd, w=w, bg=H, bold=True, rs=4)

        # ── 오른쪽 날짜 머리글 ──
        for lbl, i0, i1 in pm.month_groups(dates):
            self._cell(R, 0, 2 * i0, lbl, bg=H, bold=True, cs=2 * (i1 - i0 + 1))
        for i, d in enumerate(dates):
            fg = pm.C_WEEKEND if d.weekday() >= 5 else "black"
            self._cell(R, 1, 2 * i, pm.DAY_KO[d.weekday()], bg=H, fg=fg, bold=True, cs=2)
            self._cell(R, 2, 2 * i, str(d.day),             bg=H, fg=fg, bold=True, cs=2)
            self._cell(R, 3, 2 * i,     "ST", bg=H, fg=fg, bold=True)
            self._cell(R, 3, 2 * i + 1, "OT", bg=H, fg=fg, bold=True)
        c_st = 2 * len(dates)
        self._cell(R, 0, c_st,     "ST\nTOTAL", w=TOTAL_W, bg=H, bold=True, rs=4)
        self._cell(R, 0, c_st + 1, "OT\nTOTAL", w=TOTAL_W, bg=H, bold=True, rs=4)

        # ── 데이터 ──
        for ri, it in enumerate(items):
            r = 4 + ri
            for ci, (v, w) in enumerate(zip(
                    [str(ri + 1), it["gubun"], it["pc"], it["fc"], it["desc"]], FIXED_W)):
                self._cell(L, r, ci, v, w=w, anchor="w" if ci == 4 else "center")
            for i, d in enumerate(dates):
                has = d in it["cells"]          # 입력한 날은 0도 표시 (시스템과 동일)
                st, ot = it["cells"].get(d, (0.0, 0.0))
                self._cell(R, r, 2 * i,     pm.fmt_h(st) if has else "")
                self._cell(R, r, 2 * i + 1, pm.fmt_h(ot) if has else "", bg=pm.C_OT_BG)
            self._cell(R, r, c_st,     pm.fmt_h(it["st_total"]), w=TOTAL_W)
            self._cell(R, r, c_st + 1, pm.fmt_h(it["ot_total"]), w=TOTAL_W)

        # ── 합계 ──
        tr  = 4 + len(items)
        T   = pm.C_TOTAL_BG
        tot = pm.day_totals(items, dates)
        self._cell(L, tr, 0, "합  계", w=1, bg=T, bold=True, cs=len(FIXED_W))
        for i, d in enumerate(dates):
            st, ot = tot[d]
            self._cell(R, tr, 2 * i,     pm.fmt_h(st) if st else "", bg=T, bold=True)
            self._cell(R, tr, 2 * i + 1, pm.fmt_h(ot) if ot else "", bg=T, bold=True)
        self._cell(R, tr, c_st,     pm.fmt_h(sum(it["st_total"] for it in items)),
                   w=TOTAL_W, bg=T, bold=True)
        self._cell(R, tr, c_st + 1, pm.fmt_h(sum(it["ot_total"] for it in items)),
                   w=TOTAL_W, bg=T, bold=True)

        self.cv_left.yview_moveto(0)
        self.cv_right.yview_moveto(0)
        self.cv_right.xview_moveto(0)

    # ════════════════════════════════════════════════════════════════
    # Excel 저장
    # ════════════════════════════════════════════════════════════════
    def _save_excel(self):
        if not self._items:
            return
        d0, d1 = self._dates[0], self._dates[-1]
        path = filedialog.asksaveasfilename(
            title="Plant M/H 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile=f"PlantMH_{d0:%Y%m%d}-{d1:%Y%m%d}.xlsx",
            parent=self,
        )
        if not path:
            return
        try:
            pm.save_workbook(path, self._items, self._dates)
        except PermissionError:
            messagebox.showerror("파일 열림",
                                 f"파일이 열려 있습니다. Excel을 닫고 다시 시도하세요.\n{path}",
                                 parent=self)
            return
        except Exception as e:
            messagebox.showerror("저장 오류", str(e), parent=self)
            return
        if messagebox.askyesno("저장 완료",
                               f"저장되었습니다.\n{path}\n\n바로 여시겠습니까?", parent=self):
            try:
                os.startfile(path)
            except Exception:
                pass
