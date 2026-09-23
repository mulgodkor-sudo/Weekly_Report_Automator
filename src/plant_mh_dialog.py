"""
plant_mh_dialog.py
Plant M/H 입력 확인 다이얼로그
[실적]을 (구분 / Project Code / Func. Code) × 날짜별 ST / OT 로 정리해 표시
Plant M/H 시스템에 시간 입력 시 참고용

조회 기간:
  - 기본값은 메인 화면의 Weekly Report 조회 기간
  - 시작일/종료일을 직접 조정해 [불러오기] 하면 해당 기간을 아웃룩에서 다시 읽음
    (몇 주 / 한 달치 몰아서 입력하는 경우 대비. 시작일을 바꿔도 종료일은 유지)
"""
from __future__ import annotations
from collections import OrderedDict
from datetime import datetime
import threading

import tkinter as tk
from tkinter import ttk, messagebox

DAY_KO = ["월", "화", "수", "목", "금", "토", "일"]
_BASE_COLS = ["no", "gubun", "pc", "fc", "desc"]
_BASE_HDRS = ["No.", "구분", "Project Code", "Func. Code", "Description"]


def _to_date(d):
    return d.date() if isinstance(d, datetime) else d


class PlantMHDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, rows: list[dict] | None,
                 start: str = "", end: str = ""):
        """
        rows : 메인 화면에서 이미 불러온 행 (없으면 None → 열리자마자 자동 조회)
        start/end : 기본 조회 기간 (YYYY-MM-DD)
        """
        super().__init__(parent)
        self.title("Plant M/H 입력 확인")
        self.geometry("1200x480")
        self.minsize(700, 320)
        self.resizable(True, True)
        self.configure(bg="#F0F2F5")

        self.v_start = tk.StringVar(value=start)
        self.v_end   = tk.StringVar(value=end)

        self._build()
        if rows:
            self._show(rows)
        else:
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
                       "/ 시작일을 바꿔도 종료일은 그대로 유지",
                  foreground="gray", font=("맑은 고딕", 8)
                  ).grid(row=1, column=0, columnspan=6, sticky="w", padx=6)

        # ── Treeview 영역 ───────────────────────────────────────────
        wrap = tk.Frame(self, bg="#F0F2F5")
        wrap.pack(fill="both", expand=True, padx=12, pady=(6, 0))

        self.tv = ttk.Treeview(wrap, columns=_BASE_COLS, show="headings")
        ysb = ttk.Scrollbar(wrap, orient="vertical",   command=self.tv.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tv.xview)
        self.tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        ysb.pack(side="right",  fill="y")
        xsb.pack(side="bottom", fill="x")
        self.tv.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Treeview", rowheight=22, font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"))
        self.tv.tag_configure("odd",       background="#FFFFFF")
        self.tv.tag_configure("even",      background="#F0F4FF")
        self.tv.tag_configure("total_row", background="#FFE699",
                              font=("맑은 고딕", 9, "bold"))

        # ── 하단 ────────────────────────────────────────────────────
        fr_btn = tk.Frame(self, bg="#F0F2F5")
        fr_btn.pack(fill="x", padx=12, pady=(4, 8))
        self.lbl_stat = tk.Label(fr_btn, text="",
                                 fg="gray", bg="#F0F2F5", font=("맑은 고딕", 8))
        self.lbl_stat.pack(side="left")
        tk.Label(fr_btn,
                 text="※ ST: 평일 08:30~17:30 (하루 최대 8H)  |  OT: 그 외 시간 및 토·일",
                 fg="gray", bg="#F0F2F5", font=("맑은 고딕", 8)
                 ).pack(side="left", padx=(16, 0))
        ttk.Button(fr_btn, text="✅  닫기",
                   command=self.destroy, width=14).pack(side="right")

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

        self.btn_load.config(state="disabled")
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
                rows = build_rows(process_this_week(get_events(start, end)), [])
                self.after(0, lambda: self._on_ok(rows))
            except Exception as e:
                msg = str(e)
                self.after(0, lambda: self._on_err(msg))
            finally:
                if _com_ok:
                    try: import pythoncom; pythoncom.CoUninitialize()
                    except Exception: pass

        threading.Thread(target=task, daemon=True).start()

    def _on_ok(self, rows):
        if not self.winfo_exists():
            return
        self.progress.stop()
        self.btn_load.config(state="normal")
        self._show(rows)

    def _on_err(self, msg):
        if not self.winfo_exists():
            return
        self.progress.stop()
        self.btn_load.config(state="normal")
        self.lbl_stat.config(text=f"오류: {msg[:60]}")
        messagebox.showerror("불러오기 오류", msg, parent=self)

    # ════════════════════════════════════════════════════════════════
    # 표시
    # ════════════════════════════════════════════════════════════════
    def _show(self, rows: list[dict]):
        # [실적] 행만 (hours > 0)
        rows = [r for r in rows
                if r.get("source") == "this_week" and r.get("this_week_h", 0) > 0]

        tv = self.tv
        tv.delete(*tv.get_children())

        if not rows:
            tv["columns"] = _BASE_COLS
            for cid, hdr in zip(_BASE_COLS, _BASE_HDRS):
                tv.heading(cid, text=hdr, anchor="center")
            self.lbl_stat.config(text="해당 기간에 [실적] 데이터가 없습니다.")
            return

        dates = sorted(set(_to_date(r["date"]) for r in rows))

        # ── 집계: (gubun, pc, fc, func_name) → {date: [ST, OT]} ────
        agg: dict[tuple, dict] = OrderedDict()
        for r in rows:
            key = (
                r.get("gubun", ""),
                r.get("project_code", "") or "",
                r.get("func_code", "")    or "",
                r.get("func_name", "")    or "",
            )
            if key not in agg:
                agg[key] = {d: [0.0, 0.0] for d in dates}
            cell = agg[key][_to_date(r["date"])]
            cell[0] = round(cell[0] + r.get("this_week_st", r.get("this_week_h", 0.0)), 1)
            cell[1] = round(cell[1] + r.get("this_week_ot", 0.0), 1)

        # ── 컬럼 정의: 날짜마다 ST / OT 두 칸 ──────────────────────
        date_col_ids, date_col_hdrs = [], []
        for i, d in enumerate(dates):
            lbl = f"{DAY_KO[d.weekday()]}({d.month}/{d.day})"
            date_col_ids  += [f"d{i}st", f"d{i}ot"]
            date_col_hdrs += [f"{lbl} ST", "OT"]

        all_ids  = _BASE_COLS + date_col_ids + ["st_total", "ot_total"]
        all_hdrs = _BASE_HDRS + date_col_hdrs + ["ST TOTAL", "OT TOTAL"]
        col_widths = {"no": 40, "gubun": 85, "pc": 80, "fc": 80, "desc": 180,
                      "st_total": 70, "ot_total": 70}
        for i in range(len(dates)):
            col_widths[f"d{i}st"] = 78
            col_widths[f"d{i}ot"] = 40

        tv["columns"] = all_ids
        for cid, hdr in zip(all_ids, all_hdrs):
            w = col_widths.get(cid, 60)
            tv.heading(cid, text=hdr, anchor="center")
            tv.column(cid, width=w, minwidth=w, stretch=False,
                      anchor="w" if cid == "desc" else "center")

        def fmt(v: float) -> str:
            return f"{v:.1f}" if v > 0 else ""

        # 행 삽입
        total_by_date = {d: [0.0, 0.0] for d in dates}
        grand_st = grand_ot = 0.0

        for ri, (key, hm) in enumerate(agg.items(), 1):
            gubun, pc, fc, desc = key
            row_st = round(sum(hm[d][0] for d in dates), 1)
            row_ot = round(sum(hm[d][1] for d in dates), 1)
            grand_st = round(grand_st + row_st, 1)
            grand_ot = round(grand_ot + row_ot, 1)

            vals = [str(ri), gubun, pc, fc, desc]
            for d in dates:
                st, ot = hm[d]
                vals += [fmt(st), fmt(ot)]
                total_by_date[d][0] = round(total_by_date[d][0] + st, 1)
                total_by_date[d][1] = round(total_by_date[d][1] + ot, 1)
            vals += [fmt(row_st), fmt(row_ot)]

            tag = "even" if ri % 2 == 0 else "odd"
            tv.insert("", "end", values=vals, tags=(tag,))

        # 합계 행
        tot_vals = ["", "합  계", "", "", ""]
        for d in dates:
            tot_vals += [fmt(total_by_date[d][0]), fmt(total_by_date[d][1])]
        tot_vals += [fmt(grand_st), fmt(grand_ot)]
        tv.insert("", "end", values=tot_vals, tags=("total_row",))

        self.lbl_stat.config(
            text=f"{dates[0].month}/{dates[0].day} ~ {dates[-1].month}/{dates[-1].day}  |  "
                 f"ST {grand_st:.1f}H  /  OT {grand_ot:.1f}H")
