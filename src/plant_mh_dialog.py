"""
plant_mh_dialog.py
Plant M/H 입력 확인 다이얼로그
이번주 [실적]을 (구분 / Project Code / Func. Code) × 날짜별로 정리해 표시
Plant M/H 시스템에 시간 입력 시 참고용
"""
from __future__ import annotations
from collections import OrderedDict
from datetime import datetime, date as date_type

import tkinter as tk
from tkinter import ttk

DAY_KO = ["월", "화", "수", "목", "금", "토", "일"]


class PlantMHDialog(tk.Toplevel):
    def __init__(self, parent: tk.Tk, rows: list[dict]):
        super().__init__(parent)
        self.title("Plant M/H 입력 확인")
        self.geometry("920x420")
        self.minsize(700, 300)
        self.resizable(True, True)
        self.configure(bg="#F0F2F5")

        # 이번주 실적 행만 (hours > 0)
        this_rows = [
            r for r in rows
            if r.get("source") == "this_week" and r.get("this_week_h", 0) > 0
        ]
        self._build(this_rows)

    # ════════════════════════════════════════════════════════════════
    # UI 빌드
    # ════════════════════════════════════════════════════════════════
    def _build(self, rows: list[dict]):
        # ── 타이틀 바 ──────────────────────────────────────────────
        bar = tk.Frame(self, bg="#1F4E79", height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="  📊  Plant M/H 입력 확인  (이번주 [실적] 기준)",
                 font=("맑은 고딕", 11, "bold"),
                 fg="white", bg="#1F4E79").pack(side="left", padx=12, pady=6)

        # ── 데이터 없음 ────────────────────────────────────────────
        if not rows:
            tk.Label(self,
                     text="이번주 [실적] 데이터가 없습니다.\n먼저 주별 불러오기를 실행하세요.",
                     font=("맑은 고딕", 11), bg="#F0F2F5", fg="gray"
                     ).pack(expand=True)
            ttk.Button(self, text="✅  닫기", command=self.destroy, width=14
                       ).pack(pady=10)
            return

        # ── 날짜 목록 수집 ──────────────────────────────────────────
        def to_date(d):
            return d.date() if isinstance(d, datetime) else d

        dates = sorted(set(to_date(r["date"]) for r in rows))

        # ── 집계: (gubun, pc, fc, func_name) → {date: hours} ───────
        agg: dict[tuple, dict] = OrderedDict()
        for r in rows:
            key = (
                r.get("gubun", ""),
                r.get("project_code", "") or "",
                r.get("func_code", "")    or "",
                r.get("func_name", "")    or "",
            )
            dt = to_date(r["date"])
            if key not in agg:
                agg[key] = {d: 0.0 for d in dates}
            agg[key][dt] = round(agg[key].get(dt, 0.0) + r.get("this_week_h", 0.0), 1)

        # ── 컬럼 정의 ──────────────────────────────────────────────
        date_col_ids  = [f"d{i}" for i in range(len(dates))]
        date_col_hdrs = [f"{DAY_KO[d.weekday()]}({d.month}/{d.day})" for d in dates]

        all_ids   = ["no", "gubun", "pc", "fc", "desc"] + date_col_ids + ["total"]
        all_hdrs  = ["No.", "구분", "Project Code", "Func. Code", "Description"
                     ] + date_col_hdrs + ["합계"]
        col_widths = {"no": 40, "gubun": 85, "pc": 80, "fc": 80, "desc": 180, "total": 55}

        # ── Treeview 영역 ───────────────────────────────────────────
        wrap = tk.Frame(self, bg="#F0F2F5")
        wrap.pack(fill="both", expand=True, padx=12, pady=(6, 0))

        tv = ttk.Treeview(wrap, columns=all_ids, show="headings")
        ysb = ttk.Scrollbar(wrap, orient="vertical",   command=tv.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        ysb.pack(side="right",  fill="y")
        xsb.pack(side="bottom", fill="x")
        tv.pack(fill="both", expand=True)

        # 헤더 + 너비 설정
        for cid, hdr in zip(all_ids, all_hdrs):
            w = col_widths.get(cid, 60)
            tv.heading(cid, text=hdr, anchor="center")
            anc = "w" if cid == "desc" else "center"
            tv.column(cid, width=w, minwidth=w, anchor=anc)

        # 행 삽입
        total_by_date = {d: 0.0 for d in dates}
        grand_total   = 0.0

        for ri, (key, hm) in enumerate(agg.items(), 1):
            gubun, pc, fc, desc = key
            row_total = round(sum(hm.get(d, 0.0) for d in dates), 1)
            grand_total = round(grand_total + row_total, 1)

            vals = [str(ri), gubun, pc, fc, desc]
            for d in dates:
                v = hm.get(d, 0.0)
                vals.append(f"{v:.1f}" if v > 0 else "")
                total_by_date[d] = round(total_by_date[d] + v, 1)
            vals.append(f"{row_total:.1f}" if row_total > 0 else "")

            tag = "even" if ri % 2 == 0 else "odd"
            tv.insert("", "end", values=vals, tags=(tag,))

        # 합계 행
        tot_vals = ["", "합  계", "", "", ""]
        for d in dates:
            v = total_by_date[d]
            tot_vals.append(f"{v:.1f}" if v > 0 else "")
        tot_vals.append(f"{grand_total:.1f}" if grand_total > 0 else "")
        tv.insert("", "end", values=tot_vals, tags=("total_row",))

        # 행 색상 스타일
        style = ttk.Style()
        style.configure("Treeview", rowheight=22, font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"))
        tv.tag_configure("odd",       background="#FFFFFF")
        tv.tag_configure("even",      background="#F0F4FF")
        tv.tag_configure("total_row", background="#FFE699",
                         font=("맑은 고딕", 9, "bold"))

        # ── 하단 버튼 ──────────────────────────────────────────────
        fr_btn = tk.Frame(self, bg="#F0F2F5")
        fr_btn.pack(fill="x", padx=12, pady=(4, 8))
        tk.Label(fr_btn,
                 text="※ ST(정규시간) 기준  |  OT는 별도 확인하세요",
                 fg="gray", bg="#F0F2F5", font=("맑은 고딕", 8)
                 ).pack(side="left")
        ttk.Button(fr_btn, text="✅  닫기",
                   command=self.destroy, width=14).pack(side="right")
