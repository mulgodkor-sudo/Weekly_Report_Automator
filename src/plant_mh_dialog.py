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
import traceback

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import plant_mh as pm
from outlook_reader import FLEX_STARTS, hm_label

# 시차출퇴근 선택지: (표시, 출근시각) — None = 자동(그날 일정으로 판단)
FLEX_AUTO    = "자동 (일정 기준)"
FLEX_CHOICES = [(FLEX_AUTO, None)] + [
    (f"{h:02d}:{m:02d}~{h + 9:02d}:{m:02d}", (h, m)) for h, m in FLEX_STARTS]
C_MANUAL     = "#C55A11"   # 날짜별로 직접 바꾼 출근시각 표시
C_PICK_BG    = "#FFFFFF"   # 클릭해서 바꿀 수 있는 출근시각 칸 바탕
C_PICK_HOVER = "#DDEBF7"   # 마우스를 올렸을 때
C_TOGGLE_BG    = "#E4E4E4"   # 위쪽 접기/펼치기 줄
C_TOGGLE_HOVER = "#C9D9EE"

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
        self._starts: dict = {}          # {날짜: 출근시각} 시차출퇴근 판단 결과
        self._events: list | None = None  # 마지막으로 불러온 아웃룩 일정 (재계산용)
        self._day_override: dict = {}    # {날짜: (h, m)} 날짜별로 직접 고른 출근시각

        # 시차출퇴근 기본 선택값은 사용자별로 기억 (문서 폴더 설정)
        from config import load_config
        saved = load_config().get("mh_flex_start", "auto")
        label = next((l for l, hm in FLEX_CHOICES if hm_label(hm) == saved), FLEX_AUTO)
        self.v_flex  = tk.StringVar(value=label)
        # 항상 위에 표시 (M/H 입력 웹페이지를 보며 입력할 때 창이 뒤로 숨지 않게) — 기본 켜짐, 기억
        self.v_topmost = tk.BooleanVar(value=bool(load_config().get("mh_topmost", True)))

        self._build()
        self._apply_topmost()
        self._load()

    # ════════════════════════════════════════════════════════════════
    # UI 빌드
    # ════════════════════════════════════════════════════════════════
    def _build(self):
        # ── 타이틀 바 ──────────────────────────────────────────────
        bar = tk.Frame(self, bg="#1F4E79", height=40)
        bar.pack(fill="x")
        self._bar = bar
        bar.pack_propagate(False)
        tk.Label(bar, text="  📊  Plant M/H 입력 확인  ([실적] 기준)",
                 font=("맑은 고딕", 11, "bold"),
                 fg="white", bg="#1F4E79").pack(side="left", padx=12, pady=6)
        tk.Checkbutton(bar, text="📌 항상 위에 표시", variable=self.v_topmost,
                       command=self._on_topmost_change,
                       font=("맑은 고딕", 10, "bold"), fg="white", bg="#1F4E79",
                       activeforeground="white", activebackground="#2E75B6",
                       selectcolor="#1F4E79", cursor="hand2", bd=0, highlightthickness=0
                       ).pack(side="right", padx=14)

        # ── 조회 기간 ──────────────────────────────────────────────
        fr_d = ttk.LabelFrame(self, text="  📅  조회 기간", padding=(10, 5))
        fr_d.pack(fill="x", padx=12, pady=(8, 0))
        self._fr_d = fr_d

        P = dict(padx=6, pady=3)
        ttk.Label(fr_d, text="시작일 :").grid(row=0, column=0, sticky="e", **P)
        self.ent_start = ttk.Entry(fr_d, textvariable=self.v_start, width=14,
                  font=("맑은 고딕", 10))
        self.ent_start.grid(row=0, column=1, sticky="w", **P)
        ttk.Label(fr_d, text=" ~ 종료일 :").grid(row=0, column=2, **P)
        self.ent_end = ttk.Entry(fr_d, textvariable=self.v_end, width=14,
                  font=("맑은 고딕", 10))
        self.ent_end.grid(row=0, column=3, sticky="w", **P)

        self.btn_load = tk.Button(
            fr_d, text="🔍  불러오기", command=lambda: self._load(collapse=True),
            font=("맑은 고딕", 10, "bold"),
            bg="#1F4E79", fg="white",
            activebackground="#2E75B6", activeforeground="white",
            relief="flat", padx=14, pady=4, cursor="hand2")
        self.btn_load.grid(row=0, column=4, padx=(14, 4), pady=3)

        self.progress = ttk.Progressbar(fr_d, mode="indeterminate", length=110)
        self.progress.grid(row=0, column=5, padx=4)

        ttk.Label(fr_d,
                  text="* 기본값은 전월 25일 ~ 오늘 / 20261001 처럼 숫자만 입력해도 됨 "
                       "/ 시작일을 바꿔도 종료일은 유지 / 종료일이 속한 주의 일요일까지 표시",
                  foreground="gray", font=("맑은 고딕", 8)
                  ).grid(row=1, column=0, columnspan=6, sticky="w", padx=6)

        # ── 시차출퇴근 선택 ──
        fr_f = ttk.Frame(fr_d)
        fr_f.grid(row=2, column=0, columnspan=6, sticky="w", padx=6, pady=(4, 0))
        ttk.Label(fr_f, text="시차출퇴근 :").pack(side="left")
        cb = ttk.Combobox(fr_f, textvariable=self.v_flex, state="readonly", width=17,
                          values=[l for l, _ in FLEX_CHOICES], font=("맑은 고딕", 10))
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>", self._on_flex_change)
        ttk.Label(fr_f,
                  text="자동 = 그날 일정으로 출근시각 판단  |  표의 날짜 아래 출근시각(▾)을 클릭하면 그날만 변경",
                  foreground="gray", font=("맑은 고딕", 8)).pack(side="left", padx=8)

        # ── 표 영역: 왼쪽 고정 + 오른쪽 가로 스크롤 ─────────────────
        # ── 위쪽(제목 줄 + 조회 기간) 접기/펼치기 줄 ─────────────────
        #   처음부터 보임 (펼친 상태, ▲ 를 누르면 바로 접기 가능)
        #   [불러오기] 로 조회가 끝나면 자동으로 접어 날짜·표만 보이게 함
        #   접힌 뒤에는 가운데 ▼ 를 눌러 다시 펼치고, 펼친 상태에서는 ▲ 로 다시 접음
        self._collapsed = False
        self._collapse_pending = False
        self.fr_toggle = tk.Frame(self, bg=C_TOGGLE_BG, height=16, cursor="hand2")
        self.fr_toggle.pack_propagate(False)
        self.lbl_toggle = tk.Label(self.fr_toggle, text="▲", bg=C_TOGGLE_BG, fg="#404040",
                                   font=("맑은 고딕", 8), cursor="hand2")
        self.lbl_toggle.pack(expand=True)
        for w in (self.fr_toggle, self.lbl_toggle):
            w.bind("<Button-1>", lambda e: self._toggle_header())
            w.bind("<Enter>", lambda e: self._toggle_hover(True))
            w.bind("<Leave>", lambda e: self._toggle_hover(False))
        self.fr_toggle.pack(fill="x", pady=(4, 0))     # 처음부터 표시 (▲ = 접기)

        area = tk.Frame(self, bg="#F0F2F5")
        area.pack(fill="both", expand=True, padx=12, pady=(6, 0))
        self._area = area
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
        # 버튼을 먼저 배치 → 상태 문구가 길어져도 버튼이 화면 밖으로 밀리지 않음
        ttk.Button(fr_btn, text="✅  닫기",
                   command=self.destroy, width=12).pack(side="right")
        self.btn_excel = tk.Button(
            fr_btn, text="📄  Excel 저장", command=self._save_excel,
            font=FONT_B, bg="#1F5E20", fg="white",
            activebackground="#2E7D32", activeforeground="white",
            disabledforeground="#AAAAAA",
            relief="flat", padx=14, pady=4, cursor="hand2", state="disabled")
        self.btn_excel.pack(side="right", padx=(0, 8))

        left = tk.Frame(fr_btn, bg="#F0F2F5")
        left.pack(side="left", fill="x", expand=True)
        self.lbl_stat = tk.Label(left, text="", anchor="w",
                                 fg="#1F4E79", bg="#F0F2F5", font=FONT_B)
        self.lbl_stat.pack(side="top", fill="x")
        tk.Label(left, text=pm.NOTE, anchor="w",
                 fg="gray", bg="#F0F2F5", font=("맑은 고딕", 8)
                 ).pack(side="top", fill="x")

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
    # ── 위쪽 접기/펼치기 ─────────────────────────────────────────────
    def _toggle_hover(self, on: bool):
        bg = C_TOGGLE_HOVER if on else C_TOGGLE_BG
        self.fr_toggle.config(bg=bg)
        self.lbl_toggle.config(bg=bg)

    def _set_header(self, collapsed: bool):
        """collapsed=True: 제목 줄·조회 기간 숨김 (▼ 표시) / False: 다시 보임 (▲ 표시)"""
        self._collapsed = collapsed
        if not self.fr_toggle.winfo_manager():
            self.fr_toggle.pack(fill="x", before=self._area)
        if collapsed:
            self._bar.pack_forget()
            self._fr_d.pack_forget()
            self.lbl_toggle.config(text="▼")
            self.fr_toggle.pack_configure(pady=(0, 0))
        else:
            self._bar.pack(fill="x", before=self.fr_toggle)
            self._fr_d.pack(fill="x", padx=12, pady=(8, 0), before=self.fr_toggle)
            self.lbl_toggle.config(text="▲")
            self.fr_toggle.pack_configure(pady=(4, 0))

    def _toggle_header(self):
        self._set_header(not self._collapsed)

    def _load(self, collapse: bool = False):
        """collapse=True: [불러오기] 버튼으로 조회 — 성공하면 위쪽을 접는다 (창을 처음 열 때는 접지 않음)"""
        # 숫자만 입력해도 인식 (20261001 / 261001 / 1001 / 2026-10-01), 잘못된 날짜는 다시 입력,
        # 지난달 이전 날짜는 확인 (dateinput.py)
        import dateinput
        got = dateinput.read_fields([(self.v_start, self.ent_start, "시작일"),
                                     (self.v_end,   self.ent_end,   "종료일")], parent=self)
        if got is None:
            return
        start, end = got
        if start > end:
            messagebox.showerror("날짜 오류", "시작일이 종료일보다 늦습니다.", parent=self)
            return

        # 종료일이 속한 주의 일요일까지 (주말 OT 포함)
        show_end = datetime.combine(pm.week_sunday(end.date()), datetime.min.time())
        dates = pm.date_list(start.date(), show_end.date())

        self._collapse_pending = collapse
        self.btn_load.config(state="disabled")
        self.btn_excel.config(state="disabled")
        self.progress.start()
        self.lbl_stat.config(text="아웃룩에서 일정을 불러오는 중...")
        fixed, per_day = self._flex_fixed(), dict(self._day_override)   # 화면 값은 메인 스레드에서 읽음

        def task():
            _com_ok = False
            try:
                import pythoncom; pythoncom.CoInitialize(); _com_ok = True
            except Exception:
                pass
            try:
                from outlook_reader import get_events
                events = get_events(start, show_end)
                items, starts = self._compute(events, fixed, per_day)
                self.after(0, lambda: self._on_ok(items, dates, starts, events))
            except Exception as e:
                msg = f"{e}\n\n{traceback.format_exc()}"
                self.after(0, lambda: self._on_err(msg))
            finally:
                if _com_ok:
                    try: import pythoncom; pythoncom.CoUninitialize()
                    except Exception: pass

        threading.Thread(target=task, daemon=True).start()

    # ── 시차출퇴근 ────────────────────────────────────────────────────
    def _flex_fixed(self):
        """선택 상자 값 → 출근시각 (h, m), 자동이면 None"""
        return dict(FLEX_CHOICES).get(self.v_flex.get())

    @staticmethod
    def _compute(events, fixed, per_day):
        """아웃룩 일정 → (표 항목, {날짜: 출근시각}) — 시차출퇴근 설정 반영"""
        from outlook_reader  import assign_st_ot
        from event_processor import process_this_week, build_rows
        assign_st_ot(events, fixed=fixed, per_day=per_day)
        rows = build_rows(process_this_week(events), [])
        return pm.aggregate(rows), pm.day_starts(rows)

    def _recalc(self):
        """아웃룩을 다시 읽지 않고 시차출퇴근 설정만 바꿔 다시 계산"""
        if self._events is None:
            return
        try:
            items, starts = self._compute(self._events, self._flex_fixed(),
                                          dict(self._day_override))
        except Exception as e:
            self._on_err(f"재계산 오류: {e}\n\n{traceback.format_exc()}")
            return
        self._on_ok(items, self._dates, starts, self._events)

    # ── 항상 위에 표시 ────────────────────────────────────────────────
    def _apply_topmost(self):
        self.attributes("-topmost", bool(self.v_topmost.get()))

    def _on_topmost_change(self):
        self._apply_topmost()
        from config import load_config, save_config
        cfg = load_config()
        cfg["mh_topmost"] = bool(self.v_topmost.get())
        save_config(cfg)

    def _on_flex_change(self, _event=None):
        from config import load_config, save_config
        cfg = load_config()
        cfg["mh_flex_start"] = hm_label(self._flex_fixed()) or "auto"
        save_config(cfg)
        self._day_override.clear()       # 전체 설정을 바꾸면 날짜별 변경은 초기화
        self._recalc()

    def _flex_menu(self, event, d):
        """날짜 아래 출근시각 클릭 → 그날만 출근시각 변경"""
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=f"{d.month}/{d.day} 출근시각", state="disabled")
        m.add_separator()
        m.add_command(label=f"기본 설정 따름 ({self.v_flex.get()})",
                      command=lambda: self._set_day(d, None))
        for label, hm in FLEX_CHOICES[1:]:
            mark = "  ✔" if self._day_override.get(d) == hm else ""
            m.add_command(label=label + mark, command=lambda hm=hm: self._set_day(d, hm))
        m.tk_popup(event.x_root, event.y_root)

    def _set_day(self, d, hm):
        if hm is None:
            self._day_override.pop(d, None)
        else:
            self._day_override[d] = hm
        self._recalc()

    def _on_ok(self, items, dates, starts=None, events=None):
        if not self.winfo_exists():
            return
        self.progress.stop()
        self.btn_load.config(state="normal")
        self._items, self._dates, self._starts = items, dates, starts or {}
        if events is not None:
            self._events = events
        if self._collapse_pending:             # [불러오기] 조회 성공 → 위쪽 접기
            self._collapse_pending = False
            self._set_header(True)
        try:
            self._render()
        except Exception as e:
            # 화면 오류가 조용히 묻히지 않도록 반드시 표시
            self._items = []
            self._on_err(f"표시 중 오류: {e}\n\n{traceback.format_exc()}")
            return
        st = round(sum(it["st_total"] for it in items), 1)
        ot = round(sum(it["ot_total"] for it in items), 1)
        d0, d1 = dates[0], dates[-1]
        # Description 은 FC 엑셀(업무명)에서 채움 → 못 읽었으면 이유를 함께 표시
        from config import get_fc_status
        fc_src, fc_why = get_fc_status()
        fc_note = ""
        if fc_src == "cache":
            from config import get_fc_cache_info
            fc_note = f"   ⚠ ECM 엑셀 미연결 - 마지막 사본({get_fc_cache_info()}) 사용"
        elif fc_src == "default":
            fc_note = "   ⚠ FC 엑셀 미사용(내장 목록) - " + fc_why[:70]
        elif fc_src == "none":
            fc_note = "   ⚠ Description 없음 - " + (fc_why or "FC 엑셀 파일 경로를 설정하세요")[:70]
        self.lbl_stat.config(
            text=f"{d0.month}/{d0.day} ~ {d1.month}/{d1.day}   |   "
                 f"ST {pm.fmt_h(st)}H  /  OT {pm.fmt_h(ot)}H   |   시차출퇴근: {self.v_flex.get()}"
                 + (f" (날짜별 변경 {len(self._day_override)}일)" if self._day_override else "")
                 + ("" if items else "   (해당 기간 [실적] 없음)") + fc_note,
            fg="#C00000" if fc_note else "#1F4E79")
        self.btn_excel.config(state="normal" if items else "disabled")

    def _on_err(self, msg):
        if not self.winfo_exists():
            return
        self._collapse_pending = False         # 조회 실패 시 위쪽은 접지 않음
        self.progress.stop()
        self.btn_load.config(state="normal")
        self.lbl_stat.config(text=f"오류: {msg.splitlines()[0][:60]}")
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

        # 모든 행 높이를 고정해 왼쪽/오른쪽 표 행을 맞춤 (화면 배율 대응: 실제 글자 높이 기준)
        # ※ tkinter.font 모듈은 쓰지 않는다 — 이전 버전 exe 에는 번들되어 있지 않아
        #   패치로 실행하면 ImportError 로 표가 안 그려졌음
        probe = tk.Label(self, text="가Ag", font=FONT_B, pady=0, bd=0)
        row_h = probe.winfo_reqheight() + 4
        probe.destroy()
        HR     = 5                           # 머리글 5단: 년월/요일/일/출근/ST·OT
        n_rows = HR + len(items) + 1
        prev   = getattr(self, "_n_rows_prev", 0)
        for fr in (L, R):
            for r in range(max(n_rows, prev)):
                fr.rowconfigure(r, minsize=(row_h + 1) if r < n_rows else 0)
        self._n_rows_prev = n_rows

        H = pm.C_HDR_BG
        # ── 왼쪽 고정 머리글 (4단 병합) ──
        for ci, (hd, w) in enumerate(zip(["No.", "구분", "Project\nCode",
                                          "Func. Code", "Description"], FIXED_W)):
            self._cell(L, 0, ci, hd, w=w, bg=H, bold=True, rs=HR)

        # ── 오른쪽 날짜 머리글 ──
        for lbl, i0, i1 in pm.month_groups(dates):
            self._cell(R, 0, 2 * i0, lbl, bg=H, bold=True, cs=2 * (i1 - i0 + 1))
        for i, d in enumerate(dates):
            fg = pm.C_WEEKEND if d.weekday() >= 5 else "black"
            self._cell(R, 1, 2 * i, pm.DAY_KO[d.weekday()], bg=H, fg=fg, bold=True, cs=2)
            self._cell(R, 2, 2 * i, str(d.day),             bg=H, fg=fg, bold=True, cs=2)
            s = self._starts.get(d, "")
            flex = bool(s) and s != "08:30"
            manual = d in self._day_override
            sfg = C_MANUAL if manual else (pm.C_FLEX if flex else "#595959")
            if d.weekday() < 5 and s:
                # 평일 + 실적 있는 날: 클릭해서 그날 출근시각 변경 가능
                # → 흰 바탕 + ▾ 표시 + 마우스 올리면 하늘색 (버튼처럼 보이게)
                lbl = self._cell(R, 3, 2 * i, f"{s} ▾", bg=C_PICK_BG, fg=sfg,
                                 bold=flex or manual, cs=2)
                lbl.config(cursor="hand2")
                lbl.bind("<Button-1>", lambda e, d=d: self._flex_menu(e, d))
                lbl.bind("<Enter>", lambda e, w=lbl: w.config(bg=C_PICK_HOVER))
                lbl.bind("<Leave>", lambda e, w=lbl: w.config(bg=C_PICK_BG))
            else:
                self._cell(R, 3, 2 * i, s, bg=H, fg=sfg, bold=flex or manual, cs=2)
            self._cell(R, 4, 2 * i,     "ST", bg=H, fg=fg, bold=True)
            self._cell(R, 4, 2 * i + 1, "OT", bg=H, fg=fg, bold=True)
        c_st = 2 * len(dates)
        self._cell(R, 0, c_st,     "ST\nTOTAL", w=TOTAL_W, bg=H, bold=True, rs=HR)
        self._cell(R, 0, c_st + 1, "OT\nTOTAL", w=TOTAL_W, bg=H, bold=True, rs=HR)

        # ── 데이터 ──
        for ri, it in enumerate(items):
            r = HR + ri
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
        tr  = HR + len(items)
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
            pm.save_workbook(path, self._items, self._dates, self._starts)
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
