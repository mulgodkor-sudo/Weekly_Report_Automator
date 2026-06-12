"""
app.py  ─  Weekly Report Automator  메인 UI
레이아웃 (위→아래):
  타이틀바 / 일별추출 / 주별추출(큰 버튼) / 사전검토결과 /
  📄Excel생성(큰 버튼) / FC엑셀파일 / 저장설정 / 상태바
"""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
import threading, os

from config          import load_config, save_config, reload_fc
from outlook_reader  import get_events
from event_processor import (process_this_week, process_next_week,
                              build_rows, get_warnings)
from excel_writer    import create_excel
import overrides as ov_mod
from override_dialog  import OverrideDialog
from plant_mh_dialog  import PlantMHDialog
from monthly_dialog   import MonthlyDialog


APP_VERSION = "Ver.1.1"   # main.py 와 동기화


def _default_filename() -> str:
    today  = datetime.today()
    mon    = today - timedelta(days=today.weekday())
    return f"WeeklyReport_{mon.month}월{(mon.day-1)//7+1}주차.xlsx"


class WeeklyReportApp:
    WIN_GEO = "780x320"    # 접힌 상태 기준 초기 높이
    WIN_MIN = (780, 320)   # 최소 크기

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Weekly Report Automator  {APP_VERSION}")
        self.root.geometry(self.WIN_GEO)
        self.root.minsize(*self.WIN_MIN)
        self.root.configure(bg="#F0F2F5")

        self._rows: list[dict] | None = None
        self._cfg  = load_config()

        self._setup_vars()
        self._build_ui()
        self._init_dates()
        self._load_fc_on_startup()

    # ── 변수 ─────────────────────────────────────────────────────────
    def _setup_vars(self):
        self.v_single   = tk.StringVar()
        self.v_use_this = tk.BooleanVar(value=True)
        self.v_use_next = tk.BooleanVar(value=True)
        self.v_ts       = tk.StringVar()
        self.v_te       = tk.StringVar()
        self.v_ns       = tk.StringVar()
        self.v_ne       = tk.StringVar()
        # 이번주 시작일 변경 시 나머지 날짜 자동완성
        self.v_ts.trace_add("write", self._on_ts_changed)
        self.v_fc_path  = tk.StringVar(value=self._cfg.get("func_code_excel_path",""))
        self.v_save_dir = tk.StringVar()
        self.v_filename = tk.StringVar(value=_default_filename())
        self.v_status   = tk.StringVar(value="  준비")

    # ── 이번주 시작일 자동완성 ────────────────────────────────────────
    def _on_ts_changed(self, *_):
        """이번주 시작일(v_ts) 변경 시 나머지 날짜 자동완성.
        YYYY-MM-DD 10자리 완성 시만 동작 (입력 중 무시)."""
        s = self.v_ts.get().strip()
        if len(s) != 10:
            return
        try:
            ts   = datetime.strptime(s, "%Y-%m-%d")
            fri  = ts + timedelta(days=4)
            nmon = ts + timedelta(days=7)
            nfri = ts + timedelta(days=11)
            self.v_te.set(fri.strftime("%Y-%m-%d"))
            self.v_ns.set(nmon.strftime("%Y-%m-%d"))
            self.v_ne.set(nfri.strftime("%Y-%m-%d"))
        except ValueError:
            pass

    # ── 날짜 초기화 ───────────────────────────────────────────────────
    def _init_dates(self, *_):
        today = datetime.today()
        mon   = today - timedelta(days=today.weekday())
        fri   = mon + timedelta(days=4)
        nmon  = mon + timedelta(days=7)
        nfri  = nmon + timedelta(days=4)

        self.v_single.set(today.strftime("%Y-%m-%d"))
        self.v_ts.set(mon.strftime("%Y-%m-%d"))
        self.v_te.set(fri.strftime("%Y-%m-%d"))
        self.v_ns.set(nmon.strftime("%Y-%m-%d"))
        self.v_ne.set(nfri.strftime("%Y-%m-%d"))

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        self.v_save_dir.set(desktop)
        self.v_filename.set(_default_filename())

    # ── FC 자동 로드 ──────────────────────────────────────────────────
    def _load_fc_on_startup(self):
        path = self.v_fc_path.get().strip()
        if path and os.path.exists(path):
            try:
                reload_fc(path)
                self._set_status(f"FC 로드 완료  ({os.path.basename(path)})")
            except Exception as e:
                self._set_status(f"FC 로드 실패: {e}")
        else:
            self._set_status("FC 엑셀 파일 경로를 설정하세요")

    # ════════════════════════════════════════════════════════════════
    # UI 빌드
    # ════════════════════════════════════════════════════════════════
    # ════════════════════════════════════════════════════════════════
    # 접이식 섹션 헬퍼
    # ════════════════════════════════════════════════════════════════
    def _collapsible(self, parent, title: str, start_open: bool = False):
        """
        접이식 섹션 위젯 반환 (header_frame, content_frame, toggle_fn)
        header_frame: 항상 표시되는 토글 헤더
        content_frame: 접히는 내용 영역
        """
        wrap = tk.Frame(parent, bg="#F0F2F5")
        wrap.pack(fill="x", pady=(0, 3))

        # 헤더 행
        hdr = tk.Frame(wrap, bg="#D9E1F2", cursor="hand2")
        hdr.pack(fill="x")

        self._toggle_vars = getattr(self, "_toggle_vars", {})
        var = tk.BooleanVar(value=start_open)
        self._toggle_vars[title] = var

        arrow_lbl = tk.Label(hdr, text="▼" if start_open else "▶",
                             bg="#D9E1F2", fg="#1F4E79",
                             font=("맑은 고딕", 9, "bold"), width=2)
        arrow_lbl.pack(side="left", padx=(6, 0), pady=4)

        tk.Label(hdr, text=title, bg="#D9E1F2", fg="#1F4E79",
                 font=("맑은 고딕", 9, "bold"), anchor="w"
                 ).pack(side="left", padx=4, pady=4)

        # 내용 프레임
        content = tk.Frame(wrap, bg="#F0F2F5")
        if start_open:
            content.pack(fill="x")

        def toggle(event=None):
            if var.get():
                content.pack_forget()
                var.set(False)
                arrow_lbl.config(text="▶")
            else:
                content.pack(fill="x")
                var.set(True)
                arrow_lbl.config(text="▼")
            # 창 높이 자동 조절 (너비는 유지)
            self.root.update_idletasks()
            w = self.root.winfo_width()
            h = self.root.winfo_reqheight()
            h = max(self.WIN_MIN[1], min(h + 20, 900))
            self.root.geometry(f"{w}x{h}")

        hdr.bind("<Button-1>", toggle)
        arrow_lbl.bind("<Button-1>", toggle)
        for child in hdr.winfo_children():
            child.bind("<Button-1>", toggle)

        return content, toggle

    def _build_ui(self):
        self._title_bar()
        body = tk.Frame(self.root, bg="#F0F2F5", padx=16, pady=8)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)

        # ── 1. 상단: 주별(좌) + 일별(우) 항상 표시 ──────────────────
        top_row = tk.Frame(body, bg="#F0F2F5")
        top_row.pack(fill="x", pady=(0, 6))
        top_row.columnconfigure(0, weight=3)
        top_row.columnconfigure(1, weight=1)

        self._sec_week(top_row)
        self._sec_single(top_row)

        # ── 2. 주요 버튼 항상 표시 ───────────────────────────────────
        self._sec_gen_btn(body)

        # ── 3. 사전 검토 결과 (처음엔 접혀 있음, 로드 후 자동 펼침) ───
        self._warn_content, self._warn_toggle = self._collapsible(
            body, "  ⚠️   사전 검토 결과  ─  불러오기 후 표시", start_open=False)
        self._sec_warn(self._warn_content)

        # ── 4. 되풀이 모임 설정 (접혀 있음) ─────────────────────────
        ov_content, _ = self._collapsible(
            body, "  🔁  되풀이 모임 설정 (계획)", start_open=False)
        self._sec_override_inner(ov_content)

        # ── 5. 상세 설정: FC + 저장 (접혀 있음) ─────────────────────
        detail_content, _ = self._collapsible(
            body, "  ⚙️   상세 설정  (FC 파일 / 저장 경로)", start_open=False)
        self._sec_fc(detail_content)
        self._sec_save(detail_content)

        self._status_bar()

    # ── 타이틀 바 ─────────────────────────────────────────────────────
    def _title_bar(self):
        bar = tk.Frame(self.root, bg="#1F4E79", height=52)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text=f"  📊  Weekly Report Automator  {APP_VERSION}",
                 font=("맑은 고딕", 14, "bold"),
                 fg="white", bg="#1F4E79").pack(side="left", padx=14, pady=10)
        tk.Label(bar, text="DL이엔씨  플랜트본부 기계설계팀  ",
                 font=("맑은 고딕", 9), fg="#9DC3E6", bg="#1F4E79"
                 ).pack(side="right", padx=14)

    # ── 일별 추출 (우측 보조, grid로 배치) ──────────────────────────
    def _sec_single(self, parent):
        fr = ttk.LabelFrame(parent, text="  🗓  일별 추출 (특정 날짜)",
                            padding=(12, 10))
        fr.grid(row=0, column=1, sticky="nsew")
        fr.columnconfigure(0, weight=1)

        P = dict(padx=5, pady=6)
        ttk.Label(fr, text="날  짜 :", font=("맑은 고딕", 9)
                  ).grid(row=0, column=0, sticky="w", **P)
        ttk.Entry(fr, textvariable=self.v_single, width=14,
                  font=("맑은 고딕", 10)
                  ).grid(row=1, column=0, sticky="ew", padx=5, pady=2)
        ttk.Label(fr, text="(YYYY-MM-DD)", foreground="gray",
                  font=("맑은 고딕", 8)
                  ).grid(row=2, column=0, sticky="w", padx=5, pady=0)
        ttk.Button(fr, text="🔍  일별 불러오기",
                   command=self._load_single, width=16
                   ).grid(row=3, column=0, sticky="ew", padx=5, pady=(10, 2))

    # ── 주별 추출 (좌측 메인, grid로 부모에 배치) ────────────────────
    def _sec_week(self, parent):
        fr = ttk.LabelFrame(parent,
                            text="  📅  주별 추출  (Weekly Report 핵심 기능)",
                            padding=(14, 10))
        fr.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        fr.columnconfigure(1, weight=1)
        fr.columnconfigure(3, weight=1)

        P  = dict(padx=6, pady=6)
        Ps = dict(padx=6, pady=4)   # 작은 행 여백

        # ── 이번주: 시작일 강조 ─────────────────────────────────────
        ttk.Checkbutton(fr, text="이번주  [실적] :",
                        variable=self.v_use_this).grid(row=0, column=0, sticky="e", **P)

        # ★ 이번주 시작일: 폰트 크게, 강조 배경
        ts_entry = tk.Entry(fr, textvariable=self.v_ts,
                            width=14,
                            font=("맑은 고딕", 14, "bold"),
                            bg="#EFF6FF",          # 연한 파랑 배경
                            fg="#1F4E79",          # 네이비 글씨
                            relief="solid", bd=1,
                            insertbackground="#1F4E79")
        ts_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=8, ipady=4)

        ttk.Label(fr, text=" ~ ", font=("맑은 고딕", 10)
                  ).grid(row=0, column=2, **P)
        ttk.Entry(fr, textvariable=self.v_te, width=13,
                  font=("맑은 고딕", 10)).grid(row=0, column=3, sticky="ew", **P)

        # ── 다음주 ──────────────────────────────────────────────────
        ttk.Checkbutton(fr, text="다음주  [계획] :",
                        variable=self.v_use_next).grid(row=1, column=0, sticky="e", **Ps)
        ttk.Entry(fr, textvariable=self.v_ns, width=13,
                  font=("맑은 고딕", 10)).grid(row=1, column=1, sticky="ew", **Ps)
        ttk.Label(fr, text=" ~ ", font=("맑은 고딕", 10)
                  ).grid(row=1, column=2, **Ps)
        ttk.Entry(fr, textvariable=self.v_ne, width=13,
                  font=("맑은 고딕", 10)).grid(row=1, column=3, sticky="ew", **Ps)

        # ── 버튼 행 ─────────────────────────────────────────────────
        style = ttk.Style()
        style.configure("WeekLoad.TButton",
                        font=("맑은 고딕", 11, "bold"),
                        foreground="white",
                        background="#1F4E79")
        style.map("WeekLoad.TButton",
                  background=[("active", "#2E75B6"), ("disabled", "#AAAAAA")])
        style.configure("Reset.TButton", font=("맑은 고딕", 9))

        btn_fr = ttk.Frame(fr)
        btn_fr.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 2))
        btn_fr.columnconfigure(1, weight=1)

        ttk.Button(btn_fr, text="↺ 날짜 초기화",
                   command=self._init_dates, width=13,
                   style="Reset.TButton"
                   ).pack(side="left", padx=(0, 8))

        # ★ 주별 불러오기 - 색상 버튼
        tk.Button(btn_fr,
                  text="🔍   주별 불러오기",
                  command=self._load_week,
                  font=("맑은 고딕", 11, "bold"),
                  bg="#1F4E79", fg="white",
                  activebackground="#2E75B6", activeforeground="white",
                  relief="flat", padx=16, pady=6, cursor="hand2"
                  ).pack(side="left")

        ttk.Label(fr,
                  text="YYYY-MM-DD  |  이번주 시작일 입력 시 나머지 날짜 자동완성",
                  foreground="gray", font=("맑은 고딕", 8)
                  ).grid(row=3, column=0, columnspan=4, sticky="w", padx=4, pady=(0, 2))

    # ── 사전 검토 결과 내용 (접이식 컨텐츠 영역에 삽입) ────────────
    def _sec_warn(self, parent):
        self.txt_warn = tk.Text(
            parent, state="disabled", height=8,
            font=("맑은 고딕", 9), bg="#FFFDE7",
            relief="flat", wrap="word", bd=0,
        )
        sb = ttk.Scrollbar(parent, orient="vertical", command=self.txt_warn.yview)
        self.txt_warn.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", padx=(0, 2), pady=4)
        self.txt_warn.pack(fill="both", expand=True, padx=4, pady=4)

        self._set_warn(
            "⬅  [일별/주별 불러오기] 버튼을 클릭하면 아웃룩 일정을 가져옵니다.\n\n"
            "* Excel 생성 버튼을 누르면 Weekly Report 형식으로 출력됩니다."
        )

    # ── 섹션 4: Excel 생성 버튼 (크게, 진행률 옆) ────────────────────
    def _sec_gen_btn(self, parent):
        fr = tk.Frame(parent, bg="#F0F2F5")
        fr.pack(fill="x", pady=(0, 8))

        style = ttk.Style()
        style.configure("Gen.TButton", font=("맑은 고딕", 11, "bold"))
        style.configure("Sub.TButton", font=("맑은 고딕", 9))

        # ── 1행: 주요 기능 버튼 ─────────────────────────────────
        row1 = tk.Frame(fr, bg="#F0F2F5")
        row1.pack(fill="x", pady=(0, 4))

        self.btn_gen = tk.Button(
            row1, text="📄   Excel 파일 생성",
            command=self._generate,
            font=("맑은 고딕", 11, "bold"),
            bg="#1F5E20", fg="white",
            activebackground="#2E7D32", activeforeground="white",
            disabledforeground="#AAAAAA",
            relief="flat", padx=20, pady=7, cursor="hand2",
            state="disabled",
        )
        self.btn_gen.pack(side="left", padx=(0, 8))

        # Plant M/H 입력 확인 버튼
        self.btn_mh = tk.Button(
            row1, text="📊  Plant M/H 입력 확인",
            command=self._open_plant_mh,
            font=("맑은 고딕", 9, "bold"),
            bg="#5D4037", fg="white",
            activebackground="#6D4C41", activeforeground="white",
            disabledforeground="#AAAAAA",
            relief="flat", padx=12, pady=7, cursor="hand2",
            state="disabled",
        )
        self.btn_mh.pack(side="left", padx=(0, 8))

        # 월간업무정리 버튼
        tk.Button(row1,
                  text="📋  월간업무정리",
                  command=self._open_monthly,
                  font=("맑은 고딕", 9, "bold"),
                  bg="#4A148C", fg="white",
                  activebackground="#6A1B9A", activeforeground="white",
                  relief="flat", padx=12, pady=7, cursor="hand2",
                  ).pack(side="left", padx=(0, 8))

        self.progress = ttk.Progressbar(row1, mode="indeterminate", length=140)
        self.progress.pack(side="left", pady=4)

        # (되풀이 모임 설정 버튼은 접이식 섹션으로 이동)

    # ── FC 엑셀 파일 (접이식 상세 설정 내부) ────────────────────────
    def _sec_fc(self, parent):
        fr = ttk.LabelFrame(parent, text="  📂  Function Code 엑셀 파일", padding=(10, 5))
        fr.pack(fill="x", padx=6, pady=(6, 3))
        fr.columnconfigure(1, weight=1)

        P = dict(padx=5, pady=3)
        ttk.Label(fr, text="파일 경로 :").grid(row=0, column=0, sticky="e", **P)
        ttk.Entry(fr, textvariable=self.v_fc_path).grid(row=0, column=1, sticky="ew", **P)
        ttk.Button(fr, text="찾아보기", command=self._browse_fc, width=10
                   ).grid(row=0, column=2, **P)
        ttk.Button(fr, text="🔄 FC 재로드", command=self._reload_fc_btn, width=12
                   ).grid(row=0, column=3, padx=(4,4), pady=3)
        ttk.Label(fr, text='파일명에 "기계설계" 포함 | 읽기 전용 | 수식 결과값 기준',
                  foreground="gray"
                  ).grid(row=1, column=0, columnspan=4, sticky="w", padx=5)

    # ── 저장 설정 (접이식 상세 설정 내부) ──────────────────────────
    def _sec_save(self, parent):
        fr = ttk.LabelFrame(parent, text="  💾  저장 설정", padding=(10, 5))
        fr.pack(fill="x", padx=6, pady=(3, 8))
        fr.columnconfigure(1, weight=1)

        P = dict(padx=5, pady=3)
        ttk.Label(fr, text="저장 경로 :").grid(row=0, column=0, sticky="e", **P)
        ttk.Entry(fr, textvariable=self.v_save_dir).grid(row=0, column=1, sticky="ew", **P)
        ttk.Button(fr, text="찾아보기", command=self._browse_save, width=10
                   ).grid(row=0, column=2, **P)

        ttk.Label(fr, text="파일 이름 :").grid(row=1, column=0, sticky="e", **P)
        ttk.Entry(fr, textvariable=self.v_filename).grid(row=1, column=1, sticky="ew", **P)
        ttk.Label(fr, text=".xlsx 자동 추가", foreground="gray"
                  ).grid(row=1, column=2, sticky="w", **P)

    # ── 상태바 ───────────────────────────────────────────────────────
    def _status_bar(self):
        bar = tk.Frame(self.root, bg="#1F4E79", height=30)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        tk.Label(bar, textvariable=self.v_status,
                 fg="white", bg="#1F4E79",
                 font=("맑은 고딕", 9), anchor="w"
                 ).pack(side="left", padx=10, pady=5)
        tk.Label(bar, text="* 불편사항/개선사항은 이수신 차장에게 문의 바랍니다.  ",
                 fg="#9DC3E6", bg="#1F4E79",
                 font=("맑은 고딕", 8), anchor="e"
                 ).pack(side="right", padx=8, pady=5)

    # ════════════════════════════════════════════════════════════════
    # 핸들러
    # ════════════════════════════════════════════════════════════════

    def _open_override(self):
        OverrideDialog(self.root)

    def _sec_override_inner(self, parent):
        """접이식 섹션 내부: 되풀이 모임 설정 버튼 + 설명"""
        fr = tk.Frame(parent, bg="#F0F2F5")
        fr.pack(fill="x", padx=8, pady=8)
        tk.Button(fr,
                  text="🔁   되풀이 모임 내용 고정 설정 열기",
                  command=self._open_override,
                  font=("맑은 고딕", 10),
                  bg="#37474F", fg="white",
                  activebackground="#455A64", activeforeground="white",
                  relief="flat", padx=16, pady=6, cursor="hand2"
                  ).pack(side="left")
        tk.Label(fr,
                 text="  [계획]으로만 새로 추가되는 되풀이 약속의 상세내용을 고정 설정합니다.",
                 bg="#F0F2F5", fg="gray", font=("맑은 고딕", 8)
                 ).pack(side="left", padx=8)

    def _open_monthly(self):
        MonthlyDialog(self.root)

    def _open_plant_mh(self):
        if not self._rows:
            from tkinter import messagebox
            messagebox.showwarning("경고", "먼저 일정을 불러오세요.")
            return
        PlantMHDialog(self.root, self._rows)

    def _browse_save(self):
        p = filedialog.askdirectory(title="저장 경로 선택")
        if p: self.v_save_dir.set(p)

    def _browse_fc(self):
        p = filedialog.askopenfilename(
            title="Function Code 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if p:
            self.v_fc_path.set(p)
            self._reload_fc_btn()

    def _reload_fc_btn(self):
        path = self.v_fc_path.get().strip()
        if not path:
            messagebox.showwarning("경고", "FC 엑셀 파일 경로를 입력하세요.")
            return
        if not os.path.exists(path):
            messagebox.showerror("오류", f"파일이 없습니다:\n{path}")
            return
        try:
            reload_fc(path)
            self._cfg["func_code_excel_path"] = path
            save_config(self._cfg)
            self._set_status(f"FC 재로드 완료  ({os.path.basename(path)})")
        except Exception as e:
            messagebox.showerror("FC 로드 오류", str(e))

    def _parse_date(self, s: str) -> datetime:
        try:
            return datetime.strptime(s.strip(), "%Y-%m-%d")
        except ValueError:
            raise ValueError(f"날짜 형식 오류: '{s}'\nYYYY-MM-DD 형식으로 입력하세요.")

    def _load_single(self):
        s = self.v_single.get().strip()
        if not s:
            messagebox.showwarning("경고", "날짜를 입력하세요.")
            return
        try:
            sd = self._parse_date(s)
        except ValueError as e:
            messagebox.showerror("날짜 오류", str(e))
            return
        self._do_load(sd, sd, None, None, "일별", mode="single", single_date=sd)

    def _load_week(self):
        use_this = self.v_use_this.get()
        use_next = self.v_use_next.get()
        if not use_this and not use_next:
            messagebox.showwarning("경고", "이번주 또는 다음주 중 하나 이상 체크하세요.")
            return
        try:
            ts = self._parse_date(self.v_ts.get()) if use_this else None
            te = self._parse_date(self.v_te.get()) if use_this else None
            ns = self._parse_date(self.v_ns.get()) if use_next else None
            ne = self._parse_date(self.v_ne.get()) if use_next else None
        except ValueError as e:
            messagebox.showerror("날짜 오류", str(e))
            return
        self._do_load(ts, te, ns, ne, "주별", mode="week")

    def _do_load(self, ts, te, ns, ne, label="", mode="week", single_date=None):
        self.btn_gen.config(state="disabled")
        self.btn_mh.config(state="disabled")
        self.progress.start()
        self._set_status(f"아웃룩 {label} 일정 불러오는 중...")
        self._set_warn("⏳  아웃룩 일정을 불러오는 중입니다...\n\n잠시 기다려 주세요.")

        # 일별=8H, 주별=40H
        this_exp = 8.0 if mode == "single" else 40.0

        def task():
            # ── COM 스레드 초기화 (PyInstaller exe 필수) ─────────────────
            # 별도 스레드에서 Outlook COM 객체를 사용하려면
            # pythoncom.CoInitialize()를 반드시 먼저 호출해야 함.
            # 미호출 시 COM이 조용히 실패하여 빈 결과를 반환함.
            _com_initialized = False
            try:
                import pythoncom
                pythoncom.CoInitialize()
                _com_initialized = True
            except Exception:
                pass  # pythoncom 없으면 무시 (개발환경 등)

            try:
                this_raw = get_events(ts, te) if ts else []
                next_raw = get_events(ns, ne) if ns else []
                this_proc = process_this_week(this_raw)
                next_proc = process_next_week(next_raw)
                warns = get_warnings(this_raw, next_raw, this_expected=this_exp)

                # FC 로드 여부 확인
                from config import get_fc
                if not get_fc():
                    warns.insert(0,
                        "⚠️  Function Code 파일이 로드되지 않았습니다.\n"
                        "      하단 [FC 엑셀 파일] 경로를 확인하고 🔄FC 재로드를 눌러주세요.")

                # 0건 진단: 아웃룩 접근은 됐지만 일정이 없는 경우
                if not this_raw and not next_raw:
                    ts_str = ts.strftime("%m/%d") if ts else "?"
                    te_str = te.strftime("%m/%d") if te else "?"
                    warns.insert(0,
                        f"⚠️  아웃룩에서 해당 기간({ts_str}~{te_str}) 일정을 찾지 못했습니다.\n"
                        "      Outlook이 실행 중이고 로그인되어 있는지 확인하세요.\n"
                        "      해당 기간에 [실적]/[계획] 태그가 붙은 일정이 있는지 확인하세요.")

                built  = build_rows(this_proc, next_proc)
                ovrs   = ov_mod.load()
                built  = ov_mod.apply(built, ovrs)
                ov_cnt = ov_mod.count_applied(built, ovrs)
                self.root.after(0, lambda: self._on_load_ok(
                    this_proc, next_proc, warns, built, ov_cnt, mode, single_date, ts))
            except Exception as err:
                msg = str(err)
                self.root.after(0, lambda: self._on_err(msg, f"{label} 불러오기 오류"))
            finally:
                # COM 정리
                if _com_initialized:
                    try:
                        import pythoncom
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        threading.Thread(target=task, daemon=True).start()
    def _on_load_ok(self, this_proc, next_proc, warns, built,
                    ov_cnt: int = 0, mode: str = "week",
                    single_date=None, week_ts=None):
        self.progress.stop()
        self._rows = built

        # 사전 검토 결과 자동 펼침
        if hasattr(self, "_warn_toggle"):
            toggle_vars = getattr(self, "_toggle_vars", {})
            key = "  ⚠️   사전 검토 결과  ─  불러오기 후 표시"
            if not toggle_vars.get(key, tk.BooleanVar(value=False)).get():
                self._warn_toggle()

        # 모드에 따라 파일명 자동 업데이트
        if mode == "single" and single_date is not None:
            # 일별: YYYYMMDD 형식
            self.v_filename.set(f"WeeklyReport_{single_date.strftime('%Y%m%d')}.xlsx")
        elif mode == "week":
            # 주별: 입력한 이번주 시작일(ts) 기준 주차 산출
            anchor = week_ts or datetime.today()
            mon    = anchor - timedelta(days=anchor.weekday())
            self.v_filename.set(
                f"WeeklyReport_{mon.month}월{(mon.day-1)//7+1}주차.xlsx")

        th = sum(e["hours"] for e in this_proc)
        nh = sum(e["hours"] for e in next_proc)
        lines = []
        if this_proc: lines.append(f"✅  이번주 [실적] : {len(this_proc)}건  (합계 {th:.1f}H)")
        if next_proc: lines.append(f"✅  다음주 [계획] : {len(next_proc)}건  (합계 {nh:.1f}H)")
        if ov_cnt:    lines.append(f"🔁  되풀이 모임 override 적용 : {ov_cnt}건")
        lines.append("")
        if warns:
            lines += warns
            lines += ["", "※ 경고 항목 확인 후 아웃룩 수정 → 재로딩 권장"]
        else:
            lines.append("✅  모든 검토 항목 이상 없음")

        self._set_warn("\n".join(lines))
        if self._rows:
            self.btn_gen.config(state="normal")
            self.btn_mh.config(state="normal")
            self._set_status(f"로드 완료  →  {len(self._rows)}개 행 준비됨")
        else:
            self._set_status("로드 완료  →  [실적/계획] 항목 없음")

    def _generate(self):
        if not self._rows:
            messagebox.showwarning("경고", "먼저 일정을 불러오세요.")
            return

        save_dir = self.v_save_dir.get().strip()
        filename = self.v_filename.get().strip()
        if not save_dir:
            messagebox.showwarning("경고", "저장 경로를 입력하세요.")
            return
        if not filename:
            messagebox.showwarning("경고", "파일 이름을 입력하세요.")
            return
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
            self.v_filename.set(filename)

        out_path = os.path.join(save_dir, filename)

        if os.path.exists(out_path):
            if not messagebox.askyesno("덮어쓰기 확인",
                                       f"동일 파일이 존재합니다.\n\n{out_path}\n\n덮어쓰시겠습니까?"):
                return
            try:
                with open(out_path, "a"):
                    pass
            except PermissionError:
                messagebox.showerror("파일 열림",
                                     f"파일이 열려 있습니다. Excel을 닫고 다시 시도하세요.\n{out_path}")
                return

        self.btn_gen.config(state="disabled")
        self.progress.start()
        self._set_status("Excel 파일 생성 중...")

        def task():
            try:
                create_excel(self._rows, out_path)
                self.root.after(0, lambda: self._on_gen_ok(out_path))
            except Exception as err:
                msg = str(err)
                self.root.after(0, lambda: self._on_err(msg, "Excel 생성 오류"))

        threading.Thread(target=task, daemon=True).start()

    def _on_gen_ok(self, path: str):
        self.progress.stop()
        self.btn_gen.config(state="normal")
        self._set_status(f"저장 완료  →  {path}")
        if messagebox.askyesno("완료", f"Excel 생성 완료.\n\n{path}\n\n바로 여시겠습니까?"):
            try:
                os.startfile(path)
            except Exception:
                pass

    def _on_err(self, msg: str, title: str):
        self.progress.stop()
        self.btn_gen.config(state="disabled")
        self.btn_mh.config(state="disabled")
        self._set_status(f"오류  →  {msg[:70]}")
        self._set_warn(f"❌  {title}\n\n{msg}")
        messagebox.showerror(title, msg)

    def _set_warn(self, text: str):
        self.txt_warn.config(state="normal")
        self.txt_warn.delete("1.0", tk.END)
        self.txt_warn.insert(tk.END, text)
        self.txt_warn.config(state="disabled")

    def _set_status(self, msg: str):
        self.v_status.set(f"  {msg}")
