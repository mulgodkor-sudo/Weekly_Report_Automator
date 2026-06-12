"""
override_dialog.py
되풀이 모임 내용 고정 설정 다이얼로그 (소형 버전)
"""
from __future__ import annotations
import re
import tkinter as tk
from tkinter import ttk, messagebox
import overrides as ov_mod

# 아웃룩 제목 파싱용 정규식
_CODE_RE     = re.compile(r"\[\s*#\s*([A-Za-z0-9]{5,8})\s*#\s*([A-Z]{2}\s*\d{2}\s*-\s*\d{2})\s*\]")
_LEADING_RE  = re.compile(r"^(\s*\[[^\]]*\]\s*)+")
_TRAIL_RE    = re.compile(r"\s*\[\s*#\s*[A-Za-z0-9]{5,8}\s*#\s*[A-Z]{2}\s*\d{2}\s*-\s*\d{2}\s*\]\s*")


def _parse_outlook_subject(s: str) -> tuple[str, str, str]:
    """
    아웃룩 전체 제목에서 (project_code, func_code, clean_subject) 추출
    예) '[실적] [General] Team Meeting [#000000#GA11-25]'
        → ('000000', 'GA11-25', 'Team Meeting')
    """
    m = _CODE_RE.search(s)
    pc = m.group(1) if m else ""
    fc = m.group(2) if m else ""
    clean = _TRAIL_RE.sub(" ", s)    # 코드 블록 제거
    clean = _LEADING_RE.sub("", clean).strip()  # 앞쪽 [태그] 제거
    return pc, fc, clean


class OverrideDialog(tk.Toplevel):
    """
    되풀이 모임 내용 고정 설정 – 컴팩트 다이얼로그
    ────────────────────────────────────────────
    [빠른 파싱] 아웃룩 제목 입력 → 자동 분석
    [등록 목록] 리스트박스 (5줄)
    [입력/수정] PC / FC / 제목 / 상세(3줄)
    [버튼 행  ] 새로추가 · 저장 · 삭제 · 닫기
    """

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("되풀이 모임 내용 고정 설정")
        self.geometry("700x490")
        self.minsize(600, 430)
        self.resizable(True, True)
        self.grab_set()
        self.configure(bg="#F0F2F5")

        self._data: list[dict] = ov_mod.load()
        self._sel_idx: int | None = None

        self._build()
        self._refresh_list()

    # ════════════════════════════════════════════════════════════════
    # UI 빌드
    # ════════════════════════════════════════════════════════════════
    def _build(self):
        # 타이틀 바
        bar = tk.Frame(self, bg="#1F4E79", height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Label(bar, text="  🔁  되풀이 모임 내용 고정 설정",
                 font=("맑은 고딕", 11, "bold"),
                 fg="white", bg="#1F4E79").pack(side="left", padx=12, pady=6)

        body = tk.Frame(self, bg="#F0F2F5", padx=12, pady=8)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=0)  # 리스트
        body.rowconfigure(2, weight=1)  # 폼

        # ── 빠른 파싱 ─────────────────────────────────────────────
        fr_parse = ttk.LabelFrame(
            body, text="  ⚡  빠른 파싱 (아웃룩 제목 붙여넣기 → 자동 분석)",
            padding=(10, 5))
        fr_parse.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        fr_parse.columnconfigure(0, weight=1)

        parse_row = ttk.Frame(fr_parse)
        parse_row.pack(fill="x")
        parse_row.columnconfigure(0, weight=1)

        self.v_parse = tk.StringVar()
        ttk.Entry(parse_row, textvariable=self.v_parse,
                  font=("맑은 고딕", 9)
                  ).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(parse_row, text="🔍 분석",
                   command=self._do_parse, width=10
                   ).grid(row=0, column=1)
        ttk.Label(fr_parse,
                  text="예) [실적] [General] Team Meeting [#000000#GA11-25]",
                  foreground="gray", font=("맑은 고딕", 8)
                  ).pack(anchor="w", pady=(2, 0))

        # ── 등록된 항목 리스트 ────────────────────────────────────
        fr_list = ttk.LabelFrame(body, text="  📋  등록된 항목 (클릭하면 수정)", padding=(8, 4))
        fr_list.grid(row=1, column=0, sticky="nsew", pady=(0, 6))
        fr_list.columnconfigure(0, weight=1)

        lb_fr = tk.Frame(fr_list)
        lb_fr.pack(fill="both", expand=True)
        lb_fr.columnconfigure(0, weight=1)

        self.lb = tk.Listbox(lb_fr, height=5, selectmode="single",
                             font=("맑은 고딕", 9), activestyle="underline",
                             relief="flat", bg="white", borderwidth=1)
        sb = ttk.Scrollbar(lb_fr, orient="vertical", command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.lb.pack(fill="both", expand=True)
        self.lb.bind("<<ListboxSelect>>", self._on_select)

        # ── 입력/수정 폼 ──────────────────────────────────────────
        fr_form = ttk.LabelFrame(body, text="  ✏️  내용 입력 / 수정", padding=(10, 6))
        fr_form.grid(row=2, column=0, sticky="nsew")
        fr_form.columnconfigure(1, weight=1)
        fr_form.columnconfigure(3, weight=1)
        fr_form.rowconfigure(2, weight=1)

        P = dict(padx=5, pady=4)

        # PC / FC (1행)
        ttk.Label(fr_form, text="Project Code :").grid(row=0, column=0, sticky="e", **P)
        self.v_pc = tk.StringVar()
        ttk.Entry(fr_form, textvariable=self.v_pc, width=14
                  ).grid(row=0, column=1, sticky="w", **P)

        ttk.Label(fr_form, text="Function Code :").grid(row=0, column=2, sticky="e", **P)
        self.v_fc = tk.StringVar()
        ttk.Entry(fr_form, textvariable=self.v_fc, width=14
                  ).grid(row=0, column=3, sticky="w", **P)

        # 제목 (2행)
        ttk.Label(fr_form, text="수행업무 (제목) :").grid(row=1, column=0, sticky="e", **P)
        self.v_subj = tk.StringVar()
        ttk.Entry(fr_form, textvariable=self.v_subj
                  ).grid(row=1, column=1, columnspan=3, sticky="ew", **P)

        # 상세 (3행) - 3줄 크기
        ttk.Label(fr_form, text="수행업무 (상세) :").grid(row=2, column=0, sticky="ne", padx=5, pady=5)

        body_wrap = ttk.Frame(fr_form)
        body_wrap.grid(row=2, column=1, columnspan=3, sticky="nsew", **P)
        body_wrap.rowconfigure(0, weight=1)
        body_wrap.columnconfigure(0, weight=1)

        self.txt_body = tk.Text(
            body_wrap, font=("맑은 고딕", 9), wrap="word",
            height=3,                   # ← 3줄 크기
            relief="solid", bd=1, bg="white",
        )
        sb2 = ttk.Scrollbar(body_wrap, orient="vertical", command=self.txt_body.yview)
        self.txt_body.configure(yscrollcommand=sb2.set)
        sb2.grid(row=0, column=1, sticky="ns")
        self.txt_body.grid(row=0, column=0, sticky="nsew")

        # ── 버튼 행 (항상 보이도록 pack 마지막) ──────────────────
        fr_btn = tk.Frame(self, bg="#F0F2F5")
        fr_btn.pack(fill="x", padx=12, pady=(0, 8))

        ttk.Button(fr_btn, text="➕  새로 추가",
                   command=self._new, width=13).pack(side="left", padx=(0, 5))
        ttk.Button(fr_btn, text="💾  저장",
                   command=self._save_item, width=11).pack(side="left", padx=(0, 5))
        ttk.Button(fr_btn, text="🗑  삭제",
                   command=self._delete, width=11).pack(side="left", padx=(0, 5))

        ttk.Label(fr_btn,
                  text="제목: 태그·코드 제외한 업무명만",
                  foreground="gray", font=("맑은 고딕", 8)
                  ).pack(side="left", padx=8)

        ttk.Button(fr_btn, text="✅  닫기",
                   command=self.destroy, width=11).pack(side="right")

    # ════════════════════════════════════════════════════════════════
    # 빠른 파싱
    # ════════════════════════════════════════════════════════════════
    def _do_parse(self):
        raw = self.v_parse.get().strip()
        if not raw:
            messagebox.showwarning("입력 없음", "아웃룩 제목을 먼저 입력하세요.", parent=self)
            return
        pc, fc, clean = _parse_outlook_subject(raw)
        self.v_pc.set(pc)
        self.v_fc.set(fc)
        self.v_subj.set(clean)
        # 파싱 후 상세 입력 칸으로 포커스
        self.txt_body.focus_set()
        if not pc and not fc:
            messagebox.showinfo(
                "파싱 결과",
                f"코드 블록을 찾지 못했습니다.\n\n"
                f"제목: {clean}\n\n"
                "Project Code / Function Code를 직접 입력하세요.",
                parent=self,
            )

    # ════════════════════════════════════════════════════════════════
    # 리스트
    # ════════════════════════════════════════════════════════════════
    def _refresh_list(self, select_idx: int | None = None):
        self.lb.delete(0, tk.END)
        for item in self._data:
            pc   = item.get("project_code", "")
            fc   = item.get("func_code",    "")
            subj = item.get("subject",      "")[:45]
            self.lb.insert(tk.END, f"  [{pc} / {fc}]   {subj}")
        if select_idx is not None and 0 <= select_idx < len(self._data):
            self.lb.selection_set(select_idx)
            self.lb.see(select_idx)

    # ════════════════════════════════════════════════════════════════
    # 이벤트
    # ════════════════════════════════════════════════════════════════
    def _on_select(self, *_):
        sel = self.lb.curselection()
        if not sel:
            return
        idx = sel[0]
        self._sel_idx = idx
        item = self._data[idx]
        self.v_pc.set(item.get("project_code", ""))
        self.v_fc.set(item.get("func_code",    ""))
        self.v_subj.set(item.get("subject",    ""))
        self.txt_body.delete("1.0", tk.END)
        self.txt_body.insert("1.0", item.get("body", ""))

    def _new(self):
        self._sel_idx = None
        self.lb.selection_clear(0, tk.END)
        self.v_parse.set("")
        self.v_pc.set("")
        self.v_fc.set("")
        self.v_subj.set("")
        self.txt_body.delete("1.0", tk.END)

    def _save_item(self):
        pc   = self.v_pc.get().strip()
        fc   = self.v_fc.get().strip()
        subj = self.v_subj.get().strip()
        body = self.txt_body.get("1.0", tk.END).rstrip("\n")

        if not pc or not fc or not subj:
            messagebox.showwarning(
                "입력 오류",
                "Project Code, Function Code, 수행업무(제목)은 필수입니다.",
                parent=self,
            )
            return

        new_item = {"project_code": pc, "func_code": fc,
                    "subject": subj, "body": body}

        if self._sel_idx is not None:
            self._data[self._sel_idx] = new_item
            save_idx = self._sel_idx
        else:
            for i, d in enumerate(self._data):
                if (d.get("project_code","") == pc
                        and d.get("func_code","") == fc
                        and d.get("subject","") == subj):
                    self._data[i] = new_item
                    save_idx = i
                    break
            else:
                self._data.append(new_item)
                save_idx = len(self._data) - 1

        ov_mod.save(self._data)
        self._sel_idx = save_idx
        self._refresh_list(select_idx=save_idx)
        messagebox.showinfo("저장 완료", "저장되었습니다.", parent=self)

    def _delete(self):
        if self._sel_idx is None:
            messagebox.showwarning("선택 오류", "삭제할 항목을 먼저 선택하세요.", parent=self)
            return
        item = self._data[self._sel_idx]
        if not messagebox.askyesno(
            "삭제 확인",
            f"아래 항목을 삭제하시겠습니까?\n\n"
            f"[{item.get('project_code','')} / {item.get('func_code','')}]\n"
            f"{item.get('subject','')}",
            parent=self,
        ):
            return
        del self._data[self._sel_idx]
        ov_mod.save(self._data)
        self._sel_idx = None
        self._new()
        self._refresh_list()
