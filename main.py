"""
main.py  ─  Weekly Report Automator 런처
역할: src/ 폴더를 import 경로에 추가 후 앱 실행
      exe 옆 src/*.py 우선 → _MEIPASS/src/ fallback (패치 구조)
"""
from __future__ import annotations
import os, sys


# ── src/ 경로 설정 (가장 먼저 실행) ────────────────────────────────
def _setup_src():
    if getattr(sys, "frozen", False):
        exe_dir     = os.path.dirname(sys.executable)
        src_live    = os.path.join(exe_dir, "src")          # 패치 파일 (우선)
        src_bundled = os.path.join(sys._MEIPASS, "src")     # 번들 파일 (fallback)
        if os.path.isdir(src_live):
            sys.path.insert(0, src_live)
        if os.path.isdir(src_bundled):
            sys.path.append(src_bundled)
    else:
        # 개발 환경
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

_setup_src()

import tkinter as tk


def resource_path(rel: str) -> str:
    """assets 등 번들 리소스 경로 반환"""
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS          # _MEIPASS/assets/...
    else:
        base = os.path.join(os.path.dirname(__file__), "src")  # src/assets/...
    return os.path.join(base, rel)


# ══════════════════════════════════════════════════════════
# 스플래시 화면
# ══════════════════════════════════════════════════════════

class SplashScreen:
    DURATION_MS = 2500
    DEFAULT_W   = 600
    DEFAULT_H   = 280

    def __init__(self, root: tk.Tk, cfg: dict):
        from config import get_version_str
        self._version = get_version_str(cfg)
        self._win = tk.Toplevel(root)
        self._win.overrideredirect(True)
        self._win.lift()
        self._build(cfg)

    def _build(self, cfg: dict):
        splash_path = resource_path(cfg.get("splash_image", "assets/splash.png"))
        try:
            self._img = tk.PhotoImage(file=splash_path)
            w, h = self._img.width(), self._img.height()
            cv = tk.Canvas(self._win, width=w, height=h, bd=0, highlightthickness=0)
            cv.pack()
            cv.create_image(0, 0, anchor="nw", image=self._img)
            cv.create_text(w - 10, h - 10, anchor="se", text=self._version,
                           font=("맑은 고딕", 10, "bold"), fill="#9DC3E6")
        except Exception:
            self._img = None
            w, h = self.DEFAULT_W, self.DEFAULT_H
            frame = tk.Frame(self._win, bg="#1F4E79", width=w, height=h)
            frame.pack_propagate(False)
            frame.pack()
            tk.Label(frame, text="📊", font=("Segoe UI Emoji", 44),
                     bg="#1F4E79", fg="white").pack(pady=(36, 4))
            tk.Label(frame, text="Weekly Report Automator",
                     font=("맑은 고딕", 18, "bold"),
                     bg="#1F4E79", fg="white").pack()
            tk.Label(frame, text="DL이앤씨  플랜트본부 기계설계팀",
                     font=("맑은 고딕", 10), bg="#1F4E79", fg="#9DC3E6").pack(pady=4)
            tk.Label(frame, text=self._version, font=("맑은 고딕", 9),
                     bg="#1F4E79", fg="#5B9BD5").pack(side="bottom", pady=10)

        sw, sh = self._win.winfo_screenwidth(), self._win.winfo_screenheight()
        self._win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self._win.update()

    def destroy(self):
        self._win.destroy()


# ══════════════════════════════════════════════════════════
# 메인
# ══════════════════════════════════════════════════════════

def main():
    from config import load_config
    cfg  = load_config()
    root = tk.Tk()
    root.withdraw()

    splash = SplashScreen(root, cfg)

    def launch():
        try:
            splash.destroy()
            root.deiconify()
            icon_path = resource_path(cfg.get("icon_file", "assets/Schedule_Ico.ico"))
            if os.path.exists(icon_path):
                try: root.iconbitmap(icon_path)
                except Exception: pass
            from app import WeeklyReportApp
            WeeklyReportApp(root)
        except Exception:
            import traceback
            from tkinter import messagebox
            messagebox.showerror(
                "실행 오류",
                "프로그램 시작 중 오류가 발생했습니다:\n\n"
                + traceback.format_exc()
            )
            root.destroy()

    root.after(SplashScreen.DURATION_MS, launch)
    root.mainloop()


if __name__ == "__main__":
    main()
