"""
gui.py — 視窗版自動化流程控制介面（含步驟編輯）
執行：python gui.py
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import sys
import json
import time
import ctypes
import ctypes.wintypes
from pathlib import Path
import auto_clicker as ac


# ════════════════════════════════════════════════════════════
#  Windows 視窗列舉（ctypes，無需額外安裝套件）
# ════════════════════════════════════════════════════════════

def list_windows() -> list[tuple[int, str]]:
    """回傳所有可見且有標題的視窗：[(hwnd, title), ...]"""
    results = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    def _cb(hwnd, _):
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length:
                buf = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
                results.append((hwnd, buf.value))
        return True

    ctypes.windll.user32.EnumWindows(_cb, 0)
    return results


def get_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    """回傳視窗目前位置 (x, y, width, height)"""
    rect = ctypes.wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return (rect.left, rect.top,
            rect.right - rect.left,
            rect.bottom - rect.top)

# ════════════════════════════════════════════════════════════
#  常數
# ════════════════════════════════════════════════════════════

STEPS_FILE = "steps.json"

ACTIONS = [
    "find_and_click",
    "wait_for_image",
    "wait_for_image_gone",
    "image_exists",
    "move",
    "click_xy",
    "scroll",
    "sleep",
]
ACTION_LABELS = {
    "find_and_click":      "圖片點擊",
    "wait_for_image":      "等待出現",
    "wait_for_image_gone": "等待消失",
    "image_exists":        "確認存在",
    "move":                "移動滑鼠",
    "click_xy":            "座標點擊",
    "scroll":              "滾輪",
    "sleep":               "等待(秒)",
    "rename_pdf":          "重新命名PDF",
}
ON_FAIL_OPTIONS       = ["stop", "skip", "retry"]
ON_FAIL_DISPLAY       = ["停止", "跳過", "重試一次"]
ON_FAIL_TO_DISPLAY    = dict(zip(ON_FAIL_OPTIONS, ON_FAIL_DISPLAY))
DISPLAY_TO_ON_FAIL    = dict(zip(ON_FAIL_DISPLAY, ON_FAIL_OPTIONS))

CLICK_TYPES           = ["left", "right", "double"]
CLICK_TYPE_LABELS     = {"left": "左鍵", "right": "右鍵", "double": "雙擊"}
CLICK_TYPE_DISPLAY    = ["左鍵", "右鍵", "雙擊"]
DISPLAY_TO_CLICK_TYPE = dict(zip(CLICK_TYPE_DISPLAY, CLICK_TYPES))

# 動作：中文顯示 ↔ 內部 key
ACTION_DISPLAY        = list(ACTION_LABELS.values())          # 中文清單（combobox 用）
DISPLAY_TO_ACTION     = {v: k for k, v in ACTION_LABELS.items()}  # 中文 → key
ACTION_TO_DISPLAY     = ACTION_LABELS                          # key → 中文（同 ACTION_LABELS）

# 各動作需要的欄位群組
_NEEDS_TEMPLATE   = {"find_and_click", "wait_for_image", "wait_for_image_gone", "image_exists"}
_NEEDS_TIMEOUT    = {"find_and_click", "wait_for_image", "wait_for_image_gone", "sleep"}
_NEEDS_COORD      = {"move", "click_xy", "scroll"}
_NEEDS_SCROLL     = {"scroll"}
_NEEDS_CLICK_TYPE = {"click_xy"}
_NEEDS_FOLDER     = {"rename_pdf"}


# ════════════════════════════════════════════════════════════
#  步驟資料 — 讀寫 steps.json
# ════════════════════════════════════════════════════════════

def _empty_step(n: int) -> dict:
    return {
        "name":       f"步驟{n:02d}",
        "action":     "find_and_click",
        "template":   "",
        "on_fail":    "stop",
        "enabled":    True,
        "timeout":    10.0,
        "confidence": 0.8,
    }


def load_steps() -> list:
    p = Path(STEPS_FILE)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []


def save_steps(steps: list):
    Path(STEPS_FILE).write_text(
        json.dumps(steps, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ════════════════════════════════════════════════════════════
#  步驟執行器
# ════════════════════════════════════════════════════════════

def _rename_pdfs_in_folder(folder: str) -> bool:
    """
    掃描資料夾內所有 PDF，擷取第一行文字作為新檔名並重新命名。
    回傳 True = 至少成功處理一個；False = 沒有 PDF 或全部失敗。
    """
    try:
        import pdfplumber
    except ImportError:
        print("  ❌ 缺少套件，請執行：pip install pdfplumber")
        return False

    import re
    folder_path = Path(folder)
    if not folder_path.exists():
        print(f"  ❌ 找不到資料夾：{folder}")
        return False

    # Windows 檔名不允許的字元
    _INVALID = re.compile(r'[\\/:*?"<>|]')

    pdfs = list(folder_path.glob("*.pdf"))
    if not pdfs:
        print(f"  ⚠️  資料夾內沒有 PDF：{folder}")
        return False

    success = 0
    for pdf_path in pdfs:
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                text = pdf.pages[0].extract_text() or ""
            first_line = text.strip().splitlines()[0].strip() if text.strip() else ""
            if not first_line:
                print(f"  ⚠️  無法擷取文字，略過：{pdf_path.name}")
                continue
            new_name = _INVALID.sub("_", first_line) + ".pdf"
            new_path = pdf_path.parent / new_name
            if new_path == pdf_path:
                print(f"  ℹ️  已是正確檔名，略過：{pdf_path.name}")
                success += 1
                continue
            if new_path.exists():
                print(f"  ⚠️  目標已存在，略過：{new_name}")
                continue
            pdf_path.rename(new_path)
            print(f"  ✅ {pdf_path.name}")
            print(f"     → {new_name}")
            success += 1
        except Exception as e:
            print(f"  ❌ {pdf_path.name}：{e}")

    print(f"\n  📄 完成：{success}/{len(pdfs)} 個 PDF 重新命名")
    return success > 0


def _execute(step: dict) -> bool:
    import pyautogui
    action  = step["action"]
    tmpl    = step.get("template", "")
    conf    = float(step.get("confidence", 0.8))
    timeout = float(step.get("timeout", 10.0))
    name    = step.get("name", action)
    x       = int(step.get("x", 0))
    y       = int(step.get("y", 0))

    if action == "find_and_click":
        return bool(ac.find_and_click(tmpl, confidence=conf, wait_timeout=timeout))
    elif action == "wait_for_image":
        return ac.wait_for_image(tmpl, confidence=conf, timeout=timeout)
    elif action == "wait_for_image_gone":
        return ac.wait_for_image_gone(tmpl, confidence=conf, timeout=timeout)
    elif action == "image_exists":
        ok = ac.image_exists(tmpl, confidence=conf)
        print(f"  {'✅ 確認存在' if ok else '❌ 不存在'}：{Path(tmpl).name}")
        return ok
    elif action == "sleep":
        ac.sleep(timeout, name)
        return True
    elif action == "move":
        pyautogui.moveTo(x, y)
        print(f"  🖱️  移動到 ({x}, {y})")
        return True
    elif action == "click_xy":
        click_type = step.get("click_type", "left")
        if click_type == "right":
            pyautogui.rightClick(x, y)
        elif click_type == "double":
            pyautogui.doubleClick(x, y)
        else:
            pyautogui.click(x, y)
        label = CLICK_TYPE_LABELS.get(click_type, click_type)
        print(f"  🖱️  {label} 點擊 ({x}, {y})")
        return True
    elif action == "rename_pdf":
        return _rename_pdfs_in_folder(step.get("folder", ""))
    elif action == "scroll":
        amount = int(step.get("scroll_amount", 3))
        direction = "向上" if amount > 0 else "向下"
        if x or y:
            pyautogui.scroll(amount, x=x, y=y)
            print(f"  🖱️  滾輪 {direction} {abs(amount)} 格，位置 ({x}, {y})")
        else:
            pyautogui.scroll(amount)
            print(f"  🖱️  滾輪 {direction} {abs(amount)} 格（目前滑鼠位置）")
        return True
    return False


def _run_with_retry(step: dict) -> bool:
    on_fail = step.get("on_fail", "stop")
    try:
        if _execute(step):
            print("  ✅ 成功")
            return True

        if on_fail == "skip":
            print("  ⚠️  失敗，跳過（on_fail=skip）")
            return True
        elif on_fail == "retry":
            print("  🔄 失敗，重試一次...")
            time.sleep(1)
            if _execute(step):
                print("  ✅ 重試成功")
                return True
            print("  ❌ 重試仍失敗")
        return False

    except FileNotFoundError as e:
        print(f"  ❌ 找不到截圖：{e}")
        return on_fail == "skip"
    except Exception as e:
        print(f"  ❌ 錯誤：{e}")
        return False


# ════════════════════════════════════════════════════════════
#  stdout → Text 元件
# ════════════════════════════════════════════════════════════

class _StdoutRedirector:
    def __init__(self, widget):
        self._w = widget

    def write(self, msg: str):
        self._w.configure(state="normal")
        self._w.insert(tk.END, msg)
        self._w.see(tk.END)
        self._w.configure(state="disabled")

    def flush(self):
        pass


# ════════════════════════════════════════════════════════════
#  步驟編輯對話框
# ════════════════════════════════════════════════════════════

class StepDialog(tk.Toplevel):
    """新增 / 編輯步驟 Modal 對話框（欄位依動作動態顯示）"""

    def __init__(self, parent, step: dict | None = None):
        super().__init__(parent)
        self.title("新增步驟" if step is None else "編輯步驟")
        self.resizable(False, False)
        self.grab_set()
        self.result = None

        s = step.copy() if step else _empty_step(1)
        PAD = {"padx": 10, "pady": 3}

        # ── 固定欄位：名稱 ──
        f = tk.Frame(self); f.pack(fill=tk.X, **PAD)
        tk.Label(f, text="名稱", width=10, anchor="e").pack(side=tk.LEFT)
        self._name = tk.Entry(f, width=28)
        self._name.insert(0, s.get("name", ""))
        self._name.pack(side=tk.LEFT, padx=(4, 0))

        # ── 固定欄位：動作 ──
        f = tk.Frame(self); f.pack(fill=tk.X, **PAD)
        tk.Label(f, text="動作", width=10, anchor="e").pack(side=tk.LEFT)
        self._action_var = tk.StringVar(
            value=ACTION_TO_DISPLAY.get(s.get("action", "find_and_click"), "圖片點擊"))
        cb = ttk.Combobox(f, textvariable=self._action_var,
                          values=ACTION_DISPLAY, state="readonly", width=14)
        cb.pack(side=tk.LEFT, padx=(4, 0))
        cb.bind("<<ComboboxSelected>>", self._on_action_change)

        # ── 可切換群組：模板 ──
        self._tmpl_frame = tk.Frame(self)
        tk.Label(self._tmpl_frame, text="模板檔案", width=10, anchor="e").pack(side=tk.LEFT)
        self._tmpl = tk.Entry(self._tmpl_frame, width=24)
        self._tmpl.insert(0, s.get("template", ""))
        self._tmpl.pack(side=tk.LEFT, padx=(4, 4))
        tk.Button(self._tmpl_frame, text="瀏覽…", command=self._browse).pack(side=tk.LEFT)

        # ── 可切換群組：相似度 ──
        self._conf_frame = tk.Frame(self)
        tk.Label(self._conf_frame, text="相似度", width=10, anchor="e").pack(side=tk.LEFT)
        self._conf = tk.Spinbox(self._conf_frame, from_=0.10, to=1.00,
                                increment=0.05, format="%.2f", width=7)
        self._conf.delete(0, tk.END)
        self._conf.insert(0, str(s.get("confidence", 0.8)))
        self._conf.pack(side=tk.LEFT, padx=(4, 0))

        # ── 可切換群組：座標 X / Y ──
        self._coord_frame = tk.Frame(self)
        tk.Label(self._coord_frame, text="座標 X", width=10, anchor="e").pack(side=tk.LEFT)
        self._x = tk.Entry(self._coord_frame, width=7)
        self._x.insert(0, str(s.get("x", 0)))
        self._x.pack(side=tk.LEFT, padx=(4, 12))
        tk.Label(self._coord_frame, text="Y", width=2, anchor="e").pack(side=tk.LEFT)
        self._y = tk.Entry(self._coord_frame, width=7)
        self._y.insert(0, str(s.get("y", 0)))
        self._y.pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(self._coord_frame, text="  (0,0 = 目前位置)",
                 fg="#888").pack(side=tk.LEFT, padx=4)

        # ── 可切換群組：滾動量 ──
        self._scroll_frame = tk.Frame(self)
        tk.Label(self._scroll_frame, text="滾動量", width=10, anchor="e").pack(side=tk.LEFT)
        self._scroll_amount = tk.Spinbox(self._scroll_frame, from_=-20, to=20,
                                         increment=1, width=5)
        self._scroll_amount.delete(0, tk.END)
        self._scroll_amount.insert(0, str(s.get("scroll_amount", 3)))
        self._scroll_amount.pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(self._scroll_frame, text="格（正數=向上，負數=向下）",
                 fg="#888").pack(side=tk.LEFT, padx=4)

        # ── 可切換群組：點擊方式 ──
        self._click_type_frame = tk.Frame(self)
        tk.Label(self._click_type_frame, text="點擊方式", width=10, anchor="e").pack(side=tk.LEFT)
        self._click_type_var = tk.StringVar(
            value=CLICK_TYPE_LABELS.get(s.get("click_type", "left"), "左鍵"))
        ttk.Combobox(self._click_type_frame, textvariable=self._click_type_var,
                     values=CLICK_TYPE_DISPLAY, state="readonly",
                     width=10).pack(side=tk.LEFT, padx=(4, 0))

        # ── 可切換群組：資料夾路徑 ──
        self._folder_frame = tk.Frame(self)
        tk.Label(self._folder_frame, text="PDF 資料夾", width=10, anchor="e").pack(side=tk.LEFT)
        self._folder = tk.Entry(self._folder_frame, width=28)
        self._folder.insert(0, s.get("folder", ""))
        self._folder.pack(side=tk.LEFT, padx=(4, 4))
        tk.Button(self._folder_frame, text="瀏覽…",
                  command=self._browse_folder).pack(side=tk.LEFT)

        # ── 可切換群組：逾時 / 等待秒數 ──
        self._timeout_frame = tk.Frame(self)
        self._timeout_label_var = tk.StringVar()
        tk.Label(self._timeout_frame, textvariable=self._timeout_label_var,
                 width=10, anchor="e").pack(side=tk.LEFT)
        self._timeout = tk.Spinbox(self._timeout_frame, from_=0, to=600,
                                   increment=1, width=7)
        self._timeout.delete(0, tk.END)
        self._timeout.insert(0, str(int(s.get("timeout", 10))))
        self._timeout.pack(side=tk.LEFT, padx=(4, 0))

        # ── 固定欄位：失敗行為 ──
        f = tk.Frame(self); f.pack(fill=tk.X, **PAD)
        tk.Label(f, text="失敗行為", width=10, anchor="e").pack(side=tk.LEFT)
        self._on_fail_var = tk.StringVar(
            value=ON_FAIL_TO_DISPLAY.get(s.get("on_fail", "stop"), "停止"))
        ttk.Combobox(f, textvariable=self._on_fail_var,
                     values=ON_FAIL_DISPLAY, state="readonly",
                     width=10).pack(side=tk.LEFT, padx=(4, 0))

        # ── 固定欄位：啟用 ──
        self._enabled_var = tk.BooleanVar(value=s.get("enabled", True))
        tk.Checkbutton(self, text="啟用此步驟",
                       variable=self._enabled_var).pack(**PAD)

        # ── 按鈕 ──
        f = tk.Frame(self); f.pack(pady=8)
        tk.Button(f, text="確定", width=10, bg="#27ae60", fg="white",
                  command=self._ok).pack(side=tk.LEFT, padx=6)
        tk.Button(f, text="取消", width=10,
                  command=self.destroy).pack(side=tk.LEFT, padx=6)

        self._on_action_change()
        self.wait_window()

    # ── 動態顯示/隱藏欄位群組 ──────────────────────────────

    def _on_action_change(self, *_):
        action = DISPLAY_TO_ACTION.get(self._action_var.get(), self._action_var.get())
        PAD = {"fill": tk.X, "padx": 10, "pady": 3}

        def show(frame): frame.pack(PAD)
        def hide(frame): frame.pack_forget()

        # 模板 + 相似度
        if action in _NEEDS_TEMPLATE:
            show(self._tmpl_frame)
            show(self._conf_frame)
        else:
            hide(self._tmpl_frame)
            hide(self._conf_frame)

        # 座標
        if action in _NEEDS_COORD:
            show(self._coord_frame)
        else:
            hide(self._coord_frame)

        # 滾動量
        if action in _NEEDS_SCROLL:
            show(self._scroll_frame)
        else:
            hide(self._scroll_frame)

        # 點擊方式
        if action in _NEEDS_CLICK_TYPE:
            show(self._click_type_frame)
        else:
            hide(self._click_type_frame)

        # 資料夾
        if action in _NEEDS_FOLDER:
            show(self._folder_frame)
        else:
            hide(self._folder_frame)

        # 逾時 / 等待秒數
        if action in _NEEDS_TIMEOUT:
            self._timeout_label_var.set("等待秒數" if action == "sleep" else "逾時秒數")
            show(self._timeout_frame)
        else:
            hide(self._timeout_frame)

        # 重新計算視窗大小
        self.update_idletasks()
        self.geometry("")

    def _browse(self):
        path = filedialog.askopenfilename(
            title="選擇模板圖片",
            filetypes=[("PNG 圖片", "*.png"), ("所有檔案", "*.*")],
            initialdir="templates",
        )
        if path:
            self._tmpl.delete(0, tk.END)
            self._tmpl.insert(0, path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="選擇 PDF 資料夾")
        if path:
            self._folder.delete(0, tk.END)
            self._folder.insert(0, path)

    def _ok(self):
        self.result = {
            "name":          self._name.get().strip() or "未命名",
            "action":        DISPLAY_TO_ACTION.get(self._action_var.get(), self._action_var.get()),
            "template":      self._tmpl.get().strip(),
            "on_fail":       DISPLAY_TO_ON_FAIL.get(self._on_fail_var.get(), self._on_fail_var.get()),
            "enabled":       self._enabled_var.get(),
            "timeout":       float(self._timeout.get() or 10),
            "confidence":    float(self._conf.get() or 0.8),
            "x":             int(self._x.get() or 0),
            "y":             int(self._y.get() or 0),
            "scroll_amount": int(self._scroll_amount.get() or 3),
            "click_type":    DISPLAY_TO_CLICK_TYPE.get(self._click_type_var.get(), self._click_type_var.get()),
            "folder":        self._folder.get().strip(),
        }
        self.destroy()


# ════════════════════════════════════════════════════════════
#  主視窗
# ════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pixel Pilot — 影像辨識自動化")
        self.geometry("940x680")
        self.minsize(700, 500)

        self._steps: list[dict] = load_steps()
        self._stop_event = threading.Event()
        self._running = False
        self._loop_var = tk.BooleanVar(value=False)   # 循環執行
        self._hwnd: int | None = None          # 選定視窗的 HWND
        self._win_map: dict[str, int] = {}     # 顯示名稱 → hwnd

        self._build_ui()
        self._refresh_tree()
        self._refresh_windows()

    # ── 建構 UI ─────────────────────────────────────────────

    def _build_ui(self):
        # ── 頂部工具列 ──
        ctrl = tk.Frame(self, pady=6)
        ctrl.pack(fill=tk.X, padx=10)

        # 執行群組
        run_box = tk.LabelFrame(ctrl, text="執行", padx=6, pady=2)
        run_box.pack(side=tk.LEFT)

        tk.Label(run_box, text="從步驟").pack(side=tk.LEFT)
        self._start_var = tk.IntVar(value=1)
        self._spin = tk.Spinbox(run_box, from_=1, to=max(len(self._steps), 1),
                                width=4, textvariable=self._start_var)
        self._spin.pack(side=tk.LEFT, padx=(2, 8))

        self._btn_start = tk.Button(
            run_box, text="▶ 開始", width=8,
            bg="#27ae60", fg="white", font=("", 9, "bold"),
            command=self._on_start,
        )
        self._btn_start.pack(side=tk.LEFT)

        self._btn_stop = tk.Button(
            run_box, text="⏹ 停止", width=8,
            bg="#e74c3c", fg="white", state=tk.DISABLED,
            command=self._on_stop,
        )
        self._btn_stop.pack(side=tk.LEFT, padx=(4, 6))

        self._status = tk.Label(run_box, text="就緒", fg="#555", width=18, anchor="w")
        self._status.pack(side=tk.LEFT)

        tk.Checkbutton(run_box, text="循環執行",
                       variable=self._loop_var).pack(side=tk.LEFT, padx=(8, 0))

        # 編輯群組
        edit_box = tk.LabelFrame(ctrl, text="步驟管理", padx=6, pady=2)
        edit_box.pack(side=tk.LEFT, padx=10)

        for text, cmd in [
            ("＋ 新增", self._add_step),
            ("✏ 編輯", self._edit_step),
            ("⧉ 複製", self._copy_step),
            ("✕ 刪除", self._delete_step),
            ("↑ 上移", self._move_up),
            ("↓ 下移", self._move_down),
        ]:
            tk.Button(edit_box, text=text, width=7,
                      command=cmd).pack(side=tk.LEFT, padx=2)

        tk.Button(ctrl, text="💾 儲存", width=8,
                  bg="#2980b9", fg="white",
                  command=self._save).pack(side=tk.RIGHT)

        # ── 視窗選擇器 ──
        win_box = tk.LabelFrame(self, text="目標視窗（鎖定後只截該視窗範圍）",
                                padx=8, pady=4)
        win_box.pack(fill=tk.X, padx=10, pady=(0, 4))

        self._win_var = tk.StringVar(value="（全螢幕）")
        self._win_cb = ttk.Combobox(win_box, textvariable=self._win_var,
                                    state="readonly", width=55)
        self._win_cb.pack(side=tk.LEFT, padx=(0, 6))
        self._win_cb.bind("<<ComboboxSelected>>", self._on_window_selected)

        tk.Button(win_box, text="🔄 重新整理",
                  command=self._refresh_windows).pack(side=tk.LEFT)

        tk.Button(win_box, text="✕ 取消鎖定",
                  command=self._clear_window).pack(side=tk.LEFT, padx=6)

        self._win_status = tk.Label(win_box, text="", fg="#2980b9")
        self._win_status.pack(side=tk.LEFT, padx=8)

        # ── 步驟清單 ──
        frame_steps = tk.LabelFrame(self, text="步驟清單（雙擊列可編輯）",
                                    padx=6, pady=4)
        frame_steps.pack(fill=tk.X, padx=10, pady=(0, 4))

        cols = ("#", "啟用", "步驟名稱", "動作", "模板", "失敗行為", "狀態")
        self._tree = ttk.Treeview(frame_steps, columns=cols,
                                   show="headings", height=10)

        col_cfg = {
            "#":     (36,  "center", False),
            "啟用":  (44,  "center", False),
            "步驟名稱": (160, "w",   False),
            "動作":  (90,  "center", False),
            "模板":  (240, "w",     True),
            "失敗行為": (72, "center", False),
            "狀態":  (90,  "center", False),
        }
        for col, (w, anchor, stretch) in col_cfg.items():
            self._tree.heading(col, text=col)
            self._tree.column(col, width=w, anchor=anchor, stretch=stretch)

        self._tree.pack(fill=tk.X)
        self._tree.bind("<Double-1>", lambda _: self._edit_step())

        self._tree.tag_configure("waiting",  background="white")
        self._tree.tag_configure("running",  background="#fff9c4")
        self._tree.tag_configure("success",  background="#c8e6c9")
        self._tree.tag_configure("failed",   background="#ffcdd2")
        self._tree.tag_configure("skipped",  background="#eeeeee")
        self._tree.tag_configure("disabled", foreground="#aaaaaa")

        # ── Log ──
        frame_log = tk.LabelFrame(self, text="執行 Log", padx=6, pady=4)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        tk.Button(frame_log, text="清除",
                  command=self._clear_log).pack(anchor="ne")

        self._log = scrolledtext.ScrolledText(
            frame_log, state="disabled",
            font=("Consolas", 9), height=12,
        )
        self._log.pack(fill=tk.BOTH, expand=True)

    # ── 步驟清單工具 ────────────────────────────────────────

    def _refresh_tree(self):
        for item in self._tree.get_children():
            self._tree.delete(item)
        for i, s in enumerate(self._steps, 1):
            enabled = s.get("enabled", True)
            self._tree.insert(
                "", tk.END, iid=str(i),
                values=(
                    f"{i:02d}",
                    "✔" if enabled else "✘",
                    s["name"],
                    ACTION_LABELS.get(s["action"], s["action"]),
                    s.get("template", ""),
                    s.get("on_fail", "stop"),
                    "⏸ 等待",
                ),
                tags=("waiting" if enabled else "disabled",),
            )
        self._spin.config(to=max(len(self._steps), 1))

    def _set_row_status(self, one_based: int, label: str, tag: str):
        iid = str(one_based)
        if not self._tree.exists(iid):
            return
        vals = list(self._tree.item(iid)["values"])
        vals[6] = label
        self._tree.item(iid, values=vals, tags=(tag,))

    def _selected_index(self) -> int | None:
        """回傳 0-based index；無選取時回傳 None"""
        sel = self._tree.selection()
        return int(sel[0]) - 1 if sel else None

    # ── 編輯操作 ────────────────────────────────────────────

    def _add_step(self):
        dlg = StepDialog(self)
        if dlg.result:
            idx = self._selected_index()
            pos = len(self._steps) if idx is None else idx + 1
            self._steps.insert(pos, dlg.result)
            self._refresh_tree()
            self._tree.selection_set(str(pos + 1))

    def _edit_step(self):
        idx = self._selected_index()
        if idx is None:
            return
        dlg = StepDialog(self, self._steps[idx])
        if dlg.result:
            self._steps[idx] = dlg.result
            self._refresh_tree()
            self._tree.selection_set(str(idx + 1))

    def _copy_step(self):
        idx = self._selected_index()
        if idx is None:
            return
        import copy
        self._steps.append(copy.deepcopy(self._steps[idx]))
        self._refresh_tree()
        self._tree.selection_set(str(len(self._steps)))  # 選中最後一列

    def _delete_step(self):
        idx = self._selected_index()
        if idx is None:
            return
        name = self._steps[idx]["name"]
        if messagebox.askyesno("確認刪除", f"確定要刪除「{name}」？"):
            self._steps.pop(idx)
            self._refresh_tree()

    def _move_up(self):
        idx = self._selected_index()
        if idx is None or idx == 0:
            return
        self._steps[idx - 1], self._steps[idx] = (
            self._steps[idx], self._steps[idx - 1])
        self._refresh_tree()
        self._tree.selection_set(str(idx))      # 移動後新位置（1-based = idx）

    def _move_down(self):
        idx = self._selected_index()
        if idx is None or idx >= len(self._steps) - 1:
            return
        self._steps[idx], self._steps[idx + 1] = (
            self._steps[idx + 1], self._steps[idx])
        self._refresh_tree()
        self._tree.selection_set(str(idx + 2))  # 移動後新位置（1-based = idx+2）

    # ── 視窗選擇 ────────────────────────────────────────────

    def _refresh_windows(self):
        wins = list_windows()
        self._win_map = {}
        labels = ["（全螢幕）"]
        for hwnd, title in wins:
            label = f"{title}  [{hwnd}]"
            self._win_map[label] = hwnd
            labels.append(label)
        self._win_cb["values"] = labels
        # 保留目前選取（若視窗還存在）
        if self._win_var.get() not in labels:
            self._win_var.set("（全螢幕）")
            self._clear_window()

    def _on_window_selected(self, *_):
        label = self._win_var.get()
        if label == "（全螢幕）":
            self._clear_window()
            return
        hwnd = self._win_map.get(label)
        if not hwnd:
            return
        self._hwnd = hwnd
        self._apply_window_region()

    def _apply_window_region(self):
        """讀取目前視窗位置並更新截圖區域"""
        if not self._hwnd:
            return
        try:
            region = get_window_rect(self._hwnd)
            ac.set_capture_region(region)
            self._win_status.config(
                text=f"🪟 {region[0]},{region[1]}  {region[2]}×{region[3]}")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法取得視窗位置：{e}")

    def _clear_window(self):
        self._hwnd = None
        ac.set_capture_region(None)
        self._win_var.set("（全螢幕）")
        self._win_status.config(text="")

    def _save(self):
        save_steps(self._steps)
        self._status.config(text="✅ 已儲存")
        self.after(2000, lambda: self._status.config(text="就緒"))

    # ── 執行控制 ────────────────────────────────────────────

    def _on_start(self):
        if self._running:
            return
        if not self._steps:
            messagebox.showinfo("提示", "請先新增步驟")
            return
        self._stop_event.clear()
        self._running = True
        self._refresh_tree()
        self._btn_start.config(state=tk.DISABLED)
        self._btn_stop.config(state=tk.NORMAL)
        self._status.config(text="執行中...")
        # 執行前重新抓視窗位置（視窗可能被移動過）
        if self._hwnd:
            self._apply_window_region()
        threading.Thread(target=self._run_workflow,
                         args=(self._start_var.get(),),
                         daemon=True).start()

    def _on_stop(self):
        self._stop_event.set()
        self._status.config(text="停止中…")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete(1.0, tk.END)
        self._log.configure(state="disabled")

    # ── 背景執行緒 ──────────────────────────────────────────

    def _run_workflow(self, start_from: int):
        old_stdout = sys.stdout
        sys.stdout = _StdoutRedirector(self._log)
        try:
            loop_count = 0
            loop_mode = self._loop_var.get()

            while True:
                loop_count += 1
                active = [
                    (i + 1, s) for i, s in enumerate(self._steps)
                    if s.get("enabled", True)
                ]
                total = len(active)

                if loop_mode:
                    print(f"\n{'╔'+'═'*53}")
                    print(f"  🔁 第 {loop_count} 輪  （共 {total} 個步驟）"
                          f"  ⚠️ 緊急停止：滑鼠移到左上角")
                    print(f"{'╚'+'═'*53}\n")
                else:
                    print(f"🚀 開始執行（共 {total} 個啟用步驟）"
                          f"  ⚠️ 緊急停止：滑鼠移到螢幕左上角\n")

                # 重置步驟狀態（第二輪起才需要）
                if loop_count > 1:
                    self.after(0, self._refresh_tree)

                stopped_at = None
                _start = start_from if loop_count == 1 else 1

                for orig_idx, step in active:
                    if self._stop_event.is_set():
                        print("\n⏹ 使用者手動停止")
                        break

                    if orig_idx < _start:
                        self.after(0, self._set_row_status, orig_idx, "⏭ 略過", "skipped")
                        print(f"  ⏭️  步驟 {orig_idx:02d} 略過（已完成）")
                        continue

                    action_label = ACTION_LABELS.get(step["action"], step["action"])
                    self.after(0, self._set_row_status, orig_idx, "⏳ 執行中", "running")
                    self.after(0, self._status.config,
                               {"text": f"第{loop_count}輪 步驟{orig_idx}/{len(self._steps)}"
                                        if loop_mode else f"步驟 {orig_idx}/{len(self._steps)}"})

                    print(f"\n{'━'*55}")
                    print(f"  步驟 {orig_idx:02d}  {step['name']}  [{action_label}]")
                    print(f"{'━'*55}")

                    success = _run_with_retry(step)

                    if success:
                        self.after(0, self._set_row_status, orig_idx, "✅ 成功", "success")
                    else:
                        self.after(0, self._set_row_status, orig_idx, "❌ 失敗", "failed")
                        print(f"\n❌ 流程在步驟 {orig_idx:02d} 停止")
                        stopped_at = orig_idx
                        break

                # for 正常跑完（沒有 break）
                else:
                    print(f"\n{'═'*55}")
                    if loop_mode:
                        print(f"  ✅ 第 {loop_count} 輪完成，準備下一輪…")
                    else:
                        print(f"  ✅ 全部步驟執行完畢！")
                    print(f"{'═'*55}")

                    if not loop_mode or self._stop_event.is_set():
                        self.after(0, self._status.config, {"text": "✅ 完成"})
                        break
                    # 循環模式：繼續下一輪
                    continue

                # 有 break（步驟失敗或手動停止）
                if stopped_at:
                    self.after(0, self._status.config,
                               {"text": f"❌ 步驟 {stopped_at} 失敗"})
                elif self._stop_event.is_set():
                    self.after(0, self._status.config, {"text": "⏹ 已停止"})
                break  # 離開 while

        finally:
            sys.stdout = old_stdout
            self._running = False
            self.after(0, self._btn_start.config, {"state": tk.NORMAL})
            self.after(0, self._btn_stop.config,  {"state": tk.DISABLED})


# ════════════════════════════════════════════════════════════
#  入口
# ════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
