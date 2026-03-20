"""
auto_clicker.py — 影像辨識點擊核心函式庫

安裝：pip install pyautogui opencv-python pillow numpy
"""

import cv2
import numpy as np
import pyautogui
import time
import os
import sys
from pathlib import Path

# ── 全域安全設定 ──────────────────────────────────────────────
pyautogui.FAILSAFE = True   # 滑鼠移到左上角 → 立即停止
pyautogui.PAUSE = 0.3       # 每個操作後暫停 0.3 秒

# ── 截圖區域（None = 全螢幕；由 gui.py 透過 set_capture_region 設定）────
_capture_region: tuple | None = None  # (x, y, width, height)


def set_capture_region(region: tuple | None):
    """設定截圖範圍 (x, y, width, height)；傳 None 恢復全螢幕。"""
    global _capture_region
    _capture_region = region
    if region:
        print(f"  🪟 截圖區域：{region[0]},{region[1]}  {region[2]}×{region[3]}")
    else:
        print("  🖥️  截圖區域：全螢幕")


# ── 主要函式 ─────────────────────────────────────────────────

def find_and_click(
    template_path: str,
    confidence: float = 0.8,
    click_type: str = "left",
    offset_x: int = 0,
    offset_y: int = 0,
    wait_timeout: float = 10.0,
    debug: bool = False,
    click_color_center: bool = False,
) -> tuple | None:
    """
    在螢幕上找到指定圖示並點擊。
    click_type: "left" / "right" / "double"
    click_color_center: True 時點有顏色的重心而非幾何中心
    回傳：找到時 (x, y)，找不到 None
    """
    pos = find_only(template_path, confidence, offset_x, offset_y,
                    wait_timeout, debug, click_color_center)
    if pos is None:
        return None

    x, y = pos
    if click_type == "left":
        pyautogui.click(x, y)
    elif click_type == "right":
        pyautogui.rightClick(x, y)
    elif click_type == "double":
        pyautogui.doubleClick(x, y)

    print(f"  🖱️  {click_type} 點擊 ({x}, {y})")
    return pos


def find_only(
    template_path: str,
    confidence: float = 0.8,
    offset_x: int = 0,
    offset_y: int = 0,
    wait_timeout: float = 10.0,
    debug: bool = False,
    click_color_center: bool = False,
) -> tuple | None:
    """
    找到圖示位置但不點擊，回傳 (x, y) 中心座標。
    click_color_center: True 時回傳有顏色像素的重心而非幾何中心。
    會持續等待直到超時。
    """
    _check_file(template_path)
    template = _imread(template_path)
    if template is None:
        print(f"❌ 無法讀取圖片：{template_path}")
        return None

    h, w = template.shape[:2]
    start = time.time()

    while True:
        screen = _screenshot()
        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)

        if debug:
            _show_debug(screen, template, max_loc, max_val, confidence)

        if max_val >= confidence:
            scale = _dpi_scale()
            rx = _capture_region[0] if _capture_region else 0
            ry = _capture_region[1] if _capture_region else 0

            if click_color_center:
                raw_x, raw_y = _color_center(screen, max_loc, w, h)
                note = "顏色重心"
            else:
                raw_x = max_loc[0] + w // 2
                raw_y = max_loc[1] + h // 2
                note = ""

            cx = int(raw_x / scale) + offset_x + rx
            cy = int(raw_y / scale) + offset_y + ry
            print(f"  ✅ 找到 {Path(template_path).name}  相似度 {max_val:.0%}  "
                  f"位置 ({cx}, {cy})"
                  + (f"  [{note}]" if note else "")
                  + (f"  DPI×{scale:.2f}" if scale != 1.0 else ""))
            if debug:
                cv2.destroyAllWindows()
            return (cx, cy)

        elapsed = time.time() - start
        if elapsed >= wait_timeout:
            print(f"  ⏰ 超時 {wait_timeout}s，找不到 {Path(template_path).name}（最高相似度 {max_val:.0%}）")
            print(f"     💡 試試：降低 confidence（目前 {confidence}）或重新截圖")
            if debug:
                cv2.destroyAllWindows()
            return None

        print(f"  🔍 搜尋 {Path(template_path).name}... {max_val:.0%}/{confidence:.0%}  ({elapsed:.1f}s)")
        time.sleep(0.5)


def image_exists(template_path: str, confidence: float = 0.8) -> bool:
    """
    快速判斷圖示當前是否在螢幕上，不等待，立即回傳 True/False。
    """
    _check_file(template_path)
    template = _imread(template_path)
    if template is None:
        return False
    screen = _screenshot()
    result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, _ = cv2.minMaxLoc(result)
    return max_val >= confidence


def wait_for_image(
    template_path: str,
    confidence: float = 0.8,
    timeout: float = 30.0,
) -> bool:
    """
    等待某圖示出現（例如載入完成的按鈕）。
    出現回傳 True，超時回傳 False。
    """
    print(f"  ⏳ 等待出現：{Path(template_path).name}")
    pos = find_only(template_path, confidence=confidence, wait_timeout=timeout)
    return pos is not None


def wait_for_image_gone(
    template_path: str,
    confidence: float = 0.8,
    timeout: float = 30.0,
) -> bool:
    """
    等待某圖示消失（例如等載入動畫消失）。
    消失回傳 True，超時回傳 False。
    """
    print(f"  ⏳ 等待消失：{Path(template_path).name}")
    start = time.time()
    while True:
        if not image_exists(template_path, confidence):
            print(f"  ✅ 已消失：{Path(template_path).name}")
            return True
        elapsed = time.time() - start
        if elapsed >= timeout:
            print(f"  ⏰ 超時，{Path(template_path).name} 仍然存在")
            return False
        time.sleep(0.5)


def find_all(
    template_path: str,
    confidence: float = 0.8,
    click_all: bool = False,
) -> list:
    """
    找出畫面上所有符合的圖示（例如多個相同按鈕/checkbox）。
    回傳所有 (x, y) 座標列表。
    """
    _check_file(template_path)
    template = _imread(template_path)
    screen = _screenshot()
    h, w = template.shape[:2]

    result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
    locations = np.where(result >= confidence)
    points = list(zip(*locations[::-1]))

    # 去除重疊（距離太近視為同一個）
    filtered = []
    for pt in points:
        too_close = any(
            abs(pt[0] - ex[0]) < w // 2 and abs(pt[1] - ex[1]) < h // 2
            for ex in filtered
        )
        if not too_close:
            filtered.append((pt[0] + w // 2, pt[1] + h // 2))

    print(f"  ✅ 找到 {len(filtered)} 個 {Path(template_path).name}")

    if click_all:
        for i, (x, y) in enumerate(filtered):
            print(f"     🖱️  點擊第 {i+1} 個：({x}, {y})")
            pyautogui.click(x, y)
            time.sleep(0.3)

    return filtered


def step(name: str):
    """印出步驟標題，方便看流程進度。"""
    print(f"\n{'─'*50}")
    print(f"▶  {name}")
    print(f"{'─'*50}")


def sleep(seconds: float, reason: str = ""):
    """帶說明的等待。"""
    msg = f"（{reason}）" if reason else ""
    print(f"  💤 等待 {seconds}s {msg}")
    time.sleep(seconds)


# ── 內部工具 ─────────────────────────────────────────────────

def _screenshot():
    screenshot = pyautogui.screenshot(region=_capture_region)
    return cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)


def _dpi_scale() -> float:
    """
    回傳截圖實體像素與邏輯像素的比例。
    Windows DPI 150% → 回傳 1.5；100% → 回傳 1.0。
    截圖座標需除以此值才能正確對應 pyautogui.click() 的邏輯座標。
    """
    logical_w, _ = pyautogui.size()
    img = pyautogui.screenshot()
    physical_w = img.size[0]
    return physical_w / logical_w if logical_w else 1.0


def _check_file(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"找不到圖片檔案：{path}\n"
            f"請確認已把截圖放在 templates/ 資料夾"
        )


def _color_center(screen_bgr, top_left: tuple, w: int, h: int,
                  white_thresh: int = 230) -> tuple[int, int]:
    """
    在匹配區域內找出非白色像素的重心。
    若全為白色則回退至幾何中心。
    white_thresh: 三個 channel 都 >= 此值才視為白色。
    """
    x0, y0 = top_left
    region = screen_bgr[y0:y0 + h, x0:x0 + w]          # (h, w, 3) BGR
    # 任一 channel 低於門檻 → 有顏色
    mask = np.any(region < white_thresh, axis=2)
    if not mask.any():
        return (x0 + w // 2, y0 + h // 2)               # 全白，退回幾何中心
    ys, xs = np.where(mask)
    return (int(xs.mean()) + x0, int(ys.mean()) + y0)


def _imread(path: str):
    """讀取圖片，支援中文路徑（cv2.imread 在 Windows 上不支援 Unicode）。"""
    return cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)


def _show_debug(screen, template, max_loc, max_val, confidence):
    h, w = template.shape[:2]
    debug_img = screen.copy()
    color = (0, 255, 0) if max_val >= confidence else (0, 0, 255)
    cv2.rectangle(debug_img, max_loc,
                  (max_loc[0] + w, max_loc[1] + h), color, 2)
    cv2.putText(debug_img, f"{max_val:.0%} / {confidence:.0%}",
                (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.9, color, 2)
    cv2.imshow("Debug", debug_img)
    cv2.waitKey(1)


# ── 命令列模式 ────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="影像辨識自動點擊工具")
    parser.add_argument("template", help="目標圖片路徑，例如 templates/btn.png")
    parser.add_argument("--confidence", type=float, default=0.8)
    parser.add_argument("--timeout", type=float, default=10.0)
    parser.add_argument("--no-click", action="store_true")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--click-type", default="left",
                        choices=["left", "right", "double"])
    args = parser.parse_args()

    if args.no_click:
        pos = find_only(args.template, args.confidence,
                        wait_timeout=args.timeout, debug=args.debug)
        print(f"找到位置：{pos}" if pos else "未找到")
        sys.exit(0 if pos else 1)
    else:
        pos = find_and_click(args.template, args.confidence,
                             click_type=args.click_type,
                             wait_timeout=args.timeout, debug=args.debug)
        sys.exit(0 if pos else 1)
