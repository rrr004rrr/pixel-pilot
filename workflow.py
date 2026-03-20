"""
workflow.py — 步驟驅動的自動化流程

每個步驟是一個獨立函式，統一在 STEPS 清單裡管理。
新增步驟只要：
  1. 在 templates/ 放好截圖
  2. 新增一個 def step_XX() 函式
  3. 把它加進 STEPS 清單
"""

from auto_clicker import (
    find_and_click,
    image_exists,
    wait_for_image,
    wait_for_image_gone,
    find_all,
    sleep,
)


# ════════════════════════════════════════════════════════════
#  每個步驟定義
#  - 函式回傳 True = 成功，False = 失敗（流程停止）
#  - 每個函式只做一件事
# ════════════════════════════════════════════════════════════

def step_01_開啟主畫面():
    """等待軟體主畫面出現"""
    return wait_for_image("templates/main_window.png", timeout=15)


def step_02_點擊開始():
    return bool(find_and_click("templates/btn_start.png"))


def step_03_等待載入():
    """等載入動畫消失"""
    return wait_for_image_gone("templates/icon_loading.png", timeout=30)


def step_04_選擇項目():
    return bool(find_and_click("templates/item_target.png", confidence=0.80))


def step_05_確認彈窗():
    """彈窗不一定出現，出現才點"""
    if image_exists("templates/dialog_confirm.png"):
        return bool(find_and_click("templates/btn_ok.png"))
    return True  # 沒出現也算成功


def step_06_填寫欄位():
    pos = find_and_click("templates/input_field.png")
    if not pos:
        return False
    import pyautogui
    pyautogui.typewrite("要輸入的內容", interval=0.05)
    return True


def step_07_點擊下一步():
    return bool(find_and_click("templates/btn_next.png"))


def step_08_勾選所有選項():
    """找出所有 checkbox 並全部勾選"""
    items = find_all("templates/checkbox_unchecked.png", click_all=True)
    return len(items) > 0


def step_09_等待處理():
    sleep(2, "等待後台處理")
    return wait_for_image("templates/icon_done.png", timeout=60)


def step_10_點擊完成():
    return bool(find_and_click("templates/btn_finish.png"))


# ════════════════════════════════════════════════════════════
#  步驟清單（順序就是執行順序）
#  格式：(步驟函式, 失敗時動作)
#    失敗時動作: "stop"  = 整個流程停止（預設）
#               "skip"  = 跳過此步驟，繼續往下
#               "retry" = 重試一次再繼續
# ════════════════════════════════════════════════════════════

STEPS = [
    (step_01_開啟主畫面,  "stop"),
    (step_02_點擊開始,    "stop"),
    (step_03_等待載入,    "stop"),
    (step_04_選擇項目,    "stop"),
    (step_05_確認彈窗,    "skip"),   # 彈窗不一定有，失敗也沒關係
    (step_06_填寫欄位,    "stop"),
    (step_07_點擊下一步,  "stop"),
    (step_08_勾選所有選項,"retry"),  # 失敗先重試一次
    (step_09_等待處理,    "stop"),
    (step_10_點擊完成,    "stop"),
]


# ════════════════════════════════════════════════════════════
#  執行引擎（不需要修改）
# ════════════════════════════════════════════════════════════

def run(start_from: int = 1):
    """
    執行所有步驟。
    start_from: 從第幾步開始（用於中途失敗後跳過已完成的步驟）
    """
    total = len(STEPS)

    for i, (func, on_fail) in enumerate(STEPS, start=1):
        if i < start_from:
            print(f"  ⏭️  步驟 {i:02d}/{total} 跳過（已完成）")
            continue

        name = func.__name__.replace("step_", "").replace("_", " ", 1)
        print(f"\n{'━'*55}")
        print(f"  步驟 {i:02d}/{total}  {name}")
        print(f"{'━'*55}")

        success = _run_step(func, on_fail)

        if not success:
            print(f"\n❌ 流程在步驟 {i:02d} 停止：{name}")
            print(f"   修復後可從步驟 {i} 重新開始：")
            print(f"   run(start_from={i})")
            return False

    print(f"\n{'═'*55}")
    print(f"  ✅ 全部 {total} 個步驟執行完畢！")
    print(f"{'═'*55}")
    return True


def _run_step(func, on_fail: str) -> bool:
    """執行單一步驟，處理 retry 邏輯。"""
    try:
        result = func()
        if result:
            print(f"  ✅ 成功")
            return True

        # 失敗處理
        if on_fail == "skip":
            print(f"  ⚠️  失敗，跳過（on_fail=skip）")
            return True
        elif on_fail == "retry":
            print(f"  🔄 失敗，重試一次...")
            import time; time.sleep(1)
            result2 = func()
            if result2:
                print(f"  ✅ 重試成功")
                return True
            print(f"  ❌ 重試仍失敗")
            return False
        else:  # stop
            return False

    except FileNotFoundError as e:
        print(f"  ❌ 找不到截圖檔案：{e}")
        print(f"     請確認 templates/ 資料夾裡有對應的 PNG 截圖")
        return False
    except Exception as e:
        print(f"  ❌ 發生錯誤：{e}")
        return False


# ════════════════════════════════════════════════════════════
#  入口
# ════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys

    # 支援命令列參數：python workflow.py 5  → 從第 5 步開始
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 1

    print("🚀 影像辨識自動化流程")
    print("⚠️  緊急停止：把滑鼠移到螢幕左上角")
    print(f"   從步驟 {start} 開始，共 {len(STEPS)} 個步驟\n")

    run(start_from=start)
