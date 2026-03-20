# 影像辨識滑鼠自動化工具

> 按鈕位置會動？沒關係，截圖比對找到它再點。

---

## 安裝

```bash
pip install -r requirements.txt
```

---

## 使用流程

### 第一步：截圖放到 templates/

每個按鈕截一張 PNG，只截按鈕本身，不含多餘背景。

```
templates/
├── main_window.png     ← 主視窗（用來確認軟體開啟）
├── btn_start.png       ← 開始按鈕
├── icon_loading.png    ← 載入中圖示（等它消失）
├── btn_next.png        ← 下一步按鈕
└── btn_finish.png      ← 完成按鈕
```

**Windows 截圖：** `Win + Shift + S` 框選 → 貼到小畫家存 PNG

### 第二步：在 workflow.py 定義步驟

```python
def step_01_開啟主畫面():
    return wait_for_image("templates/main_window.png", timeout=15)

def step_02_點擊開始():
    return bool(find_and_click("templates/btn_start.png"))

def step_03_等待載入():
    return wait_for_image_gone("templates/icon_loading.png", timeout=30)
```

### 第三步：把步驟加進 STEPS 清單

```python
STEPS = [
    (step_01_開啟主畫面, "stop"),   # 失敗就停止
    (step_02_點擊開始,   "stop"),
    (step_03_等待載入,   "stop"),
    (step_04_確認彈窗,   "skip"),   # 彈窗不一定有，失敗也繼續
    (step_05_填寫欄位,   "retry"),  # 失敗重試一次
]
```

### 第四步：執行

```bash
python workflow.py        # 從頭開始
python workflow.py 5      # 從第 5 步開始（已完成的步驟跳過）
```

執行輸出範例：
```
🚀 影像辨識自動化流程
⚠️  緊急停止：把滑鼠移到螢幕左上角

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  步驟 01/10  開啟主畫面
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ✅ 找到 main_window.png  相似度 94%  位置 (640, 400)
  ✅ 成功

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  步驟 02/10  點擊開始
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  ✅ 找到 btn_start.png  相似度 91%  位置 (320, 280)
  🖱️  left 點擊 (320, 280)
  ✅ 成功
```

---

## 函式速查

```python
# 找到並點擊（等待直到超時）
find_and_click("templates/btn.png")
find_and_click("templates/btn.png", confidence=0.75)   # 降低門檻
find_and_click("templates/btn.png", click_type="right")# 右鍵
find_and_click("templates/btn.png", click_type="double")# 雙擊
find_and_click("templates/btn.png", wait_timeout=20)   # 等 20 秒

# 只找位置，不點擊
pos = find_only("templates/btn.png")   # 回傳 (x, y) 或 None

# 判斷圖示是否在畫面上（立即，不等待）
if image_exists("templates/error.png"):
    find_and_click("templates/btn_close.png")

# 等待圖示出現
wait_for_image("templates/ready.png", timeout=30)

# 等待圖示消失
wait_for_image_gone("templates/loading.png", timeout=60)

# 找出所有相同圖示並全部點擊
find_all("templates/checkbox.png", click_all=True)

# 等待 + 說明
sleep(2, "等待後台處理")
```

---

## 失敗行為對照

| STEPS 裡填 | 步驟失敗時 |
|-----------|----------|
| `"stop"`  | 整個流程停止，印出從哪步重跑 |
| `"skip"`  | 跳過此步驟，繼續往下 |
| `"retry"` | 等 1 秒後重試一次，仍失敗才停止 |

---

## 測試截圖能不能找到

```bash
# 確認截圖能被找到（不點擊）
python auto_clicker.py templates/btn_start.png --no-click

# 開啟除錯視窗，看比對結果（綠框=找到，紅框=找不到）
python auto_clicker.py templates/btn_start.png --no-click --debug

# 調低門檻測試
python auto_clicker.py templates/btn_start.png --no-click --confidence 0.75
```

---

## 常見問題

**找不到圖示（相似度很低）**
→ 重新截圖，確認截圖和執行時的螢幕縮放比例一致（Windows 設定 → 顯示 → 縮放）

**相似度接近門檻但偶爾失敗**
→ 降低 confidence：`find_and_click("...", confidence=0.75)`

**點到旁邊的位置**
→ 用 `offset_x` / `offset_y` 微調：`find_and_click("...", offset_x=5)`

**中途失敗，不想從頭來**
→ `python workflow.py 5`（從第 5 步繼續）

**緊急停止**
→ 把滑鼠移到**螢幕左上角**，程式立即停止
