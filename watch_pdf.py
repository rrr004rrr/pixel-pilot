"""
watch_pdf.py — 監控資料夾，自動將下載的 PDF 重新命名為標題內容

用法：
    python watch_pdf.py                        # 監控預設資料夾
    python watch_pdf.py "D:/TODOPDF"           # 監控指定資料夾
    python watch_pdf.py -all                   # 強制重新命名資料夾內所有 PDF，然後退出
    python watch_pdf.py "D:/TODOPDF" -all      # 指定資料夾 + 強制全部重命名

依賴：
    pip install watchdog pdfplumber
"""

import sys
import re
import uuid
import time
import logging
import warnings
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── 設定 ────────────────────────────────────────────────────
DEFAULT_FOLDER    = r"D:\TODOPDF"
DOWNLOAD_PDF_NAME = "TodoNow • 让工作快起来.pdf"
POLL_INTERVAL     = 2.0   # 輪詢間隔（秒），作為 watchdog 的備援
_INVALID          = re.compile(r'[\\/:*?"<>|]')
# ────────────────────────────────────────────────────────────


def rename_pdf(folder: Path, pdf_path: Path) -> bool:
    """擷取 PDF 第一行文字作為新檔名，先改成暫存名再改成目標名。"""
    if not pdf_path.exists():
        return False
    try:
        import pdfplumber
    except ImportError:
        log.error("缺少 pdfplumber，請執行：pip install pdfplumber")
        return False

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with pdfplumber.open(str(pdf_path)) as pdf:
                text = pdf.pages[0].extract_text() or ""
        first_line = text.strip().splitlines()[0].strip() if text.strip() else ""
        if not first_line:
            log.warning("⚠️  無法擷取文字：%s", pdf_path.name)
            return False

        new_name = _INVALID.sub("_", first_line) + ".pdf"
        new_path = folder / new_name

        if new_path == pdf_path:
            log.info("ℹ️  已是正確檔名，略過：%s", pdf_path.name)
            return True

        if new_path.exists():
            log.info("ℹ️  目標已存在，略過：%s → %s", pdf_path.name, new_name)
            return True

        pdf_path.rename(new_path)
        log.info("✅ %s → %s", pdf_path.name, new_name)
        return True

    except Exception as e:
        log.error("❌ 重新命名失敗 %s：%s", pdf_path.name, e)
        return False


def rename_all(folder: Path):
    """強制重新命名資料夾內所有 PDF。"""
    pdfs = sorted(folder.glob("*.pdf"))
    if not pdfs:
        log.info("資料夾內沒有 PDF：%s", folder)
        return
    log.info("共找到 %d 個 PDF，開始重新命名...", len(pdfs))
    ok = sum(rename_pdf(folder, p) for p in pdfs)
    log.info("完成：%d / %d 成功", ok, len(pdfs))


def _wait_until_stable(path: Path, interval: float = 0.5, max_wait: float = 30.0):
    """等到檔案大小不再變化（確認下載完成）。"""
    elapsed = 0.0
    prev_size = -1
    while elapsed < max_wait:
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            return
        if size == prev_size and size > 0:
            return
        prev_size = size
        time.sleep(interval)
        elapsed += interval


def _check_and_rename(folder: Path):
    """若目標檔案存在就處理，不存在則略過。"""
    pdf_path = folder / DOWNLOAD_PDF_NAME
    if pdf_path.exists():
        log.info("🔎 發現目標檔案，準備重新命名...")
        _wait_until_stable(pdf_path)
        rename_pdf(folder, pdf_path)


def watch(folder: Path):
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
        use_watchdog = True
    except ImportError:
        log.warning("未安裝 watchdog（pip install watchdog），改用輪詢模式")
        use_watchdog = False

    # 啟動時先掃一次，處理已存在的檔案
    _check_and_rename(folder)

    if use_watchdog:
        class _Handler(FileSystemEventHandler):
            def _handle(self, path: Path):
                if path.name == DOWNLOAD_PDF_NAME:
                    _wait_until_stable(path)
                    rename_pdf(folder, path)

            def on_created(self, event):
                if not event.is_directory:
                    self._handle(Path(event.src_path))

            def on_modified(self, event):
                if not event.is_directory:
                    self._handle(Path(event.src_path))

            def on_moved(self, event):
                if not event.is_directory:
                    self._handle(Path(event.dest_path))

        observer = Observer()
        observer.schedule(_Handler(), str(folder), recursive=False)
        observer.start()
        log.info("👀 監控中（watchdog + 輪詢）：%s", folder)
    else:
        observer = None
        log.info("👀 監控中（輪詢）：%s", folder)

    log.info("   目標檔名：%s", DOWNLOAD_PDF_NAME)
    log.info("   按 Ctrl+C 停止")

    try:
        while True:
            time.sleep(POLL_INTERVAL)
            # 輪詢備援：watchdog 萬一漏掉事件也能補救
            _check_and_rename(folder)
    except KeyboardInterrupt:
        pass
    finally:
        if observer:
            observer.stop()
            observer.join()
        log.info("⏹ 已停止")


if __name__ == "__main__":
    args = sys.argv[1:]
    force_all = "-all" in args
    args = [a for a in args if a != "-all"]

    folder_arg = args[0] if args else DEFAULT_FOLDER
    folder = Path(folder_arg).expanduser().resolve()

    if not folder.exists():
        log.error("找不到資料夾：%s", folder)
        sys.exit(1)

    if force_all:
        rename_all(folder)
    else:
        watch(folder)
