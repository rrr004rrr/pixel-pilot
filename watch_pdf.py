"""
watch_pdf.py — 監控資料夾，自動將下載的 PDF 重新命名為標題內容

用法：
    python watch_pdf.py <監控資料夾>
    python watch_pdf.py "C:/Users/xxx/Downloads/TODOPDF"

依賴：
    pip install watchdog pdfplumber
"""

import sys
import re
import uuid
import time
import logging
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

DOWNLOAD_PDF_NAME = "TodoNow • 让工作快起来.pdf"
_INVALID = re.compile(r'[\\/:*?"<>|]')


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
        with pdfplumber.open(str(pdf_path)) as pdf:
            text = pdf.pages[0].extract_text() or ""
        first_line = text.strip().splitlines()[0].strip() if text.strip() else ""
        if not first_line:
            log.warning("⚠️  無法擷取文字：%s", pdf_path.name)
            return False

        new_name = _INVALID.sub("_", first_line) + ".pdf"
        new_path = folder / new_name

        # 先改成暫存亂數名，避免目標名稱衝突
        tmp_path = folder / f"_tmp_{uuid.uuid4().hex}.pdf"
        pdf_path.rename(tmp_path)

        if new_path.exists():
            tmp_path.rename(pdf_path)  # 還原
            log.warning("⚠️  目標已存在，略過：%s", new_name)
            return False

        tmp_path.rename(new_path)
        log.info("✅ %s → %s", DOWNLOAD_PDF_NAME, new_name)
        return True

    except Exception as e:
        log.error("❌ 重新命名失敗：%s", e)
        return False


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


def watch(folder: Path):
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError:
        log.error("缺少 watchdog，請執行：pip install watchdog")
        sys.exit(1)

    class _Handler(FileSystemEventHandler):
        def _handle(self, path: Path):
            if path.name == DOWNLOAD_PDF_NAME:
                _wait_until_stable(path)
                rename_pdf(folder, path)

        def on_created(self, event):
            if not event.is_directory:
                self._handle(Path(event.src_path))

        def on_moved(self, event):
            if not event.is_directory:
                self._handle(Path(event.dest_path))

    observer = Observer()
    observer.schedule(_Handler(), str(folder), recursive=False)
    observer.start()
    log.info("👀 監控中：%s", folder)
    log.info("   目標檔名：%s", DOWNLOAD_PDF_NAME)
    log.info("   按 Ctrl+C 停止")

    try:
        while observer.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        observer.stop()
        observer.join()
        log.info("⏹ 已停止")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python watch_pdf.py <監控資料夾>")
        sys.exit(1)

    folder = Path(sys.argv[1]).expanduser().resolve()
    if not folder.exists():
        log.error("找不到資料夾：%s", folder)
        sys.exit(1)

    watch(folder)
