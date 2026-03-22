"""
watch_pdf.py — 監控資料夾，自動將下載的 PDF 重新命名為 MD5 hash

用法：
    python watch_pdf.py <監控資料夾>
    python watch_pdf.py "C:/Users/xxx/Downloads/TODOPDF"

依賴：
    pip install watchdog
"""

import sys
import hashlib
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


def rename_pdf(folder: Path, pdf_path: Path) -> bool:
    """將指定 PDF 重新命名為其內容的 MD5 hash。"""
    if not pdf_path.exists():
        return False
    try:
        md5 = hashlib.md5(pdf_path.read_bytes()).hexdigest()
        new_path = folder / f"{md5}.pdf"
        pdf_path.rename(new_path)
        log.info("✅ %s → %s", pdf_path.name, new_path.name)
        return True
    except Exception as e:
        log.error("❌ 重新命名失敗：%s", e)
        return False


def watch(folder: Path):
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError:
        log.error("缺少 watchdog，請執行：pip install watchdog")
        sys.exit(1)

    class _Handler(FileSystemEventHandler):
        def on_created(self, event):
            if event.is_directory:
                return
            path = Path(event.src_path)
            if path.name == DOWNLOAD_PDF_NAME:
                # 等檔案寫入完成再處理
                _wait_until_stable(path)
                rename_pdf(folder, path)

        def on_moved(self, event):
            # 部分下載器會先寫成暫存檔再移動
            if event.is_directory:
                return
            path = Path(event.dest_path)
            if path.name == DOWNLOAD_PDF_NAME:
                _wait_until_stable(path)
                rename_pdf(folder, path)

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


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"用法：python watch_pdf.py <監控資料夾>")
        sys.exit(1)

    folder = Path(sys.argv[1]).expanduser().resolve()
    if not folder.exists():
        log.error("找不到資料夾：%s", folder)
        sys.exit(1)

    watch(folder)
