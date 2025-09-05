import time
import logging
from pathlib import Path
import win32com.client as win32
import pywintypes
from win32com.client import constants

log = logging.getLogger(__name__)

class WordDocManager:
    def __init__(self):
        self.word = None
        self._init_word_app()

    def _init_word_app(self):
        """初始化 Word 应用程序（单例）"""
        if self.word is None:
            try:
                self.word = win32.gencache.EnsureDispatch("Word.Application")
                self.word.Visible = False
                self.word.DisplayAlerts = 0  # wdAlertsNone
                log.info("Word 应用程序初始化成功")
            except Exception as e:
                log.error(f"Word 初始化失败: {e}")
                raise

    def save_doc_file(self, record_text: str, doc_path: str) -> bool:
        """保存文本到 .doc 文件（自动重试机制）"""
        doc_path = Path(doc_path)
        doc_path.parent.mkdir(parents=True, exist_ok=True)
        doc = None

        try:
            # 创建新文档
            doc = self.word.Documents.Add()
            doc.Content.Text = record_text

            # 保存文件（带重试）
            for attempt in range(3):
                try:
                    doc.SaveAs(str(doc_path), FileFormat=constants.wdFormatDocument)
                    log.info(f"文件保存成功: {doc_path}")
                    return True
                except pywintypes.com_error as e:
                    log.warning(f"保存失败 {attempt+1}/3: {e}")
                    time.sleep(1)
            
            log.error(f"多次重试后仍保存失败: {doc_path}")
            return False

        except Exception as e:
            log.error(f"保存过程中发生异常: {e}")
            return False
        finally:
            # 只关闭文档，不退出 Word 应用
            if doc is not None:
                try:
                    doc.Close(False)
                except pywintypes.com_error as e:
                    log.warning(f"关闭文档失败: {e}")

    def close(self):
        """显式关闭 Word 应用程序（测试结束后调用）"""
        if self.word is not None:
            try:
                self.word.Quit()
                log.info("Word 应用程序已关闭")
            except pywintypes.com_error as e:
                log.error(f"关闭 Word 失败: {e}")
            finally:
                self.word = None