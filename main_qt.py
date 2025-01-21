# gui_pyside.py
import sys
import os
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QLineEdit, QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout,
    QFormLayout
)
from PySide6.QtGui import QIcon

# 假設核心功能與原先保持一致，我們匯入 core 相關函式
# 你可以按照專案結構自行改路徑
from core import (
    read_main_article,
    generate_text_and_letter,
    merge_text_and_letter,
    clean_temp_files
)
# 引入其他可能用到的函式庫來處理 doc/docx
try:
    import docx2txt
except ImportError:
    # 如果沒安裝，請先 pip install docx2txt
    pass

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("台灣存證信函對齊器")
        self.setGeometry(100, 100, 600, 300)
        self.setWindowIcon(QIcon())  # 你可以放個自訂圖示

        self.article_path = ""
        self.output_pdf_path = ""

        # 產生主視窗核心區域
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 透過 FormLayout + QHBoxLayout 讓介面更好看
        form_layout = QFormLayout()
        self.article_lineedit = QLineEdit()
        self.article_btn = QPushButton("選擇文字/Word 檔")
        self.article_btn.clicked.connect(self.browse_file)

        # 提供 PDF 輸出路徑
        self.output_lineedit = QLineEdit("output.pdf")
        self.output_btn = QPushButton("選擇輸出PDF位置")
        self.output_btn.clicked.connect(self.browse_output_pdf)

        self.sender_name = QLineEdit()
        self.sender_addr = QLineEdit()
        self.receiver_name = QLineEdit()
        self.receiver_addr = QLineEdit()

        # 「產生 PDF」按鈕
        self.generate_btn = QPushButton("產生 PDF")
        self.generate_btn.clicked.connect(self.generate_pdf)

        # 依序把控件加到介面
        form_layout.addRow("存證信函檔案：", self.article_lineedit)
        form_layout.addRow("", self.article_btn)
        form_layout.addRow("輸出 PDF 檔名：", self.output_lineedit)
        form_layout.addRow("", self.output_btn)
        form_layout.addRow("寄件人姓名：", self.sender_name)
        form_layout.addRow("寄件人地址：", self.sender_addr)
        form_layout.addRow("收件人姓名：", self.receiver_name)
        form_layout.addRow("收件人地址：", self.receiver_addr)
        form_layout.addRow("", self.generate_btn)

        central_widget.setLayout(form_layout)

    def browse_file(self):
        # 允許使用者選 .txt, .doc, .docx
        file_filter = "Text Files (*.txt);;Word Files (*.doc *.docx);;All Files (*.*)"
        selected_file, _ = QFileDialog.getOpenFileName(self, "選擇文字或Word檔", "", file_filter)
        if selected_file:
            self.article_lineedit.setText(selected_file)

    def browse_output_pdf(self):
        # 允許使用者指定 PDF 輸出檔路徑
        selected_file, _ = QFileDialog.getSaveFileName(self, "選擇輸出 PDF 路徑", "", "PDF Files (*.pdf);;All Files (*.*)")
        if selected_file:
            self.output_lineedit.setText(selected_file)

    def generate_pdf(self):
        article_file = self.article_lineedit.text().strip()
        if not article_file or not os.path.exists(article_file):
            QMessageBox.critical(self, "錯誤", "請先選擇正確的文字或 Word 檔")
            return

        out_file = self.output_lineedit.text().strip()
        if not out_file.lower().endswith(".pdf"):
            out_file += ".pdf"

        # 讀取檔案內容，若為 doc 或 docx，則改用 docx2txt 解析
        ext = Path(article_file).suffix.lower()
        if ext in [".doc", ".docx"]:
            # 透過 docx2txt 轉成純文字
            try:
                main_text = docx2txt.process(article_file)
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"無法解析 Word 檔：{e}")
                return
        else:
            # 直接讀 txt
            main_text = read_main_article(article_file)

        s_name = self.sender_name.text().strip()
        s_addr = self.sender_addr.text().strip()
        r_name = self.receiver_name.text().strip()
        r_addr = self.receiver_addr.text().strip()

        # 整理成 core.py 需要的結構
        senders = [[s_name]] if s_name else []
        senders_addr = [s_addr] if s_addr else []
        receivers = [[r_name]] if r_name else []
        receivers_addr = [r_addr] if r_addr else []

        try:
            text_path, letter_path = generate_text_and_letter(
                senders=senders,
                senders_addr=senders_addr,
                receivers=receivers,
                receivers_addr=receivers_addr,
                ccs=[],
                cc_addr=[],
                main_text=main_text
            )
            merge_text_and_letter(text_path, letter_path, out_file)
            clean_temp_files(text_path, letter_path)

            QMessageBox.information(self, "完成", f"PDF 已產生：{out_file}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"產生失敗：{e}")

def run_gui():
    app = QApplication(sys.argv)

    # 先取得應用程式預設字型
    default_font = app.font()
    # 假設它原本是 10 pt，我們想乘以 1.5 => 15 pt
    default_font.setPointSize(int(default_font.pointSize() * 1.5))
    # 設定回應用程式
    app.setFont(default_font)

    window = MainWindow()

    # 如果你想把整個視窗的大小也擴增 1.5 倍，可以在 MainWindow 初始化後再改
    # 例如：
    original_width = window.width()
    original_height = window.height()
    window.resize(int(original_width * 1.5), int(original_height * 1.5))

    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    run_gui()

