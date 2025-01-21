# main_qt.py
import sys
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton,
    QLineEdit, QFileDialog, QMessageBox, QFormLayout
)
from PySide6.QtGui import QIcon
from PySide6.QtCore import Qt

# 假設核心功能都在 core.py
from core import (
    read_main_article,
    generate_text_and_letter,
    merge_text_and_letter,
    clean_temp_files
)

try:
    import docx2txt
except ImportError:
    pass

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("台灣郵局存證信函 - PySide6 版本")
        self.setGeometry(100, 100, 600, 400)

        # 現有欄位
        self.article_lineedit = QLineEdit()
        self.output_lineedit = QLineEdit("output.pdf")
        self.sender_name = QLineEdit()
        self.sender_addr = QLineEdit()
        self.receiver_name = QLineEdit()
        self.receiver_addr = QLineEdit()

        # 新增副本收件人欄位
        self.cc_name = QLineEdit()
        self.cc_addr = QLineEdit()

        self.init_ui()

    def init_ui(self):
        layout = QFormLayout()

        btn_browse_article = QPushButton("選擇文字或 Word 檔")
        btn_browse_article.clicked.connect(self.browse_file)

        btn_browse_output = QPushButton("選擇輸出 PDF 檔")
        btn_browse_output.clicked.connect(self.browse_output_file)

        btn_generate = QPushButton("產生 PDF")
        btn_generate.clicked.connect(self.generate_pdf)

        layout.addRow("存證信函檔案：", self.article_lineedit)
        layout.addRow("", btn_browse_article)
        layout.addRow("輸出 PDF 檔案：", self.output_lineedit)
        layout.addRow("", btn_browse_output)
        layout.addRow("寄件人姓名：", self.sender_name)
        layout.addRow("寄件人地址：", self.sender_addr)
        layout.addRow("收件人姓名：", self.receiver_name)
        layout.addRow("收件人地址：", self.receiver_addr)

        # 加上副本收件人
        layout.addRow("副本收件人姓名：", self.cc_name)
        layout.addRow("副本收件人地址：", self.cc_addr)

        layout.addRow("", btn_generate)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    def browse_file(self):
        file_filter = "文字檔或 Word (*.txt *.doc *.docx);;All Files (*.*)"
        selected_file, _ = QFileDialog.getOpenFileName(self, "選擇檔案", "", file_filter)
        if selected_file:
            self.article_lineedit.setText(selected_file)

    def browse_output_file(self):
        file_filter = "PDF Files (*.pdf);;All Files (*.*)"
        selected_file, _ = QFileDialog.getSaveFileName(self, "選擇輸出 PDF 檔案", "", file_filter)
        if selected_file:
            if not selected_file.lower().endswith(".pdf"):
                selected_file += ".pdf"
            self.output_lineedit.setText(selected_file)

    def generate_pdf(self):
        article_file = self.article_lineedit.text().strip()
        if not os.path.exists(article_file):
            QMessageBox.critical(self, "錯誤", "請選擇有效的檔案路徑")
            return

        out_file = self.output_lineedit.text().strip()
        if not out_file.lower().endswith(".pdf"):
            out_file += ".pdf"

        # 讀取檔案內容
        ext = Path(article_file).suffix.lower()
        if ext in [".doc", ".docx"]:
            try:
                main_text = docx2txt.process(article_file)
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"無法讀取 Word 檔：{e}")
                return
        else:
            main_text = read_main_article(article_file)

        # 收集使用者輸入的寄件人、收件人和副本收件人
        s_name = self.sender_name.text().strip()
        s_addr = self.sender_addr.text().strip()
        r_name = self.receiver_name.text().strip()
        r_addr = self.receiver_addr.text().strip()

        # 這裡就是新增副本收件人讀取
        c_name = self.cc_name.text().strip()
        c_addr = self.cc_addr.text().strip()

        # 組合成 core.py 需要的結構
        senders = [[s_name]] if s_name else []
        senders_addr = [s_addr] if s_addr else []
        receivers = [[r_name]] if r_name else []
        receivers_addr = [r_addr] if r_addr else []
        ccs = [[c_name]] if c_name else []
        cc_addr = [c_addr] if c_addr else []

        try:
            # 呼叫generate_text_and_letter
            text_path, letter_path = generate_text_and_letter(
                senders=senders,
                senders_addr=senders_addr,
                receivers=receivers,
                receivers_addr=receivers_addr,
                ccs=ccs,
                cc_addr=cc_addr,
                main_text=main_text
            )

            merge_text_and_letter(text_path, letter_path, out_file)
            clean_temp_files(text_path, letter_path)

            QMessageBox.information(self, "完成", f"PDF 已成功產生：{out_file}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

def run_gui():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    run_gui()
