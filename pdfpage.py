# pdfpage.py (更新後的新碼)
from PyPDF2 import PdfReader, PdfWriter

class PDFPagePick:
    def __init__(self, src, output_filename):
        self.__src = PdfReader(src)  # <-- 新版 API
        self.__output = PdfWriter()  # <-- 新版 API
        self.__output_filename = output_filename

    def pick_individual_pages(self, page_num_list):
        for page_num in page_num_list:
            if self.__check_page_num(self.__src, page_num):
                # 取得第 page_num 頁
                page = self.__src.pages[page_num]
                # 加到輸出
                self.__output.add_page(page)
            else:
                print(f'pageNum {page_num} 不存在，略過')

    def insert_blank_page(self):
        # 插入空白頁，若要跟原本 PDF 同尺寸，可自行指定 width、height
        self.__output.add_blank_page()

    def save(self):
        with open(self.__output_filename, 'wb') as outputstream:
            self.__output.write(outputstream)

    def __check_page_num(self, target, page_num):
        # 新版以 len(target.pages) 檢查總頁數
        if page_num < 0 or page_num >= len(target.pages):
            print('頁碼超出範圍')
            return False
        return True


class PDFPageMerge:
    def __init__(self, src, dest, output_filename):
        self.__src = PdfReader(src)   # <-- 新版
        self.__dest = PdfReader(dest)
        self.__output_filename = output_filename
        self.__output = PdfWriter()   # <-- 新版

    def merge_src_page_to_dest_page(self, src_page_num, dest_page_num):
        if self.__check_page_num(self.__src, src_page_num) and self.__check_page_num(self.__dest, dest_page_num):
            dest_page = self.__dest.pages[dest_page_num]
            src_page = self.__src.pages[src_page_num]
            # 以新版方式合併頁面
            dest_page.merge_page(src_page)
            # 加到輸出
            self.__output.add_page(dest_page)

    def get_src_total_page(self):
        # len(self.__src.pages) 代表頁數
        return len(self.__src.pages)

    def save(self):
        with open(self.__output_filename, 'wb') as outputstream:
            self.__output.write(outputstream)

    def __check_page_num(self, target, page_num):
        if page_num < 0 or page_num >= len(target.pages):
            print('頁碼超出範圍')
            return False
        return True
