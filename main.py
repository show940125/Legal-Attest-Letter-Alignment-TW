# main.py

import argparse
import sys
import os

# 假設目前所有 py 檔都在同個資料夾
# 若有模組路徑需求，請自行調整
from core import (
    read_main_article,
    generate_text_and_letter,
    merge_text_and_letter,
    clean_temp_files
)

def main():
    parser = argparse.ArgumentParser(description='台灣郵局存證信函對齊器（Windows版）')
    parser.add_argument('article_file', nargs='?', default=None, help='存證信函全文之純文字檔路徑')
    parser.add_argument('--senderName', action='append', nargs='+', default=[], help='寄件人姓名(可多組)')
    parser.add_argument('--senderAddr', action='append', default=[], help='寄件人地址(可多組)')
    parser.add_argument('--receiverName', action='append', nargs='+', default=[], help='收件人姓名(可多組)')
    parser.add_argument('--receiverAddr', action='append', default=[], help='收件人地址(可多組)')
    parser.add_argument('--ccName', action='append', nargs='+', default=[], help='副本收件人姓名(可多組)')
    parser.add_argument('--ccAddr', action='append', default=[], help='副本收件人地址(可多組)')
    parser.add_argument('--outputFileName', default='output.pdf', help='輸出檔名')
    parser.add_argument('--gui', action='store_true', help='啟動圖形介面模式')

    args = parser.parse_args()

    # 如果指定 --gui，就直接呼叫 GUI 模式
    if args.gui:
        try:
            from gui_tk import run_gui
            run_gui()
        except ImportError:
            print("gui_tk.py 未找到或 import 失敗")
        return

    # 否則走命令列流程
    if not args.article_file:
        parser.print_help()
        sys.exit(1)

    if not os.path.exists(args.article_file):
        print(f"無法找到文章檔：{args.article_file}")
        sys.exit(1)

    main_text = read_main_article(args.article_file)
    text_path, letter_path = generate_text_and_letter(
        senders=args.senderName,
        senders_addr=args.senderAddr,
        receivers=args.receiverName,
        receivers_addr=args.receiverAddr,
        ccs=args.ccName,
        cc_addr=args.ccAddr,
        main_text=main_text
    )
    merge_text_and_letter(text_path, letter_path, args.outputFileName)
    clean_temp_files(text_path, letter_path)
    print(f"已完成，產生檔名： {args.outputFileName}")

if __name__ == '__main__':
    main()
