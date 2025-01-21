# core.py

import random
import string
import datetime
import os

from pdfpage import PDFPagePick, PDFPageMerge
from pdfpainter import PDFPainter
from constants import (
    LETTER_FORMAT_PATH,
    DEFAULT_FONT_PATH,
    LETTER_FORMAT_WIDE_HEIGHT,
    CONTENT_X_Y_BEGIN,
    CONTENT_X_Y_INTERVAL,
    CONTENT_X_Y_FIX,
    CONTENT_MAX_CHARACTER_PER_LINE,
    CONTENT_MAX_LINE_PER_PAGE,
    NAME_COORDINATE,
    ADDR_COORDINATE,
    BOX_UPPDERLEFT_X_Y,
    BOX_UPPDERRIGHT_X_Y,
    CUT_INFO_X_Y,
    QUOTE_X_Y,
    RECT_X_Y_W_H,
    CHT_IN_RECT_X_Y,
    DETAIL_START,
    DETAIL_Y_INTERVAL,
    TITLE_START,
    TITLE_Y_INTERVAL,
    CC_RECEIVER_FIX_X_Y
)

def read_main_article(filepath):
    codec_name = 'utf-8'
    bom_str = b'\xef\xbb\xbf'.decode(codec_name)
    with open(filepath, 'r', encoding=codec_name) as text_file:
        text = text_file.read()
    # 移除可能存在的 BOM
    text = text.lstrip(bom_str)
    return text

def merge_text_and_letter(text_path, letter_path, output_filename):
    print('合併頁面中...')
    page_merge = PDFPageMerge(text_path, letter_path, output_filename)
    for i in range(page_merge.get_src_total_page()):
        page_merge.merge_src_page_to_dest_page(i, i)
    page_merge.save()

def gen_filename(rand_length):
    now = datetime.datetime.now()
    prefix = ''.join(random.choice(string.ascii_lowercase) for _ in range(rand_length))
    rand_str = ''.join(random.choice(string.ascii_lowercase) for _ in range(rand_length))
    return f'{prefix}-{now.strftime("%Y%m%d%H%M%S.%f")}-{rand_str}'

def clean_temp_files(*files):
    for f in files:
        if os.path.exists(f):
            os.remove(f)

def generate_text_and_letter(
    senders, senders_addr,
    receivers, receivers_addr,
    ccs, cc_addr,
    main_text
):
    text_path = gen_filename(20) + '.pdf'
    letter_path = gen_filename(21) + '.pdf'

    # 建立 PDFPainter 產生文字頁面
    generator = PDFPainter(text_path, LETTER_FORMAT_WIDE_HEIGHT[0], LETTER_FORMAT_WIDE_HEIGHT[1])
    blank_letter_producer = PDFPagePick(LETTER_FORMAT_PATH, letter_path)

    # 偵測是否僅需在第一頁填上寄件人、收件人、副本收件人
    only_one_set = _is_only_one_name_or_address(senders, senders_addr) and \
                   _is_only_one_name_or_address(receivers, receivers_addr) and \
                   _is_only_one_name_or_address(ccs, cc_addr)

    if only_one_set:
        generator.set_font(DEFAULT_FONT_PATH, 10)
        _fill_name_address_on_1st_page(generator, senders, senders_addr, 's')
        _fill_name_address_on_1st_page(generator, receivers, receivers_addr, 'r')
        _fill_name_address_on_1st_page(generator, ccs, cc_addr, 'c')

    # 開始處理正文
    generator.set_font(DEFAULT_FONT_PATH, 20)
    _parse_main_article(generator, blank_letter_producer, main_text)

    # 若不只一筆，就需要在第二頁放「剪下區塊」的詳細寄送資訊
    if not only_one_set:
        _draw_info_box(generator, senders, senders_addr,
                       receivers, receivers_addr,
                       ccs, cc_addr)
        generator.end_this_page()
        blank_letter_producer.insert_blank_page()

    blank_letter_producer.save()
    generator.save()

    return text_path, letter_path

def _is_only_one_name_or_address(namelist, addresslist):
    if namelist and len(namelist) != 1:
        return False
    if addresslist and len(addresslist) != 1:
        return False
    return True

def _parse_main_article(painter, page_pick, main_text):
    print('處理正文內容...')
    x_begin, y_begin, line_counter, char_counter = _reset_coordinates_and_counters()

    for i, char in enumerate(main_text):
        # 處理換行或超過行字數限制
        if char == '\n' or (char_counter > CONTENT_MAX_CHARACTER_PER_LINE):
            x_begin, y_begin = _get_new_line_coordinate(y_begin)
            line_counter += 1
            char_counter = 1
            if char == '\n':
                continue

        # 判斷是否需換頁
        if line_counter > CONTENT_MAX_LINE_PER_PAGE:
            painter.end_this_page()
            page_pick.pick_individual_pages([0])  # 複製底稿頁
            x_begin, y_begin, line_counter, char_counter = _reset_coordinates_and_counters()

        # 繪製當前字元
        painter.draw_string(x_begin, y_begin, char)
        x_begin += (CONTENT_X_Y_INTERVAL[0] - CONTENT_X_Y_FIX[0])
        char_counter += 1

    painter.end_this_page()
    # 最後一頁的底稿頁
    page_pick.pick_individual_pages([0])

def _reset_coordinates_and_counters():
    return CONTENT_X_Y_BEGIN[0], CONTENT_X_Y_BEGIN[1], 1, 1

def _get_new_line_coordinate(current_y):
    new_x = CONTENT_X_Y_BEGIN[0]
    new_y = current_y - (CONTENT_X_Y_INTERVAL[1] + CONTENT_X_Y_FIX[1])
    return new_x, new_y

def _fill_name_address_on_1st_page(painter, namelist, addresslist, type_):
    if len(namelist) == 1:
        all_name = ' '.join(namelist[0])
        painter.draw_string(NAME_COORDINATE[type_+'_x_y_begin'][0],
                            NAME_COORDINATE[type_+'_x_y_begin'][1],
                            all_name)
    if len(addresslist) == 1:
        painter.draw_string(ADDR_COORDINATE[type_+'_x_y_begin'][0],
                            ADDR_COORDINATE[type_+'_x_y_begin'][1],
                            addresslist[0])

def _draw_info_box(painter,
                   sender_list, sender_addr_list,
                   receiver_list, receiver_addr_list,
                   cc_list, cc_addr_list):
    painter.set_font(DEFAULT_FONT_PATH, 8)
    painter.draw_string(CUT_INFO_X_Y[0], CUT_INFO_X_Y[1], '[請自行剪下貼上]')
    painter.draw_line(BOX_UPPDERLEFT_X_Y[0], BOX_UPPDERLEFT_X_Y[1],
                      BOX_UPPDERRIGHT_X_Y[0], BOX_UPPDERRIGHT_X_Y[1])
    painter.draw_string(QUOTE_X_Y[0], QUOTE_X_Y[1],
                        '（寄件人如為機關或公司，請加蓋單位圖章及法定代理人簽名或蓋章）')
    painter.draw_rect(RECT_X_Y_W_H[0], RECT_X_Y_W_H[1], RECT_X_Y_W_H[2], RECT_X_Y_W_H[3])
    painter.set_font(DEFAULT_FONT_PATH, 10)
    painter.draw_string(CHT_IN_RECT_X_Y[0], CHT_IN_RECT_X_Y[1], '印')

    painter.draw_string(TITLE_START[0], TITLE_START[1], '一、寄件人')
    x_begin = DETAIL_START[0]
    y_begin = DETAIL_START[1]
    y_begin = _fill_name_address_in_info_box(painter, x_begin, y_begin, sender_list, sender_addr_list)

    y_begin -= TITLE_Y_INTERVAL
    painter.draw_string(TITLE_START[0], y_begin, '二、收件人')
    y_begin = _fill_name_address_in_info_box(painter, x_begin, y_begin, receiver_list, receiver_addr_list)

    y_begin -= TITLE_Y_INTERVAL
    painter.draw_string(TITLE_START[0], y_begin, '三、')
    painter.draw_string(TITLE_START[0] + CC_RECEIVER_FIX_X_Y[0],
                        y_begin + CC_RECEIVER_FIX_X_Y[1], '副 本')
    painter.draw_string(TITLE_START[0] + CC_RECEIVER_FIX_X_Y[0],
                        y_begin - CC_RECEIVER_FIX_X_Y[1], '收件人')
    y_begin = _fill_name_address_in_info_box(painter, x_begin, y_begin, cc_list, cc_addr_list)

    painter.draw_line(BOX_UPPDERLEFT_X_Y[0], BOX_UPPDERLEFT_X_Y[1],
                      BOX_UPPDERLEFT_X_Y[0], y_begin)  
    painter.draw_line(BOX_UPPDERLEFT_X_Y[0], y_begin,
                      BOX_UPPDERRIGHT_X_Y[0], y_begin)  
    painter.draw_line(BOX_UPPDERRIGHT_X_Y[0], BOX_UPPDERRIGHT_X_Y[1],
                      BOX_UPPDERRIGHT_X_Y[0], y_begin)

def _fill_name_address_in_info_box(painter, x_begin, y_begin, namelist, addresslist):
    max_count = max(len(namelist), len(addresslist))
    if max_count == 0:
        painter.draw_string(x_begin, y_begin, '姓名：')
        y_begin -= DETAIL_Y_INTERVAL
        painter.draw_string(x_begin, y_begin, '詳細地址：')
        y_begin -= DETAIL_Y_INTERVAL
        return y_begin

    for i in range(max_count):
        name_str = ' '.join(namelist[i]) if i < len(namelist) else ''
        painter.draw_string(x_begin, y_begin, '姓名：' + name_str)
        y_begin -= DETAIL_Y_INTERVAL
        address_str = addresslist[i] if i < len(addresslist) else ''
        painter.draw_string(x_begin, y_begin, '詳細地址：' + address_str)
        y_begin -= DETAIL_Y_INTERVAL
    return y_begin
