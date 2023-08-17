import os
import docx
import xlwt
from tqdm import *
import pandas as pd
import numpy as np
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def get_key_words(xlsx_path):
    df = pd.read_excel(io=xlsx_path, header=None, sheet_name=1, keep_default_na=False)
    values = df.values
    key_words = []
    for each_row in values:
        for item in each_row:
            if item:
                key_words.append(item)
    return key_words

# 检查某行中首列文本是否出现待排除词汇，比如【公司】
def check_text(ceil_text):
    dirty_words = ["公司", "补助", "纳税"]
    for dirty_word in dirty_words:
        if dirty_word in ceil_text:
            return False
    return True


# 检查对应行是否可能为列名
def is_col_names(table, row_id, nums):
    for j in range(0, len(table.columns)):  # 检测第一行是否有重复的
        cur_item = table.cell(row_id, j).text.replace('\n', '').replace(' ', '')
        if cur_item == "" or len(table.rows) <= 2 or cur_item[0] in nums:
            return False
    return True


# 处理拿到列名和开始的行标
def get_col_names(table):
    col_names = []
    pre_item = "xxxxxxx"
    is_overlapping = False
    nums = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    second_is_col_names = False
    if len(table.rows) >= 2:
        second_is_col_names = is_col_names(table, 1, nums)
    for j in range(0, len(table.columns)):  # 检测第一行是否有重复的
        if table.cell(0, j).text == '' and j == len(table.columns) - 1:
            continue
        cur_item = table.cell(0, j).text.replace('\n', '').replace(' ', '')
        if not check_text(cur_item):  # 如果列名存在脏数据
            return [], 0, True, True
        if cur_item == "" or cur_item[0] in nums:
            return [], 0, True, False  # 表格跨页，一般首行会有列缺失或者出现数字
        if cur_item == pre_item and second_is_col_names:
            is_overlapping = True
            col_names = []
            break
        else:
            col_names.append(cur_item)
            pre_item = cur_item
    if is_overlapping:
        for j in range(0, len(table.columns)):
            col_names.append(table.cell(1, j).text.replace('\n', '').replace(' ', ''))
    begin_index = 2 if is_overlapping else 1
    return col_names, begin_index, False, False


def save_excel(total_info, target_dir, filename):
    is_empty = "空-" if len(total_info) == 0 else ""
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet("AI_info", cell_overwrite_ok=True)
    row_idx = 0
    for [col_names, per_table_result] in total_info:
        for i, col_name in enumerate(col_names):
            sheet.write(row_idx, i, col_name.replace('\n', '').replace(' ', ''))
        row_idx += 1
        for i, each_row in enumerate(per_table_result):
            for j, item in enumerate(each_row):
                sheet.write(row_idx, j, item.replace('\n', '').replace(' ', ''))
            row_idx += 1
    filename = is_empty + "结果表格-" + filename[:-4] + "xlsx"
    save_path = target_dir + "\\" + filename
    book.save(save_path)
    print(filename, "保存成功")


def check_unit(pre_paragraphs):
    for pre_paragraph in pre_paragraphs:
        if "万元" in pre_paragraph.text:
            return "万元"
    return "元"

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def deal_per_docx(docFile, key_titles):
    document = Document(docFile)
    pre_paragraphs = []
    total_info = []
    pre_col_names = []  # 前一个表格的列名，防止出现分页将表格打断
    valid_cnts = 0  # 对出现关键标题后的5个表格进行搜索
    for block in iter_block_items(document):  # block可以是Table也可以是Paragraph
        if isinstance(block, Paragraph):
            paragraph = block
            if paragraph.text == "\n" or paragraph.text == '':
                continue
            for key_title in key_titles:
                if key_title in paragraph.text:
                    valid_cnts = 5
                    break
            if (len(pre_paragraphs) > 5):
                pre_paragraphs.pop(0)
            pre_paragraphs.append(paragraph)
        elif isinstance(block, Table) and valid_cnts:
            table = block
            valid_cnts = valid_cnts - 1
            cur_table_result = []  # 当前table符合条件的行
            cur_table_unit = check_unit(pre_paragraphs)
            col_names, begin_index, is_missing_col_names, has_dirty_col = get_col_names(table)
            if has_dirty_col:
                continue
            if is_missing_col_names:  # 如果当前是缺失的列名，就使用上一个表格的列名
                col_names = pre_col_names
            else:
                pre_col_names = col_names  # 否则更新前列名，为之后的缺失列名做准备
            for i in range(begin_index, len(table.rows)):  # 遍历表格每一行，检测当前行第1列是否出现关键词
                is_vaild = False
                has_num = False  # 对出现的每一行检查是否有数字，若没有数字，即使出现关键词也进行跳过
                for j in range(0, len(table.columns)):
                    if table.cell(i, j).text.replace(' ', '').replace('\n', '')[1] in nums:
                        has_num = True
                        break
                if not has_num:
                    continue
                cur_first_item = table.cell(i, 0).text.replace(' ', '').replace('\n', '')
                if check_text(cur_first_item):  # 如果当前行首列文本没有脏数据
                    for key_word in key_words:
                        if key_word in cur_first_item or key_word.lower() in cur_first_item:
                            is_vaild = True
                            break
                if is_vaild == False:  # 没出现关键词直接跳过
                    continue
                cur_row = []  # 出现关键词搜集该行信息
                for j in range(0, len(table.columns)):
                    cur_row.append(table.cell(i, j).text.replace(' ', '').replace('\n', ''))
                cur_row.append(cur_table_unit)
                if is_missing_col_names and len(total_info) and total_info[-1][0][0] == pre_col_names[0]:
                    if total_info[-1][1][-1][-1] == " ":
                        total_info[-1][1].pop()
                    total_info[-1][1].append(cur_row)
                else:
                    cur_table_result.append(cur_row)
            if len(cur_table_result) == 0:
                continue
            col_names.append("单位")
            cur_table_result.append([" " for n in range(len(col_names))])
            total_info.append([col_names, cur_table_result])
    return total_info


# 对scr_dir目录下的所有docx文件进行搜寻，并将每个公司提取的信息存入xlsx文件中
# target_dir不指定则输出至src_dir同目录下
def search_AI(key_words, key_titles, src_dir, target_dir=None, ):
    target_dir = src_dir if target_dir == None else target_dir
    for root, dirs, files in os.walk(src_dir):
        for file in tqdm(files):
            docFile = os.path.join(root, file)
            total_info = deal_per_docx(docFile, key_titles)
            save_excel(total_info, target_dir, file)

# v1版本的代码只搜索特定标题下的表格
if __name__ == "__main__":
    key_words_excel_path = "./10-21AI导向词语词频数量.xlsx"
    src_path = "./test_v1"
    target_path = "./test_result_v1"
    key_titles = ["募集资金承诺项目情况", "开发支出", "承诺投资项目", "募集资金承诺项目",
                  "募投支出", "在研项目", "重要承诺事项", "承诺及或有事项", "募投项目"]
    key_words = get_key_words(key_words_excel_path)
    search_AI(key_words, key_titles, src_path, target_path)
