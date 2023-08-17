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
    dirty_words = ["公司"]
    for dirty_word in dirty_words:
        if dirty_word in ceil_text:
            return False
    return True


# 处理拿到列名和开始的行标
def get_col_names(table):
    col_names = []
    pre_item = "xxxxxxx"
    is_overlapping = False
    for j in range(0, len(table.columns)):  # 检测第一行是否有重复的
        cur_item = table.cell(0, j).text
        if cur_item == "" or len(table.rows) <= 2:
            return [], 0, True  # 表格跨页，一般首行会有列缺失或者行数小于等于2
        if cur_item != pre_item:
            col_names.append(cur_item)
            pre_item = cur_item
        else:
            is_overlapping = True
            col_names = []
            break
    if is_overlapping:
        for j in range(0, len(table.columns)):
            col_names.append(table.cell(1, j).text)
    begin_index = 2 if is_overlapping else 1
    return col_names, begin_index, False


def save_excel(total_info, target_dir, filename):
    is_empty = "空-" if len(total_info) == 0 else ""
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet("AI_info", cell_overwrite_ok=True)
    for [col_names, per_table_result] in total_info:
        for i, col_name in enumerate(col_names):
            sheet.write(0, i, col_name.replace('\n', '').replace('\r', ''))
        for i, each_row in enumerate(per_table_result):
            for j, item in enumerate(each_row):
                sheet.write(i + 1, j, item.replace('\n', '').replace('\r', ''))
    filename = is_empty + "结果表格" + filename[:-4] + "xlsx"
    save_path = target_dir + "\\" + filename
    book.save(save_path)
    print(filename, "保存成功")


def check_unit(pre_paragraphs):
    for pre_paragraph in pre_paragraphs:
        for run in pre_paragraph.runs:
            if "万元" in run.text:
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


def deal_per_docx(docFile):
    document = Document(docFile)  # 读入docx文件,实例化Document对象
    pre_paragraphs = []
    total_info = []
    pre_col_names = []  # 前一个表格的列名，防止出现分页将表格打断
    for block in iter_block_items(document):  # block可以是Table也可以是Paragraph
        if isinstance(block, Paragraph):
            paragraph = block
            if (len(pre_paragraphs) > 4):
                pre_paragraphs.pop(0)
            pre_paragraphs.append(paragraph)
        elif isinstance(block, Table):
            table = block
            cur_table_result = []  # 当前table符合条件的行信息
            cur_table_unit = None
            col_names, begin_index, is_missing_col_names = get_col_names(table)
            if is_missing_col_names:  # 如果当前是缺失的列名，就使用上一个表格的列名
                col_names = pre_col_names
            else:
                pre_col_names = col_names  # 否则更新前列名，为之后的缺失列名做准备
            for i in range(begin_index, len(table.rows)):  # 遍历表格每一行，检测当前行是否出现关键词
                is_vaild = False
                if check_text(table.cell(i, 0).text):  # 如果当前行首列文本没有脏数据
                    for key_word in key_words:
                        if key_word in table.cell(i, 0).text or key_word.lower() in table.cell(i, 0).text:
                            is_vaild = True
                            cur_table_unit = check_unit(pre_paragraphs)
                            break
                if is_vaild:  # 当前行出现了关键词，则进行保存
                    cur_row = []
                    for j in range(0, len(table.columns)):
                        cur_row.append(table.cell(i, j).text)
                    cur_row.append(cur_table_unit)
                    cur_table_result.append(cur_row)
            if len(cur_table_result) != 0:
                col_names.append("单位")
                cur_table_result.append([" " for n in range(len(col_names))])
                total_info.append([col_names, cur_table_result])
    return total_info


# 对scr_dir目录下的所有docx文件进行搜寻，并将每个公司提取的信息存入xlsx文件中
# target_dir不指定则输出至src_dir同目录下
def search_AI(key_words, src_dir, target_dir=None, ):
    target_dir = src_dir if target_dir == None else target_dir
    for root, dirs, files in os.walk(src_dir):
        for file in tqdm(files):
            docFile = os.path.join(root, file)
            total_info = deal_per_docx(docFile)
            total_info = []
            save_excel(total_info, target_dir, file)


# v2版本的代码使用暴力搜索，对每个表格极其前面的段落进行查找
if __name__ == "__main__":
    excel_path = "./10-21AI导向词语词频数量.xlsx"
    src_path = "./test"
    target_path = "./test_result"
    key_words = get_key_words(excel_path)
    search_AI(key_words, src_path, target_path)
