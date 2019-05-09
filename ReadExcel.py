# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/29 10:51
@Name: ReadExcel
@Author: lieyan123091
@Desc: 读取Excel表格内容
"""

import re, sys, os
import xlrd
from datetime import date, datetime

def strQ2B(ustring):
    """全角转半角"""
    rstring = ""
    for uchar in ustring:
        inside_code=ord(uchar)
        if inside_code == 12288:                              #全角空格直接转换
            inside_code = 32
        elif (inside_code >= 65281 and inside_code <= 65374): #全角字符（除空格）根据关系转化
            inside_code -= 65248

        rstring += unichr(inside_code)
    return rstring


def strB2Q(ustring):
    """半角转全角"""
    rstring = ""
    for uchar in ustring:
        inside_code=ord(uchar)
        if inside_code == 32:                                 #半角空格直接转化
            inside_code = 12288
        elif inside_code >= 32 and inside_code <= 126:        #半角字符（除空格）根据关系转化
            inside_code += 65248

        rstring += unichr(inside_code)



def get_xls_content(srcfile):
    # 打开文件
    workbook = xlrd.open_workbook(srcfile)
    # 获取所有sheet
    # print workbook.sheet_names()  # [u'sheet1', u'sheet2']
    # sheet2_name = workbook.sheet_names()[1]
    # print sheet2_name

    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始
    # sheet2 = workbook.sheet_by_name('sheet2')
    # print sheet1

    # sheet的名称，行数，列数
    # print sheet1.name, sheet1.nrows, sheet1.ncols

    # 获取整行和整列的值（数组）
    # rows = sheet1.row_values(4)  # 获取第五行内容
    # cols = sheet1.col_values(2)  # 获取第三列内容
    # print rows
    # print cols

    # 获取单元格内容
    # print sheet1.cell(3, 2).value.encode('utf-8')
    # print sheet2.cell_value(1, 0).encode('utf-8')
    # print sheet2.row(1)[0].value.encode('utf-8')

    # 获取单元格内容的数据类型
    # ctype :  0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
    # print sheet1.cell(1, 0).ctype

    # 日期处理
    # if (sheet1.cell(1, 2).ctype == 3):
    #     date_value = xlrd.xldate_as_tuple(sheet1.cell_value(rows, 3), workbook.datemode)
    #     date_tmp = date(*date_value[:3]).strftime('%Y/%m/%d')

    # 遍历表格内容：

    # 如果第四行第一列为Float，那么应该从A列开始，否则从B列开始：
    # print isinstance(sheet1.cell(3, 0).value.encode("utf-8"),str) and (len(sheet1.cell(3, 0).value.encode("utf-8").strip()) > 0)
    # print len(sheet1.cell(3, 0).value.encode("utf-8").strip())
    if isinstance(sheet1.cell(3, 0).value, float) or (isinstance(sheet1.cell(3, 0).value.encode("utf-8"),str) and (len(sheet1.cell(3, 0).value.encode("utf-8").strip()) > 0)):
        for i in range(3, sheet1.nrows):
            knowledge = sheet1.cell(i, 1).value
            types = sheet1.cell(i, 2).value
            difficult = sheet1.cell(i, 3).value
            score = sheet1.cell(i, 4).value
            question = sheet1.cell(i, 5).value.strip()
            option = sheet1.cell(i, 6).value.replace("；",";").strip("\n")
            answer = sheet1.cell(i, 7).value
            if isinstance(answer, float):
                answer = str(int(answer))
            else:
                answer = strQ2B(answer.upper()).replace("；","").replace(";","")
            if len(answer) > 0:
                print i, knowledge, types, difficult, score, question, option, answer
    else:
        for i in range(3, sheet1.nrows):
            knowledge = sheet1.cell(i, 2).value
            types = sheet1.cell(i, 3).value
            difficult = sheet1.cell(i, 4).value
            score = sheet1.cell(i, 5).value
            question = sheet1.cell(i, 6).value.strip()
            option = sheet1.cell(i, 7).value.replace("；", ";").strip("\n")
            answer = sheet1.cell(i, 8).value
            if isinstance(answer, float):
                answer = str(int(answer))
            else:
                answer = strQ2B(answer.upper()).replace("；", "").replace(";", "")
            if len(answer) > 0:
                print i, knowledge, types, difficult, score, question, option, answer



def main():
    srcfile = u"全面学习习近平新时代中国特色社会主义经济思想（上）.xls"
    srcfile = u"《中国共产党党内监督条例》解读（上）.xls"
    srcfile = u"增进人民福祉是经济发展终极目的（上）.xls"
    # srcfile = u"总体国家安全观下的中国政治安全（上）.xlsx"
    # srcfile = u"法家思想在现代社会中的价值（上）.xlsx"
    get_xls_content(srcfile)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
