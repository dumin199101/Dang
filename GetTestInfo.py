# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/28 18:21
@Name: GetTestInfo
@Author: lieyan123091
@Desc：获取试题信息
"""

import re, sys, os
import shutil
import xlrd

regex_tab = re.compile("\\t")


def write_mapping_log(logname, content):
    """
    生成mapping日志文件
    :return:
    """
    with open(logname, "a+") as f1:
        f1.write(content.encode("utf-8"))


def del_blank_char(str):
    """
    去除字符串中的空白跟换行符
    :param str:
    :return:
    """
    rep = re.compile(r'(\n|\t|\r)')
    (fstr, count) = rep.subn('', str)
    return fstr


i = 1


def get_all_xls(srcfolder):
    global i
    for name in os.listdir(srcfolder):
        srcname = srcfolder + "\\" + name
        if os.path.isdir(srcname):
            get_all_xls(srcname)
        else:
            if name.endswith("xls") or name.endswith("xlsx"):
                print i, name[:name.rfind(".")]
                write_mapping_log("2012-2014试题列表-不清晰版.txt".encode("gbk"), name[:name.rfind(".")] + "\n")
                i = i + 1


j = 1


def get_all_pname_xls(srcfolder):
    global j
    for name in os.listdir(srcfolder):
        srcname = srcfolder + "\\" + name
        if os.path.isdir(srcname):
            get_all_pname_xls(srcname)
        else:
            if name.endswith("xls") or name.endswith("xlsx"):
                name = name[:name.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                print j, name
                write_mapping_log("2012-2014试题列表-不清晰版合并前.txt".encode("gbk"), name + "\n")
                j = j + 1


def get_test_info(srcfile):
    test_set = set()
    with open(srcfile, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            name = info1[0].strip("\n")
            test_set.add(name)
    j = 1
    for obj in test_set:
        print j, obj
        write_mapping_log("2012-2014试题列表-不清晰版合并后.txt".encode("gbk"), obj + "\n")
        j = j + 1


set_file1 = set()
set_file2 = set()


def get_diff_item(file1, file2):
    """
    获取不同类型之间的差值部分
    :param file1:
    :param file2:
    :return:
    """
    with open(file1, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            set_file1.add(info1[0].strip("\n"))
    with open(file2, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            set_file2.add(info2[0].strip("\n"))

    set_diff = set_file1.difference(set_file2)
    m = 1
    for obj in set_diff:
        print m, obj
        # write_mapping_log("2012-2014试题列表-不清晰版[待审核].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2012-2018试题列表[待审核].txt".encode("gbk"), obj + "\n")
        m = m + 1


set_file1 = set()
set_file2 = set()


def get_inter_item(file1, file2):
    """
    获取不同类型之间的交集部分
    :param file1:
    :param file2:
    :return:
    """
    with open(file1, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            set_file1.add(info1[0].strip("\n"))
    with open(file2, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            set_file2.add(info2[0].strip("\n"))

    set_diff = set_file1.intersection(set_file2)
    m = 1
    for obj in set_diff:
        write_mapping_log("2012-2014试题列表-不清晰版[同数据库对应].txt".encode("gbk"), obj + "\n")
        print m, obj
        m = m + 1


def shenhe_txt(srcfile):
    j = 1
    with open(srcfile, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            course_name = info1[0]
            shiti_name = info1[1].strip("\n")
            if shiti_name.endswith("无") or shiti_name.endswith("缺"):
                continue
            print j, course_name, shiti_name
            write_mapping_log("2012-2018试题列表[已审核]--去除缺无.txt".encode("gbk"), course_name + "\t" + shiti_name + "\n")
            j = j + 1


j = 1


def get_excel_file(srcfolder, destfolder):
    """
    获取excel源文件
    :param srcfolder:
    :param destfolder:
    :return:
    """
    if not os.path.exists(destfolder):
        os.makedirs(destfolder)
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        if os.path.isdir(file):
            get_excel_file(file, destfolder)
        else:
            if filename.endswith("xls") or filename.endswith("xlsx"):
                name = filename[:filename.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                with open(u"2012-2018试题列表[已审核]--去除缺无.txt", "r") as f1:
                    global j
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        course_name = info1[0]
                        shiti_name = info1[1].strip("\n")
                        if name == shiti_name:
                            srcfile = srcfolder + "\\" + filename
                            destfile = destfolder + "\\" + filename
                            print j, course_name, filename, srcfile, destfile
                            shutil.copy2(srcfile, destfile)
                            write_mapping_log("2012-2018试题列表[最终映射].txt".encode("gbk"),
                                              course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            j = j + 1
                            break


j = 1


def get_excel_file2(srcfolder, destfolder):
    """
    获取excel源文件
    :param srcfolder:
    :param destfolder:
    :return:
    """
    if not os.path.exists(destfolder):
        os.makedirs(destfolder)
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        if os.path.isdir(file):
            get_excel_file2(file, destfolder)
        else:
            if filename.endswith("xls") or filename.endswith("xlsx"):
                name = filename[:filename.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                with open(u"2012-2018试题列表[同数据库对应].txt", "r") as f1:
                    global j
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        course_name = info1[0].strip("\n")
                        shiti_name = info1[0].strip("\n")
                        if name == shiti_name:
                            srcfile = srcfolder + "\\" + filename
                            destfile = destfolder + "\\" + filename
                            print j, course_name, filename, srcfile, destfile
                            shutil.copy2(srcfile, destfile)
                            write_mapping_log("2012-2018试题列表[最终映射]-2.txt".encode("gbk"),
                                              course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            j = j + 1
                            break


def strQ2B(ustring):
    """全角转半角"""
    rstring = ""
    for uchar in ustring:
        inside_code = ord(uchar)
        if inside_code == 12288:  # 全角空格直接转换
            inside_code = 32
        elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
            inside_code -= 65248

        rstring += unichr(inside_code)
    return rstring


def strB2Q(ustring):
    """半角转全角"""
    rstring = ""
    for uchar in ustring:
        inside_code = ord(uchar)
        if inside_code == 32:  # 半角空格直接转化
            inside_code = 12288
        elif inside_code >= 32 and inside_code <= 126:  # 半角字符（除空格）根据关系转化
            inside_code += 65248

        rstring += unichr(inside_code)


k = 1


def get_xls_content(course_name, srcfile):
    global k
    # 打开文件
    workbook = xlrd.open_workbook(srcfile)
    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始

    # 遍历表格内容：

    # 如果第四行第一列为Float，那么应该从A列开始，否则从B列开始：
    if isinstance(sheet1.cell(3, 0).value, float) or (isinstance(sheet1.cell(3, 0).value.encode("utf-8"), str) and (
                len(sheet1.cell(3, 0).value.encode("utf-8").strip()) > 0)):
        if isinstance(sheet1.cell(3, 0).value, float):
            for i in range(3, sheet1.nrows):
                knowledge = del_blank_char(sheet1.cell(i, 1).value)
                types = sheet1.cell(i, 2).value
                difficult = sheet1.cell(i, 3).value
                if isinstance(difficult, float):
                    difficult = str(int(difficult))
                score = sheet1.cell(i, 4).value
                if isinstance(score, float):
                    score = str(int(score))
                if len(score) == 0:
                    score = '5'
                question = del_blank_char(sheet1.cell(i, 5).value.strip())
                option = del_blank_char(sheet1.cell(i, 6).value.replace("；", ";"))
                answer = sheet1.cell(i, 7).value
                if isinstance(answer, float):
                    answer = str(int(answer))
                else:
                    answer = strQ2B(answer.upper()).replace("；", "").replace(";", "")
                if len(answer) > 0 and len(question) > 0:
                    print k, course_name, knowledge, types, difficult, score, question, option, answer
                    write_mapping_log("试题文件_345.txt".encode("gbk"),
                                      course_name + "\t" + knowledge + "\t" + str(types) + "\t" + str(
                                          difficult) + "\t" + str(score) + "\t" + question + "\t" + option + "\t" + str(
                                          answer) + "\n")
                    k = k + 1
        else:
            # 处理不标准的数据
            if sheet1.cell(3, 0).value.encode("utf-8").endswith("单选题"):
                for i in range(3, sheet1.nrows):
                    knowledge = "暂无知识点"
                    types = sheet1.cell(i, 0).value
                    difficult = "暂无难度"
                    score = sheet1.cell(i, 1).value
                    if isinstance(score, float):
                        score = str(int(score))
                    if len(score) == 0:
                        score = '5'
                    question = del_blank_char(sheet1.cell(i, 2).value.strip())
                    option = del_blank_char(sheet1.cell(i, 3).value.replace("；", ";"))
                    answer = sheet1.cell(i, 4).value
                    if isinstance(answer, float):
                        answer = str(int(answer))
                    else:
                        answer = strQ2B(answer.upper()).replace("；", "").replace(";", "")
                    if len(answer) > 0 and len(question) > 0:
                        print k, course_name, knowledge, types, difficult, score, question, option, answer
                        write_mapping_log("试题文件_345.txt".encode("gbk"),
                                          course_name + "\t" + knowledge + "\t" + str(types) + "\t" + str(
                                              difficult) + "\t" + str(
                                              score) + "\t" + question + "\t" + option + "\t" + str(answer) + "\n")
                        k = k + 1
            else:
                for i in range(3, sheet1.nrows):
                    knowledge = del_blank_char(sheet1.cell(i, 1).value)
                    types = sheet1.cell(i, 2).value
                    difficult = sheet1.cell(i, 3).value
                    if isinstance(difficult, float):
                        difficult = str(int(difficult))
                    score = sheet1.cell(i, 4).value
                    if isinstance(score, float):
                        score = str(int(score))
                    if len(score) == 0:
                        score = '5'
                    question = del_blank_char(sheet1.cell(i, 5).value.strip())
                    option = del_blank_char(sheet1.cell(i, 6).value.replace("；", ";"))
                    answer = sheet1.cell(i, 7).value
                    if isinstance(answer, float):
                        answer = str(int(answer))
                    else:
                        answer = strQ2B(answer.upper()).replace("；", "").replace(";", "")
                    if len(answer) > 0 and len(question) > 0:
                        print k, course_name, knowledge, types, difficult, score, question, option, answer
                        write_mapping_log("试题文件_345.txt".encode("gbk"),
                                          course_name + "\t" + knowledge + "\t" + str(types) + "\t" + str(
                                              difficult) + "\t" + str(
                                              score) + "\t" + question + "\t" + option + "\t" + str(answer) + "\n")
                        k = k + 1
    else:
        for i in range(3, sheet1.nrows):
            knowledge = del_blank_char(sheet1.cell(i, 2).value)
            types = sheet1.cell(i, 3).value
            difficult = sheet1.cell(i, 4).value
            if isinstance(difficult, float):
                difficult = str(int(difficult))
            score = sheet1.cell(i, 5).value
            if isinstance(score, float):
                score = str(int(score))
            if len(score) == 0:
                score = '5'
            question = del_blank_char(sheet1.cell(i, 6).value.strip())
            option = del_blank_char(sheet1.cell(i, 7).value.replace("；", ";"))
            answer = sheet1.cell(i, 8).value
            if isinstance(answer, float):
                answer = str(int(answer))
            else:
                answer = strQ2B(answer.upper()).replace("；", "").replace(";", "")
            if len(answer) > 0 and len(question) > 0:
                print k, course_name, knowledge, types, difficult, score, question, option, answer
                write_mapping_log("试题文件_345.txt".encode("gbk"),
                                  course_name + "\t" + knowledge + "\t" + str(types) + "\t" + str(
                                      difficult) + "\t" + str(
                                      score) + "\t" + question + "\t" + option + "\t" + str(answer) + "\n")
                k = k + 1


def get_xls_question_content(course_name, srcfile):
    """
    处理结构2的Excel
    :param course_name:
    :param srcfile:
    :return:
    """
    global k

    # 打开文件
    workbook = xlrd.open_workbook(srcfile)
    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始

    for i in range(2, sheet1.nrows):
        types = sheet1.cell(i, 1).value
        knowledge = "暂无知识点"
        difficult = "3"
        score = sheet1.cell(i, 4).value
        if isinstance(score, float):
            score = str(int(score))
        if len(score) == 0:
            score = '5'
        question = del_blank_char(sheet1.cell(i, 2).value.strip())

        if types.endswith("选择题"):
            option_list = []
            optionA = del_blank_char(str(sheet1.cell(i, 11).value))
            if len(optionA.strip()) > 0:
                option_list.append(optionA)
            optionB = del_blank_char(str(sheet1.cell(i, 12).value))
            if len(optionB.strip()) > 0:
                option_list.append(optionB)
            optionC = del_blank_char(str(sheet1.cell(i, 13).value))
            if len(optionC.strip()) > 0:
                option_list.append(optionC)
            optionD = del_blank_char(str(sheet1.cell(i, 14).value))
            if len(optionD.strip()) > 0:
                option_list.append(optionD)
            if int(sheet1.ncols) > 15:
                optionE = del_blank_char(str(sheet1.cell(i, 15).value))
                if len(optionE.strip()) > 0:
                    option_list.append(optionE)
            option = ";".join(option_list)
        else:
            option = ""

        answer = sheet1.cell(i, 3).value
        if types.endswith("判断题"):
            if answer == "A":
                answer = "1"
            elif answer == "B":
                answer = "0"

        if len(answer) > 0 and len(question) > 0:
            print k, course_name, knowledge, types, difficult, score, question, option, answer
            write_mapping_log("试题文件_Question_345.txt".encode("gbk"),
                              course_name + "\t" + knowledge + "\t" + str(types) + "\t" + str(difficult) + "\t" + str(
                                  score) + "\t" + question + "\t" + option + "\t" + str(answer) + "\n")
            k = k + 1





def get_xls_last_content(course_name, srcfile):
    """
    :param course_name:
    :param srcfile:
    :return:
    """
    global k

    # 打开文件
    workbook = xlrd.open_workbook(srcfile)
    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始

    for i in range(3, sheet1.nrows):
        knowledge = "暂无知识点"
        types = sheet1.cell(i, 2).value
        difficult = sheet1.cell(i, 3).value
        if isinstance(difficult, float):
            difficult = str(int(difficult))
        score = sheet1.cell(i, 4).value
        if isinstance(score, float):
            score = str(int(score))
        if len(score) == 0:
            score = '5'
        question = del_blank_char(sheet1.cell(i, 5).value.strip())
        option = del_blank_char(sheet1.cell(i, 6).value.replace("；", ";"))
        answer = sheet1.cell(i, 7).value
        if isinstance(answer, float):
            answer = str(int(answer))
        else:
            answer = strQ2B(answer.upper()).replace("；", "").replace(";", "")
        if len(answer) > 0 and len(question) > 0:
            print k, course_name, knowledge, types, difficult, score, question, option, answer
            write_mapping_log("试题文件_Last_345.txt".encode("gbk"),
                              course_name + "\t" + knowledge + "\t" + str(types) + "\t" + str(difficult) + "\t" + str(
                                  score) + "\t" + question + "\t" + option + "\t" + str(answer) + "\n")
            k = k + 1




n = 1


def get_excel_str(srcfile):
    SOURCE_DIR = "E:\\Goosuu\\Dang\\Script\\Excel\\All"
    global n
    with open(srcfile, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            course_name = info1[0]
            filename = info1[1]
            src_excel_file = SOURCE_DIR + "\\" + filename.encode("gbk")
            print n, course_name
            get_xls_content(course_name, src_excel_file)
            n = n + 1


def get_excel_question_str(srcfile):
    SOURCE_DIR = "E:\\Goosuu\\Dang\\Script\\Excel\\Question"
    with open(srcfile, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            course_name = info1[0]
            filename = info1[1]
            src_excel_file = SOURCE_DIR + "\\" + filename.encode("gbk")
            print course_name
            get_xls_question_content(course_name, src_excel_file)

def get_excel_last_str(srcfile):
    SOURCE_DIR = "E:\\Goosuu\\Dang\\Script\\Excel\\Last"
    with open(srcfile, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            course_name = info1[0]
            filename = info1[1]
            src_excel_file = SOURCE_DIR + "\\" + filename.encode("gbk")
            print course_name
            get_xls_last_content(course_name, src_excel_file)

def main():
    # srcfolder = u"J:\\不清晰版\\文本资料"
    # get_all_xls(srcfolder)
    # get_all_pname_xls(srcfolder)

    # srcfile = u"2012-2014试题列表-不清晰版合并前.txt"
    # get_test_info(srcfile)

    # file1 = u"tb_course.txt"
    # file2 = u"2012-2014试题列表-不清晰版合并后.txt"
    # get_diff_item(file1, file2)
    # get_diff_item(file2, file1)
    # get_inter_item(file1,file2)

    # 对三组数据进行重组，查找未找到的
    # file1 = u"2016-2018试题列表[待审核].txt"
    # file2 = u"2012-2015.txt"
    # get_diff_item(file1, file2)

    # 对审核的2012-2018试题列表去除缺失的跟无试题的
    # srcfile = u"2012-2018试题列表[已审核].txt"
    # shenhe_txt(srcfile)

    # 获取审核完毕的Excel源文件：
    # destfolder = "E:\\Goosuu\\Dang\\Script\\Excel\\Source\\1"
    # # srcfolder = u"I:\\文字资料"
    # srcfolder = u"J:\\清晰版\\文本资料"
    # # srcfolder = u"J:\\不清晰版\\文本资料"
    # get_excel_file(srcfolder,destfolder)

    # 获取无问题的Excel源文件：
    # destfolder = "E:\\Goosuu\\Dang\\Script\\Excel\\Source\\4"
    # # srcfolder = u"I:\\文字资料"
    # # srcfolder = u"J:\\清晰版\\文本资料"
    # srcfolder = u"J:\\不清晰版\\文本资料"
    # get_excel_file2(srcfolder, destfolder)

    # 读取试题内容：
    # srcfile = u"试题-第一批345本（无问题）.txt"
    # get_excel_str(srcfile)

    # srcfile = u"试题-第一批345本（有问题）.txt"
    # get_excel_question_str(srcfile)

    # 对比数据库去重后的课程名跟人工整理的课程名之间的差值
    # file2 = u"2012-2018试题数据库对应课程名.txt"
    # file1 = u"2012-2018试题对应课程名.txt"
    # get_diff_item(file1, file2)

    # 获取第三种结构
    srcfile = u"试题-第一批345本（剩下两本）.txt"
    get_excel_last_str(srcfile)



if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
