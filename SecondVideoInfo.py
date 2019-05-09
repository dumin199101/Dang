# coding=utf-8
"""
@Project: Dang
@Time: 2019/4/19 11:03
@Name: SecondVideoInfo
@Author: lieyan123091
@Desc: 加工整理第二批视频数据
"""

import os
import sys
import shutil
import re
import xlrd
from pydocx import PyDocX
from win32com import client
import ffmpeg
import subprocess
import random

regex_tab = re.compile("\\t")


def compare(str1, str2):
    """
    上变为1，下变为3，中变为2
    :param str1:
    :param str2:
    :return:
    """
    if str1.rfind("（上）") > -1 and str2.rfind("（下）") > -1:
        return -1
    elif str1.rfind("（上）") > -1 and str2.rfind("（中）") > -1:
        return -1
    elif str1.rfind("（中）") > -1 and str2.rfind("（上）") > -1:
        return 1
    elif str1.rfind("（中）") > -1 and str2.rfind("（下）") > -1:
        return -1
    elif str1.rfind("（下）") > -1 and str2.rfind("（上）") > -1:
        return 1
    elif str1.rfind("（下）") > -1 and str2.rfind("（中）") > -1:
        return 1


def write_mapping_log(logname, content):
    """
    生成mapping日志文件
    :return:
    """
    with open(logname, "a+") as f1:
        f1.write(content.encode("utf-8"))


def get_course_info_according_video_resource(srcdir):
    """
    通过转码后的视频源数据提取课程名称：将（上）（中）（下）（一）（二）这种合并成一个，整理完成后以便跟Excel表格进行校对
    :param srcdir: 视频源文件所在目录
    :return:
    """
    year = srcdir[srcdir.rfind("\\") + 1:]
    i = 1
    video_name_set = set()
    for name in os.listdir(srcdir):
        name = name[:name.rfind(".mp4")]
        if name.endswith("（上）") or name.endswith("（下）") or name.endswith("（中）"):
            if name.endswith("（上）"):
                name = name[:name.rfind("（上）")]
            elif name.endswith("（下）"):
                name = name[:name.rfind("（下）")]
            elif name.endswith("（中）"):
                name = name[:name.rfind("（中）")]
        if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                "（四）") or name.endswith("（五）") or name.endswith("（六）"):
            name = name[:name.find("（")]
        video_name_set.add(name)
    for name in video_name_set:
        print i, name
        write_mapping_log("2016-2018-Video.txt", "清晰版" + "\t" + year + "\t" + name + "\n")
        i = i + 1


def get_courseware_info_according_video_resource(srcdir):
    """
    提取课件名称
    :param srcdir: 视频源文件所在目录
    :return:
    """
    year = srcdir[srcdir.rfind("\\") + 1:]
    i = 1
    for name in os.listdir(srcdir):
        name = video_name = name[:name.rfind(".mp4")]
        if name.endswith("（上）") or name.endswith("（下）") or name.endswith("（中）"):
            if name.endswith("（上）"):
                name = name[:name.rfind("（上）")]
            elif name.endswith("（下）"):
                name = name[:name.rfind("（下）")]
            elif name.endswith("（中）"):
                name = name[:name.rfind("（中）")]
        if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                "（四）") or name.endswith("（五）") or name.endswith("（六）"):
            name = name[:name.find("（")]
        print i, name, video_name
        write_mapping_log("2016-2018-Courseware.txt", "清晰版" + "\t" + year + "\t" + name + "\t" + video_name + "\n")
        i = i + 1


def get_test_file(srcdir, pos):
    """
    :param srcdir:
    :param pos:
    :return:
    """
    i = 1
    pdir = os.path.dirname(srcdir)
    year = pdir[pdir.rfind("\\") + 1:]
    for name in os.listdir(srcdir):
        name = name[:name.rfind(".")]
        print i, year, name
        write_mapping_log((year + "-" + pos + "-试题列表.txt").encode("gbk"), name + "\n")
        i = i + 1


set_file1 = set()
set_file2 = set()


def get_diff_item(file1, file2, extract='清晰版', year='2012'):
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
            if (info1[0] == extract) and (info1[1] == year):
                set_file1.add(info1[3].strip("\n"))
    with open(file2, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            set_file2.add(info2[0].strip("\n"))

    print len(set_file1), len(set_file2)
    set_diff = set_file1.difference(set_file2)
    m = 1
    for obj in set_diff:
        print m, obj
        m = m + 1


set_file1 = set()
set_file2 = set()


def get_inter_item(file1, file2, extract='清晰版', year='2012'):
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
            if (info1[0] == extract) and (info1[1] == year):
                set_file1.add(info1[3].strip("\n"))
    with open(file2, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            set_file2.add(info2[0].strip("\n"))
    print len(set_file1), len(set_file2)
    set_diff = set_file1.intersection(set_file2)
    m = 1
    for obj in set_diff:
        print m, obj
        m = m + 1


i = 1


def get_all_xls(srcfolder):
    """
    获取所有的试题列表
    :param srcfolder:
    :return:
    """
    global i
    for name in os.listdir(srcfolder):
        srcname = srcfolder + "\\" + name
        if os.path.isdir(srcname):
            get_all_xls(srcname)
        else:
            if name.endswith("xls") or name.endswith("xlsx"):
                print i, name[:name.rfind(".")]
                # write_mapping_log("2012-2015试题列表.txt".encode("gbk"), name[:name.rfind(".")] + "\n")
                write_mapping_log("2016-2018试题列表.txt".encode("gbk"), name[:name.rfind(".")] + "\n")
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
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                print j, name
                # write_mapping_log("2012-2015试题列表-合并前.txt".encode("gbk"), name + "\n")
                write_mapping_log("2016-2018试题列表-合并前.txt".encode("gbk"), name + "\n")
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
        # write_mapping_log("2012-2015试题列表-合并后.txt".encode("gbk"), obj + "\n")
        write_mapping_log("2016-2018试题列表-合并后.txt".encode("gbk"), obj + "\n")
        j = j + 1


def get_diff_items(file1, file2):
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
            set_file1.add(info1[2].strip("\n"))
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
        # write_mapping_log("2012-2015试题列表[待审核].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2016-2018试题列表[待审核].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2012-2015PPT列表[待审核].txt".encode("gbk"), obj + "\n")
        write_mapping_log("2016-2018PPT列表[待审核].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2012-2015文稿列表[待审核].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2016-2018文稿列表[待审核].txt".encode("gbk"), obj + "\n")
        m = m + 1


def get_inter_items(file1, file2):
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
            set_file1.add(info1[2].strip("\n"))
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
        # write_mapping_log("2012-2015试题列表[同数据库对应].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2016-2018试题列表[同数据库对应].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2012-2015PPT列表[同数据库对应].txt".encode("gbk"), obj + "\n")
        write_mapping_log("2016-2018PPT列表[同数据库对应].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2012-2015文稿列表[同数据库对应].txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2016-2018文稿列表[同数据库对应].txt".encode("gbk"), obj + "\n")
        print m, obj
        m = m + 1


def get_diff_items_1(file1, file2):
    """
    获取不同类型之间的差值部分
    :param file1:
    :param file2:
    :return:
    """
    with open(file2, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            set_file2.add(info1[2].strip("\n"))
    with open(file1, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            set_file1.add(info2[0].strip("\n"))

    set_diff = set_file1.difference(set_file2)
    m = 1
    for obj in set_diff:
        print m, obj
        # write_mapping_log("2012-2015试题列表[待审核]-1.txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2012-2015PPT列表[待审核]-1.txt".encode("gbk"), obj + "\n")
        # write_mapping_log("2012-2015文稿列表[待审核]-1.txt".encode("gbk"), obj + "\n")
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
            if shiti_name.endswith("无"):
                continue
            print j, course_name, shiti_name
            # write_mapping_log("2012-2015试题列表[已审核]--去除无.txt".encode("gbk"), course_name + "\t" + shiti_name + "\n")
            # write_mapping_log("2016-2018试题列表[已审核]--去除无.txt".encode("gbk"), course_name + "\t" + shiti_name + "\n")
            # write_mapping_log("2012-2015PPT列表[已审核]--去除无.txt".encode("gbk"), course_name + "\t" + shiti_name + "\n")
            # write_mapping_log("2016-2018PPT列表[已审核]--去除无.txt".encode("gbk"), course_name + "\t" + shiti_name + "\n")
            # write_mapping_log("2012-2015文稿列表[已审核]--去除无.txt".encode("gbk"), course_name + "\t" + shiti_name + "\n")
            write_mapping_log("2016-2018文稿列表[已审核]--去除无.txt".encode("gbk"), course_name + "\t" + shiti_name + "\n")
            j = j + 1


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
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                with open(u"2016-2018试题列表[已审核]--去除无.txt", "r") as f1:
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
                            # write_mapping_log("2012-2015试题列表[最终映射].txt".encode("gbk"),
                            #                   course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            write_mapping_log("2016-2018试题列表[最终映射].txt".encode("gbk"),
                                              course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            j = j + 1
                            break


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
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                with open(u"2016-2018试题列表[同数据库对应].txt", "r") as f1:
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
                            # write_mapping_log("2012-2015试题列表[最终映射]-3.txt".encode("gbk"),
                            #                   course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            write_mapping_log("2016-2018试题列表[最终映射]-2.txt".encode("gbk"),
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


def del_blank_char(str):
    """
    去除字符串中的空白跟换行符
    :param str:
    :return:
    """
    rep = re.compile(r'(\n|\t|\r)')
    (fstr, count) = rep.subn('', str)
    return fstr


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
                    write_mapping_log("2016-2018试题文件.txt".encode("gbk"),
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
                        write_mapping_log("2016-2018试题文件.txt".encode("gbk"),
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
                        write_mapping_log("2016-2018试题文件.txt".encode("gbk"),
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
                write_mapping_log("2016-2018试题文件.txt".encode("gbk"),
                                  course_name + "\t" + knowledge + "\t" + str(types) + "\t" + str(
                                      difficult) + "\t" + str(
                                      score) + "\t" + question + "\t" + option + "\t" + str(answer) + "\n")
                k = k + 1


n = 1


def get_excel_str(srcfile):
    # SOURCE_DIR = "E:\\Goosuu\\Dang\\Script\\Excel\\Second\\Source\\3"
    # SOURCE_DIR = "E:\\Goosuu\\Dang\\Script\\Excel\\Second\\Source\\4"
    SOURCE_DIR = "E:\\Goosuu\\Dang\\Script\\Excel\\Third\\Source\\2"
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


def get_all_ppt(srcfolder):
    """
    获取所有的PPT列表
    :param srcfolder:
    :return:
    """
    global i
    for name in os.listdir(srcfolder):
        srcname = srcfolder + "\\" + name
        if os.path.isdir(srcname):
            get_all_ppt(srcname)
        else:
            if name.endswith("ppt") or name.endswith("pptx"):
                print i, name[:name.rfind(".")]
                # write_mapping_log("2012-2015PPT列表.txt".encode("gbk"), name[:name.rfind(".")] + "\n")
                write_mapping_log("2016-2018PPT列表.txt".encode("gbk"), name[:name.rfind(".")] + "\n")
                i = i + 1


def get_all_pname_ppt(srcfolder):
    global j
    for name in os.listdir(srcfolder):
        srcname = srcfolder + "\\" + name
        if os.path.isdir(srcname):
            get_all_pname_ppt(srcname)
        else:
            if name.endswith("ppt") or name.endswith("pptx"):
                name = name[:name.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                print j, name
                # write_mapping_log("2012-2015PPT列表-合并前.txt".encode("gbk"), name + "\n")
                write_mapping_log("2016-2018PPT列表-合并前.txt".encode("gbk"), name + "\n")
                j = j + 1


def get_ppt_info(srcfile):
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
        # write_mapping_log("2012-2015PPT列表-合并后.txt".encode("gbk"), obj + "\n")
        write_mapping_log("2016-2018PPT列表-合并后.txt".encode("gbk"), obj + "\n")
        j = j + 1


def move_ppt_dir(srcfolder, destfolder):
    if not os.path.exists(destfolder):
        os.makedirs(destfolder)
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        if os.path.isdir(file):
            move_ppt_dir(file, destfolder)
        else:
            if filename.endswith("ppt") or filename.endswith("pptx"):
                name = filename[:filename.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                with open(u"2016-2018PPT列表[已审核]--去除无.txt", "r") as f1:
                    global j
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        course_name = info1[0]
                        ppt_name = info1[1].strip("\n")
                        if name == ppt_name:
                            srcfile = srcfolder + "\\" + filename
                            newfolder = destfolder + "\\" + course_name.strip()
                            if not os.path.exists(newfolder):
                                os.makedirs(newfolder)
                            destfile = newfolder + "\\" + filename
                            print j, course_name, filename, srcfile, destfile
                            shutil.copy2(srcfile, destfile)
                            # write_mapping_log("2012-2015PPT列表[最终映射].txt".encode("gbk"),
                            #                   course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            write_mapping_log("2016-2018PPT列表[最终映射].txt".encode("gbk"),
                                              course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            j = j + 1
                            break


def move_ppt_dir2(srcfolder, destfolder):
    if not os.path.exists(destfolder):
        os.makedirs(destfolder)
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        if os.path.isdir(file):
            move_ppt_dir2(file, destfolder)
        else:
            if filename.endswith("ppt") or filename.endswith("pptx"):
                name = filename[:filename.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                with open(u"2016-2018PPT列表[同数据库对应].txt", "r") as f1:
                    global j
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        course_name = info1[0].strip("\n")
                        ppt_name = info1[0].strip("\n")
                        if name == ppt_name:
                            srcfile = srcfolder + "\\" + filename
                            newfolder = destfolder + "\\" + course_name.strip()
                            if not os.path.exists(newfolder):
                                os.makedirs(newfolder)
                            destfile = newfolder + "\\" + filename
                            print j, course_name, filename, srcfile, destfile
                            shutil.copy2(srcfile, destfile)
                            # write_mapping_log("2012-2015PPT列表[最终映射]-2.txt".encode("gbk"),
                            #                   course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            write_mapping_log("2016-2018PPT列表[最终映射]-2.txt".encode("gbk"),
                                              course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            j = j + 1
                            break


def get_all_doc(srcfolder):
    """
    获取所有的试题列表
    :param srcfolder:
    :return:
    """
    global i
    for name in os.listdir(srcfolder):
        srcname = srcfolder + "\\" + name
        if os.path.isdir(srcname):
            get_all_doc(srcname)
        else:
            if name.endswith("doc") or name.endswith("docx"):
                print i, name[:name.rfind(".")]
                # write_mapping_log("2012-2015文稿列表.txt".encode("gbk"), name[:name.rfind(".")] + "\n")
                write_mapping_log("2016-2018文稿列表.txt".encode("gbk"), name[:name.rfind(".")] + "\n")
                i = i + 1


def get_all_pname_doc(srcfolder):
    global j
    for name in os.listdir(srcfolder):
        srcname = srcfolder + "\\" + name
        if os.path.isdir(srcname):
            get_all_pname_doc(srcname)
        else:
            if name.endswith("doc") or name.endswith("docx"):
                name = name[:name.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                print j, name
                # write_mapping_log("2012-2015文稿列表-合并前.txt".encode("gbk"), name + "\n")
                write_mapping_log("2016-2018文稿列表-合并前.txt".encode("gbk"), name + "\n")
                j = j + 1


def get_doc_info(srcfile):
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
        # write_mapping_log("2012-2015文稿列表-合并后.txt".encode("gbk"), obj + "\n")
        write_mapping_log("2016-2018文稿列表-合并后.txt".encode("gbk"), obj + "\n")
        j = j + 1


def move_doc_dir(srcfolder, destfolder):
    if not os.path.exists(destfolder):
        os.makedirs(destfolder)
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        if os.path.isdir(file):
            move_doc_dir(file, destfolder)
        else:
            if filename.endswith("doc") or filename.endswith("docx"):
                name = filename[:filename.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                with open(u"2016-2018文稿列表[已审核]--去除无.txt", "r") as f1:
                    global j
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        course_name = info1[0]
                        ppt_name = info1[1].strip("\n")
                        if name == ppt_name:
                            srcfile = srcfolder + "\\" + filename
                            newfolder = destfolder + "\\" + course_name.strip()
                            if not os.path.exists(newfolder):
                                os.makedirs(newfolder)
                            destfile = newfolder + "\\" + filename
                            print j, course_name, filename, srcfile, destfile
                            shutil.copy2(srcfile, destfile)
                            # write_mapping_log("2012-2015文稿列表[最终映射].txt".encode("gbk"),
                            #                   course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            write_mapping_log("2016-2018文稿列表[最终映射].txt".encode("gbk"),
                                              course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            j = j + 1
                            break


def move_doc_dir2(srcfolder, destfolder):
    if not os.path.exists(destfolder):
        os.makedirs(destfolder)
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        if os.path.isdir(file):
            move_doc_dir2(file, destfolder)
        else:
            if filename.endswith("doc") or filename.endswith("docx"):
                name = filename[:filename.rfind(".")]
                if name.rfind("（下）") > -1:
                    name = name[:name.rfind("（下）")]
                if name.rfind("（上）") > -1:
                    name = name[:name.rfind("（上）")]
                if name.rfind("（中）") > -1:
                    name = name[:name.rfind("（中）")]
                if name.endswith("（一）") or name.endswith("（二）") or name.endswith("（三）") or name.endswith(
                        "（四）") or name.endswith("（五）") or name.endswith("（六）"):
                    name = name[:name.find("（")]
                with open(u"2016-2018文稿列表[同数据库对应].txt", "r") as f1:
                    global j
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        course_name = info1[0].strip("\n")
                        ppt_name = info1[0].strip("\n")
                        if name == ppt_name:
                            srcfile = srcfolder + "\\" + filename
                            newfolder = destfolder + "\\" + course_name.strip()
                            if not os.path.exists(newfolder):
                                os.makedirs(newfolder)
                            destfile = newfolder + "\\" + filename
                            print j, course_name, filename, srcfile, destfile
                            shutil.copy2(srcfile, destfile)
                            write_mapping_log("2016-2018文稿列表[最终映射]-2.txt".encode("gbk"),
                                              course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            # write_mapping_log("2012-2015文稿列表[最终映射]-2.txt".encode("gbk"),
                            #                   course_name + "\t" + filename + "\t" + srcfile + "\t" + destfile + "\n")
                            j = j + 1
                            break


def convert_doc_docx(srcfolder, srcfile, destfolder):
    """
    转换doc文件为docx文件
    :param srcfolder: 源文件夹
    :param srcfile: doc文件
    :return:
    """
    word = client.Dispatch("Word.Application")
    wordFullName = os.path.join(srcfolder, srcfile)
    doc = word.Documents.Open(wordFullName)
    name = srcfile[:srcfile.rfind(".")]
    docxFullName = os.path.join(destfolder, name)
    doc.SaveAs(docxFullName.encode("gbk") + ".docx", 16)
    doc.Close()
    word.Quit()


def covert_all_doctodocx(srcfolder, destfolder):
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        if os.path.isdir(file):
            covert_all_doctodocx(file, destfolder)
        else:
            if filename.endswith(".docx"):
                # 复制操作：
                srcfile = os.path.join(srcfolder, filename)
                pname = srcfolder[srcfolder.rfind("\\") + 1:]
                newfolder = os.path.join(destfolder, pname)
                if not os.path.exists(newfolder):
                    os.makedirs(newfolder)
                destfile = os.path.join(newfolder, filename)
                print srcfile, " Copy..."
                shutil.copy2(srcfile, destfile)
            elif filename.endswith(".doc"):
                # 转换操作：
                pname = srcfolder[srcfolder.rfind("\\") + 1:]
                newfolder = os.path.join(destfolder, pname)
                if not os.path.exists(newfolder):
                    os.makedirs(newfolder)
                convert_doc_docx(srcfolder, filename, newfolder)
                print filename, " Transfer..."


def convert_docx_html(srcfile, destdir):
    """
    转换docx文件为html文件
    :param srcfile: docx文件
    :param destdir: docx文件
    :return:
    """
    html = PyDocX.to_html(srcfile)
    name = srcfile[:srcfile.rfind(".")]
    # destfile = (destdir + "\\" + name +".html").encode("gbk")
    destfile = name.encode("gbk") + ".html"
    f = open(destfile, 'w')
    f.write(html.encode("utf-8"))
    f.close()


def convert_docx_html_from_dir(srcdir, destdir):
    if not os.path.exists(destdir):
        os.makedirs(destdir)
    i = 1
    for filename in os.listdir(srcdir):
        srcfile = srcdir + "\\" + filename
        convert_docx_html(srcfile, destdir)
        print i, srcfile
        i = i + 1


def rename_ppt_pic(srcfolder):
    """
    重命名PPT图片
    :param srcfolder:
    :return:
    """
    global i
    re_dit = re.compile("\d+")
    for filename in os.listdir(srcfolder):
        srcfile = os.path.join(srcfolder, filename)
        if os.path.isdir(srcfile):
            for name in os.listdir(srcfile):
                if name.endswith("JPG"):
                    order = re_dit.search(name).group()
                    source_file = srcfile + "\\" + name
                    newfile = srcfile + "\\" + order + ".jpg"
                    print i, source_file, newfile
                    os.rename(source_file, newfile)
                    i = i + 1


def compute_ppt_count(srcfolder):
    global i
    for root, dirs, files in os.walk(srcfolder):  # 遍历统计
        count = 0
        for each in files:
            count += 1  # 统计文件夹下文件个数
        rootname = root[root.rfind("\\") + 1:]
        if count == 0:
            continue
        print i, rootname, count  # 输出结果
        write_mapping_log("2016-2018PPT图片个数统计.txt".encode("gbk"), rootname + "\t" + str(count) + "\n")
        i = i + 1


def get_video_duration(srcfolder):
    for filename in os.listdir(srcfolder):
        srcfile = os.path.join(srcfolder, filename)
        if os.path.isdir(srcfile):
            get_video_duration(srcfile)
        else:
            if filename.endswith(".mp4"):
                print filename
                wg = subprocess.Popen(['ffprobe.exe', '-i', (srcfile.strip()).encode("gbk")], stdout=subprocess.PIPE,
                                      stderr=subprocess.STDOUT)
                (standardout, junk) = wg.communicate()
                ans = str(standardout)
                num = ans.find("Duration:")
                out = ans[num + 10:num + 18]
                print srcfile, filename[:filename.rfind(".mp4")], out
                write_mapping_log("2012-2015-Courseware-Duration.txt",
                                  filename[:filename.rfind(".mp4")] + "\t" + out + "\n")


def t2s(t):
    """
    时间转化为秒数
    :param t:
    :return:
    """
    h, m, s = t.strip().split(":")
    return int(h) * 3600 + int(m) * 60 + int(s)


def s2t(seconds):
    """
    秒数转化为时间
    :param seconds:
    :return:
    """
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return ("%02d:%02d:%02d" % (h, m, s))


def get_sum_video_duration(video_file, courseware_duration_file):
    """
    统计视频时长
    :param video_file:
    :param courseware_duration_file:
    :return:
    """
    global i
    with open(video_file, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            course_name = info1[2].strip("\n")
            sum_duration = 0
            with open(courseware_duration_file, "r") as f2:
                while True:
                    line2 = f2.readline()
                    if len(line2) < 1:
                        break
                    info2 = regex_tab.split(line2)
                    coursewarename = info2[0]
                    duration = info2[1].strip("\n")
                    if coursewarename.startswith(course_name):
                        sum_duration = sum_duration + t2s(duration)
            print i, course_name, s2t(sum_duration)
            write_mapping_log("2012-2015-Video-Duration.txt", course_name + "\t" + s2t(sum_duration) + "\n")
            i = i + 1


def create_course_code(srcfile):
    global i
    i = 250
    with open(srcfile, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            order = "%04d" % i
            course_code = "P201905071039" + str(order)
            print course_code
            play_num = random.randint(1,100)
            order_num = 999
            like_count = random.randint(1,50)
            course_period = 3
            write_mapping_log("2016-2018-Video-CourseCode.txt",info1[0]+"\t"+info1[1]+"\t"+info1[2].strip("\n")+"\t"+course_code+"\t"+course_code+".jpg"+"\t"+str(play_num)+"\t"+str(order_num)+"\t"+str(like_count)+"\t"+str(course_period)+"\n")
            i = i + 1

def copy_video(srcfolder,destfolder):
    global i
    if not os.path.exists(destfolder):
        os.makedirs(destfolder)
    for filename in os.listdir(srcfolder):
        srcfile = os.path.join(srcfolder,filename)
        if os.path.isdir(srcfile):
            newfolder = destfolder + "\\" + filename
            copy_video(srcfile,newfolder)
        else:
            if filename.endswith(".mp4"):
                destfile = destfolder + "\\" + filename
                shutil.copy2(srcfile,destfile)
                print i,srcfile,destfile
                i = i + 1

def rename_video_resource(source,srcfolder):
    global i
    for filename in os.listdir(srcfolder):
        srcfile = os.path.join(srcfolder,filename)
        if os.path.isdir(srcfile):
            rename_video_resource(source,srcfile)
        else:
            if srcfile.endswith(".mp4"):
                filename = filename[:filename.rfind(".mp4")]
                with open(source, "r") as f1:
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        courseware_name = info1[4]
                        order = info1[5]
                        course_code = info1[6].strip("\n")
                        id = info1[0]
                        year = info1[2]
                        if filename == courseware_name:
                            newname = course_code + "-" + order + ".mp4"
                            newpath = "Normal/"+ year + "/" + newname
                            write_mapping_log("Courseware_Path_2012-2015.txt",id+"\t"+newpath+"\n")
                            destfile = srcfolder + "\\" + newname
                            os.rename(srcfile,destfile)
                            # print i,srcfile,destfile
                            print i,newpath
                            i = i + 1
                            break


def rename_ppt_folder(source,srcfolder):
    global i
    for filename in os.listdir(srcfolder):
        srcfile = os.path.join(srcfolder, filename)
        if os.path.isdir(srcfile):
                with open(source, "r") as f1:
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        courseware_name = info1[2]
                        course_code = info1[3]
                        if filename == courseware_name:
                            destfile = srcfolder + "\\" + course_code
                            print i,srcfile,destfile
                            os.rename(srcfile, destfile)
                            i = i + 1
                            break

def rename_doc(source,srcfolder):
    global i
    for filename in os.listdir(srcfolder):
        srcfile = os.path.join(srcfolder, filename)
        filename = filename[:filename.rfind(".html")]
        if os.path.isfile(srcfile):
                with open(source, "r") as f1:
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        courseware_name = info1[2]
                        course_code = info1[3]
                        if filename == courseware_name:
                            destfile = srcfolder + "\\" + course_code + ".html"
                            print i,srcfile,destfile
                            os.rename(srcfile, destfile)
                            i = i + 1
                            break

def rename_cover(source,srcfolder):
    global i
    for filename in os.listdir(srcfolder):
        srcfile = os.path.join(srcfolder, filename)
        if os.path.isdir(srcfile):
            rename_cover(source, srcfile)
        else:
            if srcfile.endswith(".jpg"):
                filename = filename[:filename.rfind(".jpg")]
                with open(source, "r") as f1:
                    while True:
                        line1 = f1.readline()
                        if len(line1) < 1:
                            break
                        info1 = regex_tab.split(line1)
                        jpeg_name = info1[4]
                        course_name = info1[2]
                        if filename == course_name:
                            destfile = srcfolder + "\\" + jpeg_name
                            print i, srcfile, destfile
                            os.rename(srcfile, destfile)
                            i = i + 1
                            break

def main():
    # ①提取课程名称：将（上）（中）（下）（一）（二）（三）（四）这种合并成一个，有的（一）（二）是一组，跟（三）（四）分开
    # srcdir = u"K:\\清晰版转码-mp4\\2015"
    # srcdir = u"K:\\不清晰版\\视频资料\\2014"
    # srcdir = u"J:\\视频转码-MP4\\2016"
    # srcdir = u"J:\\视频转码-MP4\\2017"
    # srcdir = u"J:\\视频资料\\2018年成片备份"
    # get_course_info_according_video_resource(srcdir)

    # ②提取课件名称：
    # srcdir = u"K:\\清晰版转码-mp4\\2015"
    # srcdir = u"K:\\不清晰版\\视频资料\\2014"
    # srcdir = u"J:\\视频转码-MP4\\2016"
    # srcdir = u"J:\\视频转码-MP4\\2017"
    # srcdir = u"J:\\视频资料\\2018年成片备份"
    # get_courseware_info_according_video_resource(srcdir)

    """
    pass
    这里可以做单年份试题信息的校对
    # ③ 获取试题列表：
    # srcdir = u"K:\\不清晰版\\文本资料\\2014\\试题"
    # pos = "不清晰版"
    # get_test_file(srcdir,pos)

    pass
    # ④ 校对试题，看匹配的有多少
    file1 = u"2012-2015-Courseware.txt"
    file2 = u"2015-清晰版-试题列表.txt"
    # get_diff_item(file1,file2,extract = u'清晰版',year = '2015')
    # get_inter_item(file1,file2,extract = u'清晰版',year = '2015')
    """

    # ③处理试题：获取所有的试题列表
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_all_xls(srcfolder)

    # 合并前试题列表：
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_all_pname_xls(srcfolder)

    # 获取合并后的试题列表
    # srcfile = u"2012-2015试题列表-合并前.txt"
    # srcfile = u"2016-2018试题列表-合并前.txt"
    # get_test_info(srcfile)

    # 获取能跟数据库对应上的列表,跟对应不上的列表
    # file1 = u"2012-2015-Video.txt"
    # file2 = u"2012-2015试题列表-合并后.txt"
    # file1 = u"2016-2018-Video.txt"
    # file2 = u"2016-2018试题列表-合并后.txt"
    # get_inter_items(file1,file2)
    # get_diff_items(file1,file2)
    # 这个将来做给校审人员比对结果校对用
    # get_diff_items_1(file2, file1)

    # 对审核的2012-2015试题列表去除无试题的
    # srcfile = u"2012-2015试题列表[已审核].txt"
    # srcfile = u"2016-2018试题列表[已审核].txt"
    # shenhe_txt(srcfile)

    # 获取审核完毕的Excel源文件：
    # destfolder = u"E:\\Goosuu\\Dang\\Script\\Excel\\Third\\Source\\1"
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_excel_file(srcfolder,destfolder)

    # 获取无问题的Excel源文件：
    # destfolder = u"E:\\Goosuu\\Dang\\Script\\Excel\\Third\\Source\\2"
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_excel_file2(srcfolder, destfolder)

    # 读取试题内容：
    # srcfile = u"2012-2015试题列表[最终映射]-3.txt"
    # srcfile = u"Txt/2012-2015试题列表-Question.txt"
    # srcfile = u"2016-2018试题列表[最终映射]-2.txt"
    # get_excel_str(srcfile)

    # 校对人工整理试题课程跟数据库入库之后试题课程是否一致

    # ④处理PPT ：获取所有的PPT列表
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_all_ppt(srcfolder)

    # 合并前PPT列表：
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_all_pname_ppt(srcfolder)

    # 获取合并后的PPT列表
    # srcfile = u"2012-2015PPT列表-合并前.txt"
    # srcfile = u"2016-2018PPT列表-合并前.txt"
    # get_ppt_info(srcfile)

    # 获取能跟数据库对应上的列表,跟对应不上的列表
    # file1 = u"2012-2015-Video.txt"
    # file2 = u"2012-2015PPT列表-合并后.txt"
    # file1 = u"2016-2018-Video.txt"
    # file2 = u"2016-2018PPT列表-合并后.txt"
    # get_inter_items(file1,file2)
    # get_diff_items(file1,file2)
    # 这个将来做给校审人员比对结果校对用
    # get_diff_items_1(file2, file1)

    # 对审核的2012-2015PPT去除无PPT的
    # srcfile = u"2012-2015PPT列表[已审核].txt"
    # srcfile = u"2016-2018PPT列表[已审核].txt"
    # shenhe_txt(srcfile)

    # 将审核完毕每门课程下的PPT散列到对应的文件夹下
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # destfolder = u"E:\\Goosuu\\Dang\\Script\\PPT\\PowerPoint1"
    # move_ppt_dir(srcfolder,destfolder)

    # 将没问题每门课程下的PPT散列到对应的文件夹下
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # destfolder = u"E:\\Goosuu\\Dang\\Script\\PPT\\PowerPoint2"
    # move_ppt_dir2(srcfolder,destfolder)

    # 合并PPT -- 在另外一个文件中完成

    # PPT转化为图片 -- 在另外一个文件中完成

    # PPT重命名
    # srcfolder = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2012-2015\\PPT\\PPT-Picture1"
    # srcfolder = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2016-2018\\PPT\\PPT-Picture"
    # rename_ppt_pic(srcfolder)

    # 统计PPT个数
    # srcfolder = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2012-2015\\PPT\\PPT-Picture1"
    # srcfolder = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2016-2018\\PPT\\PPT-Picture"
    # compute_ppt_count(srcfolder)

    # 重命名PPT文件夹
    # srcfile = u"2012-2015-Video-CourseCode.txt"
    # srcfolder = u"J:\\党政党建视频课程\\成品数据（新添加519门课程）\\PPT\\2012-2015\\PPT-Picture"
    # rename_ppt_folder(srcfile,srcfolder)


    # ⑤处理文稿 ：获取所有的DOC列表
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_all_doc(srcfolder)

    # 合并前DOC列表：
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # get_all_pname_doc(srcfolder)

    # 获取合并后的DOC列表
    # srcfile = u"2012-2015文稿列表-合并前.txt"
    # srcfile = u"2016-2018文稿列表-合并前.txt"
    # get_doc_info(srcfile)

    # 获取能跟数据库对应上的列表,跟对应不上的列表
    # file1 = u"2012-2015-Video.txt"
    # file2 = u"2012-2015文稿列表-合并后.txt"
    # file1 = u"2016-2018-Video.txt"
    # file2 = u"2016-2018文稿列表-合并后.txt"
    # get_inter_items(file1,file2)
    # get_diff_items(file1,file2)
    # 这个将来做给校审人员比对结果校对用
    # get_diff_items_1(file2, file1)

    # 对审核的2012-2015文稿去除无文稿的
    # srcfile = u"2012-2015文稿列表[已审核].txt"
    # srcfile = u"2016-2018文稿列表[已审核].txt"
    # shenhe_txt(srcfile)

    # 将审核完毕每门课程下的文稿散列到对应的文件夹下
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # destfolder = u"E:\\Goosuu\\Dang\\Script\\HTML\\Draft1"
    # move_doc_dir(srcfolder,destfolder)

    # 将没问题每门课程下的文稿散列到对应的文件夹下
    # srcfolder = u"K:\\清晰版\\文本资料"
    # srcfolder = u"K:\\不清晰版\\文本资料"
    # srcfolder = u"J:\\文字资料"
    # destfolder = u"E:\\Goosuu\\Dang\\Script\\HTML\\Draft2"
    # move_doc_dir2(srcfolder,destfolder)


    # doc转换为docx
    # srcfolder = u"E:\\Goosuu\\Dang\\Script\\HTML\\Draft2"
    # destfolder = u"E:\\Goosuu\\Dang\\Script\\HTML\\DraftDocx2"
    # covert_all_doctodocx(srcfolder,destfolder)

    # 合并文稿：---已交给人工完成

    # 文稿转化为HTML
    # srcdir = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2012-2015\\Draft\\Combine_Docx2"
    # destdir = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2012-2015\\Draft\\HTML2"
    # srcdir = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2016-2018\\Draft\\Combine_Docx2"
    # destdir = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2016-2018\\Draft\\HTML2"
    # convert_docx_html_from_dir(srcdir,destdir)

    # 文稿重命名
    # srcfile = u"2012-2015-Video-CourseCode.txt"
    # srcfolder = u"J:\\党政党建视频课程\\成品数据（新添加519门课程）\\Draft\\HTML-2012-2015"
    # rename_doc(srcfile,srcfolder)

    # 统计单个视频时长：
    # srcfolder = u"J:\\Deal-Middle\\Videos"
    # srcfolder = u"I:\\Deal-Middle\\Videos"
    # get_video_duration(srcfolder)

    # 统计单门课程的时长
    # video_file = u"2016-2018-Video.txt"
    # courseware_duration_file = u"2016-2018-Courseware-Duration.txt"
    # video_file = u"2012-2015-Video.txt"
    # courseware_duration_file = u"2012-2015-Courseware-Duration.txt"
    # get_sum_video_duration(video_file,courseware_duration_file)

    # 生成课程码
    # srcfile = u"2016-2018-Video.txt"
    # create_course_code(srcfile)

    # 复制视频：
    # srcfolder = u"I:\\Deal-Middle\\Videos"
    # destfolder = u"J:\\党政党建视频课程\\成品数据（新添加519门课程）\\Videos\\2016-2018"
    # copy_video(srcfolder,destfolder)

    # 重命名视频
    # srcfile = u"2012-2018课程信息映射.txt"
    # srcfolder = u"J:\\党政党建视频课程\\成品数据（新添加519门课程）\\Videos\\2012-2015\\不清晰版\\2014"
    # rename_video_resource(srcfile,srcfolder)

    # 重命名封面
    # srcfile = u"2012-2015-Video-CourseCode.txt"
    # srcfolder = u"J:\\党政党建视频课程\\成品数据（新添加519门课程）\\Covers\\2012-2015\\不清晰\\2014"
    # rename_cover(srcfile,srcfolder)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
