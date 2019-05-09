# coding=utf-8
"""
@Project: Dang
@Time: 2019/4/24 14:26
@Name: Author_Format.py
@Author: lieyan123091
"""

import re, sys, os
regex_tab = re.compile(" ")

def write_mapping_log(logname, content):
    """
    生成mapping日志文件
    :return:
    """
    with open(logname, "a+") as f1:
        f1.write(content.encode("utf-8"))


def shenhe_txt(srcfile):
    j = 1
    with open(srcfile, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            author = info1[0]
            desc = info1[1].strip("\n")
            print j,author,desc
            write_mapping_log("author.txt",author+"\t"+desc+"\n")
            j = j + 1


def main():
    srcfile = u"author_format.txt"
    shenhe_txt(srcfile)

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()