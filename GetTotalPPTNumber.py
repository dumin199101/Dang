# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/8 12:38
@Name: GetTotalPPTNumber
@Author: lieyan123091
"""
import re, sys

regex_tab = re.compile("\\t")

i = 1
sum = 0

def get_total_ppt(srcfile):
    global i,sum
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            pptnum = info[10]
            sum = sum + int(pptnum)
            print i,sum
            i = i + 1





def main():
    srcfile = u"course.txt"
    get_total_ppt(srcfile)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()