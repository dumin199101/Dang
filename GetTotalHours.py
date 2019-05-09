# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/7 15:03
@Name: GetTotalHours
@Author: lieyan123091
"""
import re, sys

regex_tab = re.compile("\\t")
regex_time = re.compile(":")

i = 1
sh = sm = ss =0

def get_total_hours(srcfile):
    global i,sh,sm,ss
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            time_info = info[4]
            info1 = regex_time.split(time_info)
            h = info1[0]
            m = info1[1]
            s = info1[2]
            sh = sh + int(h)
            sm = sm + int(m)
            ss = ss + int(s)
            print i,sh,sm,ss
            i = i + 1




def main():
    srcfile = u"../Txt/courseware.txt"
    get_total_hours(srcfile)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
