# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/8 13:23
@Name: CreatePPTFileName
@Author: lieyan123091
生成PPT文件名
"""
import re, sys, os,shutil

regex_tab = re.compile("\\t")

def write_mapping_log(logname, content):
    """
    生成mapping日志文件
    :return:
    """
    with open(logname, "a+") as f1:
        f1.write(content.encode("utf-8"))

i = 1

def create_ppt_filename(srcfile):
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            id = info[0]
            name = info[1]
            course_id = info[2]
            ppt_num = int(info[10])
            for j in range(1,ppt_num+1):
                newfile = course_id + "-" + str(j)  + ".jpg"
                print i,newfile
                write_mapping_log("PPTFileName.txt",str(i) + "\t" + newfile+"\n")
                i = i + 1



def main():
    srcfile = u"course.txt"
    create_ppt_filename(srcfile)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()