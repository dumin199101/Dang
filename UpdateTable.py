# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/8 16:56
@Name: UpdateTable
@Author: lieyan123091
更新表数据
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

def update_data(srcfile):
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            id = info[0]
            path = info[4][info[4].rfind("/")+1:]
            write_mapping_log("course_id_path.txt",id+"\t"+path+"\n")





def main():
    srcfile = u"course.txt"
    update_data(srcfile)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()