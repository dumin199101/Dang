# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/8 11:42
@Name: GetCourseCover
@Author: lieyan123091
获取课程封面
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
destdir = "I:\\Deal-Middle\\Covers"

def get_cover_resource(srcfile,srcdir):
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            id = info[0]
            name = info[1]
            coverpath = info[4]
            covername = coverpath[coverpath.rfind("/")+1:]
            for filename in os.listdir(srcdir):
                if filename == covername:
                    srcfile = srcdir + "\\" + filename
                    destfolder = destdir
                    if not os.path.exists(destfolder):
                        os.makedirs(destfolder)
                    destfile = destfolder + "\\" + filename
                    shutil.copy2(srcfile,destfile)
                    # write_mapping_log("Covers.txt",id+"\t"+name+"\t"+filename+"\n")
                    print i,srcfile,destfile
                    i = i + 1


def main():
    srcfile = u"course.txt"
    srcdir = u"I:\\党政慕课资源\\课程封面图片"
    get_cover_resource(srcfile,srcdir)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()