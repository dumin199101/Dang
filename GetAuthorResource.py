# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/8 15:34
@Name: GetAuthorResource
@Author: lieyan123091
获取作者图片
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
destdir = "J:\\党政党建视频课程\\成品数据（新添加519门课程）\\Author"

def get_author_resource(srcfile,srcdir):
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            id = info[0]
            name = info[1]
            authorpath = info[3]
            for filename in os.listdir(srcdir):
                if filename == authorpath:
                    srcfile = srcdir + "\\" + filename
                    destfolder = destdir
                    if not os.path.exists(destfolder):
                        os.makedirs(destfolder)
                    destfile = destfolder + "\\" + filename
                    write_mapping_log("Author.txt",id+"\t"+name+"\t"+filename+"\n")
                    shutil.copy2(srcfile,destfile)
                    print i,srcfile,destfile
                    i = i + 1



def main():
    srcfile = u"authors.txt"
    srcdir = u"J:\\党政慕课资源\\讲师图片"
    get_author_resource(srcfile,srcdir)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()