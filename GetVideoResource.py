# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/7 15:46
@Name: GetVideoResource
@Author: lieyan123091
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
destdir = "I:\\Deal-Middle\\Videos"

def get_video_resource(srcfile,srcdir):
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            id = info[0]
            name = info[1]
            videopath = info[3]
            catname = videopath[:videopath.rfind("/")]
            videoname = videopath[videopath.rfind("/")+1:]
            for filename in os.listdir(srcdir):
                if filename == videoname:
                    srcfile = srcdir + "\\" + filename
                    destfolder = destdir + "\\" +catname
                    if not os.path.exists(destfolder):
                        os.makedirs(destfolder)
                    destfile = destfolder + "\\" + filename
                    shutil.copy2(srcfile,destfile)
                    print i,srcfile,destfile
                    i = i + 1

















def main():
    srcfile = u"courseware.txt"
    srcdir = u"I:\\党政慕课资源\\课程视频"
    get_video_resource(srcfile,srcdir)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
