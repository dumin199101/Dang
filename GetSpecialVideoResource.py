# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/7 14:33
@Name: GetVideoResource
@Author: lieyan123091
"""
# coding=utf-8

import re, sys, os,shutil

regex_tab = re.compile("\\t")

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
            guid = info[1]
            name = info[2]
            videopath = info[3][info[3].rfind("/")+1:]
            order = info[4].strip("\n")
            for filename in os.listdir(srcdir):
                if filename == videopath:
                    srcfile = srcdir + "\\" + filename
                    destfolder = destdir + "\\" + guid
                    if not os.path.exists(destfolder):
                        os.makedirs(destfolder)
                    destfile = destfolder + "\\" + order+"_"+name.encode("gbk")+".mp4"
                    shutil.copy2(srcfile,destfile)
                    print i,srcfile
                    i = i + 1

















def main():
    srcfile = u"tmp_special_video.txt"
    srcdir = u"I:\\党政慕课资源\\课程视频"
    get_video_resource(srcfile,srcdir)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
