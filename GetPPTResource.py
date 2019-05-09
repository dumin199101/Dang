# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/8 13:18
@Name: GetPPTResource
@Author: lieyan123091
获取PPT资源
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
destdir = "I:\\Deal-Middle\\PPT"

def get_ppt_resource(srcfile,srcdir):
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            pptname = info[1].strip("\n")
            course_id = pptname[:pptname.rfind("-")]
            for filename in os.listdir(srcdir):
                if filename == pptname:
                    srcfile = srcdir + "\\" + filename
                    destfolder = destdir + "\\" + course_id
                    if not os.path.exists(destfolder):
                        os.makedirs(destfolder)
                    # destfile = destfolder + "\\" + filename
                    destfile = destfolder + "\\" + pptname[pptname.rfind("-")+1:]
                    # write_mapping_log("PPT_Name_1.txt",str(i)+"\t"+filename+"\n")
                    shutil.copy2(srcfile,destfile)
                    print i,srcfile,destfile
                    i = i + 1



def main():
    srcfile = u"PPTFileName.txt"
    srcdir = u"I:\\党政慕课资源\\课程PPT"
    get_ppt_resource(srcfile,srcdir)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()