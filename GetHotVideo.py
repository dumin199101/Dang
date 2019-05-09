# coding=utf-8
"""
@Project: Dang
@Time: 2019/4/3 11:31
@Name: GetHotVideo
@Author: lieyan123091
"""
# coding=utf-8

import re, sys
import shutil
import os

def write_mapping_log(logname, content):
    """
    生成mapping日志文件
    :return:
    """
    with open(logname, "a+") as f1:
        f1.write(content.encode("utf-8"))

regex_tab = re.compile("\\t")

i = 1


# 第一步：
def update_course_category_sql(srcfile):
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            course_name = info[0]
            course_code = info[1]
            old_cat_code = info[2]
            new_cat_code = info[3].strip("\n")
            sql = "UPDATE `tb_course_1` SET category_code = \"%s\" WHERE course_code = \"%s\";"
            sql = sql % (new_cat_code,course_code)
            print i,sql
            write_mapping_log(u"update_course_category.sql",sql+"\n")
            i = i + 1



def get_update_courseware_id(srcfile):
    sql = "SELECT * FROM `tb_courseware` WHERE course_code IN ("
    ids = " "
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            course_name = info[0]
            course_code = info[1]
            old_cat_code = info[2]
            new_cat_code = info[3].strip("\n")
            ids = ids + (("\"%s\"" + ",") % (course_code))
    sql = sql + ids + " );"
    print sql


def update_courseware_sql(srcfile1,srcfile2):
    global i
    with open(srcfile1, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            id = info[0]
            course_code1 = info[2]
            video_path = info[3].strip("\n")
            with open(srcfile2, "r") as f2:
                while True:
                    line2 = f2.readline()
                    if len(line2) < 1:
                        break
                    info2 = regex_tab.split(line2)
                    course_code2 = info2[1]
                    old_cat_code = info2[2]
                    new_cat_code = info2[3].strip("\n")
                    if course_code1 == course_code2:
                        # print i,id,video_path,video_path.replace(old_cat_code,new_cat_code)
                        # sql = "UPDATE `tb_courseware_1` SET video_path = \"%s\" WHERE id = %d;"
                        # sql = sql % (video_path.replace(old_cat_code,new_cat_code), int(id))
                        # print sql
                        # write_mapping_log("update_courseware.sql",sql+"\n")
                        print i, video_path, video_path.replace(old_cat_code, new_cat_code)
                        write_mapping_log("courseware_resource.txt",video_path+"\t"+video_path.replace(old_cat_code, new_cat_code)+"\n")
                        i = i + 1



def move_files(srcfile):
    SOURCE_FOLDER = u"K:\\党政党建视频课程\\成品数据\\Videos"
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            name1 = info[0]
            name2 = info[1].strip("\n")
            sourcefile = SOURCE_FOLDER + "\\" + name1
            destfile = SOURCE_FOLDER + "\\" + name2
            if os.path.exists(sourcefile):
                destfolder = os.path.dirname(destfile)[os.path.dirname(destfile).rfind("\\")+1:]
                if not os.path.exists(SOURCE_FOLDER+"\\"+destfolder):
                    os.makedirs(SOURCE_FOLDER+"\\"+destfolder)
                write_mapping_log("move.txt",sourcefile+"\t"+destfile+"\n")
                shutil.move(sourcefile,destfile)
                print i,sourcefile,destfolder
                i = i + 1


def rename_files(srcfile):
    SOURCE_FOLDER = u"G:\\最新"
    DEST_FOLDER = u"G:\\New-Video"
    global i
    with open(srcfile, "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            name1 = info[0]
            name2 = info[2].strip("\n")
            sourcefile = SOURCE_FOLDER + "\\" + name1 + ".mp4"
            destfile = DEST_FOLDER + "\\" + name2
            if os.path.exists(sourcefile):
                destfolder = os.path.dirname(destfile)[os.path.dirname(destfile).rfind("\\") + 1:]
                if not os.path.exists(DEST_FOLDER + "\\" + destfolder):
                    os.makedirs(DEST_FOLDER + "\\" + destfolder)
                write_mapping_log("rename.txt", sourcefile + "\t" + destfile + "\n")
                shutil.copy2(sourcefile, destfile)
                print i, sourcefile, destfolder,destfile
                i = i + 1



def main():
    # 第一步
    # srcfile = u"分类更改.txt"
    # update_course_category_sql(srcfile)
    # get_update_courseware_id(srcfile)
    # srcfile1 = u"tmp_courseware.txt"
    # srcfile2 = u"分类更改.txt"
    # update_courseware_sql(srcfile1,srcfile2)
    # srcfile = u"courseware_resource.txt"
    # move_files(srcfile)

    srcfile = u"courseware_rename.txt"
    rename_files(srcfile)






if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()