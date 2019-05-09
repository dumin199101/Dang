# coding=utf-8
"""
@Project: Dang
@Time: 2019/4/17 16:11
@Name: CombinePPTFiles
@Author: lieyan123091
@Desc: 合并PPT文件
"""

import win32com.client, sys
from glob import glob
import os

DEST_DIR = "E:\\Goosuu\\Dang\\Script\\PPT\\Combine_PPT_2"


def compare(str1, str2):
    """
    上变为1，下变为3，中变为2
    :param str1:
    :param str2:
    :return:
    """
    if str1.rfind("（上）") > -1 and str2.rfind("（下）") > -1:
        return -1
    elif str1.rfind("（上）") > -1 and str2.rfind("（中）") > -1:
        return -1
    elif str1.rfind("（中）") > -1 and str2.rfind("（上）") > -1:
        return 1
    elif str1.rfind("（中）") > -1 and str2.rfind("（下）") > -1:
        return -1
    elif str1.rfind("（下）") > -1 and str2.rfind("（上）") > -1:
        return 1
    elif str1.rfind("（下）") > -1 and str2.rfind("（中）") > -1:
        return 1


newppt = None


# def join_ppt(ppt_folder):
#     global new_ppt
#     name = ppt_folder[ppt_folder.rfind("\\") + 1:]
#     print name
#
#     Application = win32com.client.Dispatch("PowerPoint.Application")
#     Application.Visible = True
#
#
#
#     # Create new presentation
#     new_ppt = Application.Presentations.Add()
#
#     files = glob(ppt_folder + "/*")
#     files = sorted(files, compare)
#     for f in files:
#         # Open and read page numbers
#         print f
#         Application = win32com.client.Dispatch("PowerPoint.Application")
#         Application.Visible = True
#         exit_ppt = Application.Presentations.Open(f)
#         page_num = exit_ppt.Slides.Count
#         exit_ppt.Close()
#         try:
#             start = 1
#             end = page_num
#             if f.rfind("（上）") > -1:
#                 start = 1
#                 end = page_num - 1
#             if f.rfind("（中）") > -1:
#                 start = 2
#                 end = page_num - 1
#             if f.rfind("（下）") > -1:
#                 start = 2
#                 end = page_num
#             num = new_ppt.Slides.InsertFromFile(f, new_ppt.Slides.Count, start, end)
#             print 1, num
#         except:
#             new_ppt = Application.Presentations.Add()
#             start = 1
#             end = page_num
#             if f.rfind("（上）") > -1:
#                 start = 1
#                 end = page_num - 1
#             if f.rfind("（中）") > -1:
#                 start = 2
#                 end = page_num - 1
#             if f.rfind("（下）") > -1:
#                 start = 2
#                 end = page_num
#             num = new_ppt.Slides.InsertFromFile(f, new_ppt.Slides.Count, start, end)
#             print 2, num
#
#     new_ppt.SaveAs(DEST_DIR + "\\" + name.strip().encode("gbk") + ".ppt")
#     Application.Quit()

def produce_ppt(srcfolder):
    """
    生成空的ppt文件
    :param srcfolder:
    :return:
    """
    i = 1
    for filename in os.listdir(srcfolder):
        file = os.path.join(srcfolder, filename)
        print filename
        if os.path.isdir(file):
            # if i <= 172:
            #     i = i + 1
            #     continue
            Application = win32com.client.Dispatch("PowerPoint.Application")
            Application.Visible = True
            new_ppt = Application.Presentations.Add()
            new_ppt.SaveAs(DEST_DIR + "\\" + filename.strip().encode("gbk") + ".ppt")
            print i, filename
            i = i + 1


def join_ppt(ppt_folder):
    global new_ppt
    name = ppt_folder[ppt_folder.rfind("\\") + 1:]
    print name

    Application = win32com.client.Dispatch("PowerPoint.Application")
    Application.Visible = True

    # Create new presentation
    print "将要打开的文件：",DEST_DIR + "\\" + name + ".ppt"
    new_ppt = Application.Presentations.Open(DEST_DIR + "\\" + name.strip() + ".ppt")

    files = glob(ppt_folder + "/*")
    files = sorted(files, compare, reverse=True)
    for f in files:
        # Open and read page numbers
        print f
        Application = win32com.client.Dispatch("PowerPoint.Application")
        Application.Visible = True
        exit_ppt = Application.Presentations.Open(f)
        page_num = exit_ppt.Slides.Count
        print page_num
        # 初始化开始跟结尾
        start = 1
        end = page_num
        if f.rfind("（上）") > -1:
            start = 1
            end = page_num - 1
        if f.rfind("（中）") > -1:
            start = 2
            end = page_num - 1
        if f.rfind("（下）") > -1:
            start = 2
            end = page_num
        # print start,end
        for i in range(start, end+1):
            exit_ppt.Slides(i).Copy()
            if (f.rfind("（上）") > -1) or (f.rfind("（中）") > -1) or (f.rfind("（下）") > -1):
                if f.rfind("（上）") > -1:
                    print i
                    new_ppt.Slides.Paste(i)
                    new_ppt.Slides(i).Design = exit_ppt.Slides(i).Design
                    new_ppt.Slides(i).ColorScheme = exit_ppt.Slides(i).ColorScheme
                    new_ppt.Slides(i).FollowMasterBackground = exit_ppt.Slides(i).FollowMasterBackground
                if f.rfind("（中）") > -1:
                    print i
                    new_ppt.Slides.Paste(i - 1)
                    new_ppt.Slides(i-1).Design = exit_ppt.Slides(i).Design
                    new_ppt.Slides(i-1).ColorScheme = exit_ppt.Slides(i).ColorScheme
                    new_ppt.Slides(i-1).FollowMasterBackground = exit_ppt.Slides(i).FollowMasterBackground
                if f.rfind("（下）") > -1:
                    print i
                    new_ppt.Slides.Paste(i - 1)
                    new_ppt.Slides(i-1).Design = exit_ppt.Slides(i).Design
                    new_ppt.Slides(i-1).ColorScheme = exit_ppt.Slides(i).ColorScheme
                    new_ppt.Slides(i-1).FollowMasterBackground = exit_ppt.Slides(i).FollowMasterBackground
            else:
                # 处理其它情况
                print i
                new_ppt.Slides.Paste(i)
                new_ppt.Slides(i).Design = exit_ppt.Slides(i).Design
                new_ppt.Slides(i).ColorScheme = exit_ppt.Slides(i).ColorScheme
                new_ppt.Slides(i).FollowMasterBackground = exit_ppt.Slides(i).FollowMasterBackground

                # new_ppt.Application.CommandBars.ExecuteMso("PasteSourceFormatting")
        exit_ppt.Close()
    new_ppt.Save()
    Application.Quit()


def combine_ppt_files(srcdir):
    i = 1
    for filename in os.listdir(srcdir):
        file = os.path.join(srcdir, filename)
        if os.path.isdir(file):
            # if i <= 165:
            #     i = i + 1
            #     continue
            join_ppt(file)
            print i, filename
            i = i + 1


def main():
    # srcfolder = u"E:\\Goosuu\\Dang\\Script\\PPT\\PowerPoint2"
    # produce_ppt(srcfolder)

    # 只能处理（上） （中） （下） 这种结构的,还存在单篇的跟（一）（二）（三）（四）这种结构的
    srcdir = u"E:\\Goosuu\\Dang\\Script\\PPT\\PowerPoint2"
    combine_ppt_files(srcdir)


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
    # join_ppt(u"E:\\Goosuu\\Dang\\Script\\PPT\\《对“协调发展”的正确解读》")
    # join_ppt(u"E:\\Goosuu\\Dang\\Script\\PPT\\当前我国宏观经济形势与调控政策趋向")
    # join_ppt(u"E:\\Goosuu\\Dang\\Script\\PPT\\PowerPoint1\\2015年“两会”热点解读——四个全面之全面推进依法治国")
