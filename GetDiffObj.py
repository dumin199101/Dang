# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/8 11:59
@Name: GetDiffObj
@Author: lieyan123091
查找数据不对应的
"""

import re, sys,os


def write_mapping_log(logname, content):
    """
    生成mapping日志文件
    :return:
    """
    with open(logname, "a+") as f1:
        f1.write(content.encode("utf-8"))

index = 1
set_items_record = set()
set_items_filelist = set()
regex_tab = re.compile("\\t")

#  校对文稿
# def get_diff_items(srcdir):
#     with open(u"Draft.txt", "r") as f1:
#         while True:
#             line = f1.readline()
#             if len(line) < 1:
#                 break
#             info = regex_tab.split(line)
#             set_items_record.add(info[2].strip("\n"))
#
#
#     for filename in os.listdir(srcdir):
#         set_items_filelist.add(filename)
#
#     set_diff = set_items_filelist.difference(set_items_record)
#
#     for obj in set_diff:
#         print obj

# 校对PPT
# def get_diff_items():
#     with open(u"PPTFileName.txt", "r") as f1:
#         while True:
#             line = f1.readline()
#             if len(line) < 1:
#                 break
#             info = regex_tab.split(line)
#             set_items_record.add(info[1].strip("\n"))
#
#     with open(u"PPT_Name.txt", "r") as f1:
#         while True:
#             line = f1.readline()
#             if len(line) < 1:
#                 break
#             info = regex_tab.split(line)
#             set_items_filelist.add(info[1].strip("\n"))
#
#     set_diff = set_items_record.difference(set_items_filelist)
#
#     global index
#     for obj in set_diff:
#         print index,obj
#         index = index + 1


# 校对作者
# def get_diff_items():
#     with open(u"authors.txt", "r") as f1:
#         while True:
#             line = f1.readline()
#             if len(line) < 1:
#                 break
#             info = regex_tab.split(line)
#             set_items_record.add(info[3].strip("\n"))
#
#     with open(u"Author.txt", "r") as f1:
#         while True:
#             line = f1.readline()
#             if len(line) < 1:
#                 break
#             info = regex_tab.split(line)
#             set_items_filelist.add(info[2].strip("\n"))
#
#     set_diff = set_items_record.difference(set_items_filelist)
#
#     global index
#     for obj in set_diff:
#         print index,obj
#         index = index + 1


#  校对封面
def get_diff_items(srcdir):
    with open(u"refer.txt", "r") as f1:
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            set_items_record.add(info[2].strip("\n"))

    i = 1
    # for name in set_items_record:
    #     print i,name
    #     i = i + 1

    for filename in os.listdir(srcdir):
        # print i,filename[:filename.rfind(".jpg")]
        set_items_filelist.add((filename[:filename.rfind(".jpg")]).encode("utf-8"))
        # i = i + 1

    # set_diff = set_items_filelist.difference(set_items_record)
    set_diff = set_items_record.difference(set_items_filelist)

    for obj in set_diff:
        print i,obj
        i = i + 1


def main():
    srcdir = u"2019-05\\图片 - 新(2)\\图片 - 新\\不清晰\\2014"
    # get_diff_items(srcdir)
    get_diff_items(srcdir)



if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()