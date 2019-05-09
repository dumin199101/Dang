# coding=utf-8
"""
@Project: Dang
@Time: 2019/3/28 9:51
@Name: GetVideoNameFromSource
@Author: lieyan123091
@Desc: 从源文件获取视频名称
"""
import re, sys, os
import collections

regex_tab = re.compile("\\t")


def write_mapping_log(logname, content):
    """
    生成mapping日志文件
    :return:
    """
    with open(logname, "a+") as f1:
        f1.write(content.encode("utf-8"))


i = 1


def get_video_name(srcdir):
    """
    从源文件夹获取所有视频名称
    :param srcdir:
    :return:
    """
    global i
    for filename in os.listdir(srcdir):
        srcfile = srcdir + "\\" + filename
        if os.path.isdir(srcfile):
            get_video_name(srcfile)
        else:
            # write_mapping_log("2016-2018-VideoName.txt", srcfile + "\t" + filename + "\n")
            write_mapping_log("2012-2014-不清晰版-VideoName.txt".encode("gbk"), srcfile + "\t" + filename + "\n")
            print i, filename, srcfile
            i = i + 1


def get_video_name_set(srcfile):
    """
    根据文件名称查找是否有重复的视频
    :param srcfile:
    :return:
    """
    i = 0
    with open(srcfile, "r") as f1:
        name_set = set()
        while True:
            line = f1.readline()
            if len(line) < 1:
                break
            i = i + 1
            info = regex_tab.split(line)
            name = info[1]
            name_set.add(name)
    print "文件去重后的总数：", len(name_set), " 文件数：", i
    if len(name_set) == i:
        print "结论：校对完成，未发现重复文件"
    else:
        print "结论：发现重复文件，请再次审核"


set_database = set()
set_srcfile = set()


def check_video_from_database(database_file, srcfile):
    """
    查找数据表中跟源文件对应不起来的部分
    :param database_file:
    :param srcfile:
    :return:
    """
    with open(database_file, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            set_database.add(info1[0].strip("\n"))

    with open(srcfile, "r") as f2:
        while True:
            line = f2.readline()
            if len(line) < 1:
                break
            info = regex_tab.split(line)
            info[1] = info[1][:info[1].rfind(".")].strip("\n")
            if info[1].rfind("（下）") > -1:
                info[1] = info[1][:info[1].rfind("（下）")]
            if info[1].rfind("（上）") > -1:
                info[1] = info[1][:info[1].rfind("（上）")]
            if info[1].rfind("（中）") > -1:
                info[1] = info[1][:info[1].rfind("（中）")]
            if info[1].endswith("上"):
                info[1] = info[1][:info[1].rfind("上")]
            if info[1].endswith("下"):
                info[1] = info[1][:info[1].rfind("下")]
            set_srcfile.add(info[1])

    # 查找匹配不上的
    set_diff_srcfile = set_srcfile.difference(set_database)
    set_diff_database = set_database.difference(set_srcfile)

    # 查找能匹配上的
    set_inter = set_srcfile.intersection(set_database)

    i = 1
    for obj in set_diff_srcfile:
        # write_mapping_log("diffVideoName.txt", obj + "\n")
        # print i, obj
        i = i + 1

    j = 1
    for obj in set_inter:
        # print j,obj
        j = j + 1

    m = 1
    for obj in set_diff_database:
        # write_mapping_log("diffVideoName.txt", obj + "\n")
        print m, obj
        m = m + 1


set_good = set()
set_bad = set()


def check_video_from_good_and_bad(good_file, bad_file):
    """
    查找清晰版跟不清晰版之间是否有交集
    :param good_file:
    :param bad_file:
    :return:
    """
    with open(good_file, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            set_good.add(info1[1].strip("\n"))
    with open(bad_file, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            set_bad.add(info2[1].strip("\n"))

    # 查找重合的
    set_diff = set_good.intersection(set_bad)
    if len(set_diff) == 0:
        print "没有重合的"
    else:
        m = 1
        for obj in set_diff:
            print m, obj
            m = m + 1


re_digit = re.compile("\d+")


def get_item_name(srcfolder, type):
    """
    按年份查找对应类型文件列表
    :param srcfolder: 原始文件夹
    :param type: 文件类型
    :return:
    """
    year = re_digit.search(srcfolder).group()
    for item in os.listdir(srcfolder):
        print item[:item.rfind(".")], year, item[item.rfind(".") + 1:], item
        write_mapping_log(year+"-"+type+".txt",item[:item.rfind(".")]+"\t"+year+"\t"+item[item.rfind(".") + 1:]+"\t"+item+"\n")

set_file1 = set()
set_file2 = set()

def get_diff_item(file1,file2):
    """
    获取不同类型之间的差值部分
    :param file1:
    :param file2:
    :return:
    """
    with open(file1, "r") as f1:
        while True:
            line1 = f1.readline()
            if len(line1) < 1:
                break
            info1 = regex_tab.split(line1)
            set_file1.add(info1[0].strip("\n"))
    with open(file2, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            set_file2.add(info2[0].strip("\n"))

    set_diff = set_file1.difference(set_file2)
    m = 1
    for obj in set_diff:
        print m, obj
        m = m + 1


def get_list_count(srcfile):
    result_list = []
    with open(srcfile, "r") as f1:
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            # 此处枚举结果不准确
            if info2[0].endswith("1") or info2[0].endswith("2") or info2[0].endswith("3") or info2[0].endswith("4"):
                if info2[0].endswith("1"):
                    result_list.append(info2[0][:info2[0].rfind("1")])
                elif info2[0].endswith("2"):
                    result_list.append(info2[0][:info2[0].rfind("2")])
                elif info2[0].endswith("3"):
                    result_list.append(info2[0][:info2[0].rfind("3")])
                elif info2[0].endswith("4"):
                    result_list.append(info2[0][:info2[0].rfind("4")])
            else:
                result_list.append(info2[0])
        b = collections.Counter(result_list)
        while True:
            line2 = f1.readline()
            if len(line2) < 1:
                break
            info2 = regex_tab.split(line2)
            src_name = info2[0]
            if src_name.endswith("1"):
                p_name = src_name[:src_name.rfind("1")]
            elif src_name.endswith("2"):
                p_name = src_name[:src_name.rfind("2")]
            elif src_name.endswith("3"):
                p_name = src_name[:src_name.rfind("3")]
            elif src_name.endswith("4"):
                p_name = src_name[:src_name.rfind("4")]


def main():
    # 1.从源文件夹获取所有视频名称
    # srcdir = u"I:\\视频资料"
    # srcdir = u"J:\\清晰版\\视频资料"
    # srcdir = u"J:\\不清晰版\\视频资料"
    # get_video_name(srcdir)

    # 2.查找视频是否有重复的
    # srcfile = "2016-2018-VideoName.txt"
    # get_video_name_set(srcfile)

    # 3.查找现有数据表中数据跟资源重复的：
    database_file = "tb_course.txt"
    srcfile = "2016-2018-VideoName.txt"
    check_video_from_database(database_file, srcfile)

    # 4.查看不清晰版跟清晰版是否有交集
    # good_file = u"2012-2015-清晰版-VideoName.txt"
    # bad_file = u"2012-2014-不清晰版-VideoName.txt"
    # check_video_from_good_and_bad(good_file,bad_file)

    # 5.按年份查找匹配视频文件数、PPT数、封面数、试题数、文稿数
    # srcfolder = u"I:\\文字资料\\2018\\文稿"
    # # srcfolder = u"J:\\不清晰版\\视频资料\\2014"
    # type = '文稿'
    # get_item_name(srcfolder, type)

    # 6.校对不同类型之间差值的部分：
    # file1 = u"2018-Video.txt"
    # file2 = u"2018-图片.txt"
    # get_diff_item(file1,file2)
    # get_diff_item(file2,file1)

    # 7.统计list出现的次数（此方法待完善）
    # srcfile = u"2018-图片.txt"
    # get_list_count(srcfile)






if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()
