# coding=utf-8
"""
@Project: Dang
@Time: 2019/4/17 14:06
@Name: ConvertDocxToHTML
@Author: lieyan123091
@Desc: 将Docx格式文档转化为HTML格式
"""

import sys,os
from pydocx import PyDocX
from win32com import client




def convert_docx_html(srcfile):
    """
    转换docx文件为html文件
    :param srcfile: docx文件
    :return:
    """
    html = PyDocX.to_html(srcfile)
    name = srcfile[:srcfile.rfind(".")]
    f = open(name.encode("gbk")+".html", 'w')
    f.write(html.encode("utf-8"))
    f.close()



def convert_doc_docx(srcfolder,srcfile):
    """
    转换doc文件为docx文件
    :param srcfolder: 源文件夹
    :param srcfile: doc文件
    :return:
    """
    word = client.Dispatch("Word.Application")
    wordFullName = os.path.join(srcfolder, srcfile)
    doc = word.Documents.Open(wordFullName)
    name = srcfile[:srcfile.rfind(".")]
    docxFullName = os.path.join(srcfolder, name)
    doc.SaveAs(docxFullName.encode("gbk")+".docx", 16)
    doc.Close()
    word.Quit()


def create_blank_docx(srcdir,destdir):
    if not os.path.exists(destdir):
        os.makedirs(destdir)
    i = 1
    for filename in os.listdir(srcdir):
        srcfile = srcdir + "\\" + filename
        if os.path.isdir(srcfile):
            print i,srcfile
            i = i + 1
            Application = client.Dispatch("Word.Application")
            Application.Visible = True
            new_word = Application.Documents.Add()
            destfile = os.path.join(destdir, filename.strip())
            new_word.SaveAs(destfile.encode("gbk") + ".docx")
            new_word.Close()






def main():
    # srcfile = u"2012中央经济工作会议解读（上）.docx"
    # srcfile = u"最新党的基层组织建设学习辅导讲座（十九大版）（上）.docx"
    # convert_docx_html(srcfile)

    # srcfolder = u"E:\\Goosuu\\Dang\\Script"
    # srcfile = u"最新党的基层组织建设学习辅导讲座（十九大版）（上）.doc"
    # convert_doc_docx(srcfolder,srcfile)

    srcdir = u"E:\\Goosuu\\Dang\\Script\\HTML\\2\\DraftDocx2"
    destdir = u"E:\\Goosuu\\Dang\\Script\\HTML\\2\\Combine_Docx2"
    create_blank_docx(srcdir,destdir)




if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()