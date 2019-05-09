# coding=utf-8
"""
@Project: Dang
@Time: 2019/4/17 15:40
@Name: ConvertPPTtoJPG
@Author: lieyan123091
@Desc: 转换PPT为图片
"""
import sys
import comtypes.client
import os

def init_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint

def ppt_to_pdf_jpg(powerpoint, inputFileName, outputFileName, formatType = 32):
    # if outputFileName[-3:] != 'pdf':
    #     outputFileName = outputFileName[0:-4] + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    # deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.SaveAs(inputFileName.rsplit('.')[0] + '.jpg', 17)
    deck.Close()

def convert_files_in_folder(powerpoint, srcfolder):
    i = 1
    files = os.listdir(srcfolder)
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    for pptfile in pptfiles:
        fullpath = os.path.join(srcfolder, pptfile)
        ppt_to_pdf_jpg(powerpoint, fullpath, fullpath)
        print i,pptfile
        i = i + 1

if __name__ == "__main__":
    reload(sys)
    sys.setdefaultencoding('utf8')
    powerpoint = init_powerpoint()
    # srcfolder = u"E:\\Goosuu\\Dang\\Script\\PPT\\Combine"
    # srcfolder = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2012-2015\\PPT\\Combine_PPT_2-2"
    srcfolder = u"I:\\党建视频库数据加工整理\\Deal-Middle\\2016-2018\\PPT\\Combine_PPT_2-2"
    convert_files_in_folder(powerpoint, srcfolder)
    powerpoint.Quit()
