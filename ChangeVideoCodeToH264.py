# coding=utf-8
"""
@Project: Dang
@Time: 2019/4/16 11:20
@Name: ChangeVideoCodeToH264
@Author: lieyan123091
@Desc: 对mpeg编码视频用H264编码
"""
import sys,os
import time

DEST_VIDEO_DIR = u"G:\\转码测试\\目标文件"


def change_mpeg_h264(srcdir,destdir):
    """
    对mpeg编码视频用H264编码
    :param srcdir: 处理前视频存储路径
    :param destdir: 处理后视频存储路径
    :return:
    """
    if not os.path.exists(destdir):
        os.makedirs(destdir)
    for name in os.listdir(srcdir):
        if os.path.isdir(srcdir+"\\"+name):
            change_mpeg_h264(srcdir+"\\"+name,destdir+"\\"+name)
        else:
            # 转码操作：
            start_time = time.time()
            srcfile = (srcdir + "\\" + name).decode("utf-8").encode("gbk")
            destfile = (destdir + "\\" + name[:name.rfind(".")]+".mp4").decode("utf-8").encode("gbk")
            os.system('ffmpeg.exe -i \"' + srcfile + '\" -c:v libx264 -strict 2 -b:v 1000k \"' + destfile + '\"')
            end_time = time.time()
            spend_time = int(end_time - start_time)
            print srcdir+name+" 转码完成,耗时：",str(spend_time)+"秒"






def main():
    srcdir = u"G:\\转码测试\\源文件"
    change_mpeg_h264(srcdir,DEST_VIDEO_DIR)



if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf8')
    main()