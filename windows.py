# -*- coding: utf-8 -*-
# author:And370
# time:2020/9/8
import os
import random
import schedule
import win32api, win32con, win32gui

from time import sleep
from itertools import cycle


class WallPaper(object):
    def __init__(self):
        self.path = input("请输入壁纸文件夹路径并回车:")
        self.pics = [os.path.join(self.path, pic) for pic in os.listdir(self.path)]
        self.run()

    def run(self, mode="ascend"):
        """
        :param mode: "ascend","descend","random"
        """
        if mode == "ascend":
            self.pics.sort()
        elif mode == "descend":
            self.pics.sort(reverse=True)
        else:
            random.shuffle(self.pics)
        first = True
        for pic in cycle(self.pics):
            if first:
                self.change(pic)
                first = False
                continue
            schedule.every().hour.do(self.change, pic)
            sleep(1)

    def change(self, pic):
        key = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,
                                    "Control Panel\\Desktop", 0, win32con.KEY_SET_VALUE)
        win32api.RegSetValueEx(key, "WallpaperStyle", 0, win32con.REG_SZ, "2")
        # 2拉伸适应桌面,0桌面居中
        win32api.RegSetValueEx(key, "TileWallpaper", 0, win32con.REG_SZ, "0")
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, pic, 1 + 2)


if __name__ == "__main__":
    wall = WallPaper()
