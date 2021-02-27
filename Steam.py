# -*- coding: utf-8 -*-
# author:And370
# time:2021/2/27


import os
import json
import tqdm
import win32com.client as client
from itertools import chain
from vals import *


class WallPaper(object):
    def __init__(self):
        self.orgin_path = input("请输入 Wallpaper Engine 内容文件夹路径并回车:")
        self.target_path = input("请输入快捷方式格式化目标文件夹路径并回车:")
        for path in [self.orgin_path, self.target_path]:
            if not os.path.exists(path):
                raise FileExistsError(path)

        self.allowed_types = input(
            """
输入文件类型:
a:视频
b:音频
c:图片
如:ab
含义:提取视频和图片
            """
        )

        self.allowed_extensions = self.file_extensions()
        self.shell = client.Dispatch("WScript.Shell")
        self.count = 0
        self.faild_list = []
        self.run()
        print("成功添加了%s个快捷方式，失败了%s个。" % (self.count, len(self.faild_list)))

    def file_extensions(self):
        """解析允许的文件类型"""
        maps = {"a": video_extensions,
                "b": audio_extensions,
                "c": picture_extensions}

        return set(chain.from_iterable(maps[char] for char in self.allowed_types))

    def get_files(self, dir_path):
        allowed_files = []
        for root, dirs, files in os.walk(dir_path):
            allowed_files += [file for file in files
                              if os.path.splitext(file)[1].replace(".", "") in self.allowed_extensions]
        return allowed_files

    def create_shortcut(self, dir_path, filename, title):
        """filename should be abspath, or there will be some strange errors"""
        clean_filename = self.file_name_clean(filename)
        title = self.file_name_clean(title)

        # 原文件路径
        origin_file = os.path.join(self.orgin_path, dir_path, filename)

        # 快捷方式路径
        lnkname = os.path.join(self.target_path,
                               "【%s】" % title.replace("【", "").replace("】", "") + clean_filename + ".lnk")
        try:
            shortcut = self.shell.CreateShortCut(lnkname)
        except Exception as e:
            print(e)
            self.shell = client.Dispatch("WScript.Shell")
            lnkname = os.path.join(self.target_path,
                                   "【%s】" % title.replace("【", "").replace("】", "") + ".lnk")
            shortcut = self.shell.CreateShortCut(lnkname)

        shortcut.TargetPath = origin_file
        shortcut.WorkingDirectory = os.path.join(self.orgin_path, dir_path)
        shortcut.save()

    def file_name_clean(self, file):
        for char in forbidden_chars:
            file = file.replace(char, "")
        return file

    def run(self):

        for dir_path in tqdm.tqdm(os.listdir(self.orgin_path)):
            # 取title值
            title = ""
            if "project.json" in os.listdir(os.path.join(self.orgin_path, dir_path)):
                with open(os.path.join(self.orgin_path, dir_path, "project.json"),
                          "r", encoding="utf-8") as project_json:
                    title = json.loads(project_json.read())["title"]

            files = self.get_files(os.path.join(self.orgin_path, dir_path))

            # 创建快捷方式
            for file in files:
                self.count += 1
                try:
                    self.create_shortcut(dir_path, file, title)
                except Exception as e:
                    print(e)
                    self.faild_list.append(file)


if __name__ == '__main__':
    wallpaper = WallPaper()
