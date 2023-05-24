# -*- coding: utf-8 -*-
"""
@Time ： 2023/1/18 16:22
@Auth ： 大雄
@File ：model.py
@IDE ：PyCharm
@Email:3475228828@qq.com
"""
import threading

thread_max_num = 1000
class Custom:
    def __init__(self):
        self.display = None
        self.mouse = None
        self.keyboard = None
    def clear(self):
        self.__init__()
def tdi():
    return [Custom() for _ in range(thread_max_num)]
td_info = tdi()
