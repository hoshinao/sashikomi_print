import win32api
import sys
import os
import time

from auto_print import file_check

path = "C:/Users/obi21703/PycharmProjects/差し込み印刷/print_list"
class Print_auto:
    # 印刷処理
    def auto_print_module(path):
        if __name__ == '__main__':
            win32api.ShellExecute(0, "print", path, None, ".", 0)
        print("Printed:" + path)

    # フォルダ内のファイル読み込み
    def file_check(path):
        if os.path.isdir(path):
            files = os.listdir(path)
            for file in files:
                file_check(path + "\\" + file)
        else:
            path.auto_print_module(path)
            time.sleep(3)


    print_path = r"./print_list"
    file_check(print_path)
    print("Process finished")
