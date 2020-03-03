# -*- coding: utf-8 -*-

import os
from shutil import copyfile, rmtree
import sys
from time import time
while True:
    try:
        from comtypes.client import CreateObject
        break
    except Exception as e:
        print(e)
        os.system("pip install comtypes")
        input("请重新运行程序。按回车退出。")
        exit(1)


def process_detection(process_name):
    process = len(os.popen('tasklist | findstr ' + process_name).readlines())
    return process


def select_file_folder(select_type, selcet_path):
    file_folder_list = (os.listdir(selcet_path))
    if cur_file_name in file_folder_list:
        file_folder_list.remove(cur_file_name)
    folder_list = []
    file_list = []
    for i in file_folder_list:
        if "." in i:
            file_list.append(i)
        else:
            folder_list.append(i)
    if select_type == "file":
        return file_list
    else:
        return folder_list


def file_process(path):
    global process_num
    process_num = 0
    if choice == "1":
        process_num = sort_by_extension(path)
    elif choice == "2":
        process_num = sort_by_filename(path)
    elif choice == "3":
        process_num = dismiss_folder(path)
    elif choice == "4":
        process_num = word_ppt2pdf(choice, path)
    elif choice == "5":
        process_num = word_ppt2pdf(choice, path)
    return process_num


def copy_file(source, target):
    try:
        copyfile(source, target)
    except IOError as e:
        print("Unable to copy file. %s" % e)
        exit(1)


def del_old_file(file):
    os.remove(file)


def del_old_floder(folder):
    rmtree(folder)


def print_list(list):
    for var in list:
        print(var)


def init():
    global cur_file_name, choice, cur_path
    cur_file_name = os.path.basename(sys.argv[0])
    while True:
        cur_path = os.getcwd()
        print("当前工作目录：{}".format(cur_path))
        if cur_path.endswith(("Desktop", "desktop")):
            print("当前目录为桌面，如果需要切换目录请选择0")
        print("请选择:")
        print("1.按文件类型整理\t2.按文件名整理\t\t3.文件夹解散")
        print("4.Word=>PDF\t\t\t5.PPT=>PDF")
        print("0.更改当前目录\n")
        while True:
            choice = input()
            if choice == "0":
                new_path = input("输入新的工作目录：\n")
                while True:
                    try:
                        os.chdir(new_path)
                        break
                    except Exception as e:
                        print(e)
                        new_path = input("输入有误。请重新输入：\n")
                    continue
                break
            elif choice in ("1", "2", "3", "4", "5"):
                return
            else:
                print("输入有误。请重新输入：")


def sort_by_extension(path):
    count = 0
    file_list = select_file_folder("file", path)
    for i in file_list:
        cur_file_extension = os.path.splitext(i)[-1]
        cur_file_extension = (cur_file_extension.replace(".", "")).upper()
        if not os.path.exists(cur_file_extension + "文件"):
            os.mkdir(cur_file_extension + "文件")
        copy_file(i, cur_file_extension + "文件\\" + i)
        del_old_file(i)
        count += 1
        print("已整理", i)
    return count


def sort_by_filename(path):
    count = 0
    file_list = select_file_folder("file", path)
    for i in file_list:
        cur_file_without_extension = os.path.splitext(i)[0]
        if not os.path.exists(cur_file_without_extension):
            os.mkdir(cur_file_without_extension)
        copy_file(i, cur_file_without_extension + "\\" + i)
        del_old_file(i)
        count += 1
        print("已整理", i)
    return count


def dismiss_folder(path):
    count = 0
    folder_list = select_file_folder("folder", path)
    for i in folder_list:
        file_list = select_file_folder("file", path + "\\" + i)
        for j in file_list:
            copy_file(i + "\\" + j, j)
            count += 1
        del_old_floder(i)
        print("已解散", i)
    return count


def word_ppt2pdf(covert_choice, path):
    global start
    count = 0
    file_list = select_file_folder("file", path)
    if covert_choice == "4":
        wd_list = [f for f in file_list if f.endswith((".doc", ".docx"))]
        if len(wd_list) > 0:
            while process_detection("WINWORD.EXE"):
                input("检测到Word已经打开，请保存当前文件并关闭程序。按回车键继续。")
                os.system('TASKKILL /F /IM "WINWORD.EXE"')
            print("按回车键开始转换")
            print_list(wd_list)
            input()
            start = time()
            word = CreateObject("Word.Application")
            word.Visible = 0
            for i in wd_list:
                print("正在转换{}为PDF文件".format(i))
                new_pdf = word.Documents.Open(cur_path + "\\" + i)
                new_pdf.SaveAs(cur_path + "\\" + os.path.splitext(i)[0] + ".pdf", FileFormat=17)
                new_pdf.Close()
                count += 1
                print("已转换{}为PDF文件".format(i))
            os.system('TASKKILL /F /IM "WINWORD.EXE"')
            return count
        else:
            return count
    elif covert_choice == "5":
        ppt_list = [f for f in file_list if f.endswith((".ppt", ".pptx"))]
        if len(ppt_list) > 0:
            while process_detection("POWERPNT.EXE"):
                input("检测到PowerPoint已经打开，请保存当前文件并关闭程序。按回车键继续。")
                os.system('TASKKILL /F /IM "POWERPNT.EXE"')
            print("可能耗时较长，请耐心等待。建议先关闭PowerPoint的所有加载项。\n若出现卡死，请手动结束PowerPoint。")
            print("按回车键开始转换")
            print_list(ppt_list)
            input()
            start = time()
            ppt = CreateObject("Powerpoint.Application")
            ppt.Visible = 1
            for i in ppt_list:
                print("正在转换{}为PDF文件".format(i))
                new_pdf = ppt.Presentations.Open(cur_path + "\\" + i)
                new_pdf.SaveAs(cur_path + "\\" + os.path.splitext(i)[0] + ".pdf", FileFormat=32)
                new_pdf.Close()
                count += 1
                print("已转换{}为PDF文件".format(i))
            os.system('TASKKILL /F /IM "POWERPNT.EXE"')
            return count
        else:
            return count


def exit_program(choice, process_num, cost_time):
    print("\n\n")
    if process_num == 0:
        input("未发现可以处理的文件。\n按回车键继续。")
        return
    else:
        if choice == "1" or choice == "2":
            print("所有文件整理完成，", end="")
        elif choice == "3":
            print("所有文件夹解散完成，", end="")
        elif choice == "4" or choice == "5":
            print("所有文件转换为PDF完成，", end="")
    input("耗时{:.2f}秒。请刷新。\n按回车键继续。若要结束程序请关闭此窗口。\n".format(cost_time))


def main():
    global end, start
    print("********************************************************")
    print("***********************FileMaster***********************")
    print("********https://github.com/BenjiaH/FileMaster***********")
    print("********************************************************\n")
    while True:
        init()
        start = time()
        process_sum_num = file_process(cur_path)
        end = time()
        exit_program(choice, process_sum_num, end - start)


if __name__ == '__main__':
    main()
