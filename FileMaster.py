# -*- coding: utf-8 -*-

import os
from shutil import copyfile, rmtree
import sys
from time import time
import hashlib


def main():
    global end, start
    print("********************************************************")
    print("***********************FileMaster***********************")
    print("********https://github.com/BenjiaH/FileMaster***********")
    print("********************************************************\n")
    while True:
        init()
        start = time()
        process_sum_num = file_process(cur_working_path)
        end = time()
        exit_program(choice, process_sum_num, end - start)


def init():
    global cur_file_name, choice, cur_working_path, cur_prog_path
    cur_file_name = os.path.basename(sys.argv[0])
    cur_prog_path = sys.path[0]
    cur_working_path = os.getcwd()
    os.chdir(cur_working_path)
    while True:
        cur_working_path = os.getcwd()
        print("当前工作目录：{}".format(cur_working_path))
        if cur_working_path.endswith(("Desktop", "desktop")):
            print("当前目录为桌面，如果需要切换目录请选择0")
        print("请选择:")
        print("{name0:<{len0}}\t{name1:<{len1}}\t{name2:<{len2}}".format(name0="1.按文件类型整理", len0=get_print_len(
            "1.按文件类型整理"), name1="2.按文件名整理", len1=get_print_len("2.按文件名整理"), name2="3.文件夹解散", len2=get_print_len("3.文件夹解散")))
        print("{name0:<{len0}}\t{name1:<{len1}}".format(name0="4.Word=>PDF", len0=get_print_len(
            "4.Word=>PDF"), name1="5.PPT=>PDF", len1=get_print_len("5.PPT=>PDF")))
        print("{name0:<{len0}}\t{name1:<{len1}}".format(name0="6.计算MD5", len0=get_print_len(
            "6.计算MD5"), name1="7.计算SHA1", len1=get_print_len("7.计算SHA1")))
        print("0.更改当前目录\n")
        while True:
            choice = input()
            if choice in ("1", "2", "3", "4", "5", "6", "7"):
                return
            elif choice not in ("1", "2", "3", "4", "5", "6", "7", "0"):
                print("输入有误。请重新输入：")
            elif choice == "0":
                new_working_path = input("输入新的工作目录：\n")
                while True:
                    try:
                        os.chdir(new_working_path)
                        break
                    except Exception as e:
                        print(e)
                        new_working_path = input("输入有误。请重新输入：\n")
                    continue
                break


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
        process_num = word2pdf(path)
    elif choice == "5":
        process_num = ppt2pdf(path)
    elif choice == "6":
        process_num = get_MD5(path)
    elif choice == "7":
        process_num = get_SHA1(path)
    return process_num


def get_folder_list(path):
    file_folder_list = (os.listdir(path))
    if cur_file_name in file_folder_list:
        file_folder_list.remove(cur_file_name)
    folder_list = []
    for i in file_folder_list:
        if os.path.isdir(path + "\\" + i):
            folder_list.append(i)
    return folder_list


def get_file_list(path):
    file_folder_list = (os.listdir(path))
    if cur_file_name in file_folder_list:
        file_folder_list.remove(cur_file_name)
    file_list = []
    for i in file_folder_list:
        if os.path.isfile(path + "\\" + i):
            file_list.append(i)
    return file_list


def get_file_folder_list(path):
    file_folder_list = (os.listdir(path))
    if cur_file_name in file_folder_list:
        file_folder_list.remove(cur_file_name)
    return file_folder_list


def sort_by_extension(path):
    count = 0
    file_list = get_file_list(path)
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
    file_list = get_file_list(path)
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
    folder_list = get_folder_list(path)
    for i in folder_list:
        file_folder_list = get_file_folder_list(path + "\\" + i)
        for j in file_folder_list:
            copy_file(i + "\\" + j, j)
            count += 1
        del_old_floder(i)
        print("已解散", i)
    return count


def word2pdf(path):
    global start
    count = 0
    file_list = get_file_list(path)
    wd_list = [f for f in file_list if f.endswith((".doc", ".docx"))]
    if len(wd_list) <= 0:
        return count
    elif len(wd_list) > 0:
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
            new_pdf = word.Documents.Open(cur_working_path + "\\" + i)
            new_pdf.SaveAs(cur_working_path + "\\" + os.path.splitext(i)
                           [0] + ".pdf", FileFormat=17)
            new_pdf.Close()
            count += 1
            print("已转换{}为PDF文件".format(i))
        os.system('TASKKILL /F /IM "WINWORD.EXE"')
        return count


def ppt2pdf(path):
    global start
    count = 0
    file_list = get_file_list(path)
    ppt_list = [f for f in file_list if f.endswith((".ppt", ".pptx"))]
    if len(ppt_list) <= 0:
        return count
    elif len(ppt_list) > 0:
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
            new_pdf = ppt.Presentations.Open(cur_working_path + "\\" + i)
            new_pdf.SaveAs(cur_working_path + "\\" + os.path.splitext(i)
                           [0] + ".pdf", FileFormat=32)
            new_pdf.Close()
            count += 1
            print("已转换{}为PDF文件".format(i))
        os.system('TASKKILL /F /IM "POWERPNT.EXE"')
        return count


def get_SHA1(path):
    count = 0
    file_list = get_file_list(path)
    sha1 = hashlib.sha1()
    print("计算中......")
    for i in file_list:
        with open(cur_working_path + "\\" + i, 'rb') as f:
            sha1.update(f.read())
        print("{:10} {}".format("Filename:", i))
        print("{:10} {}\n".format("SHA1:", sha1.hexdigest()))
        count += 1
    return count


def get_MD5(path):
    count = 0
    file_list = get_file_list(path)
    md5 = hashlib.md5()
    print("计算中......")
    for i in file_list:
        with open(cur_working_path + "\\" + i, 'rb') as f:
            md5.update(f.read())
        print('{:10} {}'.format("Filename:", i))
        print('{:10} {}\n'.format("MD5:", md5.hexdigest()))
        count += 1
    return count


def exit_program(choice, process_num, cost_time):
    print("\n\n")
    if process_num == 0:
        input("未发现可以处理的文件。\n按回车键继续。")
        return
    elif process_num > 0:
        if choice == "1" or choice == "2":
            print("所有文件整理完成，", end="")
        elif choice == "3":
            print("所有文件夹解散完成，", end="")
        elif choice == "4" or choice == "5":
            print("所有文件转换为PDF完成，", end="")
    input("耗时{:.2f}秒。请刷新。\n按回车键继续。若要结束程序请关闭此窗口。\n".format(cost_time))


def copy_file(source, target):
    try:
        copyfile(source, target)
    except IOError as e:
        print("Unable to copy file. %s" % e)
        exit(1)


def del_old_file(file):
    os.remove(file)


# NOT SAFE
# TODO move files to the recycle bin
def del_old_floder(folder):
    rmtree(folder)


def print_list(list):
    for var in list:
        print(var)


def process_detection(process_name):
    process = len(os.popen("tasklist | findstr " + process_name).readlines())
    return process


def get_print_len(string):
    return 16-len(string.encode('GBK'))+len(string)


if __name__ == '__main__':
    main()
