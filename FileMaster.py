import os
from shutil import copyfile, rmtree
import sys
from time import time


def get_file_list(path):
    return os.listdir(path)


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
    global file_num, folder_num
    file_num = 0
    folder_num = 0
    if choice == "1":
        file_num = sort_by_extension(path)
    elif choice == "2":
        file_num = sort_by_filename(path)
    elif choice == "3":
        folder_num = dismiss_folder(path)


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


def init():
    global cur_file_name, choice, cur_path
    print("***********************FileMaster***********************")
    print("********https://github.com/BenjiaH/FileMaster***********\n")
    cur_file_name = os.path.basename(sys.argv[0])
    cur_path = os.getcwd()
    os.chdir(cur_path)
    flag_0 = 1
    while flag_0:
        print("当前工作目录：{}".format(cur_path))
        choice = input("请选择:\n1.按文件类型整理\t2.按文件名整理\t\t3.文件夹解散\n4.更改当前目录\n")
        if choice == "4":
            cur_path = input("输入新的工作目录：\n")
            flag_1 = 1
            while flag_1:
                try:
                    os.chdir(cur_path)
                    flag_1 = 0
                    print()
                except Exception as e:
                    print(e)
                    cur_path = input("输入有误请重新输入：\n")
                continue
        else:
            flag_0 = 0


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


def exit_program(cost_time):
    if file_num == 0 | folder_num == 0:
        input("\n\n未发现可以整理的文件。\n按任意键退出。")
    else:
        input("\n\n所有文件（夹）整理（解散）完成，耗时{:.2f}秒。请刷新。\n按任意键退出。".format(cost_time))


def main():
    global end, start
    init()
    start = time()
    file_process(cur_path)
    end = time()
    exit_program(end - start)


if __name__ == '__main__':
    main()
