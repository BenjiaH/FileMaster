import os
import shutil
import sys
import time


def get_file_list(path):
    return os.listdir(path)


def file_process(file_list):
    global file_num, folder_num
    file_num = 0
    folder_num = 0
    if choice == "1":
        file_num = sort_by_extension(file_list)
    elif choice == "2":
        file_num = sort_by_filename(file_list)
    elif choice == "3":
        folder_num = dismiss_folder(file_list)


def copy_file(source, target):
    try:
        shutil.copyfile(source, target)
    except IOError as e:
        print("Unable to copy file. %s" % e)
        exit(1)


def del_old_file(file):
    os.remove(file)


def del_old_floder(folder):
    shutil.rmtree(folder)


def init():
    global choice, cur_path
    print("***************FileMaster***************")
    print("**https://github.com/BenjiaH/FileMaster**\n")
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


def sort_by_extension(file_list):
    count = 0
    for i in file_list:
        cur_file_extension = os.path.splitext(i)[-1]
        if i == (os.path.basename(sys.argv[0])):
            pass
        elif cur_file_extension == "":
            pass
        else:
            cur_file_extension = (cur_file_extension.replace(".", "")).upper()
            if os.path.exists(cur_file_extension + "文件"):
                copy_file(i, cur_file_extension + "文件\\" + i)
                del_old_file(i)
                count += 1
                print("已整理", i)
            else:
                os.mkdir(cur_file_extension + "文件")
                copy_file(i, cur_file_extension + "文件\\" + i)
                del_old_file(i)
                count += 1
                print("已整理", i)
    return count


def sort_by_filename(file_list):
    count = 0
    for i in file_list:
        cur_file_extension = os.path.splitext(i)[-1]
        cur_file_without_extension = os.path.splitext(i)[0]
        if i == (os.path.basename(sys.argv[0])):
            pass
        # 文件夹
        elif cur_file_extension == "":
            pass
        else:
            if os.path.exists(cur_file_without_extension):
                copy_file(i, cur_file_without_extension + "\\" + i)
                del_old_file(i)
                count += 1
                print("已整理", i)
            else:
                os.mkdir(cur_file_without_extension)
                copy_file(i, cur_file_without_extension + "\\" + i)
                del_old_file(i)
                count += 1
                print("已整理", i)
    return count


def dismiss_folder(file_list):
    count = 0
    for i in file_list:
        # for i in range(1):
        cur_file_extension = os.path.splitext(i)[-1]
        # 文件夹
        if cur_file_extension == "":
            dismissed_folder = get_file_list(i)
            for j in dismissed_folder:
                copy_file(i + "\\" + j, j)
                count += 1
            del_old_floder(i)
    return count


def exit_program():
    if file_num == 0 | folder_num == 0:
        input("\n\n未发现可以整理的文件。\n按任意键退出。")
    else:
        input("\n\n所有文件（夹）整理（解散）完成，耗时{:.2f}秒。请刷新。\n按任意键退出。".format(end - start))


def main():
    global end, start
    init()
    start = time.time()
    cur_file_list = get_file_list(cur_path)
    file_process(cur_file_list)
    end = time.time()
    exit_program()


if __name__ == '__main__':
    main()
