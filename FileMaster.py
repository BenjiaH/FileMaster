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
    global choice
    choice = input("请选择整理方式:\n1.按文件类型\t2.按文件名\t3.文件夹解散\n")


def sort_by_extension(file_list):
    count = 0
    for i in file_list:
        cur_file_extension = os.path.splitext(i)[-1]
        if i == (os.path.basename(sys.argv[0])).replace("py", "exe"):
            pass
        elif i == os.path.basename(sys.argv[0]):
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
        if i == (os.path.basename(sys.argv[0])).replace("py", "exe"):
            pass
        elif i == os.path.basename(sys.argv[0]):
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
        input("\n\n所有文件（夹）整理（解散）完成，耗时{:.2f}秒。\n按任意键退出。".format(end - start))


def main():
    global end, start
    init()
    start = time.time()
    cur_path = os.getcwd()
    os.chdir(cur_path)
    cur_file_list = get_file_list(os.getcwd())
    file_process(cur_file_list)
    end = time.time()
    exit_program()


if __name__ == '__main__':
    main()
