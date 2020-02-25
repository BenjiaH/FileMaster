import os
import shutil
import time


def get_file_list(path):
    return os.listdir(path)


def file_process(file_list):
    global file_num
    if choice == "1":
        file_num = sort_by_extension(file_list)
    elif choice == "2":
        file_num = sort_by_filename(file_list)


def copy_file(source, target):
    try:
        shutil.copyfile(source, target)
    except IOError as e:
        print("Unable to copy file. %s" % e)
        exit(1)


def del_old_file(file):
    os.remove(file)


def init():
    global choice
    choice = input("请选择整理方式:\n1.按文件类型\t2.按文件名\n")


def sort_by_extension(file_list):
    count = 0
    for i in range(len(file_list)):
        cur_file_extension = os.path.splitext(file_list[i])[-1]
        if file_list[i] == (os.path.basename(__file__)).replace("py", "exe"):
            pass
        elif file_list[i] == os.path.basename(__file__):
            pass
        elif cur_file_extension == "":
            pass
        else:
            cur_file_extension = cur_file_extension.replace(".", "")
            if os.path.exists(cur_file_extension + "文件"):
                copy_file(file_list[i], cur_file_extension + "文件\\" + file_list[i])
                del_old_file(file_list[i])
                count += 1
                print("已整理", file_list[i])
            else:
                os.mkdir(cur_file_extension + "文件")
                copy_file(file_list[i], cur_file_extension + "文件\\" + file_list[i])
                del_old_file(file_list[i])
                count += 1
                print("已整理", file_list[i])
    return count


def sort_by_filename(file_list):
    count = 0
    for i in range(len(file_list)):
        cur_file_extension = os.path.splitext(file_list[i])[-1]
        cur_file_without_extension = os.path.splitext(file_list[i])[0]
        if file_list[i] == (os.path.basename(__file__)).replace("py", "exe"):
            pass
        elif file_list[i] == os.path.basename(__file__):
            pass
        # 文件夹
        elif cur_file_extension == "":
            pass
        else:
            os.mkdir(cur_file_without_extension)
            copy_file(file_list[i], cur_file_without_extension + "\\" + file_list[i])
            del_old_file(file_list[i])
            count += 1
            print("已整理", file_list[i])
    return count


def exit_program():
    if file_num == 0:
        input("\n\n未发现可以整理的文件。\n按任意键退出。")
    else:
        input("\n\n所有文件整理完成，耗时{:.2f}秒。\n按任意键退出。".format(end - start))


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
