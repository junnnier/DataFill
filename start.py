import PIL
import os
import sys
from generate_excel import execution


def check_file(path):
    if not os.path.exists(path):
        print("The file is not exist: {}".format(path))
    elif not os.path.isfile(path):
        print("Path is not file: {}".format(path))
    else:
        return path
    input("Press Enter to exit...")
    sys.exit()


def main():
    print("----------Excel Data Fill Automatically----------")
    data_path = check_file(input("数据路径："))
    template_path = check_file(input("模板路径："))
    config_path = check_file(input("配置路径："))
    save_dir = input("结果保存目录：")
    execution(data_path, template_path, config_path, save_dir)
    print("end")


if __name__ == '__main__':
    main()
    input("Press Enter to exit...")