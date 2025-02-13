import os


def find_excel_files(directory):
    """
    查找指定目录及其子目录下的所有Excel文件(.xlsx 和 .xls)。
    :param directory: 要搜索的根目录
    :return: 包含所有找到的Excel文件路径的列表
    """
    excel_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".xlsx") or file.endswith(".xls"):
                excel_files.append(os.path.join(root, file))
    return excel_files



# 使用示例
if __name__ == "__main__":

    dir_path = "2024年"  # 替换为你的目标文件夹路径
    excel_paths = find_excel_files(dir_path)
    print(excel_paths)