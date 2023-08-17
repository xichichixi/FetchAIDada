import os


# 将A文件夹下所有文件夹，复制一份空且同名的文件夹到B文件夹下
def copy_empty_dirs(src_dir, target_dir):
    for root, dirs, files in os.walk(src_dir):
        for dir in dirs:
            new_root = target_dir + root[49:]+"\\"
            new_dir = new_root + dir
            os.makedirs(new_dir)

if __name__ == "__main__":
    src = r"D:\BaiduNetdiskDownload\2017年a股所有行业上市公司年报pdf"
    target = r"D:\Desktop\word文档\2017年a股所有行业上市公司年报pdf"
    copy_empty_dirs(src, target)
