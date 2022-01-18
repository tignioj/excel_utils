import os
from pathlib import Path
# https://www.geeksforgeeks.org/how-to-get-file-extension-in-python/
def getFileExtension(fname):
    split_tup = os.path.splitext(fname)
    file_extension = split_tup[1]
    return file_extension

def getFileName(fname):
    return Path(fname).stem

def getParentFolder(absoluteFile):
    """
    获取文件所在目录
    :param absoluteFile:
    :return:
    """
    p = Path(absoluteFile)
    parent_folder = p.parent
    return str(parent_folder)

