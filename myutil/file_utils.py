import os
# https://www.geeksforgeeks.org/how-to-get-file-extension-in-python/
def getFileExtension(fname):
    split_tup = os.path.splitext(fname)
    file_extension = split_tup[1]
    return file_extension

def getFileName(fname):
    split_tup = os.path.splitext(fname)
    file_name = split_tup[0]
    return file_name
