import os
import shutil

destDir = os.getcwd()

def copy_file(src):
    for i in  range(0,10):
        destPath =destDir + os.path.sep + '4_1_' + str(i) +'.py'

        if os.path.exists(src) and not os.path.exists(destPath):
            shutil.copyfile(src, destPath)
        
if __name__ == '__main__':
    copy_file('D:/Alice/Project/Python/cutfile.py')
