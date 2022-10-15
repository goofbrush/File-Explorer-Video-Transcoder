# credit to https://www.youtube.com/watch?v=jS2LuG1p8Vw
# make sure you run this in admin !

import os, sys
import winreg as reg


try:
    python_exe = sys.executable # Get path of python.exe
    cwd = os.getcwd()   # Get path of current working directory and python.exe
    
    # all file types wanted to have the option for
    fileType = ["mp4", "mkv", "mov", "wmv", "avi"]

    for type in fileType:

        # path of the context menu (right-click menu) for the file type
        key_path = r'SystemFileAssociations\\.' + type + r'\\shell\\compress\\'

        # Create outer key
        key = reg.CreateKey(reg.HKEY_CLASSES_ROOT, key_path)
        reg.SetValue(key, '', reg.REG_SZ, 'Compress Video')

        # create inner key
        key1 = reg.CreateKey(key, r"command")
        reg.SetValue(key1, '', reg.REG_SZ, python_exe + f' "{cwd}\\handbrake_script.py" %1')

except Exception as e:
    print(e , "\n Run this in Admin")
