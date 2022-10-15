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

        # video compress function
        key_path = r'SystemFileAssociations\\.' + type + r'\\shell\\compress\\' # path of the right-click menu for the file type
        
        key = reg.CreateKey(reg.HKEY_CLASSES_ROOT, key_path) # Create outer key
        reg.SetValue(key, '', reg.REG_SZ, 'Compress Video')

        key1 = reg.CreateKey(key, r"command") # create inner key
        reg.SetValue(key1, '', reg.REG_SZ, python_exe + f' "{cwd}\\handbrake_script.py" %1')



    # compress config function
        key_path2 = r'SystemFileAssociations\\.' + type + r'\\shell\\compresscfg\\' # path of the right-click menu for the file type

        keyb = reg.CreateKey(reg.HKEY_CLASSES_ROOT, key_path2) # Create outer key
        reg.SetValue(keyb, '', reg.REG_SZ, 'Compress Config')

        keyb1 = reg.CreateKey(keyb, r"command") # create inner key
        reg.SetValue(keyb1, '', reg.REG_SZ, "notepad.exe" + f' "{cwd}\\config.cfg"')

except Exception as e:
    print(e , "\n Run this in Admin")
