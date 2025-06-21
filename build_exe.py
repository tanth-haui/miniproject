import PyInstaller.__main__
import PyInstaller.config
import os

print("path:", str(os.getcwd()))

distpath ="--distpath=" + r"D:\Code_learn\Build_exe"
workpath = "--workpath=" + r"D:\Code_learn\Build_exe\tempo"
PyInstaller.__main__.run([
    "--onedir",
    # "--onefile",
    r"D:\Code_learn\Project_test\main.py",
    "-nSignal_Processing",
    "--windowed",
    distpath,
    workpath,
    "--clean",
    "-y"
])