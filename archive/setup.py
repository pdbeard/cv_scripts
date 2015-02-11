from cx_Freeze import setup, Executable

setup(
        name = "IV_parse",
        version = "0.1",
        description = "pdbeard",
        executables = [Executable("ImageViewer.py")])
