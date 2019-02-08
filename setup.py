from cx_Freeze import setup, Executable

base = None    

executables = [Executable("main.py", base=base)]

packages = ["idna", "prettytable", "openpyxl"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "ExcelBook",
    options = options,
    version = "<any number>",
    description = '<any description>',
    executables = executables
)