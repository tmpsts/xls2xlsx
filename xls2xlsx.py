# xls2xlsx.py - Matthew Denton
# For use as terminal applicaiton with user input
# Paste and open the file in the highest directory of the .xls files
# Follow the prompts to convert all .xls files to .xlsx files

import os
import time
from pathlib import Path
import win32com.client as win32

s = "./"
d = "./xls_old/"
os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
def convert_xls_to_xlsx(save, path: Path) -> None:

    start = time.perf_counter()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path.absolute())

    # FileFormat=51 is for .xlsx extension
    wb.SaveAs(str(path.absolute().with_suffix(".xlsx")), FileFormat=51)
    wb.Close()
    excel.Application.Quit()

    if save:
        if os.path.exists("./xls_old"):
            pass
        else:
            os.mkdir("./xls_old")
        os.rename(s + str(path), d + str(path))

    else:
        os.remove(s + str(path))

    end = time.perf_counter()

    print('File converted successfully: ', path, ' in ', round(end-start, 2), ' seconds')

def question_save():
    print("\nDo you wish to save the original .xls files in ./xls_old ? (y/n)")
    input_ = input()
    if input_ == "n":
        return False
    elif input_ == "y":
        return True
    else:
        print("Invalid input. Please try again.")
        question_save()

def question_subfolders():
    print("\nDo you wish to convert files from all subfolders as well? (y/n)")
    input_ = input()
    if input_ == "n":
        return False
    elif input_ == "y":
        return True
    else:
        print("Invalid input. Please try again.")
        question_subfolders()

def search_subfolders(save):    
    for folder in Path(os.getcwd()).iterdir():
            if folder.is_dir() and folder.name != "xls_old":
                os.chdir(folder)
                for file in Path('.').glob('*.xls'):
                    convert_xls_to_xlsx(save, file)
                search_subfolders(save)
                os.chdir('..')

if __name__ == "__main__":
    print()
    print("="*76)
    print("-"*76)
    print()
    print("  ██╗  ██╗██╗     ███████╗    ██████╗     ██╗  ██╗██╗     ███████╗██╗  ██╗\n  ╚██╗██╔╝██║     ██╔════╝    ╚════██╗    ╚██╗██╔╝██║     ██╔════╝╚██╗██╔╝\n   ╚███╔╝ ██║     ███████╗     █████╔╝     ╚███╔╝ ██║     ███████╗ ╚███╔╝\n   ██╔██╗ ██║     ╚════██║    ██╔═══╝      ██╔██╗ ██║     ╚════██║ ██╔██╗ \n  ██╔╝ ██╗███████╗███████║    ███████╗    ██╔╝ ██╗███████╗███████║██╔╝ ██╗\n  ╚═╝  ╚═╝╚══════╝╚══════╝    ╚══════╝    ╚═╝  ╚═╝╚══════╝╚══════╝╚═╝  ╚═╝")
    print("  ____  _  _    _  _   __  ____  ____    ____  ____  __ _  ____  __   __ _ \n (  _ \\( \\/ )  ( \\/ ) / _\\(_  _)(_  _)  (    \\(  __)(  ( \\(_  _)/  \\ (  ( \\ \n  ) _ ( )  /   / \\/ \\/    \\ )(    )(     ) D ( ) _) /    /  )( (  O )/    /\n (____/(__/    \\_)(_/\\_/\\_/(__)  (__)   (____/(____)\\_)__) (__) \\__/ \\_)__)")
    print()
    print("-"*76)
    print("="*76)
    save = question_save()
    subfolders = question_subfolders()
    print("\nConverting all .xls files to .xlsx files in the current directory...")

    if subfolders:
        for file in Path('.').glob('*.xls'):
            convert_xls_to_xlsx(save, file)
        search_subfolders(save)
    
    else:
        for file in Path('.').glob('*.xls'):
            convert_xls_to_xlsx(save, file)

    print("\nAll .xls files have been converted to .xlsx files!")
    print("\nPress enter to close...")
    input()