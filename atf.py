import tkinter as tk
from tkinter import filedialog
from handles import *
import time
import sys
import os

old_atf_loc = "**"

def cls():
    os.system('cls' if os.name=='nt' else 'clear' )

def progress(count, total, suffix=''):
    bar_len = 60
    filled_len = int(round(bar_len * count / float(total)))

    percents = round(100.0 * count / float(total), 1)
    bar = '=' * filled_len + '-' * (bar_len - filled_len)
    cls()
    sys.stdout.write('[%s] %s%s ...%s\r' % (bar, percents, '%', suffix))



def read_atf(f, filename, name,savelocation):
    trash = f.readline()
    word = ""
    while word != "[EndofFile]":
        word = f.readline()
        if word == "[EndOfFile]\n":
            return
        if word == "":
            return
        if word == "[BeginTestComment]\n":
            word = f.readline()
            word = f.readline()
            word = f.readline()
            handle_program_lines(f, filename, name,savelocation)


# noinspection PyBroadException
def handle_files(content_list_def, path_def,savelocation):
    if len(content_list_def) == 0:
        quit()

    for index in range(len(content_list_def)):
        if content_list_def[index].encode().lower().endswith(r'.atf'):
            progress(index, len(content_list_def), suffix='Progress')
            if index == len(content_list_def):
                return 0
            name = content_list[index]
            if not os.path.isfile(savelocation + "/" + name + ".xlsx"):
                filename = path_def + "/"  + name
                f = open(filename, 'r')
                read_atf(f, filename, name,savelocation)
                f.close()
                # noinspection PyBroadException
                # os.remove(filename)
                continue
                # Do nothing
            else:
                filename = path_def + "/" + name
                exceltime = os.path.getmtime(savelocation + "/" + name + ".xlsx")
                atftime = os.path.getmtime(filename)

                if exceltime < atftime:

                    f = open(filename, 'r')
                    read_atf(f, filename, name,savelocation)
                    f.close()
                    # noinspection PyBroadException
                    os.remove(filename)
                    continue
                        # Do nothing
                else:
                    print("atf is younger then excel file.")
                    filename = path_def +"/" + name
                    # noinspection PyBroadException

                    os.remove(filename)
                    continue
                    # Do nothing

print("Pick out the folder that contains the atf files.")

root = tk.Tk()
root.withdraw()
file_path_string = filedialog.askdirectory()
root.destroy()

try:
    content_list = os.listdir(file_path_string)
except WindowsError:
    # User exited out of filedialog
    exit(0)

savelocation = file_path_string

if savelocation is None:
    print("Please select a save location.")
    time.sleep(5)
    exit(0)

if not os.path.exists(savelocation):
    savelocation = "./excel"
    if not os.path.exists("./excel/"):
        os.makedirs("./excel/")

if len(content_list) > 0:
    files = handle_files(content_list, file_path_string,savelocation)
    if files is None:
        content_list = []
    print "Done!----------------------------------"
    abspath = os.path.abspath(savelocation + "/")
    os.system(r'start explorer.exe ' + abspath)
    time.sleep(3)
    quit()
else:
    print "No ATF in the Work Order Folder"


