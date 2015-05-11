from Tkinter import *
from tkMessageBox import *
from FileDialog import *
import xlrd

def test():
    """Simple test program."""
    root = Tk()
    root.withdraw()
    fd = LoadFileDialog(root)
    loadfile = fd.go(key="test")
    loadfile_name = loadfile.split("/")[-1][:-4]
    new_file = loadfile_name + "_convertido.txt"
    
    book = xlrd.open_workbook(loadfile)
    sh = book.sheet_by_index(0)
    
    with open(new_file, "w") as f:
        f.write(str(sh.cell_value(rowx=16, colx=3)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=9, colx=3)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=28, colx=3)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=14, colx=3)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=22, colx=3)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=21, colx=3)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=8, colx=3)))
        f.write("\n")
        
        f.write(str(sh.cell_value(rowx=16, colx=5)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=9, colx=5)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=28, colx=5)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=14, colx=5)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=22, colx=5)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=21, colx=5)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=8, colx=5)))
        f.write("\n")
        
        f.write(str(sh.cell_value(rowx=16, colx=7)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=9, colx=7)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=28, colx=7)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=14, colx=7)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=22, colx=7)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=21, colx=7)))
        f.write("\t")
        f.write(str(sh.cell_value(rowx=8, colx=7)))
        f.write("\n")
    f.close()


if __name__ == '__main__':
    test()