from Tkinter import *
from tkMessageBox import *
from FileDialog import *
import xlrd

# def answer():
#     showerror("Answer", "Sorry, no answer available")
# 
# def callback():
#     if askyesno('Verify', 'Really quit?'):
#         showwarning('Yes', 'Not yet implemented')
#     else:
#         showinfo('No', 'Quit has been cancelled')
# 
# Button(text='Quit', command=callback).pack(fill=X)
# Button(text='Answer', command=answer).pack(fill=X)
# mainloop()


def test():
    """Simple test program."""
    root = Tk()
    root.withdraw()
    fd = LoadFileDialog(root)
    loadfile = fd.go(key="test")
    loadfile_name = loadfile.split("/")[-1][:-4]
    new_file = loadfile_name + "_convertido.txt"
    
#     print loadfile
    
    book = xlrd.open_workbook(loadfile)
    sh = book.sheet_by_index(0)
    
    with open(new_file, "w") as f:
        f.write(sh.cell_value(rowx=16, colx=3))
        f.write("\t")
        f.write(sh.cell_value(rowx=9, colx=3))
        f.write("\t")
        f.write(sh.cell_value(rowx=28, colx=3))
        f.write("\t")
        f.write(sh.cell_value(rowx=14, colx=3))
        f.write("\t")
        f.write(sh.cell_value(rowx=22, colx=3))
        f.write("\t")
        f.write(sh.cell_value(rowx=21, colx=3))
        f.write("\t")
        f.write(sh.cell_value(rowx=8, colx=3))
        f.write("\n")
        
        f.write(sh.cell_value(rowx=16, colx=5))
        f.write("\t")
        f.write(sh.cell_value(rowx=9, colx=5))
        f.write("\t")
        f.write(sh.cell_value(rowx=28, colx=5))
        f.write("\t")
        f.write(sh.cell_value(rowx=14, colx=5))
        f.write("\t")
        f.write(sh.cell_value(rowx=22, colx=5))
        f.write("\t")
        f.write(sh.cell_value(rowx=21, colx=5))
        f.write("\t")
        f.write(sh.cell_value(rowx=8, colx=5))
        f.write("\n")
        
        f.write(sh.cell_value(rowx=16, colx=7))
        f.write("\t")
        f.write(sh.cell_value(rowx=9, colx=7))
        f.write("\t")
        f.write(sh.cell_value(rowx=28, colx=7))
        f.write("\t")
        f.write(sh.cell_value(rowx=14, colx=7))
        f.write("\t")
        f.write(sh.cell_value(rowx=22, colx=7))
        f.write("\t")
        f.write(sh.cell_value(rowx=21, colx=7))
        f.write("\t")
        f.write(sh.cell_value(rowx=8, colx=7))
        f.write("\n")
    f.close()
    
#     print "The number of worksheets is", book.nsheets
#     print "Worksheet name(s):", book.sheet_names()
#     sh = book.sheet_by_index(0)
#     print sh.name, sh.nrows, sh.ncols
#     print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
#     for rx in range(sh.nrows):
#         print sh.row(rx)
    
#     fd = SaveFileDialog(root)
#     savefile = fd.go(key="test")
#     print loadfile, savefile


if __name__ == '__main__':
    test()