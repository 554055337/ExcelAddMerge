import os
from tkinter import filedialog
from Workbook import Workbook
from Merge3 import Merge3
import traceback
import tkinter.messagebox
from ParseError import ParseError

list1 = []
try:
    default_dir = r"D:\\"
    mergeFiles = filedialog.askopenfilenames(title=u'选择合并文件', initialdir=(os.path.expanduser(default_dir)))
    if mergeFiles == '': raise Exception("至少选择一个合并文件!!!")

    for file in mergeFiles: list1.append(Workbook().init(file))
    merge = Merge3().init(list1)

    saveFile = filedialog.askopenfilename(title=u'选择保存文件', initialdir=(os.path.expanduser(default_dir)))
    if saveFile == '': raise Exception("未选择保存文件!!!")

    merge.save(saveFile)

    tkinter.messagebox.showinfo('结果', "完成")
except (Exception, ParseError) as e:
    traceback.print_exc()
    tkinter.messagebox.showerror('结果', e)
finally:
    for wb in list1:
        wb.close()
