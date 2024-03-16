import office
import os
import sys
import tkinter
import tkinter.filedialog
from tkinter import ttk

window = tkinter.Tk()
window.title('Xlsx to Pptx Converter')
# window.geometry('770x650')
window.resizable(False, False)

xlsxPath = tkinter.StringVar()
xlsxPathinfo = tkinter.Entry(window, width=40, textvariable=xlsxPath)
xlsxPathinfo.grid(column=1, row=1, sticky='we')

def choosexlsxPath():
    global xlsxPath
    xlsxPath = tkinter.filedialog.askopenfilename(initialdir=os.path.dirname(__file__), filetypes=[('Xlsx File', 'xlsx')])
    # 把文件路径显示在输入框中
    xlsxPathinfo.delete(0, tkinter.END)
    xlsxPathinfo.insert(0, xlsxPath)


choosexlsxPathBtn = tkinter.Button(window, text="选择Xlsx文件", command=choosexlsxPath)
choosexlsxPathBtn.grid(column=0, row=1, sticky='w')

pptPath = tkinter.StringVar()
pptPathinfo = tkinter.Entry(window, width=40, textvariable=pptPath)
pptPathinfo.grid(column=1, row=2, sticky='we')

def choosepptPath():
    global pptPath
    pptPath = tkinter.filedialog.askopenfilename(initialdir=os.path.dirname(__file__), filetypes=[('Pptx File', 'pptx')])
    # 把文件路径显示在输入框中
    pptPathinfo.delete(0, tkinter.END)
    pptPathinfo.insert(0, pptPath)

choosepptPathBtn = tkinter.Button(window, text="选择Pptx模板文件", command=choosepptPath)
choosepptPathBtn.grid(column=0, row=2, sticky='w')

imageFolderPath = tkinter.StringVar()
imageFolderPathinfo = tkinter.Entry(window, width=40, textvariable=imageFolderPath)
imageFolderPathinfo.grid(column=1, row=3, sticky='we')

def chooseimageFolderPath():
    global imageFolderPath
    imageFolderPath = tkinter.filedialog.askdirectory(initialdir=os.path.dirname(__file__))
    # 把文件路径显示在输入框中
    imageFolderPathinfo.delete(0, tkinter.END)
    imageFolderPathinfo.insert(0, imageFolderPath)

chooseimageFolderPathBtn = tkinter.Button(window, text="选择图片文件夹", command=chooseimageFolderPath)
chooseimageFolderPathBtn.grid(column=0, row=3, sticky='w')

# 创建一个下拉框
# 下拉框的值
values = ['图片名称已按顺序命名', '图片名称需要按名称中时间重新命名', '图片名称需要按创建时间重新命名']
# 创建一个下拉框
combo = ttk.Combobox(window, values=values, state='readonly')
# 设置默认值
combo.set(values[0])
combo.grid(column=1, row=4, sticky='we')

# exe_path = os.path.dirname(sys.executable)

def createPptx():
    try:
        office.open_file(os.path.join(os.getcwd(), "output.pptx"), pptPath).fill(xlsxPath).save()
        print("Done",os.path.join(os.getcwd(), "output.pptx"))
    except Exception as e:
        print(f"An error occurred: {e}")

createBtn = tkinter.Button(window, text="生成Pptx", command=createPptx)
createBtn.grid(column=0, row=4, sticky='w')

# exe_path = os.path.dirname(sys.executable)

# try:
#     office.open_file(os.path.join(exe_path, "output.pptx"), os.path.join(exe_path, "template.pptx")).fill(os.path.join(exe_path, "datafile.xlsx")).save()
#     print("Done")
# except Exception as e:
#     print(f"An error occurred: {e}")

window.mainloop()