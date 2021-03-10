import excel
import get_express
import os
from openpyxl.drawing.image import Image
from tkinter import StringVar
import tkinter
import time
global list_box 
global label_1
global win
global var
win =tkinter.Tk()
var = StringVar()
var.set("查询日期")
def main():
    global list_box 
    global label_1
    global win

    win.geometry('165x170')
    sc_time = tkinter.Scrollbar(win)
    list_box = tkinter.Listbox(win,height=6,width=20,yscrollcommand=sc_time.set)
    sc_time.config(command=list_box.yview)
    for item in os.listdir(os.getcwd()+"\\date"):
        list_box.insert(tkinter.END,item)
    label_1 = tkinter.Label(win,textvariable=var)
    button_express = tkinter.Button(text="查询",width=20,command=button_click)
    list_box.grid(row=0,column=0)
    sc_time.grid(row=0,column=1,ipady=30)
    label_1.grid(row=1,column=0)
    button_express.grid(row=2,column=0,rowspan=1,columnspan=6)
    win.mainloop()
    pass
def button_click():
    global list_box 
    global label_1
    global var
    path = os.getcwd()+"\\date\\"+list_box.get(list_box.curselection())
    file_name = list_box.get(list_box.curselection())+".xlsx"
    print(path,file_name)
    renew(path,file_name, var)
    pass
def renew(path,file_name,var):
    var.set("开始查询")
    print(path+"\\"+file_name)
    sheet,excel_ = excel.main(path+"\\"+file_name)
    image_list = [i for i in os.listdir(path+"\\image")]
    e_i = 0
    for express in sheet['L']:
        e_i += 1
        if express.value != '单号':
            data = get_express.get_data(express.value)
            print(e_i)
            for name in image_list:
                if str(e_i) in name :
                    break
            r_image = Image(path+"\\image\\"+name)
            if data['message'] == 'ok':
                excel.wri_excel(sheet,index=e_i,y='M',text=data['data'][0]['context'])
                if len(data['data'][0]['context']) >= 5:
                    excel.set_color(sheet,x=e_i)
            else:
                excel.wri_excel(sheet,index=e_i,y='M',text=data['message'])
            excel.wri_excel(sheet,index=e_i,y='A',image=r_image)

    excel_.save(path+"\\"+file_name)
    var.set("完成")


if __name__ == "__main__":
    main()
    pass