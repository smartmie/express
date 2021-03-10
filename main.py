import tkinter
import excel
import os
import re
import get_express
from PIL import Image,ImageTk
import windnd
from tkinter import ttk
import datetime
import shutil
from tkinter import messagebox
global raw_image
global list_time
list_time = ''
global windows 
global new_windows
global main_sheet
global sumbit_image
global sumbit_result
global main_path
global main_excel
def main():
    global windows
    windows = tkinter.Tk()
    windows.title("查询数据")
    windows.geometry('720x520')
    windows.resizable(width=0,height=0)
    windows.columnconfigure(0, weight=1)
    windows.rowconfigure(0,weight=1)
    frame1 = tkinter.Frame(windows)
    frame2 = tkinter.Frame(windows)
    image_canvas = tkinter.Canvas(frame1,bg='white',bd=2,height=120,width=170,relief = 'solid')
    sc_time = tkinter.Scrollbar(frame1)
    list_box_time = tkinter.Listbox(frame1,height=6,width=20,yscrollcommand=sc_time.set)
    sc_time.config(command=list_box_time.yview)
    for item in os.listdir(os.getcwd()+"\\date"):
        list_box_time.insert(tkinter.END,item)
    
    button_update = tkinter.Button(frame1,text="查询",height=6,width=18,command=lambda :  update_list_data(list_box_time.get(list_box_time.curselection()),list_box_data))
    button_set_date = tkinter.Button(frame1,text="添加",height=6,width=18,command= lambda : new_toplevel())
    button_build_excel = tkinter.Button(frame1,text="创建",height=6,width=6, command=lambda :build_excel(windows,list_box_time))

    sc_data_x = tkinter.Scrollbar(frame2,orient=tkinter.HORIZONTAL)
    sc_data_y = tkinter.Scrollbar(frame2,orient=tkinter.VERTICAL)
    list_box_data = tkinter.Listbox(frame2,height=20,width=97,yscrollcommand=sc_data_y.set,xscrollcommand=sc_data_x.set)
    sc_data_x.config(command=list_box_data.xview)
    sc_data_y.config(command=list_box_data.yview)
    
    list_box_data.bind('<<ListboxSelect>>',lambda event: list_box_data_click(list_box_data.get(list_box_data.curselection()),image_canvas))

    image_canvas.grid(row=0,column=0)
    list_box_time.grid(row=0,column=1)
    sc_time.grid(row=0,column=2,ipady=30)
    button_update.grid(row=0,column=3,padx=10)
    button_set_date.grid(row=0,column=4,padx=10)
    button_build_excel.grid(row=0,column=5)
    list_box_data.grid(row=0,column=0)
    sc_data_x.grid(row=1,column=0,ipadx=320)
    sc_data_y.grid(row=0,column=1,ipady=155)
    frame1.place(x=10,y=0)
    frame2.place(x=10,y=130)
    windows.mainloop()
    return list_box_data,list_box_time,image_canvas

def build_excel(w,list_):
    time_yeas = datetime.date.today().year
    time_mon = datetime.date.today().month
    time_day = datetime.date.today().day
    new_path = "{0}-{1}-{2}".format(time_yeas,time_mon,time_day)
    if not(os.path.exists(os.getcwd()+"\\date\\"+new_path)):
        list_.insert(tkinter.END,new_path)
        os.mkdir(os.getcwd()+"\\date\\"+new_path)
        os.mkdir(os.getcwd()+"\\date\\"+new_path+"\\image")
        shutil.copy(os.getcwd()+"\\formwork.xlsx",os.getcwd()+"\\date\\"+new_path+"\\"+new_path+".xlsx")

    pass

def new_toplevel():
    global new_windows
    new_windows = tkinter.Toplevel()
    new_windows.title("加入数据")
    new_windows.geometry('610x500')
    new_windows.resizable(width=0,height=0)
    frame1 = tkinter.Frame(new_windows)
    frame2 = tkinter.Frame(new_windows)
    frame3 = tkinter.Frame(new_windows)
    text = tkinter.Text(frame1,width=30,height=36)
    image_shoes = tkinter.Canvas(frame3,bg='white',bd=2,height=120,width=170,relief = 'solid')
    windnd.hook_dropfiles(image_shoes,func=lambda files:dragged_file(files,image_shoes))
    
    text1 = tkinter.Text(frame2,width=25,height=3)
    text2 = tkinter.Text(frame2,width=25,height=3)
    text3 = tkinter.Text(frame2,width=25,height=3)
    text4 = tkinter.Text(frame2,width=25,height=3)
    text5 = tkinter.Text(frame2,width=25,height=3)
    text6 = tkinter.Text(frame2,width=25,height=3)
    text7 = tkinter.Text(frame2,width=25,height=3)
    text8 = tkinter.Text(frame2,width=25,height=3)
    text9 = tkinter.Text(frame2,width=25,height=3)
    text10 = tkinter.Text(frame2,width=25,height=3)
    text11 = tkinter.Text(frame2,width=25,height=3)
    label_1 = tkinter.Label(frame2,text="码数")
    label_2 = tkinter.Label(frame2,text="姓名")
    label_3 = tkinter.Label(frame2,text="电话")
    label_4 = tkinter.Label(frame2,text="代理")
    label_5 = tkinter.Label(frame2,text="地址")
    label_6 = tkinter.Label(frame2,text="邮费")
    label_7 = tkinter.Label(frame2,text="价格")
    label_8 = tkinter.Label(frame2,text="成本")
    label_9 = tkinter.Label(frame2,text="利润")
    label_10 = tkinter.Label(frame2,text="日期")
    label_11 = tkinter.Label(frame2,text="单号")


    button_convert = tkinter.Button(frame3,text="转换",width=23,height=9,command=lambda : re_text(text.get('0.0',tkinter.END),text1, text2, text3, text4 ,text5 ,text6 ,text7 ,text8 ,text9 ,text10,text11))
    button_sumbit = tkinter.Button(frame3,text="提交",width=23,height=10,command=lambda : button_sumbit_click(text1, text2, text3, text4 ,text5 ,text6 ,text7 ,text8 ,text9 ,text10,text11))

    frame1.place(x=10,y=10)
    frame2.place(x=220,y=10)
    frame3.place(x=430,y=10)
    text.grid(column=0,row=0)
    text1.grid(column=1,row=0)
    text2.grid(column=1,row=1)
    text3.grid(column=1,row=2)
    text4.grid(column=1,row=3)
    text5.grid(column=1,row=4)
    text6.grid(column=1,row=5)
    text7.grid(column=1,row=6)
    text8.grid(column=1,row=7)
    text9.grid(column=1,row=8)
    text10.grid(column=1,row=9)
    text11.grid(column=1,row=10)

    label_1.grid(column=0,row=0)
    label_2.grid(column=0,row=1)
    label_3.grid(column=0,row=2)
    label_4.grid(column=0,row=3)
    label_5.grid(column=0,row=4)
    label_6.grid(column=0,row=5)
    label_7.grid(column=0,row=6)
    label_8.grid(column=0,row=7)
    label_9.grid(column=0,row=8)
    label_10.grid(column=0,row=9)
    label_11.grid(column=0,row=10)
    
    image_shoes.grid(row=0,column=0)
    # combobox_list.grid(row=1,column=0)
    button_convert.grid(row=2,column=0)
    button_sumbit.grid(row=3,column=0)
    new_windows.mainloop()
def button_sumbit_click(*args):
    global sumbit_result
    global main_sheet
    global list_time
    global main_excel
    global raw_image
    result_text = []
    if list_time == '':
        messagebox.showinfo(title="提醒", message="选择查询的单号")
        return 
    for text_content in args :
        if text_content.get('1.0',tkinter.END) == "\n" :
            messagebox.showinfo(title="提醒", message="内容为空,检查内容")
            return 
        else:
            result_text.append(text_content.get('1.0',tkinter.END))
    
    x_count = main_sheet.max_row+1
    for y_count in range(2, main_sheet.max_column):
        main_sheet.cell(x_count,y_count).value = result_text[y_count-2]
    path = os.getcwd()+"\\date\\"+list_time
    raw_image.save(path+"\\image\\"+str(x_count)+".jpeg")
    red_image = excel.list_image(path+"\\image")

    for x_count in range(2,main_sheet.max_row+1):
        excel.wri_excel(main_sheet,index=x_count,y='A',text="",image=red_image[str(x_count)])
    main_excel.save(path+"\\"+list_time+".xlsx")
    pass
def dragged_file(files,image_canvas):
    for f in files:
        break
    global raw_image
    global sumbit_image 
    sumbit_image = f
    raw_image = Image.open(f)
    convertImage =raw_image.convert("RGBA")
    resize_image = convertImage.resize((170,120),Image.ANTIALIAS)
    imTK = ImageTk.PhotoImage(resize_image)
    image_canvas.create_image(0,0,anchor = tkinter.NW,image=imTK)

    new_windows.mainloop()
    pass
def re_text(text,*args):
    global sumbit_result 
    text = re.sub(r'\s|\t|\n','\n',text)
    regex_phone   =  r'1\d{10}'
    regex_address =  r'([^(\d|\s)$省]+省|.+自治区)?([^市]+市)([^县]+县|.+区)?(.*)'
    regex_shoes = r'\d{2}[$码]'
    regex_name = r'([^\d]{3,4})'
    regex_date = r'\d{4}年\d{1,2}月\d{1,2}号'
    regex_money = r'^\d{2,3}$'
    print(text)
    t_split = text.split('\n')
    money = []
    for t in t_split:
        if '省' in t:
            # print(t)
            address = re.findall(regex_address,t)
        tt = re.findall(regex_name,t)
        if tt and 2 <= len(t) <= 4 :
            name = tt
        if re.match(r'[0-9]',t) and len(t) == 11:
            phone = re.findall(regex_phone,t)
        money.append(re.findall(regex_money,t))
        if re.match(r'\d{13}',t):
            express_number = t
    
    money = [int(e[0]) for e in list(filter(None,money))]
    money.sort()
    date = re.findall(regex_date,text)
    shoes = re.findall(regex_shoes,text)
    date = date if date else ['']
    print(address)
    address = address[0][0]+address[0][1]+address[0][2]+address[0][3]
    result = [shoes[0],name[0],phone[0],name[0],address,money[0],money[1],money[2],money[2]-money[1],date[0],express_number]
    count = 0
    for tk_text in args:
        tk_text.delete('1.0',tkinter.END)
        tk_text.insert(tkinter.END,result[count])
        count += 1
    sumbit_result = result
    pass
def list_box_data_click(list_obj_text,image_obj):
    global windows
    now_path = os.getcwd()+"\\date\\"+list_time+"\\image\\"
    split_text = list_obj_text.split(" ")
    # print()
    for f in os.listdir(now_path):
        # print(f)
        if split_text[0] == re.findall(r'\d+',f)[0] :
            red_image = Image.open(os.path.join(now_path,f))
            print(os.path.join(now_path,f))
            # cv_image = cv2.imread(onw_path+"\\"+f)
            
            break

    
    # cv_convert = cv2.cvtColor(cv_image, cv.COLOR_BGR2RGBA)
    convertImage =red_image.convert("RGBA")
    resize_image = convertImage.resize((170,120),Image.ANTIALIAS)
    imTK = ImageTk.PhotoImage(resize_image)
    image_obj.create_image(0,0,anchor = tkinter.NW,image=imTK)
    windows.title("无物流" if split_text[-2] == 'None' else split_text[-2])
    # for t in split_text:
    #     # print(t)
    #     pass
    windows.update()
    windows.mainloop()

    pass
def update_list_data(t,list_box_data):
    global list_time 
    global main_sheet
    global main_excel
    list_time = t
    list_box_data.delete(0,tkinter.END)
    path = os.getcwd()+"\\date\\"+t+"\\"+t+".xlsx"
    sheet ,excel_ = excel.main(path)
    main_sheet = sheet
    main_excel = excel_

    text = ""
    count = 0
    for i in range(2,sheet.max_row+1):
        for x in sheet[i]:
            # print(x.value)
            count += 1
            # print(count)
            text =  str(i)+" " if count == 2 else text + str(x.value) +" "
            # text =  text + str(x.value) +" "
        list_box_data.insert(tkinter.END,text)
        count = 0
        text = ""
    

if __name__ == "__main__":
    main()
    pass
