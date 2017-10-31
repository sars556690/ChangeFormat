# -*- coding: utf-8 -*-
import xlrd
import xlsxwriter
import datetime
from Tkinter import *
ticks = datetime.datetime.now()
filename = "" 

class Windows:
    def __init__(self, master):     
        self.filename=""
        master.title("日期時間轉換")
        master.resizable(0,0)
        self.var = StringVar()
        msg = Label( root, textvariable=self.var).grid(row=4, column=0 , columnspan=3, sticky=W)
        Lable_file=Label(master, text="File").grid(row=1, column=1)
        self.Entry_file=Entry(master)
        self.Entry_file.grid(row=1, column=2)

        self.chkbtn_value_time_col = IntVar(value=1)
        self.chkbtn_time_col=Checkbutton(master ,text="時間列數", command = self.chkbtn_command_time_col ,variable=self.chkbtn_value_time_col)
        self.chkbtn_time_col.grid(row=2, column=0 ,columnspan=2, sticky=E)

        self.Entry_time_col=Entry(master,width=5)
        self.Entry_time_col.grid(row=2, column=2 , sticky=W)
        self.Entry_time_col.insert(0, 3)
        
        self.chkbtn_value_date_col = IntVar(value=1)
        self.chkbtn_date_col=Checkbutton(master , text="日期列數",  command = self.chkbtn_command_date_col , variable=self.chkbtn_value_date_col)
        self.chkbtn_date_col.grid(row=3, column=0 ,columnspan=2, sticky=E)

        self.Entry_date_col=Entry(master ,width=5)
        self.Entry_date_col.grid(row=3, column=2, sticky=W)
        self.Entry_date_col.insert(0, 4)
        self.cbutton= Button(master, text="完成" , command=self.create_new_excel)
        self.cbutton.grid(row=2, column=3, sticky = W + E)
        self.bbutton= Button(master, text="瀏覽", command=self.browse)
        self.bbutton.grid(row=1, column=3)

    def chkbtn_command_time_col(self):
        if(not(self.chkbtn_value_time_col.get())):
            self.Entry_time_col['state'] = 'disabled'
        else:
            self.Entry_time_col['state'] = 'normal'

    def chkbtn_command_date_col(self):
        if(not(self.chkbtn_value_date_col.get())):
            self.Entry_date_col['state'] = 'disabled'
        else:
            self.Entry_date_col['state'] = 'normal'

    def browse(self):
        from tkFileDialog import askopenfilename
        opts = {}
        opts['filetypes'] = [('Excel','.xlsx'),('all files','.*')]
        self.filename = askopenfilename(**opts)
        self.Entry_file.delete(0, END)
        self.Entry_file.insert(0, self.filename)
        
    def create_new_excel(self):
        mistake = 0

        if(len(self.Entry_file.get())==0):
            self.var.set("無選擇檔案")
        elif(not(self.check_int(self.Entry_time_col.get()) and self.check_int(self.Entry_date_col.get()))):
            self.var.set("指定時間及日期列數錯誤")
        else:
            try:  
                time_col = int(self.Entry_time_col.get())-1
                date_col = int(self.Entry_date_col.get())-1
        
                data = xlrd.open_workbook(self.filename)
                table = data.sheet_by_index(0)
                
                workbook = xlsxwriter.Workbook(str(ticks.year)+'-'+str(ticks.month)+'-'+str(ticks.day) + '.xlsx')
                worksheet = workbook.add_worksheet()
                
                row_count = table.nrows
                col_count = table.ncols
                for i in range(0 , row_count):
                    for j in range(0 , col_count):
                        worksheet.write(i, j, table.cell(i,j).value)

                if(self.chkbtn_value_time_col.get()):
                    for i in range(1 , row_count):
                        time_format = self.change_time_format(table.cell(i,time_col).value)
                        if(time_format[0]):
                            worksheet.write(i, time_col, time_format[1])
                        else:
                            mistake+=1
                            format = workbook.add_format({'font_color': 'red'})
                            worksheet.write(i, time_col, str(time_format[1]),format)

                if(self.chkbtn_value_date_col.get()):
                    for i in range(1 , row_count):
                        date_format = self.change_date_format(table.cell(i,date_col).value)
                        if(date_format[0]):
                            worksheet.write(i, date_col, date_format[1])
                        else:
                            mistake+=1
                            format = workbook.add_format({'font_color': 'red'})
                            worksheet.write(i, date_col, str(date_format[1]) ,format)

                workbook.close()
                self.var.set("OK , 有"+str(mistake)+"筆錯誤")
            except Exception as inst:
                self.var.set("轉換失敗")

    def change_time_format(self ,time):
        try:
            time_hour = int(time)/100
            time_min = int(time)%100
            t_bool = self.getTime(time_hour , time_min)[0]
            dt = self.getTime(time_hour , time_min)[1]
            if(t_bool):
                return [True, dt.strftime("%H:%M")]
            else:
                return [False, str(int(time))]
        except Exception as inst:
            return [False, str(time)]

    def change_date_format(self , date):
        try:
            date_month = int(date)/100
            date_day = int(date)%100
            year = ticks.year
            t_bool = self.getDate(year , date_month , date_day)[0]
            
            if(t_bool):
                if(self.getDate(year , date_month , date_day)[1] > ticks):
                    year-=1
                dt = self.getDate(year , date_month , date_day)[1]
                return [True, dt.strftime("%Y-%m-%d")]
            else:
                return [False, str(int(date))]
        except Exception as inst:
            return [False, str(date)]

    def check_int(self , num):
        try:
            int(num)
            return True
        except ValueError:
            pass
        return False

    def getDate(self , y , m , d ):
        try:
            d = datetime.datetime.strptime( str(y)+'-'+str(m)+'-'+str(d) , "%Y-%m-%d")
        except:
            return [False , None]
        return [True , d]

    def getTime(self ,h ,M ):
        try:
            t = datetime.datetime.strptime( str(h)+":"+str(M), "%H:%M")
        except:
            return [False , None]
        return [True , t]
    
root = Tk()
window=Windows(root)
root.mainloop()  


