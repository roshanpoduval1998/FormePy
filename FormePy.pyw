import os
import re
import time
import random
import xlsxwriter
import tkinter as tk
from tkinter import *
from time import sleep
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as tmessage
from datetime import datetime
import tkinter.messagebox as tm
from validate_email import validate_email

LARGE_FONT = ("Verdana", 52)
FONT1 = ("Verdana",12)
FONT2 = ("Banschrift",16)
time1 = ''

user_name = os.getlogin()
Documents_folder = "C:\\Users\\{}\\Documents".format(user_name)
try:
	os.mkdir(Documents_folder+"\\FormePy")
except:
	pass
document_ = Documents_folder+"\\FormePy"
loco = os.path.dirname(os.path.realpath(__file__))

class main_frame(tk.Tk):

    def __init__(self, *args, **kwargs):

        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand = True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (IntroPage1,IntroPage2,IntroPage3):

            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(IntroPage1)

    def show_frame(self, navigation): 

        frame = self.frames[navigation]
        frame.tkraise()

class IntroPage1(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)
        def open_form():
        	
        	name = open("{}/Name_database.txt".format(document_),"w")
        	email = open("{}/Email_database.txt".format(document_),"w")
        	phone = open("{}/Phone_number_database.txt".format(document_),"w")
        	amount = open("{}/Amount_database.txt".format(document_),"w")
        	DOB = open("{}/Date_of_birth_database.txt".format(document_),"w")
        	gender = open("{}/Gender_database.txt".format(document_),"w")
        	submission_number = open("{}/Submission_number.txt".format(document_),"w")
        	address_file = open("{}/Address_database.txt".format(document_),"w")
        	controller.show_frame(IntroPage2)
        canvas = tk.Canvas(self, bg="white", width=1320,height=700, borderwidth=0)
        canvas.place(x=0,y=0)
        Start = tk.Button(self, command=open_form, borderwidth=0, highlightthickness=0)
        Start.img = tk.PhotoImage(file="{}/buttons__/start.png".format(loco))
        Start['image'] = Start.img
        Start.place(x=600,y=280)

class IntroPage2(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)
        canvas = tk.Canvas(self, bg="white", width=1320,height=700, borderwidth=0)
        canvas.place(x=0,y=0)
        canvas = tk.Canvas(self, bg="gray98", width=5,height=700, borderwidth=0)
        canvas.place(x=670,y=0)
        canvas = tk.Canvas(self, bg="gray98", width=5,height=700, borderwidth=0)
        canvas.place(x=0,y=0)
        canvas = tk.Canvas(self, bg="gray98", width=5,height=700, borderwidth=0)
        canvas.place(x=1310,y=0)
        canvas = tk.Canvas(self, bg="gray98", width=1310,height=5, borderwidth=0)
        canvas.place(x=0,y=0)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=50)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=130)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=210)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=290)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=370)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=450)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=530)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=610)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=690)
        canvas = tk.Canvas(self, bg="gray98", width=1320,height=5, borderwidth=0)
        canvas.place(x=0,y=770)
        clock = Label(self, font=('Banschrift', 20), bg='white')
        clock.place(x=100, y=8)
        def tick():
        	global time1
        	time2 = time.strftime("%H:%M:%S")
        	if time2 != time1:
        		time1 = time2
        		clock.config(text=time2)
        	clock.after(200, tick)
        tick()
        first_Name = tk.Label(self,text="First Name", width="10",
        	font=("Banschrift",15,"italic"),bg="white", anchor=W)
        first_Name.place(x=100,y=70)
        middle_Name = tk.Label(self, bg="white",text="Middle Name", width="10",
        	font=("Banschrift",15,"italic"), anchor=W)
        middle_Name.place(x=100,y=150)
        last_Name = tk.Label(self, bg="white",text="Last Name", width="10",
        	font=("Banschrift",15,"italic"), anchor=W)
        last_Name.place(x=100,y=230)
        first_Name_Entry = tk.Entry(self, width="30",
        	font=("Banschrift",15),bg="gray95")
        first_Name_Entry.place(x=310,y=70)
        self.first_Name_Entry = first_Name_Entry
        middle_Name_Entry = tk.Entry(self, width="30",
        	font=("Banschrift",15),bg="gray95")
        middle_Name_Entry.place(x=310,y=150)
        self.middle_Name_Entry = middle_Name_Entry
        last_Name_Entry = tk.Entry(self, width="30",
        	font=("Banschrift",15),bg="gray95")
        last_Name_Entry.place(x=310,y=230)
        self.last_Name_Entry = last_Name_Entry
        #-------------------------------------------------------------------------------------
        EMAIL = tk.Label(self, bg="white",text="Email-ID", width="10",
        	font=("Banschrift",15,"italic"), anchor=W)
        EMAIL.place(x=100,y=310)
        self.EMAIL = EMAIL
        Email_Entry = tk.Entry(self, width="30",
        	font=("Banschrift",15),bg="gray95")
        Email_Entry.place(x=310,y=310)
        self.Email_Entry = Email_Entry
        #-------------------------------------------------------------------------------------
        Phone_number = tk.Label(self, bg="white",text="Phone no.", width="10",
        	font=("Banschrift",15,"italic"), anchor=W)
        Phone_number.place(x=100,y=390)
        self.Phone_number = Phone_number
        Phone_number_Entry = tk.Entry(self, width="30",
        	font=("Banschrift",15),bg="gray95")
        Phone_number_Entry.place(x=310,y=390)
        #-------------------------------------------------------------------------------------
        self.Phone_number_Entry = Phone_number_Entry
        Amount = tk.Label(self, bg="white",text="Alt. Phone no.", width="15",
        	font=("Banschrift",15,"italic"), anchor=W)
        Amount.place(x=100,y=470)
        self.Amount = Amount
        Amount_Entry = tk.Entry(self, width="30",
        	font=("Banschrift",15),bg="gray95")
        Amount_Entry.place(x=310,y=470)
        self.Amount_Entry = Amount_Entry
        #-------------------------------------------------------------------------------------
        Date_of_birth = tk.Label(self, bg="white",text="D - O - B", width="10",
        	font=("Banschrift",15,"italic"), anchor=W)
        Date_of_birth.place(x=700,y=70)
        self.Date_of_birth = Date_of_birth
        #-------------------------------------------------------------------------------------
        year_value = [1930, 1931, 1932, 1933, 1934, 1935, 1936, 1937, 
        1938, 1939, 1940, 1941, 1942, 1943, 1944, 1945, 1946, 1947, 1948, 
        1949, 1950, 1951, 1952, 1953, 1954, 1955, 1956, 1957, 1958, 1959, 
        1960, 1961, 1962, 1963, 1964, 1965, 1966, 1967, 1968, 1969, 1970, 
        1971, 1972, 1973, 1974, 1975, 1976, 1977, 1978, 1979, 1980, 1981, 
        1982, 1983, 1984, 1985, 1986, 1987, 1988, 1989, 1990, 1991, 1992,
        1993, 1994, 1995, 1996, 1997, 1998, 1999, 2000, 2001, 2002, 2003, 
        2004,2005,2006]
        year_initial = StringVar(self)
        year_initial.set("Year")
        year = ttk.Combobox(self, textvariable=year_initial, values = year_value)
        year.config(width="5")
        year.config(font=("Banschrift",15))
        year.place(x=990,y=70)
        self.year_initial = year_initial
        #--------------------------------------------------------------------------------------
        month_value = ["01","02","03","04","05","06","07","08","09","10","11","12"]
        month_initial = StringVar(self)
        month_initial.set("Month")
        month = ttk.Combobox(self, textvariable=month_initial, values = list(month_value))
        month.config(width="5")
        month.config(font=("Banschrift",15))
        month.place(x=910,y=70)
        self.month_initial = month_initial
        #-------------------------------------------------------------------------------------
        day_value = ["01","02","03","04","05","06","07","08","09","10","11","12","13","14"
        ,"15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"]
        Day_initial = StringVar(self)
        Day_initial.set("Day")
        Day = ttk.Combobox(self, textvariable=Day_initial, values = list(day_value))
        Day.config(width="5")
        Day.config(font=("Banschrift",15))
        Day.place(x=830,y=70)
        self.Day_initial = Day_initial
        #-------------------------------------------------------------------------------------
        gender = tk.Label(self, text="Gender", bg="white", width="10",
        	font=("Banschrift",15,"italic"), anchor=W)
        gender.place(x=700,y=150)
        self.gender = gender
        gender_value_list = ["Male","Female","Other"]
        gender_value = StringVar(self)
        gender_value.set("Select")
        gen = ttk.Combobox(self, textvariable=gender_value, values = list(gender_value_list))
        gen.config(width="15")
        gen.config(font=("Banschrift",15))
        gen.place(x=830,y=150)
        self.gender_value = gender_value
        #--------------------------------------------------------------------------------------
        address_line_1 = tk.Label(self, text="Address Line 1", width="30",
            font=("Banschrift",15),bg="white",anchor=W)
        address_line_1.place(x=700,y=230)
        address_line_1_Entry = tk.Entry(self, width="30",
            font=("Banschrift",15),bg="gray95")
        address_line_1_Entry.place(x=870,y=230)
        self.address_line_1_Entry = address_line_1_Entry
        #--------------------------------------------------------------------------------------
        address_line_2 = tk.Label(self, text="Address Line 2", width="30",
            font=("Banschrift",15),bg="white",anchor=W)
        address_line_2.place(x=700,y=310)
        address_line_2_Entry = tk.Entry(self, width="30",
            font=("Banschrift",15),bg="gray95")
        address_line_2_Entry.place(x=870,y=310)
        self.address_line_2_Entry = address_line_2_Entry
        #--------------------------------------------------------------------------------------
        address_line_3 = tk.Label(self, text="Address Line 3", width="30",
            font=("Banschrift",15),bg="white",anchor=W)
        address_line_3.place(x=700,y=390)
        address_line_3_Entry = tk.Entry(self, width="30",
            font=("Banschrift",15),bg="gray95")
        address_line_3_Entry.place(x=870,y=390)
        self.address_line_3_Entry = address_line_3_Entry
        #-------------------------------------------------------------------------------------
        Zip_code = tk.Label(self, text="Zip code", width="30",
            font=("Banschrift",15),bg="white",anchor=W)
        Zip_code.place(x=700,y=470)
        Zip_code_Entry = tk.Entry(self, width="30",
            font=("Banschrift",15),bg="gray95")
        Zip_code_Entry.place(x=870,y=470)
        self.Zip_code_Entry = Zip_code_Entry
        #-------------------------------------------------------------------------------------
        submit = tk.Button(self, command=self.data_acquisition)# text="Submit",width="8", font=FONT1,
        submit.img = tk.PhotoImage(file="{}/buttons__/Submit_entry.png".format(loco))
        submit['image'] = submit.img
        submit.config(borderwidth=0, highlightthickness=0)
        submit.place(x=700,y=625)
        self.submit = submit
        clear = tk.Button(self,command=self.clear_Entry)# text="Clear",width="8", font=FONT1, 
        clear.img = tk.PhotoImage(file="{}/buttons__/clear.png".format(loco))
        clear['image'] = clear.img
        clear.config(borderwidth=0, highlightthickness=0)
        clear.place(x=600,y=625)
        self.clear = clear
        def back(event=None):
        	message = tm.askyesnocancel("restart?"
        		,"Do you want to restart project\nNOTE : All the data will be lost")
        	if message == YES:controller.show_frame(IntroPage1)
        back = tk.Button(self, command=back)#text="Back", font=FONT2)
        back.img = tk.PhotoImage(file="{}/buttons__/back.png".format(loco))
        back['image'] = back.img
        back.config(borderwidth=0, highlightthickness=0)
        back.place(x=8,y=8)
        self.back = back
        def final_database(event=None):
        	message = tm.askyesnocancel("Create"
        		,"Are you sure you want to create Excel Database")
        	if message == YES:
        		controller.show_frame(IntroPage3)
        final_database = tk.Button(self, command=final_database)#text="Save as Database", width="18", font=FONT2)
        final_database.img = tk.PhotoImage(file="{}/buttons__/database_logo.png".format(loco))
        final_database['image'] = final_database.img
        final_database.config(borderwidth=0, highlightthickness=0)
        final_database.place(x=55,y=8)
        self.final_database = final_database

    def clear_Entry(self):
    	self.first_Name_Entry.delete(0, END)
    	self.middle_Name_Entry.delete(0, END)
    	self.last_Name_Entry.delete(0, END)
    	self.Email_Entry.delete(0, END)
    	self.Phone_number_Entry.delete(0, END)
    	self.Amount_Entry.delete(0, END)
    	self.year_initial.set("Year")
    	self.month_initial.set("Month")
    	self.Day_initial.set("Day")
    	self.gender_value.set("Select")
    	self.address_line_2_Entry.delete(0,END)
    	self.address_line_1_Entry.delete(0,END)
    	self.address_line_3_Entry.delete(0,END)
    	self.Zip_code_Entry.delete(0,END)

    def data_acquisition(self):
        address1 = self.address_line_1_Entry.get()
        address2 = self.address_line_2_Entry.get()
        address3 = self.address_line_3_Entry.get()
        address4 = self.Zip_code_Entry.get()
        address = open("{}/Address_database.txt".format(document_),"a+")
        address.write("{} {} {} {},".format(address1,address2,address3,address4))
        address.close()
        first_Name_value = self.first_Name_Entry.get()
        def first_Name_val(x):
        	if not re.match("^[A-Za-z]+$",x):
        		tm.showerror("Error","Invalid First Name\nPlease Enter valid data")
        		return None
        	else:
        		return x
        first_Name = first_Name_val(first_Name_value)
        middle_Name_value = self.middle_Name_Entry.get()
        def middle_Name_val(x):
        	if not re.match("^[A-Za-z]+$",x):
        		tm.showerror("Error","Invalid middle Name\nPlease Enter valid data")
        		return None
        	else:
        		return x
        middle_Name = middle_Name_val(middle_Name_value)
        last_Name_value = self.last_Name_Entry.get()
        def last_Name_val(x):
        	if not re.match("^[A-Za-z]+$",x):
        		tm.showerror("Error","Invalid Last Name\nPlease Enter valid data")
        		return None
        	else:
        		return x
        last_Name = last_Name_val(last_Name_value)
        Email_value = self.Email_Entry.get()
        def Email_verify(x):
        	s = validate_email(x)
        	if s == True:
        		return x
        	else:
        		tm.showerror("Error","Incorrect Email-ID\nPlease Enter valid data")
        		return None
        Email = Email_verify(Email_value)
        Phone = self.Phone_number_Entry.get()
        def phone(x):
        	try:
        		for i in range(int(x)):
        			line = x
        			if re.match(r"^[4-9]\d{9}$",line):
        				return x
        			else:
        				tm.showerror("Error","Incorrect Phone number\nPlease Enter valid data")
        				return None
        	except:
        		tm.showerror("Error","Incorrect Phone number\nPlease Enter valid data")
        Phone_number = phone(Phone)
        donation_Amount = self.Amount_Entry.get()
        def amount(x):
        	try:
        		for i in range(int(x)):
        			line = x
        			if re.match(r"^[4-9]\d{9}$",line):
        				return x
        			else:
        				tm.showerror("Error","Incorrect Alt.Phone Number\nPlease Enter valid data")
        				return None
        	except:
        		tm.showerror("Error","Incorrect Alt.Phone Number\nPlease Enter valid data")
        Amount = amount(donation_Amount)
        year = self.year_initial.get()
        month = self.month_initial.get()
        day = self.Day_initial.get()
        gender = self.gender_value.get()
        DOB = year+"-"+month+"-"+day
        def dob(x):
        	try:
        		if x != datetime.strptime(x, '%Y-%m-%d').strftime('%Y-%m-%d'):
        			raise ValueError
        		return DOB
        	except ValueError:
        		tm.showerror("Error","Incorrect Date Of Birth\nPlease Enter valid data")
        		return None
        dob(DOB)
        date = DOB
        submit_value = {last_Name,middle_Name,first_Name,Email,
        Amount,year,month,day,date,gender,Phone_number}
        submit_value_list = [last_Name,middle_Name,first_Name,Email,
        Amount,year,month,day,date,gender,Phone_number]
        def error_box():
        	j = 0
        	for i in submit_value_list:
        		if (i == ""):
        			j+=1
        		if (i == "Year"):
        			j+=1
        		if (i == "Month"):
        			j+=1
        		if (i == "Day"):
        			j+=1
        		if (i == None):
        			j+=1
        		if (i == "select"):
        			j+=1
        		d = j
        	if d > 0:
        		tm.showerror("Missing Entry","Unable to get {} Entries".format(d))
        	if d == 0:
        		Name_file = open("{}/Name_database.txt".format(document_),"a+")
        		Name_file.write("{} {} {},".format(first_Name,middle_Name,last_Name))
        		Name_file.close()
        		Email_file = open("{}/Email_database.txt".format(document_),"a+")
        		Email_file.write("{},".format(Email))
        		Email_file.close()
        		Phone_number_file = open("{}/Phone_number_database.txt".format(document_),"a+")
        		Phone_number_file.write("{},".format(Phone_number))
        		Phone_number_file.close()
        		Amount_file = open("{}/Amount_database.txt".format(document_),"a+")
        		Amount_file.write("{},".format(Amount))
        		Amount_file.close()
        		date_of_birth = open("{}/Date_of_birth_database.txt".format(document_),"a+")
        		date_of_birth.write("{},".format(date))
        		date_of_birth.close()
        		gender_file = open("{}/Gender_database.txt".format(document_),"a+")
        		gender_file.write("{},".format(gender))
        		gender_file.close()
        		submission_file = open("{}/Submission_number.txt".format(document_),"a+")
        		submission_file.write("1,")
        		submission_file = open("{}/Submission_number.txt".format(document_),"r")
        		x = submission_file.readline(-1)
        		sub = "{}".format(x)
        		submission = sub.split(",")
        		lists = []
        		for i in range(0,len(submission)):
        			if submission[i] == "":
        				i+=1
        			else:
        				j = int(submission[i])
        				lists.append(j)
        				i+=1
        		total_entry = sum(lists)
        		total_entry_label = tk.Label(self, text="Total Entries = {}".format(total_entry), font=("Banschrift",20), bg="white")
        		total_entry_label.place(x=310,y=10)
        error_box()
        self.clear_Entry()

class IntroPage3(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)
        Canvas = tk.Canvas(self, width=1320,height=700, bg="white")
        Canvas.place(x=0,y=0)
        file = open("{}/DataBase_Form.xlsx".format(document_),"w")
        save = tk.Button(self,command=self.savefile)#text="save", font=FONT1, 
        save.img = tk.PhotoImage(file="buttons__/save.png".format(loco))
        save['image'] = save.img
        save.config(borderwidth=0,highlightthickness=0)
        save.place(x=600,y=275)

        def restart():
        	Message = tm.askyesnocancel("{}/restart","Do you want to go back")
        	if Message == YES:
        		controller.show_frame(IntroPage2)
        	else:
        		pass

        restart = tk.Button(self,command=restart)#text="restart", font=FONT1, 
        restart.img = tk.PhotoImage(file="{}/buttons__/restart.png".format(loco))
        restart['image'] = restart.img
        restart.config(borderwidth=0,highlightthickness=0)
        restart.place(x=500,y=280)

        def aboutus():
            os.startfile("https://github.com/roshanpoduval1998/FormePy")

        about_us = tk.Button(self,command=aboutus)
        about_us.img = tk.PhotoImage(file="{}/buttons__/about_us.png".format(loco))
        about_us['image'] = about_us.img
        about_us.config(borderwidth=0, highlightthickness=0)
        about_us.place(x=800, y=280)

    def savefile(self):
        try:
            output_file = filedialog.asksaveasfile(mode="w", defaultextension='xlsx', filetypes=(("Microsoft Excel File","*xlsx*"),("All files","*.*")))
            name_file = open("{}/Name_database.txt".format(document_),"r")
            name = name_file.readlines()
            x = name[0]
            name_val = x.split(",")
            email_file = open("{}/Email_database.txt".format(document_),"r")
            email = email_file.readlines()
            y = email[0]
            email_val = y.split(",")
            Phone_file = open("{}/Phone_number_database.txt".format(document_),"r")
            Phone = Phone_file.readlines()
            z = Phone[0]
            Phone_val = z.split(",")
            amount_file = open("{}/Amount_database.txt".format(document_),"r")
            amount = amount_file.readlines()
            a = amount[0]
            amount_val = a.split(",")
            DOB_file = open("{}/Date_of_birth_database.txt".format(document_),"r")
            DOB = DOB_file.readlines()
            b = DOB[0]
            DOB_val = b.split(",")
            gender_file = open("{}/Gender_database.txt".format(document_),"r")
            gender = gender_file.readlines()
            c = gender[0]
            gender_val = c.split(",")
            Address_file = open("{}/Address_database.txt".format(document_),"r")
            Address = Address_file.readlines()
            v = Address[0]
            Address_val = v.split(",")
            Address_entity = Address_val[0]
            workbook = xlsxwriter.Workbook(output_file.name)
            worksheet = workbook.add_worksheet()
            bold = workbook.add_format({'bold':True})
            title=["NAME","EMAIL-ID","Phone Number","Alt. Phone Number","Date of Birth","Phone Number"]
            list_column = ['A1','B1','C1','D1','E1','F1']
            for item in range(len(title)):
                worksheet.write(list_column[item],title[item],bold)
            row = 1
            for item in name_val:
                worksheet.write(row,0,item)
                row+=1
            row = 1
            for item in email_val:
                worksheet.write(row,1,item)
                row+=1
            row = 1
            for item in Phone_val:
                worksheet.write(row,2,item)
                row+=1         
            row = 1 
            for item in amount_val:
                worksheet.write(row,3,item)
                row+=1
            row = 1
            for item in DOB_val:
                worksheet.write(row,4,item)
                row+=1
            row = 1
            for item in gender_val:
                worksheet.write(row,5,item)
                row+=1
            row = 1
            for item in Address_val:
                worksheet.write(row,6,item)
                row+=1
            workbook.close()

        except IndexError:
            message = tm.showinfo("Invalid Form","Form cannot be created without Entries")

app = main_frame()
windowWidth = 1320
windowHeight = 700
x_ = int(app.winfo_screenwidth()/2 - windowWidth/2)
y_ = int(app.winfo_screenheight()/2 - windowHeight/2)
app.geometry("1320x700+{}+{}".format(x_,y_))
def close_app():
    try:
        x = tmessage.askyesnocancel("Exit","Are you sure you want to Exit")
        if x == YES:
            app.quit()
        else:
            pass
    except:
        pass
app.configure(bg="white")
app.title("FormePy")
app.iconbitmap("{}/icons__/icon.ico".format(loco))
app.resizable(width=0,height=0)
app.wm_attributes('-alpha',0.9)
app.protocol('WM_DELETE_WINDOW',close_app)
app.mainloop()