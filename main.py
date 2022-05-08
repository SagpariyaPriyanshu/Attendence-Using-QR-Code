# pip install tkinter
# pip install tkcalendar
# pip install numpy
# pip install pyzbar
# pip install pymysql

import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import messagebox as mess
from tkcalendar import DateEntry
from datetime import date
from mysql.connector import Error
import mysql.connector
import pyzbar.pyzbar as pyzbar
import numpy as np
import cv2
from pymysql import*
import xlwt
import pandas.io.sql as sql
from tkinter import font as tkFont
import datetime
import time
import pyttsx3
import re

global e1,e2,e3,e4,Name,Subject,Class,db_table_name,new_date,message,root

def ok():
    global db_table_name
    Subject=e1.get()
    Name=e2.get()
    Class=e3.get()
    Date=e4.get()
    new_date=Date.replace('/','_')
    new_name=Name.replace(' ','_')
    new_subject=Subject.replace(' ','_')
    new_class=Class.replace(' ','_') 
    db_table_name = new_subject+"_"+new_name+"_"+new_class+"_"+new_date

    if Name=='Select Faculty' or Subject=='Select Subject' or Class=='Select Class':
        message.set("Enter full details")
    else:
        main()

def main():
    try:
        connection = mysql.connector.connect(host='localhost',database='qr',user='root',password='')
        if connection.is_connected():
            cursor = connection.cursor()

    except Error as e:
        print("Error while connecting to MySQL", e)

    try:
        cursor.execute("CREATE TABLE "+ db_table_name +" (Id INT AUTO_INCREMENT, Enrollment VARCHAR(12),Name VARCHAR(30),PRIMARY KEY( Id ))")
    except mysql.connector.Error as error:
        print("Failed to create table in MySQL: {}".format(error))

    cap = cv2.VideoCapture(0)
    cap.set(3,1280)
    cap.set(4,1024)
    font = cv2.FONT_HERSHEY_SIMPLEX

    while(cap.isOpened()):
        ret, frame = cap.read()
        im = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)         
        data = pyzbar.decode(im)

        for i in data: 
            points = i.polygon     
            # If the points do not form a quad, find convex hull
            if len(points) > 4 : 
                hull = cv2.convexHull(np.array([point for point in points], dtype=np.float32))
                hull = list(map(tuple, np.squeeze(hull)))
            else : 
                hull = points;         
            # Number of points in the convex hull
            n = len(hull)     
            # Draw the convext hull
            for j in range(0,n):
                cv2.line(frame, hull[j], hull[ (j+1) % n], (0,255,0), 4)


            barCode = str(i.data)
            enrollment = barCode[2:14]
            name = barCode[14:-1]
            cv2.putText(frame, name, (i.rect.left,i.rect.top), font, 1, (255,60,), 2, cv2.LINE_AA)

            cursor.execute("select Enrollment from "+db_table_name+" where Enrollment=%s", (enrollment,))
            record = cursor.fetchone()

            cnf=r"20SOE[CI][ET]110[0-9][0-9]"
            if not(re.fullmatch(cnf,enrollment)):
                k = pyttsx3.init()
                k.setProperty("rate", 150)
                k.say(f"sorry your id is not found, please contact to admin ")
                k.runAndWait()

            elif not record: 
                cursor.execute("INSERT INTO "+db_table_name+"(Enrollment,Name) VALUES (%s,%s)",(enrollment,name))
                j = pyttsx3.init()
                j.setProperty("rate", 150)
                j.say(f"{name} your attendance is taken")
                j.runAndWait()
                connection.commit()
                
            else:
                k = pyttsx3.init()
                k.setProperty("rate", 150)
                k.say(f"{name} your attendance is already taken")
                k.runAndWait()  

        cv2.imshow('frame',frame)
        key = cv2.waitKey(1)
        if key & 0xFF == ord('q'):
            break   

    cap.release()
    cv2.destroyAllWindows()

    if connection.is_connected():
        df=sql.read_sql('select * from '+db_table_name+"",connection)
        df.to_excel(db_table_name+'.xls')
        cursor.close()
        connection.close()
        print("Successfully complete")
        root.destroy()

def tick():
    time_string = time.strftime('%d-%m-%Y | %I:%M:%S %p')
    clock.config(text=time_string)
    clock.after(200,tick)

def contact():
    mess._show(title='Contact us', message="Please contact us on : jaynandasana14@gmail.com ")
def about():
    mess._show(title='About', message="Created by Jay Nandasana")
#######################################################     GUI     #######################################################  
# root window
root = tk.Tk()
root.geometry('555x480')
root.title('Attandance System')
root.resizable(0, 0)
root.configure(bg='lightgreen')
helv12 = tkFont.Font(family='Helvetica', size=11)

menubar = tk.Menu(root,relief='ridge')
filemenu = tk.Menu(menubar,tearoff=0)
filemenu.add_command(label='Contact Us', command = contact)
filemenu.add_command(label='About', command = about)
filemenu.add_command(label='Exit',command = root.destroy)
menubar.add_cascade(label='Help',font=('times', 29, ' bold '),menu=filemenu)

b= Label(root, text="QR Code Based Attendance System",font=("Helvetica", 15))
b.config(background='lightgreen',foreground='blue')
b.grid(columnspan=2,row=0,  padx=7, pady=7)

clock = tk.Label(root ,bg="lightgreen",height=2,font=('times', 14, ' bold '))
clock.grid(columnspan=2, row=1)
tick()

# configure the grid
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=2)

# subject
subject = ttk.Label(root, text="Subject  :",font=("Helvetica", 12))
subject.config(background="lightgreen")
subject.grid(column=0, row=2,  padx=5, pady=20)
e1= StringVar()
e1.set("Select Subject")
x=OptionMenu(root, e1, "Java","Python","OS","IP","Maths","DCN","NEN","CPI")
x.config(bg="white",font=helv12)
x['menu'].config(bg="LightSteelBlue2",font=helv12)
x.grid(column=1, row=2,  padx=50, sticky=tk.EW)

# username
name = ttk.Label(root, text="Faculty Name  :",font=("Helvetica", 12))
name.config(background='lightgreen')
name.grid(column=0, row=3,  padx=5, pady=20)
e2= StringVar()
e2.set("Select Faculty")
x=OptionMenu(root, e2, "Snehal Sathwara","Jasmin Jasani","Paresh Tanna","Nikunj Vadher","Neha Chauhan","Shivangi Patel","Alpa Makwana","Sweta Patel","Hiren Kathiriya","Ankit Pandya","Purvangi Butani","Mahesh Parasaniya","Dushyant Joshi")
x.config(bg="white",font=helv12)
x['menu'].config(bg="LightSteelBlue2",font=helv12)
x.grid(column=1, row=3,  padx=50, sticky=tk.EW) 

# class
Class = ttk.Label(root, text="Class  :",font=("Helvetica", 12))
Class.config(background='lightgreen')
Class.grid(column=0, row=4,  padx=5, pady=20)
e3= StringVar()
e3.set("Select Class")
x=OptionMenu(root, e3, "4CEA","4CEB","4CEC")
x.config(bg="white",font=helv12)
x['menu'].config(bg="LightSteelBlue2",font=helv12)
x.grid(column=1, row=4,  padx=50, sticky=tk.EW)

# date
today = date.today()
y = int(today.strftime("%y"))
m = int(today.strftime("%m"))
d = int(today.strftime("%d"))
date = ttk.Label(root, text="Date  :",font=("Helvetica", 12))
date.config(background='lightgreen')
# date.grid(column=0, row=5,  padx=5, pady=20)

e4 = DateEntry(root, width=12, year=y, month=m, day=d,background='deep sky blue', borderwidth=2,font=helv12,date_pattern="dd/mm/yy")

# e4.grid(column=1, row=5,  padx=50, sticky=tk.EW)

# Take Attendance button
message=StringVar()
b= Label(root, text="",font=("Helvetica", 11),textvariable=message)
b.config(background='lightgreen',foreground='red')
b.grid(columnspan=2,  padx=7, pady=7)
button = tk.Button(root, text="Take Attendance",bg='orange',command=ok,font=helv12)
button.grid(columnspan=2,  padx=5, pady=5)

root.configure(menu=menubar)
root.mainloop()