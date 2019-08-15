#skilloxide invoice generator
#from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import tkinter as tk
import tkinter.messagebox
import inflect
from tkinter import *
import tkinter.ttk as T
import tkinter.font as font
import time
#from tkinter.ttk import *
#root = Tk()
def clear():
    pass
##    name.delete(0, 'end')    
##    amount.delete(0, 'end')
##    course.delete(0, 'end') 
##    res = ""
##    namel.configure(name= res)
##    amountl.configure(amount= res)
##    coursel.configure(course= res)

def selectedCity(event):
    global selected
    selected=numberChosen.get()
    print(selected)



def send_entry():
    f=open('count.txt','r')
    content = f.read()
    print(content)
    f.close()
    template = "receipt - Copy.docx"
    n=amount.get()
    p = inflect.engine()
    a = p.number_to_words(n)
    #print(a)
    document = MailMerge(template)
    document.merge(
    inumber = ('00'+content),
    Client_Name = name.get(),
    Amount_in_words = a,
    Amount = n,
    payment_mode = selected,
    Course = course.get(),
     
    date='{:%d-%b-%Y}'.format(date.today())
    )
    f = open('count.txt','w')
    f.write(str(int(content.strip())+1))
    f.close()
   # print(x)
    m = name.get()+ ".docx"
    document.write(m)
    
#tk.mainloop()

def comb():
    send_entry()
    clear()

#GUI test
master = tk.Tk()
master.title("Payment Receipt Generator")
master.geometry('1280x720')
master.configure(background='white')

master.attributes('-fullscreen', True)


master.grid_rowconfigure(0, weight=1)
master.grid_columnconfigure(0, weight=1)
message = tk.Label(master, text="Payment Receipt Generation" ,bg="White"  ,fg="black"  ,width=50  ,height=3,font=('times', 30, 'italic bold '))

message.place(x=200, y=20)

namel=tk.Label(master, text="Full Name",width=20  ,height=2  ,fg="black"  ,bg="white" ,font=('times', 15, ' bold ') )
namel.place(x=400, y=200)

name = tk.Entry(master,width=20  ,bg="white" ,fg="black",font=('times', 15, ' bold '))
name.place(x=700, y=215)

amountl=tk.Label(master, text="Amount",width=20  ,height=2  ,fg="black"  ,bg="white" ,font=('times', 15, ' bold ') )
amountl.place(x=400, y=300)

amount = tk.Entry(master,width=20  ,bg="white" ,fg="black",font=('times', 15, ' bold '))
amount.place(x=700, y=315)

coursel=tk.Label(master, text="Course",width=20  ,height=2  ,fg="black"  ,bg="white" ,font=('times', 15, ' bold ') )
coursel.place(x=400, y=400)

course = tk.Entry(master,width=20  ,bg="white" ,fg="black",font=('times', 15, ' bold '))
course.place(x=700, y=415)


mode = tk.Label(master, text="Mode of Payment:",width=20  ,height=2  ,fg="black"  ,bg="white" ,font=('times', 15, ' bold '))
mode.place(x=400, y=500)
number = tk.StringVar()                     
numberChosen = T.Combobox(master,width=30, height=30, textvariable=number)
numberChosen['values'] = ('cash', 'card', 'PayTM', 'Bank transfer')     
#numberChosen.current(0)
numberChosen.pack()

numberChosen.bind("<<ComboboxSelected>>",selectedCity)
numberChosen.place(x=700, y=515)
submit = tk.Button(master, text="Submit", command=send_entry  ,fg="blue"  ,bg="white"  ,width=20  ,height=3, activebackground = "Red" ,font=('times', 15, ' bold '))
submit.place(x=500, y=600)
quitWindow = tk.Button(master, text="Quit", command=master.destroy  ,fg="blue"  ,bg="white"  ,width=20  ,height=3, activebackground = "Red" ,font=('times', 15, ' bold '))
quitWindow.place(x=800, y=600)
master.mainloop()
time.sleep(15)
