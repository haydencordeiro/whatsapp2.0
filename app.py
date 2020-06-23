import tkinter as tk
import pymsgbox
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from tkinter import BOTH, END, LEFT
from script import whatsapp_login,Load_excel

filename = ''
df=''
OptionList = [''] 
noOfVar=0
varList=[]#list of the var columns in order
variableCols=''
DropDownCols=' '

def GetFileName():
  global filename
  Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
  filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
  Refresh()
  

def Refresh():
  global df,filename,OptionList
  tk.Label(master, text="{}".format(filename), bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=0,column=1)
  df = pd.read_excel((filename))
  OptionList=(list(df.columns))
  print(OptionList)
  if(len(varList)>0):
  	tk.Label(master, text="{}".format(varList), bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)
  else:
  	tk.Label(master, text="", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)
  DropDownNumber()
  VariableColumsDropDown()


def AddVarList():
  global varList
  varList.append(str(variableCols.get()))
  dispVarList=''
  for i in varList:
    if(len(dispVarList)<1):
      dispVarList+=i
    else:
      dispVarList+=','+i

  tk.Label(master, text='                                                                                                                                       ' ,bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)

  tk.Label(master, text='{}'.format(dispVarList) ,bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)

def DelVarList():
  global varList
  try:
    # print(varList)
    temp=varList.pop()
    # print(temp,varList)
  except:
    pass
  dispVarList=''
  for i in varList:
    if(len(dispVarList)<1):
      dispVarList+=i
    else:
      dispVarList+=','+i
  # tk.Label(master, text=spaceText ,bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)
  tk.Label(master, text='                                                                                                                                       ' ,bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)

  # varList+=str(variableCols.get())+','
  # tk.Label(master, text='                                            ' ,bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)
  tk.Label(master, text="{}".format(dispVarList), bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)
  
def SendMessage():
  global df,varList,filename
  dispVarList=''
  for i in varList:
    if(len(dispVarList)<1):
      dispVarList+=i
    else:
      dispVarList+=','+i
  Load_excel(df,MsgBox.get('1.0', END),dispVarList,len(varList),str(DropDownCols.get()))





def DropDownNumber():
  global DropDownCols
  DropDownCols = tk.StringVar(master)
  DropDownCols.set(OptionList[0])
  opt = tk.OptionMenu(master, DropDownCols, *OptionList)
  opt.config(width=15, font=('Helvetica', 12))
  opt.grid(row=2, column=1, padx=10, pady=10)


def VariableColumsDropDown():
  global variableCols
  variableCols = tk.StringVar(master)
  variableCols.set(OptionList[0])
  opt2 = tk.OptionMenu(master, variableCols, *OptionList)
  opt2.config(width=15, font=('Helvetica', 12))
  opt2.grid(row=4, column=1, padx=10, pady=10)

#labels
master = tk.Tk()
master.configure(bg='white')


#Filename
tk.Label(master, text="File Name :", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=0)
tk.Label(master, text="None Selected", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=0,column=1)

#message Box
tk.Label(master, text="Message", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=1)
MsgBox = tk.Text(master,height=10)
MsgBox.grid(row=1, column=1, padx=10, pady=10)


#Mobile Number
tk.Label(master, text="Mobile Numbers", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=2)
DropDownNumber()


#Variable Columns List
tk.Label(master, text="Variable Order", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3)
tk.Label(master, text="", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=3,column=1)
tk.Label(master, text="Select New Variable", bg='white',fg="#0A5688", font=("Arial", 12, "bold")).grid(row=4)
VariableColumsDropDown()

tk.Button(master, text='+', command=AddVarList,bg='#F9D162', fg="#0A5688", font=("Helvetica", 10, "bold"), height=3, width=3, activebackground="#F3954F").grid(row=4, column=2, padx=10, pady=20)

tk.Button(master, text='-', command=DelVarList,bg='#F9D162', fg="#0A5688", font=("Helvetica", 10, "bold"), height=3, width=3, activebackground="#F3954F").grid(row=4, column=3, padx=10, pady=20)



tk.Button(master,text='Send',command=SendMessage, bg='#F9D162', fg="#0A5688", font=("Helvetica", 10, "bold"), height=3, width=10, activebackground="#F3954F").grid(row=5,column=0,padx=10, pady=20)
tk.Button(master, text='Select File', command=GetFileName,bg='#F9D162', fg="#0A5688", font=("Helvetica", 10, "bold"), height=3, width=10, activebackground="#F3954F").grid(row=5, column=1, padx=10, pady=20)
tk.Button(master, text='Login', command=whatsapp_login,bg='#F9D162', fg="#0A5688", font=("Helvetica", 10, "bold"), height=3, width=10, activebackground="#F3954F").grid(row=5, column=2, padx=10, pady=20)
# tk.Button(master, text='Plot Pie', command=Plot_Pie,bg='#F9D162', fg="#0A5688", font=("Helvetica", 10, "bold"), height=3, width=10, activebackground="#F3954F").grid(row=5, column=2, padx=10, pady=20)



master.mainloop()

tk.mainloop()