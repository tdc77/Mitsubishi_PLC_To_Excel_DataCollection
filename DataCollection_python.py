######################################################
#Author: Terry Carlisle
#Date: 6/6/24
#Name of Program: DataCollection for R and D
######################################################

import pymcprotocol # for read from Mitsubishi PLC
import openpyxl # to send to Excel
import time # for datetime
import schedule # to schedule email send time
import numpy as np
from tkinter import *
from openpyxl import Workbook
import os
import datetime
from openpyxl.utils.dataframe import dataframe_to_rows

#remember to install above modules.

root = Tk()

root.geometry("900x600")#gui size
root.title("R&D Data Collection")

#change or remove what you dont need.
global IpAddress
global read_time
global sid
global stopprog
global slurryID
global dataheader2
global nametxt
global data


stopprog = 1
read_time = 20000 #data collect interval
IpAddress = '192.168.128.34'#default IP

pymc3e = pymcprotocol.Type3E()
#mitsubishi QJ71E71-100 setup                                 Data PC Address
#TCP	Fullpassive	Send	Procedure Exist	Disable	Confirm	 1027	192.168.128.205	65535 # connection config on Mit Ethernet card

# these are excel headers, change to what you need
dataheader = ['Slurry ID','LowRPM','LowAmp','LowHZ','HighRPM','HighAmp','HighHZ','Temp','DateTime']#headers for new sheet


#change folder location when get up and running to save location of file to be sent
file =  "Path to where you want file saved remeber that slashes are //" + datetime.datetime.now().strftime("%Y_%m_%d")+".xlsx"

def GetData() :
  
  global data
  
  Label1.config(text="Data collection running", fg='green')
  dt = datetime.datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
  sid = slurryID
  pymc3e.connect(IpAddress, 1027) #PLC Ethernet connection
  #D registers to read, change to what you need
  word_values1, dword_values = pymc3e.randomread(word_devices=["D3456","D3457","D409","D3458","D3459","D413", "D3404"], dword_devices=[]) # registers to read
  data = word_values1
  Convertexcel()# add decimal places only if needed!!! comment out if you dont.
  data.append(dt) # append date time to end of list
  data.insert(0,sid)#insert slurry ID at beginning. comment out if you dont need to force a number at beginning of datastring.
  try:
   if os.path.isfile(file):  # if file already exists append to existing file
     wb = openpyxl.load_workbook(file)  # load workbook if already exists
     ws = wb.active
     # append the dataframe results to the current excel file
     ws.append(data)
     wb.save(file)  # save workbook
     wb.close() 
     lbl10.config(text=data)#print data to screen
     print(data)
     data.pop() # remove datetime from list so new can be added at next read.
     pymc3e.close()# close connection when done.
    
   else:  # create the excel file if doesn't already exist
        wb = Workbook()
        ws = wb.active
        ws.append(dataheader)
        #ws.append(dataheader2)not used
        ws.append(data) 
        wb.save(file)  # save workbook
        wb.close()
        lbl10.config(text=data)#print data to screen
        data.pop()#remove date time so next adds new
        pymc3e.close()
        
  except:
    print("Connection Failure")
    pymc3e.close()
    Label1.config(text="Failed to Connect to PLC", font="arial", fg='red')        
           

def Data_time():
  GetData()
  root.after(read_time, CheckPLCRdy) #resume after read_time ms
  

def CheckPLCRdy():
 try:
  if stopprog  == 0: # this is for stopping prog with gui still up.
   Data_time()
   
  else:
    Label1.config(text="Waiting for PLC ready to read", font="arial", fg='orange')
    root.after(7000, CheckPLCRdy)  # If not running check again after 7 seconds
 except: 
    Label1.config(text="Waiting for PLC ready to read", font="arial", fg='orange')
    pymc3e.close()
    root.after(7000, CheckPLCRdy)
    

def Interval_time():
  global read_time
  read_time = int(str(tb2.get())) * 1000
  tb2.delete(0,END)
  tb2.insert(0, read_time)
  
     
def AddName():
  global nametxt
  global dataheader2
  nametxt = tb3.get()
  lb8.config(text=nametxt)
  tb3.delete(0,END)
  dataheader2 = ["Operator: " +  str(nametxt) ]
  
  
def ADDSlurry():
  global slurryID
  slurryID =  tb1.get()
  lb5.config(text=slurryID) 
  tb1.delete(0,END)
  
  
def StartLog():
  global stopprog
  global slurryID
  try:
   if slurryID != '':
    stopprog = 0
    CheckPLCRdy()
  except:
    lb5.config(text='No Slurry ID') 
  
 
def StopLogging():
  global stopprog
  stopprog = 1
  Label1.config(text="Data Logging Stopped", fg='red')
  pymc3e.close()
  
def ChangeIP():
  global IpAddress
  newAddress = tb4.get()
  IpAddress = newAddress
  tb4.delete(0,END)
  CheckPLCRdy()


#------------------Had to split and multiply datalist because I needed floats, if you dont need 
#------------------comment out above so it doesnt do the calculations.
#------------------obviously you may need to do some fixing here.
def Convertexcel():  #convert to decimal data it comes from PLC in INT not FLOAT
  
  global data
  
  numberhundred = 100
  
  hlistconv = []
  highrpm = []
  
  tenlist = [data] #lowrpm data /10
  tenlist = data[0] / 10
    
  tenlist1 = [data]#temp data /10
  tenlist1 = data[6] / 10
  
  highrpm = [data]
  highrpm = data[3]
  
  hundredlist = [data]
  hundredlist = data[1:3] 
  
  hundredlist2 = [data]
  hundredlist2 = data[4] / 100
 
  
  hundredlist3 = [data]
  hundredlist3 = data[5] / 100
   
  #divide by 100 for 2 decimal places
  for val in hundredlist:
    hlistconv.append(val/numberhundred)
    
 
  hlistconv.insert(0,tenlist)#add low rpm to the beginning
  hlistconv.insert(3,highrpm)
  hlistconv.insert(4,hundredlist2)
  hlistconv.insert(5,hundredlist3)
  hlistconv.insert(7,tenlist1)#add temp to end of list
  
  data = hlistconv
  
       
#set up GUI layout
p1 = PanedWindow()
p1.pack(fill=BOTH, expand=1)

Label1 = Label(p1, text="Starting up")
Label1.grid(row=1, column=1, columnspan=2)

lb2 = Label(p1, text="Slurry ID#:")
lb2.grid(row=3, column=1, pady=15)

tb1 = Entry(p1, width=30 )
tb1.grid(row=3, column=2, pady=15)

btn2 = Button(p1, text="ENTER",width=12, command= ADDSlurry)
btn2.grid(row=3, column=3)

btn = Button(p1, text="Change", width=12, command= Interval_time)
btn.grid(row=2, column=3, padx=3, pady=15)

lb3 = Label(p1, text="Interval time:")
lb3.grid(row=2, column=1, pady=10)

tb2 = Entry(p1, width=25)
tb2.insert(0,read_time)
tb2.grid(row=2, column=2, pady=10) 

tb3 = Entry(p1, width=30)  
tb3.grid(row=4, column=2, pady=15) 

lb4 = Label(p1, text="Operator:")
lb4.grid(row=4, column=1, padx=5, pady=15)

btn1 = Button(p1, text="Submit", width= 12,command= AddName)
btn1.grid(row=4, column=3)

btn3 = Button(p1, text="START LOGGING", command= StartLog)
btn3.grid(row=5, column=2)

btn4 = Button(p1, text='STOP LOGGING', command=StopLogging)
btn4.grid(row=5, column=3)

lb5 = Label(p1, font="arial")
lb5.grid(row=7, column=2, padx=10, pady=25)

lb6 = Label(p1, text="SLURRY ID:")
lb6.grid(row=7, column=1, padx=10, pady=25)
  
lb7 = Label(p1, text="Operator:") 
lb7.grid(row=8,column=1,padx=2, pady=10)

lb8 = Label(p1, font="arial")  
lb8.grid(row=8, column=2, pady=10)  

tb4 = Entry(p1, width=35)
tb4.grid(row=2, column=4, padx=55)

lbl9 = Label(p1,text="IPADDRESS:")
lbl9.grid(row=1, column=4)

btn5 = Button(p1, text='Change', width=12, command=ChangeIP)
btn5.grid(row=2, column=5)

labIP = Label(p1, text='IPADDRESS:')
labIP.grid(row=9, column=1, pady=15)

labIPAd = Label(p1, font='arial')
labIPAd.grid(row=9, column=2, pady=15)
labIPAd.config(text=IpAddress)

lbl10 = Label(p1, font='arial', fg='green')
lbl10.grid(row=10,column=1, columnspan=6,pady=10)




root.mainloop()


  
 



