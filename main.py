import pymcprotocol  # for read from Mitsubishi PLC
import openpyxl  # to send to Excel
import time  # for datetime
from tkinter import *
from tkinter import messagebox
from openpyxl import Workbook
import os
import datetime
import re

root = Tk()

root.geometry("900x600")  # gui size
root.title("Mitsubishi to Excel")

#---------------Frames for easier grid placement------------------#

frame1 = Frame(root, width=250, height=25)
frame1.grid(row=0, column=0, sticky='w')

frame2 = Frame(root, width=250, height=25)
frame2.grid(row=1, column=0, sticky='w')

frame3 = Frame(root, width=300, height=25)
frame3.grid(row=2, column=0, sticky='w')

frame4 = Frame(root, width=300, height=25)
frame4.grid(row=3, column=0, sticky='w')

frame5 = Frame(root, width=350, height=25)
frame5.grid(row=4, column=0, sticky='w')

#------------------------------------------------------------------#

global read_time
global stop_prog
global data
global IpAddress

#----------Defaults-----------------#
stop_prog = 1
#IpAddress = ''  

#-----------------------------------#


pymc3e = pymcprotocol.Type3E()

# Mitsubishi QJ71E71-100 Card settings                              Data PC Address
# TCP	Fullpassive	Send	Procedure Exist	Disable	Confirm	 1027	192.168.128.205	65535 

#headers for your excel sheet(column names--change to what you need)
data_header = ['header1', 'header2', 'header3', 'header4', 'header5', 'header6', 'header7', 'header8','DateTime']  # headers for new sheet


# change folder location when get up and running to save location of file to be sent
file = "C:\\Users\\terry\\Documents\\test_excel" + datetime.datetime.now().strftime( "%Y_%m_%d") + ".xlsx"

#-----------------------Methods/Functions----------------------------#
def getdata():
    global data
    lblmsg.config(text="Data collection running")
    dt = datetime.datetime.now().strftime("%Y-%m-%d_%H:%M:%S")
    pymc3e.connect(IpAddress, 1027)  # PLC Ethernet connection
    #testdata = ['123', '234', '333', '422', '15', '398', 'test'] # for testing
    word_values, dword_values = pymc3e.randomread(word_devices=["D3456", "D3457", "D409", "D3458", "D3459", "D413", "D3404"],dword_devices=[])  # registers to read
    data_values = word_values
    #data_values = testdata for testing
    data = data_values  # Move Data_Values to data so can write to excel
    data.append(dt)  # append date time to end of list comment out if not needed
    lb8.config(text=data)
    try:
        if os.path.isfile(file):  # if file already exists append to existing file
            wb = openpyxl.load_workbook(file)  # load workbook if already exists
            ws = wb.active
            # append the data results to the current excel file
            ws.append(data)
            wb.save(file)  # save workbook
            wb.close()  # close workbook
            #print(data) # for testing print to console


        else:  # create the excel file if doesn't already exist
            wb = Workbook()
            ws = wb.active
            ws.append(data_header)
            ws.append(data)
            wb.save(file)  # save workbook
            wb.close()


    except Exception as e:
        lblmsg.config(text=f"Failed to connect to plc: {e}", font="arial")



def data_time():
    getdata()
    root.after(read_time, check_plc_rdy)  # resume after read_time ms --keep getting data every(read_time seconds)


def check_plc_rdy():
    global read_time
    intread_time = int(read_time)# need to convert for if statement > sign, could have left a string but wanted int here.
    try:
        if stop_prog == 0:  # this is for stopping prog with gui still up.
            if intread_time > 0  and IpAddress != '': # make sure ip is set and interval > 0
              pymc3e.connect(IpAddress, 1027)
              ok_to_read = pymc3e.batchread_bitunits(headdevice="M100", readsize=1)
              index = 0
              bit_to_read = ok_to_read[index]
              if bit_to_read == 1:  # IF M100 is on ok to read data.
                # Call the function to start the timer for data collection
                data_time()

            else:
                lblmsg.config(text="Cannot connect to plc, is your interval and ip set?", font="arial")
                root.after(7000, check_plc_rdy)  # If not running check again after 7 seconds
    except Exception as e:
        lblmsg.config(text=f"Waiting for PLC ready to read, check Ip address: {e}", font="arial")
       #print(e) console test error
        root.after(7000, check_plc_rdy)


def interval_time():
    global read_time
    read_time = tbreadtime.get()
    tbreadtime.delete(0, END)
    lblint_time = Label(frame3, text=read_time, font='Arial')
    lblint_time.grid(row=0, column=4)

def start_log():
    global stop_prog
    stop_prog = 0
    check_plc_rdy()


def stop_logging():
    global stop_prog
    stop_prog = 1
    lblmsg.config(text="Data Logging Stopped")


def changeip():
    global IpAddress
    
    new_address = tb_ipadd.get()
    tb_ipadd.delete(0, END)# clear tb
    
    # Regular expression for validating an IP address
    ip_pattern = re.compile(r'^(\d{1,3}\.){3}\d{1,3}$')
    
    if ip_pattern.match(new_address):
        IpAddress = new_address
        labIPAd = Label(frame3, text=IpAddress, font='arial')
        labIPAd.grid(row=0, column=1, pady=15)
        check_plc_rdy()
    else:
        messagebox.showerror("Invalid IP", "Please enter a valid IP address.")


#------------ set up GUI layout ---------------#

lblmsg = Label(frame1, text="Starting up")
lblmsg.grid(row=0, column=0)

btn_inttime = Button(frame2, text="Change", width=12, command=interval_time)
btn_inttime.grid(row=0, column=3, padx=3, pady=15, sticky='w')

lb3 = Label(frame2, text="Interval time(ms): ")
lb3.grid(row=0, column=1, pady=10)

tbreadtime = Entry(frame2, width=25)
tbreadtime.grid(row=0, column=2, pady=10)

btn_startlog = Button(frame4, text="START LOGGING", command=start_log)
btn_startlog.grid(row=0, column=1, padx=20, pady=20, sticky='w')

btn_stoplog = Button(frame4, text='STOP LOGGING', command=stop_logging)
btn_stoplog.grid(row=0, column=2, padx=20, pady=20, sticky='w')

lb8 = Label(frame5, font="arial")
lb8.grid(row=0, column=2, padx=100, pady=10)

tb_ipadd = Entry(frame2, width=35)
tb_ipadd.grid(row=0, column=5, padx=0, sticky='w')

lbl9 = Label(frame2, text="IPADDRESS:")
lbl9.grid(row=0, column=4, padx=20, sticky='w')

btn_setip = Button(frame2, text='Set IP', width=12, command=changeip)
btn_setip.grid(row=0, column=6, sticky='w')

labIP = Label(frame3, text='IPADDRESS:')
labIP.grid(row=0, column=0, pady=15, sticky='w')


lblread = Label(frame3, text='READ INTERVAL(ms): ')
lblread.grid(row=0, column=3, padx=40, sticky='w')

root.mainloop()







