#### Define Import Libraries
import win32com.client #for outlook invoke
import datetime #for validity check
from uuid import getnode as get_mac #for mac check
import xlrd #for worksheet read
import sys #for stopping the program
from tkinter import * #for GUI interface
from tkinter import messagebox
from tkinter import filedialog
import time #for sleep/pause


##### Module 1 Compare Time
def valid():
	expdate = datetime.datetime(2016, 11, 30, 23, 59, 59, 999999)
#	expdate = datetime.datetime(2016, 10, 8, 23, 59, 59, 999999)
	nowdate = datetime.datetime.now()
	return nowdate < expdate
 
#####  Module 2 validate MAC
def check_mac():
	nowmac = '%012x' % get_mac()
	return nowmac=='00ff20c5930b' or nowmac=='80000b17bf92' or nowmac=='645106ff12fd' or nowmac=='80000b17bf96'
#	return nowmac=='00ff20c5930b' or nowmac=='80000b17bf92' or nowmac=='645106ff12fd' or nowmac=='80000b17bf97'

##### Module 3 Read from Excel Sheet
def get_mailid(i, path):
	wb= xlrd.open_workbook(path)
	wb.sheet_names()
	sh=wb.sheet_by_index(0)
	try:
		if sh.cell(i,0).value != None:
			mailid = sh.cell(i,0).value
			name = sh.cell(i,1).value
		else :
			sys.exit()
	except IndexError:
		sys.exit()
	return mailid, name

##### Module 5 Send Mail with Outlook
def send_mail(mailid, name, sub):
	olMailItem = 0x0
#	s = win32com.client.Dispatch("Mapi.Session")
	o = win32com.client.Dispatch("Outlook.Application")
#	s.Logon("Outlook2003")
	Msg = o.CreateItem(olMailItem)
	Msg.To = mailid
	Msg.Subject = sub
	Msg.GetInspector
	mailBody1 = Msg.HTMLBody
	index = mailBody1.find('>', mailBody1.find('<body'))
	mailBody2 = mailBody1[:index + 1] + name + "\n" + mailBody1[index + 1:] 
	Msg.HTMLBody = mailBody2
	Msg.Send()
	return ("Mail Sent SuccessFully")
	
	
##### Module 6 for GUI to Browse File Selection 
def open_file():
    global file_path
    file_path = filedialog.askopenfilename()
    mlabel0_1 = Label(text = file_path, fg="blue").grid(row=6, column=0)
    return file_path

##### Module 7 MAIN Process for excution
def main():
	sub = var1.get()
	message = var2.get()
	path = file_path
#	print (path)
#	print (sub)
#	print (message)
	i = 0 #defined for indexing excell sheet
	mlabel4 = Label(mGui, text ="YOUR MAILS ARE BEING SENT; PLS DON'T PRESS SEND AGAIN", fg="red", bg="yellow").place(x=10, y=400)
	messagebox.showinfo("Yout Mails are been Sent...With Details", "Subject:"+sub+"\nBody:\n"+message)
	for i in range (0, 1000):
		mailid,name = get_mailid(i, path)
#		print(mailid)
#		print (sub)
#		print (message)
		Info = send_mail(mailid, name, sub)
		time.sleep(3)
		i += 1
		
	sys.exit()



##### Uncomment below line to enable mac validation
#while valid() and check_mac():
while valid():
	mGui = Tk()
	mGui.geometry('450x450+500+20')
	mGui.title('eDispatcher    1.0v')
	#Trail Label
	mlabel = Label(text="***THIS IS A TRIAL VERSION, VALID UNTILL 31st OCT'2016***", fg="red").grid(row=0, column=0)
	##### Define Variables for GUI
	var1 = StringVar()
	var2 = StringVar()
	file_path = StringVar()

	###Display Labels on Window
	mlabel0 = Label(text="Open the File:").grid(row=5, column=0)
	#mfile0 = filedialog.askopenfilename()
	mbutton0 = Button(mGui, text="Browse", command =open_file).grid(row=6, column=0, sticky=E)
	#mlabel0_1 = Label(text = file_path, fg="blue").grid(row=6, column=0)

	mlabel1 = Label(text="Enter the SUBJECT:").grid(row=9, column=0)
	mEntry1 = Entry(mGui, width = 70, textvariable= var1).grid(row=11, column=0)
	mlabel2 = Label(text="Enter ths Message Body:").grid(row=13, column=0)
	mlabel2x3 = Label(text="   ").grid(row=14, column=0)
	mlabel3 = Label(text="Please Make Sure, Correct Signature is Selected\n             Before Pressing Send...  \n", bg = "yellow", fg="blue").grid(row=16, column=0)
#	mText = Text(mGui, height=3, width=40)
#	mText.grid(row=16, column=0)
#	mText.insert(END, "Please Make Sure, Correct Signature is Selected\n             Before Pressing Send...  \n")
#	mText.config(state=DISABLED)
#	mEntry2 = Entry(mGui, width = 70, textvariable= var2).grid(row=16, column=0)

	mbutton1 = Button (mGui, text = 'SEND', command= main).grid(row=29, column=0)
	mlabel4 = Label(text="Designed by AYNTECH SOLN. Contact:+91 996031515").place(x=150, y=425)

	mGui.mainloop()