import os
from tkinter import *
import time
import xlrd
import xlwt
from xlutils.copy import copy
import tkinter.messagebox
import tkinter as tk
from tkinter import ttk

## just my PR
root=Tk()
root.title("welcome")


#             ''''  submit_function  ''''

def submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar):

	for i in range(0,205,10):
		progress_bar['value']=i
		time.sleep(0.05)
		root_cice.update_idletasks()
	progress_bar['value']=0


	if(a==1):
		rb=xlrd.open_workbook('cice.xlsx')
		wb=copy(rb)
	elif(a==2):
			rb=xlrd.open_workbook('ucr.xlsx')
			wb=copy(rb)
	elif(a==3):
		rb=xlrd.open_workbook('Knuth.xlsx')
		wb=copy(rb)
	elif(a==4):
		rb=xlrd.open_workbook('IEEE.xlsx')
		wb=copy(rb)
	elif(a==5):
		rb=xlrd.open_workbook('Thespian.xlsx')
		wb=copy(rb)
	else:
		rb=xlrd.open_workbook('Graphicas.xlsx')
		wb=copy(rb)


	name=rb.sheet_by_index(0)
	val_name=name.nrows

	enroll=rb.sheet_by_index(0)
	val_enroll=enroll.nrows

	mobile=rb.sheet_by_index(0)
	val_mobile=mobile.nrows

	email=rb.sheet_by_index(0)
	val_email=email.nrows

	branch=rb.sheet_by_index(0)
	val_branch=branch.nrows

	w_sheet=wb.get_sheet(0)
	w_sheet.write(val_name,0,entry_name.get())
	w_sheet.write(val_enroll,1,entry_enrollment.get())
	w_sheet.write(val_mobile,2,entry_mobile.get())
	w_sheet.write(val_email,3,entry_email.get())
	w_sheet.write(val_branch,4,entry_branch.get())

	if(a==1):
		wb.save('cice.xlsx')
	elif(a==2):
		wb.save('ucr.xlsx')
	elif(a==3):
		wb.save('Knuth.xlsx')
	elif(a==4):
		wb.save('IEEE.xlsx')
	elif(a==5):
		wb.save('Thespian.xlsx')
	else:
		wb.save('Graphicas.xlsx')
	#progress_bar.destroy()
	label_name=Label(topframe,text="Name")
	label_enrollment=Label(topframe,text="ENROL. NO.")
	label_mobile=Label(topframe,text="Mobile No.")
	label_email=Label(topframe,text="Email id")
	label_branch=Label(topframe,text="Branch")

	entry_name=Entry(topframe)
	entry_enrollment=Entry(topframe)
	entry_mobile=Entry(topframe)
	entry_email=Entry(topframe)
	entry_branch=Entry(topframe)

	label_name.grid(row=0,sticky="E")
	label_enrollment.grid(row=2,sticky="E")
	label_mobile.grid(row=4,sticky="E")
	label_email.grid(row=6,sticky="E")
	label_branch.grid(row=8,sticky="E")

	entry_name.grid(row=0,column=1)
	entry_enrollment.grid(row=2,column=1)
	entry_mobile.grid(row=4,column=1)
	entry_email.grid(row=6,column=1)
	entry_branch.grid(row=8,column=1)

	submit_button=Button(bottomframe,text="Submit",fg="blue",command= lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))  #  ''' submit button '''
	reset_button=Button(bottomframe,text="Reset",fg="red",command= lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))     #  ''' reset button '''

	submit_button.grid(row=0,column=0)
	reset_button.grid(row=0,column=1)




#              ''''  CICE FUNCTION  '''''

def CICE(a,root_search):
	root_search.destroy()
	root_cice=Tk()
	if(a==1):
		root_cice.title("CICE")
	elif(a==2):
		root_cice.title("ucr")
	elif(a==3):
		root_cice.title("Knuth")
	elif(a==4):
		root_cice.title("IEEE")
	elif(a==5):
		root_cice.title("Thespian")
	else:
		root_cice.title("Graphicas")


	topframe=Frame(root_cice)
	bottomframe=Frame(root_cice)

	topframe.pack()
	bottomframe.pack()

	label_name=Label(topframe,text="Name")
	label_enrollment=Label(topframe,text="ENROL. NO.")
	label_mobile=Label(topframe,text="Mobile No.")
	label_email=Label(topframe,text="Email id")
	label_branch=Label(topframe,text="Branch")

	entry_name=Entry(topframe)
	entry_enrollment=Entry(topframe)
	entry_mobile=Entry(topframe)
	entry_email=Entry(topframe)
	entry_branch=Entry(topframe)

	label_name.grid(row=0,sticky="E")
	label_enrollment.grid(row=2,sticky="E")
	label_mobile.grid(row=4,sticky="E")
	label_email.grid(row=6,sticky="E")
	label_branch.grid(row=8,sticky="E")

	entry_name.grid(row=0,column=1)
	entry_enrollment.grid(row=2,column=1)
	entry_mobile.grid(row=4,column=1)
	entry_email.grid(row=6,column=1)
	entry_branch.grid(row=8,column=1)

	progress_bar=ttk.Progressbar(bottomframe,orient="horizontal",length=200,mode="determinate")


	if(a==1):
		submit_button=Button(bottomframe,text="Submit",fg="blue",command = lambda :submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))  #  ''' submit button '''
		reset_button=Button(bottomframe,text="Reset",fg="red",command = lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))     #  ''' reset button '''
	elif(a==2):
		submit_button=Button(bottomframe,text="Submit",fg="blue",command = lambda :submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))  #  ''' submit button '''
		reset_button=Button(bottomframe,text="Reset",fg="red",command = lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))     #  ''' reset button '''
	elif(a==3):
		submit_button=Button(bottomframe,text="Submit",fg="blue",command = lambda :submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))  #  ''' submit button '''
		reset_button=Button(bottomframe,text="Reset",fg="red",command = lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))     #  ''' reset button '''
	elif(a==4):
		submit_button=Button(bottomframe,text="Submit",fg="blue",command = lambda :submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))  #  ''' submit button '''
		reset_button=Button(bottomframe,text="Reset",fg="red",command = lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))     #  ''' reset button '''
	elif(a==5):
		submit_button=Button(bottomframe,text="Submit",fg="blue",command = lambda :submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))  #  ''' submit button '''
		reset_button=Button(bottomframe,text="Reset",fg="red",command = lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))     #  ''' reset button '''
	else:
		submit_button=Button(bottomframe,text="Submit",fg="blue",command = lambda :submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))  #  ''' submit button '''
		reset_button=Button(bottomframe,text="Reset",fg="red",command = lambda : submit_function(a,entry_name,entry_enrollment,entry_mobile,entry_email,entry_branch,root_cice,topframe,bottomframe,progress_bar))     #  ''' reset button '''

	submit_button.grid(row=0,column=0)
	reset_button.grid(row=0,column=1)
	progress_bar.grid(row=1,columnspan=2)



	root_cice.mainloop()

def search_anything(a,entry):
	s=entry.get()
	if(a==1):
		wb=xlrd.open_workbook('cice.xlsx')
		name=wb.sheet_by_index(0)
	elif(a==2):
		wb=xlrd.open_workbook('ucr.xlsx')
		name=wb.sheet_by_index(0)
	elif(a==3):
		wb=xlrd.open_workbook('Knuth.xlsx')
		name=wb.sheet_by_index(0)
	elif(a==4):
		wb=xlrd.open_workbook('IEEE.xlsx')
		name=wb.sheet_by_index(0)
	elif(a==5):
		wb=xlrd.open_workbook('.xlsx')
		name=wb.sheet_by_index(0)
	else:
		wb=xlrd.open_workbook('Graphias.xlsx')
		name=wb.sheet_by_index(0)
	row=name.nrows
	cols=name.ncols
	if(s.isalpha()):
		for i in range(row):
			if(s==name.cell(i,0).value):
				s1="Name : "+name.cell(i,0).value + "\n" + "Enrollment number : " +name.cell(i,1).value + "\n"+"Mobile : " +name.cell(i,2).value+ "\n"+"Email : " +name.cell(i,3).value+ "\n" + "Branch : " +name.cell(i,4).value
				tkinter.messagebox.showinfo("details",s1)
	elif(s.isdigit()):
		for i in range(row):
			if(s==name.cell(i,1).value or s==name.cell(i,2).value):
				s1="Name : "+name.cell(i,0).value + "\n" + "Enrollment number : " +name.cell(i,1).value + "\n"+"Mobile : " +name.cell(i,2).value+ "\n"+"Email : " +name.cell(i,3).value+ "\n" + "Branch : " +name.cell(i,4).value
				tkinter.messagebox.showinfo("details",s1)
				break
	else:
		for i in range(row):
			if(s==name.cell(i,3) or s==name.cell(i,4)):
				s1="Name : "+name.cell(i,0).value + "\n" + "Enrollment number : " +name.cell(i,1).value + "\n"+"Mobile : " +name.cell(i,2).value+ "\n"+"Email : " +name.cell(i,3).value+ "\n" + "Branch : " +name.cell(i,4).value
			tkinter.messagebox.showinfo("details",s1)

def search_function(a,root_search):
	root_search.destroy()
	root_search_function=Tk()
	root_search_function.title("Search")

	topframe=Frame(root_search_function)
	bottomframe=Frame(root_search_function)
	topframe.pack()
	bottomframe.pack()
	label=Label(topframe,text="Enter anything")
	label.grid(row=0,column=0)

	entry=Entry(topframe)
	entry.grid(row=0,column=1)

	search_button=Button(bottomframe,text="Search",command=lambda:search_anything(a,entry))
	search_button.grid(row=1,column=0,sticky=E)

	back_button=Button(bottomframe,text="Back",command=lambda:search(a))
	back_button.grid(row=1,column=1)
	progress_bar=ttk.Progressbar(bottomframe,orient="horizontal",length=200,mode="determinate")
	progress_bar.grid(row=2)
	progress_bar.start()

	root_search_function.mainloop()



def search(a):
	root.destroy()
	root_search=Tk()
	root_search.title("Welcome")

	topframe=Frame(root_search)
	bottomframe=Frame(root_search)

	topframe.pack()
	bottomframe.pack()

	if(a==1):
		label=Label(topframe,text="welcome to cice",font=("Helvetica",30),bg="white",fg="powderblue")
		label.pack()
	elif(a==2):
		label=Label(topframe,text="welcome to ucr",font=("Helvetica",30),bg="white",fg="powderblue")
		label.pack()
	elif(a==3):
		label=Label(topframe,text="welcome to Knuth",font=("Helvetica",30),bg="white",fg="powderblue")
		label.pack()
	elif(a==4):
		label=Label(topframe,text="welcome to IEEE",font=("Helvetica",30),bg="white",fg="powderblue")
		label.pack()
	elif(a==5):
		label=Label(topframe,text="welcome to Thespian",font=("Helvetica",30),bg="white",fg="powderblue")
		label.pack()
	else:
		label=Label(topframe,text="welcome to Graphicas",font=("Helvetica",30),bg="white",fg="powderblue")
		label.pack()

	button1=Button(bottomframe,text="Search",bg="powderblue",height=3,width=7,command = lambda : search_function(a,root_search))
	button2=Button(bottomframe,text="Submit",bg="powderblue",height=3,width=7,command = lambda: CICE(a,root_search))

	button1.grid(row=0,column=0,sticky=N)
	button2.grid(row=0,column=1,sticky=N)

	root_search.mainloop()

##              ''''  MAIN FUNCTION ''''
def main_fun():
	##            '''  TOP FRAME '''

	topframe=Frame(root,bg="powderblue")
	topframe.pack()

	photo=PhotoImage(file="download.png") # ''' jaypee logo object '''


	label1=Label(topframe,text="JAYPEE INSTITUTE OF TECHNOLOGY",bg="white",font=("Helvetica",44),fg='black') ## ''' label 1 for name of jaypee '''
	label1.pack()

	label2=Label(topframe,image=photo) ## ''' label 2 for photo '''
	label2.pack()

	label3=Label(topframe,text="\n click on your repective hub !!! \n",font=("Helvetica",30),fg='black',bg="powderblue")
	label3.pack()

	##           '''  BOTTOM FRAME '''
	bottomframe=Frame(root,bg="powderblue")
	bottomframe.pack()

	## buttons for groups

	button1=Button(bottomframe,text="CICE",height=3,width=10,fg="red",command = lambda : search(1))
	button2=Button(bottomframe,text="ucr",height=3,width=10,fg="red",command = lambda : search(2))
	button3=Button(bottomframe,text="Knuth",height=3,width=10,fg="red",command = lambda : search(3))
	button4=Button(bottomframe,text="IEEE",height=3,width=10,fg="red",command = lambda : search(4))
	button5=Button(bottomframe,text="Thespian",height=3,width=10,fg="red",command = lambda : search(5))
	button6=Button(bottomframe,text="Graphicas",height=3,width=10,fg="red",command = lambda : search(6))

	button1.grid(row=0,column=0,sticky=EW)
	button2.grid(row=2,column=0,sticky=EW)
	button3.grid(row=0,column=6,sticky=EW)
	button4.grid(row=2,column=6,sticky=EW)
	button5.grid(row=0,column=12,sticky=EW)
	button6.grid(row=2,column=12,sticky=EW)

	root.mainloop()

main_fun()
