'''
Jim Tracker Application

Dependencies:
    Python 2.7.5
    Tkinter 8.5.3

Author: Tyler Weaver
'''

from log_data import LoggerDialog,NewCustomerException
from customer import NewCustomerDialog,Customers

from Tkinter import Menubutton,Menu
from ttk import Frame, Label, Combobox, LabelFrame, Button
from tkMessageBox import showerror
from datetime import datetime

class TrackerWindow(Frame):
	def __init__(self):
		Frame.__init__(self)
		self.master.title("Jim Tracker")
		self.pack()

		self.customers = Customers()
		self.newc_diag = NewCustomerDialog(self, self.customers)
		self.logger_diag = LoggerDialog(self,self.customers)

		self.menubar = Menu(self)

		menu = Menu(self.menubar, tearoff=0)
		self.menubar.add_cascade(label="File", menu=menu)
		menu.add_command(label="Set Data Folder")
		menu.add_command(label="Quit")

		menu = Menu(self.menubar, tearoff=0)
		self.menubar.add_cascade(label="Help", menu=menu)
		menu.add_command(label="Help")
		menu.add_command(label="About")
		self.master.config(menu=self.menubar)
		
		cust_group = LabelFrame(self,text="Customer")
		cust_group.pack(padx=50, pady=10)
		Button(cust_group,text="Check In",width=15,command=self.check_in).pack(padx=5, pady=5)

		admin_group = LabelFrame(self,text="Admin")
		admin_group.pack(padx=10, pady=10)
		Button(admin_group,text="New Customer",width=15,command=self.new_customer).pack(padx=5, pady=5)
		Button(admin_group,text="New Payment",width=15).pack(padx=5, pady=5)

		report_group = LabelFrame(self,text="Report")
		report_group.pack(padx=10, pady=10)
		Label(report_group,text="Year",width=7).grid(row=0,column=0,pady=5,padx=5)
		Label(report_group,text="Month",width=7).grid(row=1,column=0,pady=5,padx=5)
		Combobox(report_group,width=10).grid(row=0,column=1,pady=5,padx=5)
		Combobox(report_group,width=10).grid(row=1,column=1,pady=5,padx=5)
		Button(report_group,text="Save Report",width=15).grid(row=2,columnspan=2,pady=5,padx=5)

		self.bind('<<NewCustomer>>',self.new_customer)

	def check_in(self):
		self.logger_diag.show()

	def new_customer(self, event=None):
		#if evert this came as a new customer entry through check in
		if event:
			temp = self.logger_diag.name.get().split(' ')
			self.newc_diag.fname.set(temp[0])
			if len(temp) == 2:
				self.newc_diag.lname.set(temp[1])
			elif len(temp) == 3:
				self.newc_diag.mname.set(temp[1])
				self.newc_diag.lname.set(temp[2])
			elif len(temp) > 3:
				self.newc_diag.mname.set(temp[1])
				self.newc_diag.lname.set(' '.join(temp[2:4]))

		self.newc_diag.show()

		name = ' '.join([self.newc_diag.fname.get(),self.newc_diag.mname.get(),
			self.newc_diag.lname.get()])

		## TODO: Ask "Enter first payment?"

		# clean up new customer dialog for next time lname value
		self.newc_diag.fname.set('')
		self.newc_diag.lname.set('')
		self.newc_diag.mname.set('')

if __name__ == '__main__':
	TrackerWindow().mainloop()