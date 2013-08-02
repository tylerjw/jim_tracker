'''
Jim Tracker Application

Dependencies:
    Python 2.7.5
    Tkinter 8.5.3

Author: Tyler Weaver
'''

from log_data import LoggerWindow
from customer import NewCustomerDialog

from ttk import Frame, Label, Combobox, LabelFrame, Button
from tkMessageBox import showerror
from datetime import datetime

class TrackerWindow(Frame):
	def __init__(self):
		Frame.__init__(self)
		self.master.title("Jim Tracker")
		self.pack()
	
		cust_group = LabelFrame(self,text="Customer")
		cust_group.pack(padx=50, pady=10)
		Button(cust_group, text="Check In", width=15).pack(padx=5, pady=5)

		admin_group = LabelFrame(self,text="Admin")
		admin_group.pack(padx=10, pady=10)
		Button(admin_group, text="New Customer", width=15).pack(padx=5, pady=5)
		Button(admin_group, text="New Payment", width=15).pack(padx=5, pady=5)

		report_group = LabelFrame(self,text="Report")
		report_group.pack(padx=10, pady=10)
		Label(report_group,text="Year",width=7).grid(row=0,column=0,pady=5,padx=5)
		Label(report_group,text="Month",width=7).grid(row=1,column=0,pady=5,padx=5)
		Combobox(report_group,width=10).grid(row=0,column=1,pady=5,padx=5)
		Combobox(report_group,width=10).grid(row=1,column=1,pady=5,padx=5)
		Button(report_group,text="Save Report",width=15).grid(row=2,columnspan=2,pady=5,padx=5)


if __name__ == '__main__':
	TrackerWindow().mainloop()