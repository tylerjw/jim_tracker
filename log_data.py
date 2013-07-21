'''
Log Data - dialog for logging workouts

Dependencies:
    Python 2.7.5
    Tkinter 8.5.3

Author: Tyler Weaver

Revision History:
(20 June 2013)
    Logger class
(25 June 2013)
    Initial GUI design
(30 June 2013)
    Remove email field for middle name
(01 July 2013)
    Update customer list after new customer has been added.
    Accomidate for suffixes (Jr, III, etc.)
(02 July 2013)
    Add date field
(03 July 2013)
    Disable/Enable date text field
    Fixed set workout bug

TODO:
    Automatically sellect correct workout.
    Allow setting date to other than today.
    
'''

from Tkinter import StringVar,E,W
from ttk import Frame, Button, Entry, Label, Combobox
from tkMessageBox import showerror
from datetime import datetime, date, timedelta
from customer import NewCustomerDialog, Customers
from schedule import Schedule

from openpyxl import Workbook, load_workbook
from openpyxl.style import Border, Fill
from time import strftime
from datetime import date, time

from pprint import pprint

class LoggerWindow(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.master.title("Log Workout")
        self.pack()

        self.name = StringVar() # variable for customer
        self.customers = Customers() # customers object
        self.names = []
        self.populate_names()
        self.workout = StringVar()
        self.workouts = []
        self.workouts_form = []
        self.schedule = Schedule()
        self.date = StringVar()
        self.date.set(strftime("%m/%d/%Y"))
        self.populate_workouts()
        self.newc_diag = NewCustomerDialog(self, self.customers)
        self.logger = Logger()
        self.refresh_time = 15 # in minutes
        self.output = '' # for the output label at the bottom

        inf = Frame(self)
        inf.pack(padx=10,pady=10)
        Label(inf, text="Name:").grid(row=0,column=0,sticky=E,ipady=2,pady=2,padx=10)
        Label(inf, text='Date:').grid(row=1,column=0,sticky=E,ipady=2,pady=2,padx=10)
        Label(inf, text="Workout:").grid(row=2,column=0,sticky=E,ipady=2,pady=2,padx=10)

        self.name_cb = Combobox(inf, textvariable=self.name, width=30,
                                values=self.names)
        self.name_cb.grid(row=0,column=1,sticky=W,columnspan=2)
        self.date_ent = Entry(inf, textvariable=self.date)
        self.date_ent.grid(row=1,column=1,sticky=W)
        self.date_ent.bind('<FocusOut>', self.update_workouts)
        Button(inf,text='Edit', command=self.enable_date_ent).grid(row=1,column=2,sticky=E)
        self.workout_cb = Combobox(inf, textvariable=self.workout, width=30,
                                   values=self.workouts_form)
        self.workout_cb.grid(row=2,column=1,sticky=W,columnspan=2)

        self.log_btn=Button(inf,text="Log Workout",command=self.log)
        self.log_btn.grid(row=3,column=0,columnspan=2,pady=4)
        self.out_ent = Label(inf, text=self.output)
        self.out_ent.grid(row=4,column=0,columnspan=3,pady=4)

        self.master.bind('<Return>',self.log)
        self.name_cb.focus_set()  # set the focus here when created

        #disable the date field
        self.disable_date_ent()

        #start time caller
        self.time_caller()

    def log(self, e=None):
        #check to see if name is blank
        if self.name.get() == '':
            showerror("Error!", "Please enter or select a name.")
        elif self.workout.get() not in self.workouts_form:
            showerror("Error!", "Please select a workout from the list.")
        elif self.name.get() not in self.names: # new customer
            #parse name and preset values in the new customer dialog
            temp = self.name.get().split(' ')
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
            # clean up new customer dialog lname value
            # note, fname will always get updated
            self.newc_diag.lname.set('')
            self.newc_diag.mname.set('')

            #update the names list
            self.update_names()
            
        else: # log the workout
            self.logger.log(self.workouts[self.workout_cb.current()][0],
                            self.workouts[self.workout_cb.current()][1],
                            self.name.get(), day=datetime.strptime(str(self.date.get()),'%m/%d/%Y'))

            self.update_time_now()
            self.set_workout_now()
            self.update_workouts()
            

    def disable_date_ent(self, e=None):
        self.date_ent['state'] = 'disabled'

    def enable_date_ent(self, e=None):
        self.date_ent['state'] = 'normal'
        
    def time_caller(self):
        #updates every 15 min automatically
        msec = self.refresh_time * 60 * 100

        self.update_time_now() #update time to current time
        self.set_workout_now()
        self.update_workouts() #update the workouts
        
        self.after(msec, self.time_caller) #call again

    def update_time_now(self):
        self.enable_date_ent()
        self.date.set(strftime("%m/%d/%Y"))

    def set_workout_now(self):
        #set workout field
        if len(self.workouts) == 0:
            self.disable_date_ent()
            return #no workouts
        index = 0
        now = datetime.today()
        for i, workout in enumerate(self.workouts):
            test = datetime.combine(date.today(),workout[0])
            if now < (test - timedelta(minutes=30)):
                index = i
                break
        self.workout_cb.current(index)
        self.disable_date_ent()
            
    def update_workouts(self, e=None):
        try:
            self.populate_workouts()
            self.workout_cb['values'] = []
            self.workout_cb['values'] = self.workouts_form
        except ValueError:
            self.workout.set(' Enter Valid Date ')
        if len(self.workouts) > 0 and e:
            self.workout_cb.current(0)
            
    def populate_workouts(self):
        today = datetime.strptime(str(self.date.get()), "%m/%d/%Y") #get date
        dow = self.schedule.weekday_to_str(today.weekday()) #get dow string

        self.workouts = self.schedule.get_wkday(dow)
        self.workouts_form = []
        for w in self.workouts:
            self.workouts_form.append(w[0].strftime("%H:%M") + ' - ' + w[1])
        if len(self.workouts) == 0:
            self.workout.set(' No workouts today ')

    def update_names(self):
        self.populate_names()
        self.name_cb['values'] = self.names
        
    def populate_names(self):
        clist = self.customers.get_list()
        clist.sort(key = lambda x: ', '.join(x[0:3]).lower())
        self.names = []
        for line in clist:
            self.names.append(' '.join([line[1],line[2],line[0]]))

    def find_line(self, name):
        [fname, mname, lname] = name.split(' ')
        return self.customers.find(lname, fname, mname)

class Logger:
    def __init__(self, filename='jim_data.xlsx'):
        self.month = strftime("%B")
        self.wb = load_workbook(filename)
        self.filename = filename
        self.sh = self.wb.get_sheet_by_name(self.month)
        if not self.sh:
            self.wb.create_sheet(title=self.month)
            print "Created new month log: " + self.month
            self.sh = self.wb.get_sheet_by_name(self.month)
            self.sh.append(['Date','Time','Class Type','Customer'])
            for col in range(4):
                cell = self.sh.cell(row=0, column=col).style
                cell.fill.fill_type = Fill.FILL_SOLID
                cell.fill.start_color.index = "FFDDD9C4"
                cell.borders.bottom.border_style = Border.BORDER_THIN

        self.sh.garbage_collect()
        self.wb.save(self.filename)

    def log(self, hour, class_type, customer, day=date.today()):
        line = [day, hour.strftime('%H:%M'), class_type, customer]
        self.sh.garbage_collect()
        self.sh.append(line)
        self.sh.cell(row=(self.sh.get_highest_row()-1),
                         column=0).style.number_format.format_code='m/d/yyyy'
        self.wb.save(self.filename)
        

if __name__ == '__main__':
    LoggerWindow().mainloop()
