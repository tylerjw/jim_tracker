'''
admin frame for jim tracker notebook
'''
from reports import find_years_months,generate_info_file
from customer import Customers 
from payment import Payments 

from ttk import Frame,LabelFrame,Label,Combobox,Button,Entry
from Tkinter import Tk,StringVar
import tkFileDialog

from os import chdir, getcwd, listdir
from time import strftime

class AdminFrame(Frame):
    def __init__(self,master,customers,payments,output_text,refresh,root_dir):
        Frame.__init__(self,master)
        self.customers = customers
        self.payments = payments
        self.output_text = output_text
        self.refresh = refresh
        self.root_dir = root_dir

        self.year_months = find_years_months(getcwd())

        self.year = StringVar()
        self.month = StringVar()
        self.years = sorted(self.year_months.keys())
        self.months = []

        self.root_directory = StringVar()
        self.root_directory.set(getcwd())

        self.columnconfigure(0,weight=1)

        lf = LabelFrame(self,text="Copy Punch Cards into Current Month")
        lf.grid(padx=5,pady=5,row=0,column=0,sticky='ew')
        lbl = Label(lf,text="Select month to copy from: ")
        lbl.grid(row=0,column=0,columnspan=4,sticky='w',padx=10,pady=(10,2))
        Label(lf,text="Year: ").grid(row=1,column=0,sticky='e',padx=(10,0),pady=2)
        self.year_cb = Combobox(lf,textvariable=self.year,width=12,values=self.years,state='readonly')
        self.year_cb.grid(row=1,column=1,sticky='ew',pady=2)
        Label(lf,text="Month: ").grid(row=1,column=2,sticky='e',pady=2)
        self.month_cb = Combobox(lf,textvariable=self.month,width=12,values=self.months,state='readonly')
        self.month_cb.grid(row=1,column=3,sticky='ew',padx=(0,10),pady=2)
        Button(lf,text="Copy",command=self.copy).grid(row=2,column=3,sticky='ew',padx=(0,10),pady=(2,10))
        for i in range(4):
            lf.columnconfigure(i,weight=1)

        lf = LabelFrame(self,text="File Management")
        lf.grid(padx=5,pady=5,row=1,column=0,sticky='ew')
        lbl = Label(lf,text="Data Directory:")
        lbl.grid(row=0,column=0,sticky='w',padx=10,pady=(10,2))
        Entry(lf,textvariable=self.root_directory,state='readonly').grid(row=1,column=0,sticky='ew',padx=(10,0),pady=2)
        Button(lf,text="Browse",command=self.browse).grid(row=1,column=1,sticky='e',padx=(0,10),pady=2)
        btn = Button(lf,text="Generate Schedule/Customer Info File",command=self.generate)
        btn.grid(row=2,column=0,columnspan=2,sticky='ew',padx=10,pady=(20,10))
        for i in range(1):
            lf.columnconfigure(i,weight=1)

        self.year_cb.bind('<<ComboboxSelected>>',self.year_selected)

        self.update() #update the values

    def generate(self):
        files = listdir(getcwd())
        if "jim_info.xlsx" in files:
            self.output_text("! - jim_info.xlsx already in working directory.")
        else:
            generate_info_file()
            self.output_text("A - Generated empty jim_info.xlsx file into working directory.")
            self.refresh()

    def browse(self):
        '''
        change the current working directory
        '''
        directory = tkFileDialog.askdirectory(parent=self,initialdir=getcwd())
        if directory != '': 
            try:
                fh = open(self.root_dir + '/config.ini', 'w')
                fh.write(directory)
                fh.close()
            except IOError:
                self.output_text("! - config.ini file open in another program, close and try again.")
            else:
                # user selected one, not canceled, and config file not open
                self.root_directory.set(directory)
                chdir(directory)
                self.refresh()

                #check for jim_info file
                files = listdir(getcwd())
                if "jim_info.xlsx" not in files:
                    self.output_text("! - jim_info.xlsx not found in working directory.\n")
                    self.generate()


    def copy(self):
        if self.month.get() is '':
            self.output_text("! - Select a month to copy from\n")
        else:
            cards = self.payments.update_cards(self.month.get(),self.year.get())
            if cards:
                self.output_text("A - Copied " + str(cards) + " Punch Cards from " + self.month.get() + ' ' + self.year.get() + '\n')
            else:
                # no cards copied
                self.output_text("! - No Punch Cards to copy from " + self.month.get() + ' ' + self.year.get() + '\n')

    def year_selected(self,event=None):
        '''
        run when year is year is selected 
        copied from reprots setup
        '''
        self.months = self.year_months[self.year.get()]
         # don't copy from the existing month !
        if self.year.get() == strftime("%Y"):
            try:
                self.months.remove(strftime("%B"))
            except:
                pass #it'll throw an exception if the month isn't there, that's ok

        self.month_cb['values'] = self.months
        self.month_cb.current(0)

    def update(self):
        '''
        method for updating values when things change
        copied from reports setup
        '''
        self.year_months = find_years_months(getcwd()) # use cwd, this should be set
        self.years = sorted(self.year_months.keys())
        self.months = []
        self.year_cb['values'] = self.years
        self.month_cb['values'] = self.months
        self.root_directory.set(getcwd())

def output_text(text):
    print text,

def refresh():
    print "refreshing...."

if __name__ == '__main__':
    root = Tk()
    c = Customers()
    p = Payments()

    af = AdminFrame(root,c,p,output_text,refresh)
    af.grid(sticky='nsew')
    root.rowconfigure(0,weight=1)
    root.columnconfigure(0,weight=1)

    root.mainloop()