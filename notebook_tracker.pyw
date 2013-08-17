'''
ttk.notebook ui

'''

#tkinter
from ttk import Notebook,Frame,Label
from Tkinter import Text,Menu
from ScrolledText import ScrolledText
#jim tracker
from customer import CustomerFrame, Customers
from payment import PaymentFrame, Payments
from log_data import CheckInFrame

class JimNotebook(Frame):
    def __init__(self, name='notebookdemo'):
        Frame.__init__(self, name=name)
        self.pack(expand=True, fill='both')
        self.master.title('Jim Tracker')

        #variables
        self.customers = Customers()
        self.payments = Payments()

        #menu
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

        #notebook
        self.nb = Notebook(self, name='notebook')

        #frames
        self.ci_frame = CheckInFrame(self.nb, self.customers, self.payments)
        self.pt_frame = PaymentFrame(self.nb, self.customers, self.payments, self.output_text, self.refresh)
        self.cu_frame = CustomerFrame(self.nb, self.customers, self.output_text, self.refresh)

        #add to notebook
        self.nb.add(self.ci_frame, text="Check In")
        self.nb.add(self.cu_frame, text="Customers")
        self.nb.add(self.pt_frame, text="Payments")

        #pack notebook
        self.nb.pack(expand=True,fill='both',side='top')

        #output log
        stf = Frame(self)
        stf.pack(fill='x',side='top')
        self.scrolled_text = ScrolledText(stf,height=10,width=50,wrap='word',state='disabled')
        self.scrolled_text.pack(expand=True,fill='both')

    def output_text(self,outstr):
        self.scrolled_text['state'] = 'normal'
        self.scrolled_text.insert('end',outstr)
        self.scrolled_text.see('end')
        self.scrolled_text['state'] = 'disabled'

    def refresh(self):
        self.pt_frame.update_names() # update names in payments drop down boxes
        self.cu_frame.reset_values() # clear out the name value and reset date in new customer

if __name__ == '__main__':
    JimNotebook().mainloop()