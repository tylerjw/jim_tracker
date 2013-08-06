'''
ttk.notebook ui

'''

from ttk import Notebook,Frame,Label
from Tkinter import Text,Menu
#jim tracker
from customer_frame import NewCustomerFrame, Customers
from payment import PaymentFrame, Payments

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

        #admin frame
        admin_frame = Frame(self.nb, name='admin')
        #new customer
        nc_frame = NewCustomerFrame(admin_frame, self.customers)
        nc_frame.pack(padx=5,pady=5,ipadx=30,ipady=5,fill='x')
        #payment
        pt_frame = PaymentFrame(admin_frame, self.customers, self.payments)
        nc_frame.pack(padx=5,pady=5,ipadx=5,ipady=5)
        admin_frame.pack()
        self.nb.add(admin_frame, text="Admin",sticky='ew')

        #pack notebook
        self.nb.pack()

if __name__ == '__main__':
    JimNotebook().mainloop()