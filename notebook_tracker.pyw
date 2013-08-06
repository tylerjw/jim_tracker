'''
ttk.notebook ui

'''

from ttk import Notebook,Frame,Label
from Tkinter import Text,Menu
from customer_frame import NewCustomerFrame, Customers

class JimNotebook(Frame):
    def __init__(self, name='notebookdemo'):
        Frame.__init__(self, name=name)
        self.pack(expand=True, fill='both')
        self.master.title('Jim Tracker')

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
        self.customers = Customers()
        admin_frame = Frame(self.nb, name='admin')
        nc_frame = NewCustomerFrame(admin_frame, self.customers)
        nc_frame.pack(padx=10,pady=10)
        admin_frame.pack(padx=10,pady=10)
        f2 = Frame(self.nb, name='textbox')
        f2.pack()
        txt = Text(f2, wrap='word', width=40, height=10)
        txt.pack(fill='both', expand=True)
        self.nb.add(admin_frame, text="Admin")
        self.nb.add(f2, text="frame2")
        self.nb.pack()

if __name__ == '__main__':
    JimNotebook().mainloop()