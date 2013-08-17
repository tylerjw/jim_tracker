""""
Dialog base class.  Example of use below...

Author: Tyler Weaver
Date: 23 June 2013

(25 June 2013)
    Added ErrorDialog
"""

from Tkinter import Toplevel
from ttk import Button, Frame, Label

class Dialog:
    def __init__(self, master, title, class_=None, relx=0.5, rely=0.3):
        self.master = master
        self.title = title
        self.class_ = class_
        self.relx = relx
        self.rely = rely

    def setup(self):
        if self.class_:
            self.root = Toplevel(self.master, class_=self.class_)
        else:
            self.root = Toplevel(self.master)

        self.root.title(self.title)
        self.root.iconname(self.title)

    def change_title(self,title):
        self.title = title
        self.root.title(self.title)
        self.root.iconname(self.title)

    def enable(self):
        ### enable
        self.root.protocol('WM_DELETE_WINDOW', self.wm_delete_window)
        self._set_transient(self.relx, self.rely)
        self.root.wait_visibility()
        self.root.grab_set()
        self.root.mainloop()
        self.root.destroy()

    def _set_transient(self, relx=0.5, rely=0.3):
        widget = self.root
        widget.withdraw() # Remain invisible while we figure out the geometry
        widget.transient(self.master)
        widget.update_idletasks() # Actualize geometry information
        if self.master.winfo_ismapped():
            m_width = self.master.winfo_width()
            m_height = self.master.winfo_height()
            m_x = self.master.winfo_rootx()
            m_y = self.master.winfo_rooty()
        else:
            m_width = self.master.winfo_screenwidth()
            m_height = self.master.winfo_screenheight()
            m_x = m_y = 0
        w_width = widget.winfo_reqwidth()
        w_height = widget.winfo_reqheight()
        x = m_x + (m_width - w_width) * relx
        y = m_y + (m_height - w_height) * rely
        if x+w_width > self.master.winfo_screenwidth():
            x = self.master.winfo_screenwidth() - w_width
        elif x < 0:
            x = 0
        if y+w_height > self.master.winfo_screenheight():
            y = self.master.winfo_screenheight() - w_height
        elif y < 0:
            y = 0
        widget.geometry("+%d+%d" % (x, y))
        widget.deiconify() # Become visible at the desired location

    def wm_delete_window(self):
        self.root.quit() 

class ErrorDialog(Dialog):
    def __init__(self, master, title, text, class_=None, relx=0.5, rely=0.3):
        Dialog.__init__(self, master, title,
                        class_, relx, rely)

        self.text = text

    def show(self):
        self.setup()

        #contents of dialog
        f = Frame(self.root)
        Label(f, text=self.text, width=40).pack()
        Button(f, text='Close', width=10, command=self.close).pack(pady=15)
        f.pack(padx=5, pady=5)

        self.root.bind("<Return>", self.close)

        self.enable()

    def close(self, event=None):
        self.wm_delete_window()
        
class TestDialog(Dialog):
    def __init__(self, master, class_=None, relx=0.5, rely=0.3):
        Dialog.__init__(self, master, "Test Dialog Title",
                        class_, relx, rely)

        #initialize variables...

    def show(self):
        self.setup()

        #contents of dialog
        f = Frame(self.root)
        Label(f, text='Warning, Testing Dialog Boxes!', width=40).pack()
        Button(f, text='Close', width=10, command=self.wm_delete_window).pack(pady=15)
        f.pack(padx=5, pady=5)

        self.enable()

if __name__ == '__main__':
    root = Frame()
    test = TestDialog(root)
    error = ErrorDialog(root, "Error Dialog!!", "You did something bad, \nshame on you")
    Button(root,text='Test Dialog',command=test.show).pack()
    Button(root,text='Error Dialog',command=error.show).pack()
    root.pack()
    root.mainloop()
