'''
Payment objects for jim tracker

jim_payments<year>.xlsx (book)
    <month> (sheet)
        Monthly (column 0,1)
            Customer - 0
            Date - 1
        Punch Card (column 2,3,4)
            Customer - 2
            Date - 3
            Remaining - 4
        Drop In (column 5,6)
            Customer - 5
            Date - 6

TODO: forward dated entries
'''
#standard python
from time import strftime
from datetime import datetime
#openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.shared.exc import InvalidFileException
from openpyxl.style import Border, Fill
from openpyxl.cell import get_column_letter
#tkinter
from ttk import Frame,Label,Entry,Combobox,LabelFrame,Button
from Tkinter import StringVar
#jim tracker
from customer import Customers

class PaymentFrame(Frame):
    """docstring for PaymentFrame"""
    def __init__(self, master, customers, payments, output_text, refresh):
        Frame.__init__(self, master)
        self.refresh = refresh
        self.master = master
        self.output_text = output_text
        self.customers = customers
        self.payments = payments
        self.pname = StringVar()
        self.pnames = []
        self.mname = StringVar()
        self.mnames = []
        self.date = StringVar()
        self.date.set(strftime("%m/%d/%Y"))

        # Monthly Customers
        monthly_lf = LabelFrame(self, text="Monthly Customers Payment")
        monthly_lf.pack(padx=5,pady=5,ipadx=5,ipady=5,side='top')
        
        Label(monthly_lf,text="Name:").grid(row=0,column=0,sticky='e',padx=10)
        Label(monthly_lf,text="Date:").grid(row=0,column=2,sticky='e',padx=10)

        self.mname_cb = Combobox(monthly_lf,textvariable=self.mname,width=20,values=self.mnames,
            state='readonly')
        self.mname_cb.grid(row=0,column=1,sticky='w')

        Entry(monthly_lf,textvariable=self.date).grid(row=0,column=3,sticky='w')

        Button(monthly_lf,text='Reset Values',width=15).grid(row=3,column=0,columnspan=2,sticky='w',padx=10,pady=3)
        Button(monthly_lf,text='Submit',width=15,command=self.monthly_payment).grid(row=3,column=3,sticky='e')

        # Punch Card Customers
        puch_lf = LabelFrame(self, text="Punch Card Customers (Purchace Card)")
        puch_lf.pack(padx=5,pady=5,ipadx=5,ipady=5,side='top')
        
        Label(puch_lf,text="Name:").grid(row=0,column=0,sticky='e',padx=10)
        Label(puch_lf,text="Date:").grid(row=0,column=2,sticky='e',padx=10)

        self.pname_cb = Combobox(puch_lf,textvariable=self.pname,width=20,values=self.pnames,
            state='readonly')
        self.pname_cb.grid(row=0,column=1,sticky='w')

        Entry(puch_lf,textvariable=self.date).grid(row=0,column=3,sticky='w')

        Button(puch_lf,text='Reset Values',width=15).grid(row=3,column=0,columnspan=2,sticky='w',padx=10,pady=3)
        Button(puch_lf,text='Submit',width=15,command=self.new_punchcard).grid(row=3,column=3,sticky='e')

        self.pack(padx=10,pady=10,expand=True,fill='both')

        self.update_names()

    def monthly_payment(self):
        try:
            self.payments.monthly_payment(self.mname.get())
            self.output_text("$ - Monthly Payment: " + self.mname.get() + "\n")
        except IOError:
            showerror("Error writting to file", "Please close " + self.payments.filename + " and press OK.")

    def new_punchcard(self):
        try:
            self.payments.new_punchcard(self.pname.get())
            self.output_text("$ - New Puncard: " + self.mname.get() + "\n")
        except IOError:
            showerror("Error writting to file", "Please close " + self.payments.filename + " and press OK.")

    def update_names(self):
        self.populate_names()
        self.mname_cb['values'] = self.mnames
        self.mname_cb.current(0)
        self.pname_cb['values'] = self.pnames
        self.pname_cb.current(0)
        
    def populate_names(self):
        # try:
        clist = self.customers.get_list()
        # except IOError:
        #     self.output_text("! - " + self.customers.filename + " open in another application.\n")
        #     return
        clist.sort(key = lambda x: ', '.join(x[0:3]).lower())
        self.mnames = [' '.join([line[1],line[2],line[0]]) for line in clist if line[3]=='Monthly']
        self.pnames = [' '.join([line[1],line[2],line[0]]) for line in clist if line[3]=='Punch Card']


#jim_tracker
class Payments:
    def __init__(self):
        '''
        Constructor for Payments.
        Throws IOError if file is open in another application.
        '''
        self.month = strftime("%B")
        self.year = strftime("%Y")
        self.highest_row = 0
        self.open_workbook()
        self.open_sheet()
        
    def open_workbook(self):
        '''
        helper method for initalization
        Throws IOError if file is open in another application.
        '''
        self.filename = 'jim_payments' + str(self.year) + '.xlsx'
        try:
            self.wb = load_workbook(self.filename)
        except InvalidFileException: #create new file!
            self.wb = Workbook()
            sh = self.wb.get_sheet_by_name("Sheet")
            self.wb.remove_sheet(sh)

    def open_sheet(self):
        '''
        helper method for initalization
        '''
        self.sh = self.wb.get_sheet_by_name(self.month)
        if not self.sh: # new month!
            self.wb.create_sheet(title=self.month)
            self.sh = self.wb.get_sheet_by_name(self.month)
            self.sh.append(['Monthly', '', 'Punch Card', '', '', 'Drop In', ''])
            self.sh.append(['Customer', 'Date', 'Customer', 'Date', 'Remaining', 'Customer', 'Date'])
            for row in range(2):
                for col in range(7):
                    cell = self.sh.cell(row=row, column=col).style
                    cell.fill.fill_type = Fill.FILL_SOLID
                    cell.fill.start_color.index = "FFDDD9C4"
                    cell.borders.bottom.border_style = Border.BORDER_THIN

            # set column widths
            column_widths = [20,12,20,12,12,20,12]
            for i, column_width in enumerate(column_widths):
                self.sh.column_dimensions[get_column_letter(i+1)].width = column_width
        
            # auto card import
            today = datetime.today()
            old = None
            if today.month == 1: #january
                old = datetime(today.year-1, 12, 1)
            else: # other months
                old = datetime(today.year, today.month-1, 1)

            self.update_cards(old.strftime("%B"), old.strftime("%Y"))

        self.format_save()

    def monthly_payment(self, customer, date = datetime.today()):
        '''
        enters a new monthly Payment
        '''
        #find the next empty line (row value)
        row = 2
        cust_column = 0
        while self.sh.cell(row=row,column=cust_column).value != None:
            row += 1

        self.sh.cell(row=row,column=0).value = customer
        self.sh.cell(row=row,column=1).value = date
        self.sh.cell(row=row,column=1).style.number_format.format_code = 'm/d/yyyy'

        self.format_save()

    def new_punchcard(self, customer, date = datetime.today(), punches = 10):
        '''
        creates a new punch card entry
        '''
        #find the next empty line (row value)
        row = 2
        cust_column = 2
        while self.sh.cell(row=row,column=cust_column).value != None:
            row += 1

        self.sh.cell(row=row,column=2).value = customer
        self.sh.cell(row=row,column=3).value = date
        self.sh.cell(row=row,column=3).style.number_format.format_code = 'm/d/yyyy'
        self.sh.cell(row=row,column=4).value = punches
        self.sh.cell(row=row,column=4).style.number_format.format_code = '0'

        self.format_save()

    def punch(self, customer):
        '''
        punches a punchcard 
        1. finds punch card with remaining punches
        2. punches
        3. returns remaining punchs

        returns None if card does not exist or no remaining punches
        '''
        #find the punchcard (customer) with remaining punches
        row = 2
        cust_column = 2
        punch_column = 4
        for row in range(self.highest_row):
            if self.sh.cell(row=row,column=cust_column).value == customer:
                if int(self.sh.cell(row=row,column=punch_column).value) > 0:
                    break
            if self.sh.cell(row=row,column=cust_column).value == None: # not found
                return None
            
        punch = int(self.sh.cell(row=row,column=punch_column).value) - 1
        self.sh.cell(row=row,column=punch_column).value = punch

        self.format_save()

        return punch

    def drop_in(self, customer, date = datetime.today()):
        '''
        enters a drop in Payment
        '''
        #find the next empty line (row value)
        row = 2
        cust_column = 5
        while self.sh.cell(row=row,column=cust_column).value != None:
            row += 1

        self.sh.cell(row=row,column=5).value = customer
        self.sh.cell(row=row,column=6).value = date
        self.sh.cell(row=row,column=6).style.number_format.format_code = 'm/d/yyyy'

        self.format_save()

    def update_cards(self, from_month, year = None):
        '''
        copies in old punches from previous month
        can copy in from old years (different filename)

        returns None if month or year does not exist
        returns number of cards copied otherwise
        '''
        if not year:
            year = self.year

        cards = []

        if year != self.year: #different year
            from_filename = 'jim_payments' + str(year) + '.xlsx'
            try:
                from_wb = load_workbook(from_filename)
            except InvalidFileException: #file does not exist
                return None
        else:
            from_wb = self.wb

        from_sh = from_wb.get_sheet_by_name(from_month)
        if not from_sh: # month does not exist
            return None

        #get the data
        row = 2
        cust_column = 2
        date_column = 3
        punch_column = 4
        while from_sh.cell(row=row,column=cust_column).value != None:
            if from_sh.cell(row=row,column=punch_column).value != 0:
                cards.append([from_sh.cell(row=row,column=cust_column).value,
                    from_sh.cell(row=row,column=date_column).value,
                    from_sh.cell(row=row,column=punch_column).value])
            row += 1

        # add the values to current sheet (tests if they exist first in current sheet (date validation))
        current_cards = {}
        row = 2
        while self.sh.cell(row=row,column=cust_column).value != None:
            current_cards[self.sh.cell(row=row,column=cust_column).value] = self.sh.cell(row=row,column=date_column).value
            row += 1
        
        cards_copied = 0
        for card in cards:
            if current_cards.haskey(card[0]):
                if current_card[card[0]] == card[1]: #same punch card
                    continue
            self.new_punchcard(card[0],card[1],card[2])
            cards_copied += 1

        if cards_copied:
            self.format_save()

        return cards_copied

    def has_paid_monthly(self, customer):
        '''
        Returns true if customer has paid monthly dues
        '''
        row = 2
        cust_column = 0
        while self.sh.cell(row=row,column=cust_column).value != None:
            if customer == self.sh.cell(row=row,column=cust_column).value:
                return True
            row += 1

        return False

    def format_save(self):
        self.sh.garbage_collect()
        self.format_borders()
        self.wb.save(self.filename)

    def format_borders(self):
        columns = [1, 4, 6]
        for row in range(self.highest_row, self.sh.get_highest_row()):
            for col in columns:
                cell = self.sh.cell(row=row, column=col).style
                cell.borders.right.border_style = Border.BORDER_THIN

        self.highest_row = self.sh.get_highest_row()

def test1():
    p = Payments()

    monthly_customers = ['Tyler J Weaver', 'Marcus T Weaver']
    p.monthly_payment(monthly_customers[0])
    p.monthly_payment(monthly_customers[1])

    # for c in monthly_customers:
    #     print c
    #     if p.has_paid_monthly(c): 
    #         print "Paid"
    #     else:
    #         print "Unpaid"

    # punch_card_customers = ['Brad Bradley', 'Sam P Frank']

    # for c in punch_card_customers:
    #     p.new_punchcard(c)

    # for x in range(4):
    #     print punch_card_customers[0], p.punch(punch_card_customers[0])

    # for x in range(7):
    #     print punch_card_customers[1], p.punch(punch_card_customers[1])

    # drop_in_customers = ['Tom R Jones', 'Rodgers P Smith']

    # p.drop_in(drop_in_customers[0])
    # p.drop_in(drop_in_customers[1])

def output_text(text):
    print text

def refresh():
    print "refresh..."

if __name__ == '__main__':
    root = Frame()
    root.pack()
    PaymentFrame(root, Customers(), Payments(), output_text, refresh).pack()
    root.mainloop()
    # test1()