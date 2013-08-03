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
'''
#standard python
from time import strftime
#openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.shared.exc import InvalidFileException
from openpyxl.style import Border, Fill

class Payments:
    def __init__(self):
        '''
        Constructor for Payments.
        Throws IOError if file is open in another application.
        '''
        self.month = strftime("%B")
        self.year = strftime("%Y")
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
                    cell = self.sh.cell(row=0, column=col).style
                    cell.fill.fill_type = Fill.FILL_SOLID
                    cell.fill.start_color.index = "FFDDD9C4"
                    cell.borders.bottom.border_style = Border.BORDER_THIN

            # auto card import
            today = datetime.today()
            old = None
            if today.month == 1: #january
                old = datetime(today.year-1, 12, 1)
            else: # other months
                old = datetime(today.year, today.month-1, 1)

            self.update_cards(old.strftime("%B"), old.strftime("%Y"))

        self.sh.garbage_collect()
        self.wb.save(self.filename)

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

        self.sh.garbage_collect()
        self.wb.save(self.filename)

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
        self.sh.cell(row=row,column=4).value = punches

        self.sh.garbage_collect()
        self.wb.save(self.filename)

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
        while (self.sh.cell(row=row,column=cust_column).value != customer and
            self.sh.cell(row=row,column=punch_column).value != 0:
            if self.sh.cell(row=row,column=cust_column) == None: # not found
                return None
            row += 1
            
        self.sh.cell(row=row,column=punch_column).value -= 1
        return self.sh.cell(row=row,column=punch_column).value

        self.sh.garbage_collect()
        self.wb.save(self.filename)

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

        self.sh.garbage_collect()
        self.wb.save(self.filename)

    def update_cards(self, from_month, year = self.year):
        '''
        copies in old punches from previous month
        can copy in from old years (different filename)

        returns None if month or year does not exist
        returns number of cards copied otherwise
        '''
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
            self.sh.garbage_collect()
            self.wb.save(self.filename)

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