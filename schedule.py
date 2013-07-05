from openpyxl import Workbook, load_workbook
from datetime import time

class Schedule:
    def __init__(self, filename='jim_info.xlsx', sheet_name='Schedule'):
        self.wb = load_workbook(filename)
        self.filename = filename
        self.sh = self.wb.get_sheet_by_name(sheet_name)
        if not self.sh:
            print "Error opening " + sh_name + " sheet."

        self.weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

    def weekday_to_str(self, idx):
        return self.weekdays[idx]

        
    def get_dict(self):
        output = {}
        max_row = self.sh.get_highest_row()
        range_str = "A1:B"+str(max_row)
        
        for i, col in enumerate(range(0,14,2)):
            output[self.weekdays[i]] = {}
            times = self.sh.range(range_str,column=col)
            for row in range(1,max_row):
                if times[row][0].value  == None:
                    continue
                output[self.weekdays[i]][times[row][0].value] = str(times[row][1].value)

        return output

    def get_wkday(self, wkday):
        full_list = self.get_dict()
        short_list = full_list[wkday]
        output = []
        for key in short_list:
            output.append([key, short_list[key]])
        output.sort(key=lambda x: x[0])
        return output

if __name__ == '__main__':
    s = Schedule()
    
    for i in range(7):
        print s.weekdays[i]
        print(s.get_wkday(s.weekdays[i]))
