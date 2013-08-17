'''
scripts for generating reports for jim tracker
'''

from openpyxl import Workbook, load_workbook
from openpyxl.style import Border, Fill
from openpyxl.cell import get_column_letter
from pprint import pprint
from calendar import Calendar
from time import strftime

from customer import Customers
from payment import Payments

def workouts_this_month(customer,log_file,month=strftime("%B")):
    '''
    laod the log file and count number of workouts done by customer
    '''
    wb = load_workbook(log_file)
    sh = wb.get_sheet_by_name(month)

    data = get_data(sh)
    count = len( [row for row in data if row[3] == customer] )
    return count

def get_data(sh):
    '''
    converts sheet into array of arrays (rows of columns)
    '''
    #get the data
    data = [[cell.value for cell in row] for row in sh.rows]
    return data

def month_report(log_file,month,output_file,customers,payments):
    #read the log
    wb = load_workbook(log_file)
    sh = wb.get_sheet_by_name(month)

    if not sh:
        # error - sheet doesn't exist
        return False

    #get the data
    data = get_data(sh)


    wb = Workbook()
    sh = wb.get_active_sheet()
    sh.title = "Customers"
    customers_report(data,sh,customers,payments)

    sh = wb.create_sheet(title='Classes')
    class_sheet = class_report(data,sh)
    
    #write new workbook
    
    wb.save(output_file)

def label_format(sh,columns,row=0,border='bottom'):
    #format top row
    for col in range(columns):
        cell = sh.cell(row=row, column=col).style
        cell.fill.fill_type = Fill.FILL_SOLID
        cell.fill.start_color.index = "FFDDD9C4"
        if border == 'bottom':
            cell.borders.bottom.border_style = Border.BORDER_THIN
        if border == 'top':
            cell.borders.top.border_style = Border.BORDER_THIN

def customers_report(data,sh,c,p):
    '''
    Build the Customers, # of workouts report
    data, log data 
    sh, output sheet
    c, customer class  
    p, payments class
    '''

    #create a dictionary of customers this month
    customers = dict.fromkeys(set([str(x[3]) for x in data[2:]]))
    for key in customers:
        customers[key] = [row for row in data if row[3] == key]

    report_data = [('Customer','# of Workouts','Type','Note')] + \
                  sorted([(key,str(len(value)),c.get_type(key),customer_note(key,c,p)) for (key,value) in customers.items() if c.get_type(key) != "Inactive"],
                         key=lambda pair: int(pair[1]), reverse=True)

    #write data
    for values in report_data:
        sh.append(values)
        color = None
        if values[3] is "Unpaid":
            color = "00ff0000"
        elif values[3] is "Paid":
            color = "0000ff00"

        if color:
            for col in range(4):
                cell = sh.cell(row=sh.get_highest_row()-1,column=col).style
                cell.fill.fill_type = Fill.FILL_SOLID
                cell.fill.start_color.index = color

    #format labels
    label_format(sh,4)
    # set column width

    for i, column_width in enumerate([25, 13, 18, 25]):
        sh.column_dimensions[get_column_letter(i+1)].width = column_width

def customer_note(name,customers,payments):
    '''
    name - customer name
    customers - Customers object
    payments - Payments object
    '''
    ctype = customers.get_type(name) # the customer type
    if ctype == "Customer Not Found":
        return "Add customer to info file."

    if ctype == "Monthly":
        if payments.has_paid_monthly(name):
            return "Paid"
        else:
            return "Unpaid"

    if ctype == "Punch Card":
        return "Remaining Punches: " + str(payments.get_remaining_punches(name))

    if ctype == "Inactive":
        return "Error!"

    return ""


def class_report(data,sh):
    dates = map(lambda x: x.date(), sorted(list(set([x[0] for x in data[2:]]))))
    report_data = dict.fromkeys(dates)
    for day in report_data:
        report_data[day] = dict.fromkeys(set([(x[1],x[2]) for x in data[2:] if x[0].date()==day]))
        for workout in report_data[day]:
            report_data[day][workout] = len([x for x in data[2:] if (x[0].date()==day and (x[1],x[2])==workout)])

    weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    line = []
    for wkd in weekdays:
        line += [wkd,'Class','# Att']
    sh.append(line)
    label_format(sh,len(line))
    first = True
    line = []
    workouts = {}
    for day in Calendar().itermonthdates(dates[0].year,dates[1].month):
        if day.weekday() == 0 and not first:
            sh.append(line)
            label_format(sh,len(line),sh.get_highest_row()-1,'top')
            if workouts:
                row = sh.get_highest_row()
                for wkd in workouts:
                    for i,workout in enumerate(workouts[wkd]):
                        sh.cell(row=row+i,column=wkd*3).value = workout[0].strftime("%H:%M")
                        sh.cell(row=row+i,column=wkd*3+1).value = workout[1]
                        sh.cell(row=row+i,column=wkd*3+2).value = workout[2]
            line = []
            workouts = {}
            
        if day in dates:
            line.append(day.day)
            line.append('')
            line.append('')
            date_data = [day.day,'','']
            for (time,workout),num in report_data[day].items():
                if workouts.has_key(day.weekday()):
                    workouts[day.weekday()].append([time,workout,num])
                else:
                    workouts[day.weekday()] = [[time,workout,num]]
        else:
            line.append(day.day)
            line.append('')
            line.append('')

        first = False
    sh.append(line)
    label_format(sh,len(line),sh.get_highest_row()-1,'top')
    if workouts:
        row = sh.get_highest_row()+1
        for wkd in workouts:
            for i,workout in enumerate(workouts[wkd]):
                sh.cell(row=row+i,column=wkd*3).value = workouts[wkd][0].strftime("%H:%M")
                sh.cell(row=row+i,column=wkd*3+1).value = workouts[wkd][1]
                sh.cell(row=row+i,column=wkd*3+2).value = workouts[wkd][2]

    for col in range(0,sh.get_highest_column()+1,3):
        for row in range(sh.get_highest_row()):
            cell = sh.cell(row=row, column=col).style
            cell.borders.left.border_style = Border.BORDER_THIN

    auto_column_width(sh)
                

def auto_column_width(worksheet):
    raw_data = worksheet.range(worksheet.calculate_dimension())
    data = [[str(x.value) for x in row] for row in raw_data]
    column_widths = []
    for row in data:
        for i, cell in enumerate(row):
            if len(column_widths) > i:
                if len(cell) > column_widths[i]:
                    column_widths[i] = len(cell)
            else:
                column_widths += [len(cell)]

    for i, column_width in enumerate(column_widths):
        worksheet.column_dimensions[get_column_letter(i+1)].width = column_width

if __name__ == '__main__':
    inputf = 'jim_data2013.xlsx'
    month = 'August'
    c = Customers()
    p = Payments()
    outputf = month + '_report.xlsx'
    month_report(inputf,month,outputf,c,p)
    # print workouts_this_month("Dave L Sanders", inputf, month) 