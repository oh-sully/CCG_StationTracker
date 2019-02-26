import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import date, timedelta
from excel_stylings import wb, Colors, etStyles, dtfStyles

def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + timedelta(days=4)
    return next_month - timedelta(days=next_month.day)

#Calculates the number of months that occur during these two dates
def months_between(start_date, end_date):
    return (end_date.year - start_date.year) * 12 + end_date.month - start_date.month + 1

#Writes the day and date for the month of the given beginning date from begin to end of month
def writeDates(sheet, begin, end, row_start, col, fmtArr, count, port):

    days_left = (end - begin).days
    fmt = fmtArr[0]
    if days_left < 10:
        sheet.merge_range(row_start, col, row_start + days_left, col, begin.strftime('%B')[0:4].upper(), fmt)
    else:
        sheet.merge_range(row_start, col, row_start + days_left, col, begin.strftime('%B').upper(), fmt)
    new_row_start = row_start
    fmt = None
    for day in range(days_left + 1):
        if day == 0 and port:
            fmt = fmtArr[1]
        elif day == 0 and not port:
            fmt = fmtArr[2]
        elif count < 14 and port:
            fmt = fmtArr[3]
        elif not port:
            fmt = fmtArr[4]
        else:
            raise ValueError('Wrong format for day/date column')
        sheet.write(row_start + day, col + 1, (begin + timedelta(day)).strftime('%a'), fmt)
        sheet.write(row_start + day, col + 2, (begin + timedelta(day)).day, fmt)
        if count == 13:
            port = False
            count += 1
        elif count == 27:
            count = 0
            port = True
        else:
            count += 1
        new_row_start += 1
    return (new_row_start, begin + timedelta(days_left + 1), count, port)

#Writes the formula for the rows of the various Totals column in Distance-Time-Fuel Worksheet
def writeTotalRow(sheet, row_start, col_start, cols, day, duration, fmt_arr):
    for i in range(cols):
        shift = 2 if i == 1 else 0
        #if start of the week
        if (day % duration) == 0:
            formula = '={0}'.format(xl_rowcol_to_cell(row_start + day, 3 + i))
            fmt = fmt_arr[0 + shift]
        #if end of the week
        elif (day % duration) == (duration - 1):
            formula = '={0}+{1}'.format(
                xl_rowcol_to_cell(row_start + day, 3 + i),
                xl_rowcol_to_cell(row_start + day - 1, col_start + i))
            fmt = fmt_arr[1 + shift]
        #else its the middle of the week
        else:
            formula = '={0}+{1}'.format(
                xl_rowcol_to_cell(row_start + day, 3 + i),
                xl_rowcol_to_cell(row_start + day - 1, col_start + i))
            fmt = fmt_arr[0 + shift]
        sheet.write_formula(row_start + day, col_start + i, formula, fmt)

def writeDataRow(sheet, row_start, col_start, cols, day, fmt_arr):
    for i in range(cols):
        sheet.write(row_start + day, col_start + i, "", fmt_arr[i])

def formatBuffers(sheet, start_row, end_row, cols, fmt):
    for col in cols:
        sheet.merge_range(start_row, col, end_row-1, col, "", fmt)
        sheet.set_column(col, col, 3)

#len(measurements) must be > 0;  
def formatHeaders(sheet, start_col, intervals, measurements, measure_fmt, interval_fmt):
    num_m = len(measurements)
    jump = num_m + 1
    for i, interval in enumerate(intervals):
        sheet.merge_range(1, start_col + jump*i, 1, start_col + jump*i + (num_m-1), interval, interval_fmt[i])
        for m, measurement in enumerate(measurements):
            col = start_col + m + jump*i
            sheet.write(2, col, measurement, measure_fmt[m])
            sheet.set_column(col, col, len(measurement) + 1)

#Adds the Expense Tracker Worksheet
def addExpenseTracker(wb, budget, num_shifts = 8):

    et = wb.add_worksheet('Expense Tracker')

    sum_header = ['Budget', 'Consumed', 'Remaining']
    txt_total = "Total"

    #List of strings describing the shift number
    col_headers = ['Shift {0}'.format(i) for i in range(1, num_shifts + 1)]
    #total number of columns used
    tot_col = num_shifts*2 + 1
    #length of the largest of the budget categories or var 'txt_total'
    len_max_budget_headers = max( max( len(key) for key in budget ), len(txt_total) )
    #length of the largest budget + the length of '$,.00'
    len_col_width = max(len("$,.00") + max(len(str(value)) for value in budget.values()),
                        len("Week #"),
                        max(len(header) for header in sum_header)) 
    
    #Lists of various stylings
    fmt_header_budget_arr = [etStyles.header_budget_top, etStyles.header_budget_mid, etStyles.header_budget_bottom]
    fmt_header_budgetcat_arr = [etStyles.header_budgetcat_left, etStyles.header_budgetcat_mid, etStyles.header_budgetcat_right]
    fmt_protected_sum_left_arr = [etStyles.protected_sum_tleft, etStyles.protected_sum_mleft, etStyles.protected_sum_bleft]
    fmt_protected_sum_mid_arr = [etStyles.protected_sum_tmid, etStyles.protected_sum_mmid, etStyles.protected_sum_bmid]
    fmt_protected_sum_right_arr = [etStyles.protected_sum_tright, etStyles.protected_sum_mright, etStyles.protected_sum_bright]

    #Locks the worksheet
    et.set_column('A:XFD', None, etStyles.locked)
    #set width of first column to max size of content    
    et.set_column(0, 0, len_max_budget_headers, etStyles.fmt_sheet)
    #set width of first column to max size of budgets
    et.set_column(1, tot_col, len_col_width, etStyles.fmt_sheet)
    #Set row seperating the summary and weekly tracker to size 15
    #Prevents the row from being hidden when all unused rows are hidden
    et.set_row(len(budget)*2 + 2, 15)

    #Create the main column headers
    et.merge_range(0, 0, 1, 0, "")
    for col, header in enumerate(col_headers):
        et.merge_range(0, 1 + col*2, 0, 2 + col*2, "")
        if col % 2 == 0:
            et.write(0, 1 + 2*col, header, etStyles.header_port)
        else:
            et.write(0, 1 + 2*col, header, etStyles.header_stbd)
        et.write(1, 1 + 2*col, 'Week 1', etStyles.header_week)
        et.write(1, 2 + 2*col, 'Week 2', etStyles.header_week)

    #Create the summary column headers
    for col, header in enumerate(sum_header):
        et.write(3 + len(budget)*2, 1 + col, header, fmt_header_budgetcat_arr[col])

    #Create all row headers and data
    for idx, key in enumerate(budget):
        budget_row = 2 + 2*idx
        total_row = budget_row + 1
        sum_budget_row = 4 + len(budget)*2 + idx
        #Position for formating (top middle bottom)
        position = 0 if idx == 0 else 1 if idx < len(budget) - 1 else 2
        formula_consumed = '=SUM({0}:{1})'.format(xl_rowcol_to_cell(budget_row, 1),
                                                xl_rowcol_to_cell(budget_row, num_shifts*2 + 2))
        formula_remaining = '={0}-{1}'.format(xl_rowcol_to_cell(sum_budget_row, 1),
                                            xl_rowcol_to_cell(sum_budget_row, 2))
        #Create the main row data
        for col in range(1, tot_col):
            #shift totals cell formula
            formula_shift_totals = '=SUM({0}:{1})'.format(xl_rowcol_to_cell(total_row-1, col-1),
                                                        xl_rowcol_to_cell(total_row-1, col))
            #initialize weekly amounts to 0 and format cells
            et.write(budget_row, col, 0, etStyles.costs)
            et.data_validation(budget_row, col, budget_row, col, {'validate': 'decimal', 'criteria': '!=', 'value': -4521834994})
            #Write to the budget shift totals cells
            if (col % 2) == 0:
                et.write(total_row, col, formula_shift_totals, etStyles.protected_tot)
            else:
                et.write(total_row, col, '', etStyles.protected_blank)

        #Write to the 'Budget' column
        et.write(sum_budget_row, 1, budget[key], fmt_protected_sum_left_arr[position])
        #Write to the 'Consumed' column
        et.write_formula(sum_budget_row, 2, formula_consumed, fmt_protected_sum_mid_arr[position])
        #Write to the 'Remaining' column
        et.write_formula(sum_budget_row, 3, formula_remaining, fmt_protected_sum_right_arr[position])

        #Create row headers
        et.write(budget_row, 0, key, etStyles.header_budget_top)
        et.write(total_row, 0, txt_total, etStyles.header_budget_bottom)
        et.write(sum_budget_row, 0, key, fmt_header_budget_arr[position])

    et.set_default_row(hide_unused_rows=True)
    et.hide_gridlines(2)
    col_XFD = 16384 #number of columns in excel
    et.set_column(tot_col, col_XFD - 1, None, cell_format=None, options={'hidden': 1})
    #et.protect()

#Adds the Distance-Time-Fuel Worksheet
def addDistanceTimeFuel(wb, start_date, end_date):
        
    dtf = wb.add_worksheet('Distance-Time-Fuel')
    calHeadTxt = "Calendar Date"
    monthTxt = "Mo"
    dayTxt = "Day"
    dateTxt = "Date"
    measurements = ['Dist. (NM)', 'Time (H:m)', 'Fuel (L)']
    intervals = ['Daily Totals', 'Running Weekly Totals', 'Running Shift Totals', 'Running Monthly Totals', 'Running Season Total']
    num_months = months_between(start_date, end_date)
    num_days = (end_date - start_date).days + 1
    row_days_start = 3 #0 indexed
    col_intervals_start = 3 #0 indexed
    last_row = row_days_start + num_days
    #array of column indices for the buffer columns
    buffer_cols = [col_intervals_start + len(measurements) + i*(len(measurements) + 1) 
                    for i in range(0, len(intervals))]
    #total number of active columns in the sheet
    tot_col = 3 + len(measurements)*len(intervals) + len(buffer_cols)

    #Lists of various stylings
    fmt_date_arr = [dtfStyles.month, dtfStyles.date_first_port, dtfStyles.date_first_stbd, dtfStyles.date_rest_port, dtfStyles.date_rest_stbd]
    fmt_intervals = [dtfStyles.daily, dtfStyles.weekly, dtfStyles.shift, dtfStyles.monthly, dtfStyles.season]
    fmt_header_measurements = [dtfStyles.header_distance, dtfStyles.header_time, dtfStyles.header_fuel]
    fmt_measurements = [dtfStyles.distance, dtfStyles.time, dtfStyles.fuel]
    fmt_arr = [[dtfStyles.notcol, dtfStyles.col_wk, dtfStyles.notcol_time, dtfStyles.col_wk_time],
    [dtfStyles.notcol, dtfStyles.col_shft, dtfStyles.notcol_time, dtfStyles.col_shft_time],
    [dtfStyles.notcol, dtfStyles.col_mo, dtfStyles.notcol_time, dtfStyles.col_mo_time],
    [dtfStyles.notcol, dtfStyles.col_seas, dtfStyles.notcol_time, dtfStyles.col_seas_time]]

    dtf.set_column('A:A', max(4, len(monthTxt) + 2))
    dtf.set_column('B:B', len("Thurs"))
    dtf.set_column('C:C', max(3, len(dateTxt) + 1))

    #Write the formulas for the totaling columns
    for day in range(num_days): 
        writeDataRow(dtf, row_days_start, 3, len(measurements), day, fmt_measurements)
        writeTotalRow(dtf, row_days_start, 7, len(measurements), day, 7, fmt_arr[0])
        #need to investigate when there aren't a multiple of 7 days
        if ((int((num_days-1) / 7) % 2) == 0):
            #normal case, where there are an even number of weeks in the season
            writeTotalRow(dtf, row_days_start, 11, len(measurements), day, 14, fmt_arr[1])            
        else:
            #case for if there are an odd number of weeks in the season
            pass
        writeTotalRow(dtf, row_days_start, 19, len(measurements), day, num_days, fmt_arr[3])

    #Write the months, dates and days according to the start and end dates
    #Also writes the formula for the monthly total columns
    port = True
    count = 0
    for _ in range(num_months - 1):
        days_in_month = (last_day_of_month(start_date) - start_date).days + 1
        for day in range(days_in_month):
            writeTotalRow(dtf, row_days_start, 15, len(measurements), day, days_in_month, fmt_arr[2])
        row_days_start, start_date, count, port = writeDates(dtf, start_date, last_day_of_month(start_date), row_days_start, 0, fmt_date_arr, count, port)
    days_in_month = (end_date - start_date).days + 1
    for day in range(days_in_month):
        writeTotalRow(dtf, row_days_start, 15, len(measurements), day, days_in_month, fmt_arr[2])
    row_days_start, start_date, count, port = writeDates(dtf, start_date, end_date, row_days_start, 0, fmt_date_arr, count, port)

    #Write the title cell
    dtf.merge_range(0, 0, 0, tot_col, '{0} {1} - {2} {3}'.format(start_date.strftime("%B"), start_date.year, end_date.strftime("%B"), end_date.year), dtfStyles.title)

    #Write date headers
    dtf.merge_range('A2:C2', calHeadTxt, dtfStyles.header_cal)
    dtf.write('A3', monthTxt, dtfStyles.header_fuel)
    dtf.write('B3', dayTxt, dtfStyles.header_time)
    dtf.write('C3', dateTxt, dtfStyles.header_fuel)

    #Merge and format buffer columns
    formatBuffers(dtf, 1, last_row, buffer_cols, dtfStyles.buffer)

    #Write time interval and measurement headers
    formatHeaders(dtf, col_intervals_start, intervals, measurements, fmt_header_measurements, fmt_intervals)

    dtf.set_default_row(hide_unused_rows=True)
    dtf.hide_gridlines(2)
    col_XFD = 16384 #number of columns in excel
    dtf.set_column(tot_col, col_XFD - 1, None, cell_format=None, options={'hidden': 1})
    dtf.freeze_panes(3, 3)
    #dtf.protect()

#Adds the Shift Summary Worksheet
def addShiftSummary(wb):
    pass

def main():

    num_shifts = 8
    start_date = date(2018, 5, 16)
    end_date = date(2018, 9, 5)
    #budget values ideally need to be integers, or MAX 2 decimal places. If .00, just type integer
    budget = {'Fuel': 50000, 'Groceries': 4000, 'Base Supplies': 8000, 'SAR Snacks': 320, 'Medical Gear': 1000, 'Other': 0}

    addExpenseTracker(wb, budget, num_shifts)
    addDistanceTimeFuel(wb, start_date, end_date)
    addShiftSummary(wb)

    wb.close()
    
    # # opens the excel file when this code is run
    # from win32com.client import Dispatch
    # xl = Dispatch("Excel.Application")
    # xl.Visible = True # otherwise excel is hidden
    # # newest excel does not accept forward slash in path
    # xl.Workbooks.Open(r'C:\Users\brend\Documents\Coding\Side Projects\CCG_StationTracker\Example_Output\CCG-ShiftTracker-StationName.xlsx')

if __name__ == "__main__":
    main()