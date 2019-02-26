import xlsxwriter
filepath = 'Example_Output/'
filename = 'CCG-ShiftTracker-StationName.xlsx'
wb = xlsxwriter.Workbook(filepath + filename) 

#color variables
class Colors:
    port = '#CC0000' #Red
    stbd = '#008800' #Green
    daily = '#6F8F8F' #Slategray-ish
    weekly = 'FFBB00' #Yellow-Orange
    shift = '#1E90FF' #Doger Blue
    monthly = '#DD00DD' #Purple/Pink
    season = '#D55A5A' #Light Brown/Red
    protect_light = '#E6E6E6' #background gray
    protect_normal = '#C8C8C8' #light gray
    protect_dark = 'gray' #dark gray
    border_light = 'gray' #dark gray

#Expense Tracker Styles
class etStyles:
    
    locked = wb.add_format({'locked': True})
    fmt_sheet = wb.add_format({'bg_color': Colors.protect_light})
    
    header_week = wb.add_format({
        'bold': True,
        'align': 'center',
        'bottom': 2,
        'right': 1,
        'left': 1,
        'top': 1,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'top_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_port = wb.add_format({
        'bold': True,
        'align': 'center',
        'font_color': 'white',
        'bottom': 1,
        'right': 1,
        'left': 1,
        'top': 2,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.port})
    header_stbd = wb.add_format({
        'bold': True,
        'align': 'center',
        'font_color': 'white',
        'bottom': 1,
        'right': 1,
        'left': 1,
        'top': 2,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.stbd})
    header_budget_top = wb.add_format({
        'bold': True,
        'align': 'left',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 2,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_budget_mid = wb.add_format({
        'bold': True,
        'align': 'left',
        'top': 1,
        'right': 2,
        'bottom': 1,
        'left': 1,
        'top_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_budget_bottom = wb.add_format({
        'bold': True,
        'align': 'left',
        'top': 1,
        'right': 2,
        'bottom': 2,
        'left': 2,
        'top_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_budgetcat_left = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 1,
        'bottom': 2,
        'left': 2,
        'right_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_budgetcat_mid = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 1,
        'bottom': 2,
        'left': 1,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_budgetcat_right = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 2,
        'bottom': 2,
        'left': 1,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})

    protected_blank = wb.add_format({
        'top': 1,
        'right': 1,
        'bottom': 2,
        'left': 1,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_tot = wb.add_format({
        'num_format': '$#,##0.00',
        'bold': True,
        'align': 'center',
        'top': 1,
        'right': 1,
        'bottom': 2,
        'left': 1,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_dark})
    costs = wb.add_format({
        'num_format': '$#,##0.00', 
        'align': 'center',
        'top': 2,
        'right': 1,
        'bottom': 1,
        'left': 1,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'left_color': Colors.border_light,
        'locked': False})
    protected_sum_tleft = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 2,
        'right': 1,
        'bottom': 1,
        'left': 2,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_mleft = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 1,
        'right': 1,
        'bottom': 1,
        'left': 2,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_bleft = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 1,
        'right': 1,
        'bottom': 2,
        'left': 2,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_tmid = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 2,
        'right': 1,
        'bottom': 1,
        'left': 1,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_mmid = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 1,
        'right': 1,
        'bottom': 1,
        'left': 1,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_bmid = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 1,
        'right': 1,
        'bottom': 2,
        'left': 1,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_tright = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 1,
        'bottom_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_mright = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 1,
        'right': 2,
        'bottom': 1,
        'left': 1,
        'top_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    protected_sum_bright = wb.add_format({
        'num_format': '$#,##0.00',
        'top': 1,
        'right': 2,
        'bottom': 2,
        'left': 1,
        'top_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})

#DistanceTimeFuel Styles
class dtfStyles:
    buffer = wb.add_format({
        'top': 2,
        'right': 1,
        'bottom': 1,
        'left': 1,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_dark})

    month = wb.add_format({
        'font_size': 24,
        'rotation': -90,
        'align': 'center',
        'valign': 'vcenter',
        'top': 6,
        'right': 2,
        'bottom': 6,
        'left': 2,
        'bg_color': Colors.protect_normal})
    date_first_port = wb.add_format({
        'font_color': Colors.port,
        'align': 'center',
        'top': 6,
        'right': 1,
        'bottom': 1,
        'left': 2,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    date_first_stbd = wb.add_format({
        'font_color': Colors.stbd,
        'align': 'center',
        'top': 6,
        'right': 1,
        'bottom': 1,
        'left': 2,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    date_rest_port = wb.add_format({
        'font_color': Colors.port,
        'align': 'center',
        'top': 1,
        'right': 1,
        'bottom': 1,
        'left': 2,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    date_rest_stbd = wb.add_format({
        'font_color': Colors.stbd,
        'align': 'center',
        'top': 1,
        'right': 1,
        'bottom': 1,
        'left': 2,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_normal})

    header_cal = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 2,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.protect_dark})
    title = wb.add_format({
        'bold': True,
        'align': 'center',
        'border': 2,
        'bg_color': Colors.protect_normal})

    daily = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 2,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.daily})
    weekly = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 2,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.weekly})
    shift = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 2,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.shift})
    monthly = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 2,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.monthly})
    season = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 2,
        'right': 2,
        'bottom': 1,
        'left': 2,
        'bottom_color': Colors.border_light,
        'bg_color': Colors.season})

    header_distance = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 1,
        'right': 1,
        'bottom': 2,
        'left': 2,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_time = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 1,
        'right': 1,
        'bottom': 2,
        'left': 1,
        'top_color': Colors.border_light,
        'right_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    header_fuel = wb.add_format({
        'bold': True,
        'align': 'center',
        'top': 1,
        'right': 2,
        'bottom': 2,
        'left': 1,
        'top_color': Colors.border_light,
        'left_color': Colors.border_light,
        'bg_color': Colors.protect_normal})

    distance = wb.add_format({
        'num_format': '#,##0.00',
        'border': 1,
        'border_color': Colors.border_light})
    time = wb.add_format({
        'num_format': '[h]:mm',
        'border': 1,
        'border_color': Colors.border_light})
    fuel = wb.add_format({
        'num_format': '#,##0.00',
        'border': 1,
        'border_color': Colors.border_light})

    notcol = wb.add_format({
        'num_format': '#,##0.00',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    notcol_time = wb.add_format({
        'num_format': 'h:mm',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.protect_normal})
    col_wk = wb.add_format({
        'num_format': '#,##0.00',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.weekly})
    col_wk_time = wb.add_format({
        'num_format': '[h]:mm',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.weekly})
    col_shft = wb.add_format({
        'num_format': '#,##0.00',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.shift})
    col_shft_time = wb.add_format({
        'num_format': '[h]:mm',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.shift})
    col_mo = wb.add_format({
        'num_format': '#,##0.00',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.monthly})
    col_mo_time = wb.add_format({
        'num_format': '[h]:mm',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.monthly})
    col_seas = wb.add_format({
        'num_format': '#,##0.00',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.season})
    col_seas_time = wb.add_format({
        'num_format': '[h]:mm',
        'align': 'center',
        'border': 1,
        'border_color': Colors.border_light,
        'bg_color': Colors.season})
