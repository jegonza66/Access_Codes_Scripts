import pandas as pd
import xlrd
import tkinter as tk
from tkinter import filedialog

def read_low_no_file():
    # Read files
    print('Please use the dialog window to go to the location of the Access Code Notification.xlsx file '
          'and open it.')
    root = tk.Tk()
    root.withdraw()
    No_notification_path = filedialog.askopenfilename()

    columns = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Number Codes Needed', 'Access Code URL']
    columns_no_codes = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Access Code URL']
    New_file = pd.read_excel(io=No_notification_path)[columns]
    # Drop duplicated rows
    New_file.drop_duplicates(inplace=True)
    # Take only rows with bigger number of codes needed for repeated isbn schools and catalogs
    New_file = New_file.sort_values('Number Codes Needed', ascending=False).drop_duplicates(columns_no_codes, keep='first')
    # Make New_file ISBN columns type object (in case no ISBNS have '-' or 'R' it will be type int, and cannot compare
    # to Old_file columns of type object)
    New_file = New_file.astype({"Billing Isbn": object})

    return New_file


def read_low_files():
    # Read files
    print('Please use the dialog window to go to the location of the Old Low Access Code Notification.xls file '
          '(the one with highlighted cells), and open it.\nRemember to change the extension to .xls, you can do so by\n'
          'opening the file with Excel -> Save as -> Select the .xls extension -> Save.')
    # Open file explorer dialog
    root = tk.Tk()
    root.withdraw()
    Low_notification_old_path = filedialog.askopenfilename()

    root = tk.Tk()
    root.withdraw()
    print('Please use the dialog window to go to the location of the New Low Access Code Notification.xls file '
          'and open it.\nRemember to change the extension to .xls.')
    Low_notification_new_path = filedialog.askopenfilename()

    columns = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Number Codes Needed', 'Access Code URL']
    columns_no_codes = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Access Code URL']
    Old_file = pd.read_excel(io=Low_notification_old_path)[columns]
    New_file = pd.read_excel(io=Low_notification_new_path)[columns]
    # Drop duplicated rows
    New_file.drop_duplicates(inplace=True)
    # Take only rows with bigger number of codes needed for repeated isbn schools and catalogs
    New_file = New_file.sort_values('Number Codes Needed', ascending=False).drop_duplicates(columns_no_codes,
                                                                                            keep='first')
    # Make New_file ISBN columns type object (in case no ISBNS have '-' or 'R' it will be type int, and cannot compare
    # to Old_file columns of type object(
    New_file = New_file.astype({"Billing Isbn": object})
    Old_file.drop_duplicates(inplace=True)

    return Old_file, New_file, Low_notification_old_path


def get_old_file_colors(Old_file, Low_notification_old_path):
    # Read excel format for taking higlight colors
    book = xlrd.open_workbook(Low_notification_old_path, formatting_info=True)
    sheet = book.sheet_by_index(0)
    rows, cols = sheet.nrows, sheet.ncols

    # Make dataframes from rows with different colours
    columns_no_codes = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Access Code URL']

    Cell_colours = {}
    Cell_colours['Red'] = pd.DataFrame(columns=columns_no_codes)
    Cell_colours['Blue'] = pd.DataFrame(columns=columns_no_codes)
    Cell_colours['Green'] = pd.DataFrame(columns=columns_no_codes)
    Cell_colours['White'] = pd.DataFrame(columns=columns_no_codes)
    colors = []

    for row in range(1, rows):
        col = 0
        if not sheet.cell_type(row, col) == xlrd.XL_CELL_EMPTY:
            # could get 'dump', 'value', 'xf_index'
            xfx = sheet.cell_xf_index(row, col)
            xf = book.xf_list[xfx]
            bgx = xf.background.pattern_colour_index
            colors.append(bgx)
            if bgx == 29 or bgx == 10:
                Cell_colours['Red'] = Cell_colours['Red'].append(Old_file.iloc[row - 1])
            elif bgx == 44 or bgx == 40:
                Cell_colours['Blue'] = Cell_colours['Blue'].append(Old_file.iloc[row - 1])
            elif bgx == 50:
                Cell_colours['Green'] = Cell_colours['Green'].append(Old_file.iloc[row - 1])
            elif bgx == 64:
                Cell_colours['White'] = Cell_colours['White'].append(Old_file.iloc[row - 1])
        else:
            break

    return Cell_colours, colors


def get_new_repeated(Old_file, New_file):
    columns_no_codes = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Access Code URL']
    # mark which cases are new and repeated
    common_right = Old_file.merge(New_file, on=columns_no_codes, how='right',
                                        indicator=True)

    # take new and repeated cases indexes
    new_cases_index = common_right['_merge'] == 'right_only'
    repeated_cases_index = common_right['_merge'] == 'both'

    new_cases = common_right[new_cases_index].drop('_merge', axis=1)
    repeated_cases = common_right[repeated_cases_index].drop('_merge', axis=1)

    return new_cases, repeated_cases


def append_not_red(new_cases, repeated_cases, Cell_colours):
    columns_no_codes = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Access Code URL']

    # Change this for check if not red
    try:
        not_red = repeated_cases.merge(Cell_colours['Red'], on=columns_no_codes, how='left',
                                       indicator=True)
        not_red_index = not_red['_merge'] == 'left_only'
        not_red = not_red[not_red_index]

        # append not red cases to new cases
        new_cases = new_cases.append(not_red)
        new_cases = new_cases.drop('_merge', axis=1)
    except:
        pass
    return new_cases


def check_old_colors(repeated_cases, Cell_colours):
    columns_no_codes = ['Login', 'Name', 'Name.1', 'Title', 'Billing Isbn', 'Access Code URL']

    # Check if repeated cases were in white in old file
    Common = {}
    if Cell_colours['White'].shape[0]:
        common_white = repeated_cases.merge(Cell_colours['White'], on=columns_no_codes, how='left',
                                            indicator=True)
        common_white_index = common_white['_merge'] == 'both'
        common_white = common_white[common_white_index]
        Common['White'] = common_white

    # Check if repeated cases were in green in old file
    if Cell_colours['Green'].shape[0]:
        common_green = repeated_cases.merge(Cell_colours['Green'], on=columns_no_codes, how='left',
                                            indicator=True)
        common_green_index = common_green['_merge'] == 'both'
        common_green = common_green[common_green_index]
        Common['Green'] = common_green

    # Check if repeated cases were in blue in old file
    if Cell_colours['Blue'].shape[0]:
        common_blue = repeated_cases.merge(Cell_colours['Blue'], on=columns_no_codes, how='left',
                                           indicator=True)
        common_blue_index = common_blue['_merge'] == 'both'
        common_blue = common_blue[common_blue_index]
        Common['Blue'] = common_blue

    # Check if repeated cases were in red in old file
    if Cell_colours['Red'].shape[0]:
        common_red = repeated_cases.merge(Cell_colours['Red'], on=columns_no_codes, how='left',
                                          indicator=True)
        common_red_index = common_red['_merge'] == 'both'
        common_red = common_red[common_red_index]
        Common['Red'] = common_red

    return Common