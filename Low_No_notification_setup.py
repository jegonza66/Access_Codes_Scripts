import numpy as np

import Functions
import Low_Notification_Excel_Manage

def run_low_notification_setup(Credentials):
    # Read files
    Old_file, New_file, Low_notification_old_path = Low_Notification_Excel_Manage.read_low_files()

    # get cells highlight colors from old file
    Cell_colours, colors = Low_Notification_Excel_Manage.get_old_file_colors(Old_file=Old_file, Low_notification_old_path=Low_notification_old_path)

    # Get new and repeated cases from new and old file
    new_cases, repeated_cases = Low_Notification_Excel_Manage.get_new_repeated(Old_file=Old_file, New_file=New_file)

    # append repeated cases not in red to new cases
    new_cases = Low_Notification_Excel_Manage.append_not_red(new_cases=new_cases, repeated_cases=repeated_cases, Cell_colours=Cell_colours)

    # Get color of repeated cases in old file
    Common = Low_Notification_Excel_Manage.check_old_colors(repeated_cases=repeated_cases, Cell_colours=Cell_colours)

    # Take new cases from new file with number of codes needed
    New_cases_codes = new_cases.merge(New_file[['Login', 'Name', 'Billing Isbn']], how='right', indicator=True)
    New_cases_codes_index = New_cases_codes['_merge'] == 'both'
    New_file_codes = New_file[New_cases_codes_index.values]

    # Getting ready for ruby run
    Billing_ISBNs = list(New_file_codes['Billing Isbn'].values)
    VBIDs = [Billing_ISBN if type(Billing_ISBN) == int else Billing_ISBN.split('R')[0] for Billing_ISBN in Billing_ISBNs]

    # Define quantities (rounding up), schools and catalogs from dataframe
    quantities = np.array(list(New_file_codes['Number Codes Needed'].values))
    quantities = np.ceil(np.round(quantities+5.1, -1)).astype(int)
    Verba_Schools = list(New_file_codes['Login'].values)
    Catalogs = list(New_file_codes['Name'].values)
    URLs = list(New_file_codes['Access Code URL'].values)
    Publishers = list(New_file_codes['Name.1'].values)
    Titles = list(New_file_codes['Title'].values)

    # Load school names list and translate to School names
    School_names_path = Functions.School_Name_path(Credentials)
    Schools = Functions.translate_verba_school_name(School_names_path=School_names_path, Verba_Schools=Verba_Schools)

    # Check Billing_ISBNs and VBIDs have same size
    if len(Billing_ISBNs) == len(VBIDs) == len(quantities) == len(Schools) == len(Catalogs) == len(URLs) == len(Publishers) == len(Titles):
        run = True
    else:
        run = False

    return Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, Publishers, Titles, URLs, run


def run_no_notification_setup(Credentials):
    # Read files
    New_file = Low_Notification_Excel_Manage.read_no_file()

    # Getting ready for ruby run
    Billing_ISBNs = list(New_file['Billing Isbn'].values)
    VBIDs = [Billing_ISBN if type(Billing_ISBN) == int else Billing_ISBN.split('R')[0] for Billing_ISBN in Billing_ISBNs]

    # Define quantities (rounding up), schools and catalogs from dataframe
    quantities = np.array(list(New_file['Number Codes Needed'].values))
    quantities = np.ceil(np.round(quantities + 5.1, -1)).astype(int)
    Verba_Schools = list(New_file['Login'].values)
    Catalogs = list(New_file['Name'].values)
    URLs = list(New_file['Access Code URL'].values)

    # Load school names list and translate to School names
    School_names_path = Functions.School_Name_path(Credentials)
    Schools = Functions.translate_verba_school_name(School_names_path=School_names_path, Verba_Schools=Verba_Schools)

    # Check Billing_ISBNs and VBIDs have same size
    if len(Billing_ISBNs) == len(VBIDs) == len(quantities) == len(Schools) == len(Catalogs) == len(URLs):
        run = True
    else:
        run = False

    return Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, URLs, run
