import numpy as np

import Functions
import Excel_Files_Manager


def low_notification(Credentials):
    # Read cases file
    New_file = Excel_Files_Manager.load_file()

    # Ask if dismiss cases
    New_file = Excel_Files_Manager.dismiss_repeated(New_file=New_file)

    # Getting ready for ruby run
    Billing_ISBNs = list(New_file['Billing Isbn'].values)
    VBIDs = [Billing_ISBN if type(Billing_ISBN) == int else Billing_ISBN.split('R')[0] for Billing_ISBN in
             Billing_ISBNs]

    # Define quantities (rounding up), schools and catalogs from dataframe
    quantities = np.array(list(New_file['Number Codes Needed'].values))
    quantities = np.ceil(np.round(quantities + 5.1, -1)).astype(int)
    Verba_Schools = list(New_file['Login'].values)
    Catalogs = list(New_file['Name'].values)
    URLs = list(New_file['Access Code URL'].values)
    Publishers = list(New_file['Name.1'].values)
    Titles = list(New_file['Title'].values)

    # Load school names list and translate to School names
    School_names_path = Functions.School_Name_path(Credentials)
    Schools = Functions.translate_verba_school_name(School_names_path=School_names_path, Verba_Schools=Verba_Schools)

    # Check Billing_ISBNs and VBIDs have same size
    if len(Billing_ISBNs) == len(VBIDs) == len(quantities) == len(Schools) == len(Catalogs) == len(URLs) == len(
            Publishers) == len(Titles):
        run = True
    else:
        run = False

    return Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, Publishers, Titles, URLs, run

def code_reveal(Credentials):
    # Read DD file and take Code Reveal Cases
    DD_Code_Reveal, DD = Excel_Files_Manager.Load_DD_Code_Reveal()

    # Define Billing ISBN and VBIDS from SKU, exception only in case SKU contains R character
    Billing_ISBNs = list(DD_Code_Reveal['SKU.1'])
    VBIDs = [Billing_ISBN if type(Billing_ISBN) == int else Billing_ISBN.split('R')[0] for Billing_ISBN in
             Billing_ISBNs]

    # Define quantities (rounding up), schools and catalogs from dataframe
    quantities = [np.ceil(np.round(Total_Enrollments + 5.1, -1)).astype(int) if not np.isnan(Total_Enrollments) else 0
                  for Total_Enrollments in DD_Code_Reveal['Total Estimated Enrollments']]
    Schools = DD_Code_Reveal['School'].values
    Catalogs = DD_Code_Reveal['Catalog'].values

    # Load school names list and translation to verba names
    School_names_path = Functions.School_Name_path(Credentials=Credentials)
    Verba_Schools = Functions.translate_school_name(School_names_path=School_names_path, Schools=Schools)

    # Check Billing_ISBNs and VBIDs have same size
    if len(Billing_ISBNs) == len(VBIDs) == len(quantities) == len(Schools) == len(Catalogs):
        run = True
    else:
        run = False

    return Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, run, DD


def fake_code_reveal(Credentials, DD):
    # Read DD file and take Code Reveal Cases
    DD_Fake_Code_Reveal = Excel_Files_Manager.Load_DD_Fake_Code_Reveal(DD=DD)

    # Define Billing ISBN and VBIDS from SKU, exception only in case SKU contains R character
    Billing_ISBNs = list(DD_Fake_Code_Reveal['SKU.1'])
    VBIDs = [Billing_ISBN if type(Billing_ISBN) == int else Billing_ISBN.split('R')[0] for Billing_ISBN in
             Billing_ISBNs]

    # Define quantities (rounding up), schools and catalogs from dataframe
    quantities = [np.ceil(np.round(Total_Enrollments + 5.1, -1)).astype(int) if not np.isnan(Total_Enrollments) else 0
                  for Total_Enrollments in DD_Fake_Code_Reveal['Total Estimated Enrollments']]
    Schools = DD_Fake_Code_Reveal['School'].values
    Catalogs = DD_Fake_Code_Reveal['Catalog'].values

    # Load school names list and translation to verba names
    School_names_path = Functions.School_Name_path(Credentials=Credentials)
    Verba_Schools = Functions.translate_school_name(School_names_path=School_names_path, Schools=Schools)

    # Check Billing_ISBNs and VBIDs have same size
    if len(Billing_ISBNs) == len(VBIDs) == len(quantities) == len(Schools) == len(Catalogs):
        run = True
    else:
        run = False

    return Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, run


# OLD FUUNCTIONS OUT OF USE
def run_low_notification_setup_DISMISSED(Credentials):
    # Read files
    Old_file, New_file, Low_notification_old_path = Excel_Files_Manager.read_low_files()

    # get cells highlight colors from old file
    Cell_colours, colors = Excel_Files_Manager.get_old_file_colors(Old_file=Old_file, Low_notification_old_path=Low_notification_old_path)

    # Get new and repeated cases from new and old file
    new_cases, repeated_cases = Excel_Files_Manager.get_new_repeated(Old_file=Old_file, New_file=New_file)

    # append repeated cases not in red to new cases
    new_cases = Excel_Files_Manager.append_not_red(new_cases=new_cases, repeated_cases=repeated_cases, Cell_colours=Cell_colours)

    # Get color of repeated cases in old file
    Common = Excel_Files_Manager.check_old_colors(repeated_cases=repeated_cases, Cell_colours=Cell_colours)

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
