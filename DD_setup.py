import numpy as np

import BNC_DD_Load
import Functions

def run_code_reveal(Credentials):
    # Read DD file and take Code Reveal Cases
    DD_Code_Reveal, DD = BNC_DD_Load.DD_Code_Reveal()

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


def run_fake_code_reveal(Credentials, DD):
    # Read DD file and take Code Reveal Cases
    DD_Fake_Code_Reveal = BNC_DD_Load.DD_Fake_Code_Reveal(DD=DD)

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
