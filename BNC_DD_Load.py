import pandas as pd
import tkinter as tk
from tkinter import filedialog


def DD_Code_Reveal(sheet_name='April 2021 -', Case='Code Reveal', columns=None):

    print('Please use the dialog window to go to the location of the BNC Daily Delta.xlsx file and open it.')
    # Open file explorer dialog for
    root = tk.Tk()
    root.withdraw()
    BNC_DD_path = filedialog.askopenfilename()

    print('Loading BNC DD file...')
    # Read file
    DD = pd.read_excel(io=BNC_DD_path, sheet_name=sheet_name)

    # Keep only Code Reveal Cases and columns of interest
    if columns is None:
        columns = ['School', 'Catalog', 'Total Estimated Enrollments', 'SKU.1', 'New Item Work complete', 'Notes',
                   'Publisher']
    DD_Code_Reveal = DD.loc[(DD['New Item Work complete'] == Case) & (DD['Notes'] == Case)][columns]

    # Drop duplicate rows
    DD_Code_Reveal.drop_duplicates(inplace=True)
    print('Done')

    return DD_Code_Reveal, DD


def DD_Fake_Code_Reveal(DD, Case='Fake Code Reveal', columns=None):

    # Keep only Code Reveal Cases and columns of interest
    if columns is None:
        columns = ['School', 'Catalog', 'Total Estimated Enrollments', 'SKU.1', 'New Item Work complete', 'Notes',
                   'Publisher']
    DD_Code_Reveal = DD.loc[(DD['New Item Work complete'] == Case) & (DD['Notes'] == Case)][columns]

    # Drop duplicate rows
    DD_Code_Reveal.drop_duplicates(inplace=True)

    return DD_Code_Reveal