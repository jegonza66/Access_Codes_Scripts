import os
import pickle
import tkinter as tk
from tkinter import filedialog


def API_Paths(Credentials_file = 'Credentials/Credentials.pkl'):
    try:
        f = open(Credentials_file, 'rb')
        Credentials = pickle.load(f)
        f.close()
    except:
        Credentials = {}
        # Get API Key
        Credentials['API_Key'] = input('Please enter your API Key\nAPI Key:')
        # Get save path
        Credentials['csv_save_path'] = input(
            'Please copy and paste the path to the save folder (usually the BNED folder '
            'in OneDrive).\nPath:').replace('\\', '/') + '/'
        print('Please use the dialog window to go to the location of the School_Names.xlsx file '
              'and open it.')
        # Get School_Names file
        root = tk.Tk()
        root.withdraw()
        Credentials['School_Name_file'] = filedialog.askopenfilename()

        # Save API Key and BNED OneDrive path in credentials file
        os.makedirs('Credentials', exist_ok=True)
        f = open(Credentials_file, 'wb')
        pickle.dump(Credentials, f)
        f.close()
    return Credentials


def Verba_Credentials(Credentials, Credentials_file = 'Credentials/Credentials.pkl'):

    try:
        Credentials['Verba_Username'] and Credentials['Verba_Password']
        return Credentials
    except:
        print('No Verba Connect Credentials found. Please Enter your username and password.\n')
        Credentials['Verba_Username'] = input('Username:')
        Credentials['Verba_Password'] = input('Password:')

        # Save Verba credentials in Credentials file
        f = open(Credentials_file, 'wb')
        pickle.dump(Credentials, f)
        f.close()

    return Credentials