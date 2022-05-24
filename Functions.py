import time
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import pickle


def choose_process():
    process = False
    while not process:
        Answer = input('\nWhich process would you like to run?\n'
                       '1. BNED Daily Delta\n'
                       '2. Low Notification\n'
                       'Enter the corresponding number or name and press Enter.')

        DD_selection = {'1', 'Daily Delta', 'dd', 'DD', 'BNED'}
        Low_notification_selection = {'2', 'low', 'Low Notification', 'no', 'low-no', 'Low-No'}
        if Answer in DD_selection:
            process = 'Daily Delta'
        elif Answer in Low_notification_selection:
            process = 'Low Notification'
        if not process:
            print('Please select one of the options given!')

    return process


def School_Name_path(Credentials):
    try:
        School_names_path = Credentials['School_Name_file']
    except:
        print('\nPlease use the new window explorer and open to find and open the School_Names.xlsx file.')
        # Open file explorer dialog to get ruby path
        root = tk.Tk()
        root.withdraw()
        Credentials['School_Name_file'] = filedialog.askopenfilename()
        School_names_path = Credentials['School_Name_file']

        # Save updated Credentials with Ruby path
        Credentials_file = 'Credentials/Credentials.pkl'
        f = open(Credentials_file, 'wb')
        pickle.dump(Credentials, f)
        f.close()

    return School_names_path


def translate_school_name(School_names_path, Schools):
    School_names_list = pd.read_csv(School_names_path)
    Verba_Schools = []
    Update_School_Names_file = False
    for School in Schools:
        try:
            Verba_School = School_names_list.loc[School_names_list['School Name'] == School]['Verba School Name'].values[0]
            Verba_Schools.append(Verba_School)
        except:
            Update_School_Names_file = True
            Verba_School = input('\nCould not find the School name for {}.\n Please enter the school name '
                                'used in Connect and press Enter.'.format(School))
            Verba_Schools.append(Verba_School)
            School_names_list = School_names_list.append(pd.DataFrame([[School, Verba_School]], columns=School_names_list.columns))

    if Update_School_Names_file:
        School_names_list.to_csv(School_names_path, index=False)
    return Verba_Schools


def translate_verba_school_name(School_names_path, Verba_Schools):
    School_names_list = pd.read_csv(School_names_path)
    Schools = []
    Update_School_Names_file = False
    for Verba_School in Verba_Schools:
        try:
            School = School_names_list.loc[School_names_list['Verba School Name'] == Verba_School]['School Name'].values[0]
            Schools.append(School)
        except:
            Update_School_Names_file = True
            School = input('\nCould not find the School name for {}.\n Please enter the school name '
                                'and press Enter.'.format(Verba_School))
            Schools.append(School)
            School_names_list = School_names_list.append(pd.DataFrame([[School, Verba_School]], columns=School_names_list.columns))

    if Update_School_Names_file:
        School_names_list.to_csv(School_names_path, index=False)
    return Schools


def ruby_directory(Credentials):
    try:
        ruby_path = Credentials['Ruby_path']
        os.chdir(ruby_path)
    except:
        print('\nPlease use the new window explorer and open to find and open the ruby program.')
        # Open file explorer dialog to get ruby path
        root = tk.Tk()
        root.withdraw()
        Credentials['Ruby_path'] = '/'.join(filedialog.askopenfilename().split('/')[:-1])

        # Save updated Credentials with Ruby path
        Credentials_file = 'Credentials/Credentials.pkl'
        f = open(Credentials_file, 'wb')
        pickle.dump(Credentials, f)
        f.close()

        # Change directory to ruby path
        os.chdir(Credentials['Ruby_path'])


def check_file(df, csv_file, quantity, URL, Title):
    Error = False

    if not pd.isna(URL):
        df['URL'] = URL

    codes_generated = int(quantity - df['Access Code'].isnull().values.sum())
    url_generated = int(quantity - df['URL'].isnull().values.sum())
    Missing_codes = int(quantity - codes_generated)
    Missing_url = int(quantity - url_generated)

    if Missing_codes and Missing_url:
        Error = 'Missing Access Codes and URL'
        if codes_generated and codes_generated == url_generated:
            Error = 'Run out of codes'
            df = df.head(codes_generated)
        if 'eVP' in Title:
            print('eVP Title.\nChange codes column to message.')
            df['Access Code'] = 'No access code is needed for this content, ' \
                                'please navigate to the publisher integration via your LMS.'
            Error = False
            Missing_codes = 0
            df.to_csv(csv_file, index=False)

    elif df['Access Code'].isnull().values.any():
        Error = 'Missing Access Codes'
        if codes_generated:
            Error = 'Run out of codes'
            df = df.head(codes_generated)
        if 'eVP' in Title:
            print('eVP Title.\nChange codes column to message.')
            df['Access Code'] = 'No access code is needed for this content, ' \
                                'please navigate to the publisher integration via your LMS.'
            Error = False
            Missing_codes = 0
            df.to_csv(csv_file, index=False)

    elif df['URL'].isnull().values.any():
        Error = 'Missing URL'
        if 'eVP' in Title:
            print('eVP Title.\nChange codes column to message.')
            df['Access Code'] = 'No access code is needed for this content, ' \
                                'please navigate to the publisher integration via your LMS.'
            Error = False
            Missing_codes = 0
            df.to_csv(csv_file, index=False)

    return df, Error, Missing_codes


def append_to_report(Report, File_imported, Error, Missing_codes, School, Catalog, Publisher, Title, Billing_ISBN, quantity):
    if File_imported:
        Report[File_imported].append([School, Catalog, Publisher, Title, Billing_ISBN, quantity-Missing_codes])
    else:
        Report[Error].append([School, Catalog, Publisher, Title, Billing_ISBN, Missing_codes])
    return Report


def write_report(Report, save_path, process):
    new_dir = save_path + time.strftime('Access-Codes/%Y/%m/%d/{}/Reports/'.format(process))
    os.makedirs(new_dir, exist_ok=True)
    file_name = time.strftime(new_dir + '{} Report %H_%M.txt'.format(process))
    text = ''
    if len(Report['OK']):
        text += 'The following {} files were correctly uploaded to connect:' \
                '\n{}'.format(len(Report['OK']), '\n'.join(str(line) for line in Report['OK']))
    if len(Report['Failed Import']):
        text += '\n\nThe following {} files were generated correctly but could not be uploaded to connect:' \
                '\n{}'.format(len(Report['Failed Import']),
                              '\n'.join(str(line) for line in Report['Failed Import']))
    if len(Report['Dismiss']):
        text += '\n\nThe following {} requests were dismissed because already available codes:' \
                '\n{}'.format(len(Report['Dismiss']), '\n'.join(str(line) for line in Report['Dismiss']))
    if len(Report['Missing Access Codes and URL']):
        text += '\n\nThe following {} files were missing Access Codes and URL (# missing codes):' \
                '\n{}'.format(len(Report['Missing Access Codes and URL']),
                              '\n'.join(str(line) for line in Report['Missing Access Codes and URL']))
    if len(Report['Missing Access Codes']):
        text += '\n\nThe following {} files were missing Access Codes (# missing codes):' \
                '\n{}'.format(len(Report['Missing Access Codes']),
                              '\n'.join(str(line) for line in Report['Missing Access Codes']))
    if len(Report['Missing URL']):
        text += '\n\nThe following {} files were missing URL:' \
                '\n{}'.format(len(Report['Missing URL']), '\n'.join(str(line) for line in Report['Missing URL']))
    if len(Report['Run out of codes']):
        text += '\n\nThe following {} files Run out of Access Codes. Must upload and request missing codes:' \
                '\n{}'.format(len(Report['Run out of codes']),
                              '\n'.join(str(line) for line in Report['Run out of codes']))
    if len(Report['eCampus Content Holding']):
        text += '\n\nThe following {} files have eCampus Content Holding Publisher with no available codes.' \
                '\nContact Maka for new codes:' \
                '\n{}'.format(len(Report['eCampus Content Holding']),
                              '\n'.join(str(line) for line in Report['eCampus Content Holding']))

    f = open(file_name, 'w')
    f.write(text)
    f.close()


def write_final_report(Report, save_path, process):
    for key_name in Report.keys():
        Report[key_name] = sorted(Report[key_name], key=lambda t: (t[0], t[1]))
    new_dir = save_path + time.strftime('Access-Codes/%Y/%m/%d/{}/'.format(process))
    os.makedirs(new_dir, exist_ok=True)
    file_name = time.strftime(new_dir + '{} Final Report %H_%M.txt'.format(process))
    text = ''
    if len(Report['OK']):
        text += 'The following {} files were correctly uploaded to connect:' \
                '\n{}'.format(len(Report['OK']), '\n'.join(str(line) for line in Report['OK']))
    if len(Report['Failed Import']):
        text += '\n\nThe following {} files were generated correctly but could not be uploaded to connect:' \
                '\n{}'.format(len(Report['Failed Import']),
                              '\n'.join(str(line) for line in Report['Failed Import']))
    if len(Report['Dismiss']):
        text += '\n\nThe following {} requests were dismissed because already available codes:' \
                '\n{}'.format(len(Report['Dismiss']), '\n'.join(str(line) for line in Report['Dismiss']))
    if len(Report['Missing Access Codes and URL']):
        text += '\n\nThe following {} files were missing Access Codes and URL (# missing codes):' \
                '\n{}'.format(len(Report['Missing Access Codes and URL']),
                              '\n'.join(str(line) for line in Report['Missing Access Codes and URL']))
    if len(Report['Missing Access Codes']):
        text += '\n\nThe following {} files were missing Access Codes (# missing codes):' \
                '\n{}'.format(len(Report['Missing Access Codes']),
                              '\n'.join(str(line) for line in Report['Missing Access Codes']))
    if len(Report['Missing URL']):
        text += '\n\nThe following {} files were missing URL:' \
                '\n{}'.format(len(Report['Missing URL']), '\n'.join(str(line) for line in Report['Missing URL']))
    if len(Report['Run out of codes']):
        text += '\n\nThe following {} files Run out of Access Codes. Must upload and request missing codes:' \
                '\n{}'.format(len(Report['Run out of codes']),
                              '\n'.join(str(line) for line in Report['Run out of codes']))
    if len(Report['eCampus Content Holding']):
        text += '\n\nThe following {} files have eCampus Content Holding Publisher with no available codes.' \
                '\nContact Maka for new codes:' \
                '\n{}'.format(len(Report['eCampus Content Holding']),
                              '\n'.join(str(line) for line in Report['eCampus Content Holding']))

    f = open(file_name, 'w')
    f.write(text)
    f.close()

    print(text)


def move_csv_file(Error, Check_file, School, Catalog, save_path, access_codes_file, process):
    if (not Error) & (not Check_file):
        # Make new directory to save file without errors
        new_dir = time.strftime('Access-Codes/%Y/%m/%d/{}/Good/{}/{}/'.format(process, School, Catalog))
        os.makedirs(save_path + new_dir, exist_ok=True)
        os.rename(str(access_codes_file), save_path + new_dir + str(access_codes_file))
        print('File moved to {}\n'.format(new_dir))
    elif Error:
        # Make new directory to save file with errors
        new_dir = time.strftime('Access-Codes/%Y/%m/%d/{}/{}/{}/{}/'.format(process, Error, School, Catalog))
        os.makedirs(save_path + new_dir, exist_ok=True)
        os.rename(str(access_codes_file), save_path + new_dir + str(access_codes_file))
        print('File moved to {}\n'.format(new_dir))
    elif Check_file:
        # Make new directory to save file for manual check
        new_dir = time.strftime('Access-Codes/%Y/%m/%d/{}/Ruby Error/{}/{}/'.format(process, School, Catalog))
        os.makedirs(save_path + new_dir, exist_ok=True)
        os.rename(str(access_codes_file), save_path + new_dir + str(access_codes_file))
        print('\nFile moved to {}\n'.format(new_dir))
