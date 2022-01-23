import time
import os

import pandas as pd
import tkinter as tk
from tkinter import filedialog
import pickle


def choose_process():
    Answer = input('\nWhat process would you like to run?\n'
                   '1. BNED Daily Delta\n'
                   '2. Low Notification\n'
                   '3. No Notification\n'
                   'Enter the corresponding number or name and press Enter.')

    DD_selection = {'1', 'Daily Delta', 'dd', 'DD'}
    Low_notification_selection = {'2', 'low', 'Low Notification'}
    No_notification_selection = {'3', 'No', 'No Notification'}
    if Answer in DD_selection:
        process = 'Daily Delta'
    elif Answer in Low_notification_selection:
        process = 'Low Notification'
    elif Answer in No_notification_selection:
        process = 'No Notification'

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


def check_file(df, quantity, URL):
    Error = False

    if not pd.isna(URL):
        df['URL'] = URL

    codes_generated = int(quantity - df['Access Code'].isnull().values.sum())
    url_generated = int(quantity - df['URL'].isnull().values.sum())
    Missing_codes = int(quantity - codes_generated)
    Missing_url = int(quantity - url_generated)

    if Missing_codes and Missing_url:
        Error = 'Access Codes and URL'
        if codes_generated and codes_generated == url_generated:
            Error = 'Run out of codes'
            df = df.head(codes_generated)
    elif df['Access Code'].isnull().values.any():
        Error = 'Access Codes'
        if codes_generated:
            Error = 'Run out of codes'
            df = df.head(codes_generated)
    elif df['URL'].isnull().values.any():
        Error = 'URL'

    return df, Error, Missing_codes


def append_to_report(Report, File_imported, Error, Missing_codes, School, Catalog, Billing_ISBN, quantity):
    if File_imported:
        Report[File_imported].append([School, Catalog, Billing_ISBN, quantity-Missing_codes])
    else:
        Report[Error].append([School, Catalog, Billing_ISBN, Missing_codes])
    return Report


def write_report(Report, save_path, process):
    new_dir = save_path + time.strftime('Access-Codes/%Y/%m/%d/Reports/')
    os.makedirs(new_dir, exist_ok=True)
    file_name = time.strftime(new_dir + '{} Report %H_%M.txt'.format(process))
    text = 'The following {} files where correctly uploaded to connect: {}' \
           '\n\nThe following {} files where generated correctly but could not be uploaded to connect: {}' \
           '\n\nThe following files where missing information, please contact triage after checking there ' \
           'where no previous request on the ISBNs.' \
           '\n\n{}: Missing Access Codes and URL (# missing codes): {}' \
           '\n\n{}: Missing Access Codes (# missing codes): {}' \
           '\n\n{}: Missing URL: {}' \
           '\n\n{}: Run out of Access Codes (# missing codes): {}'\
        .format(len(Report['OK']), Report['OK'], len(Report['Failed Import']), Report['Failed Import'],
                len(Report['Access Codes and URL']), Report['Access Codes and URL'], len(Report['Access Codes']),
                Report['Access Codes'], len(Report['URL']), Report['URL'], len(Report['Run out of codes']),
                Report['Run out of codes'])

    f = open(file_name, 'w')
    f.write(text)
    f.close()


def write_final_report(Report, save_path, process):
    new_dir = save_path + time.strftime('Access-Codes/%Y/%m/%d/')
    os.makedirs(new_dir, exist_ok=True)
    file_name = time.strftime(new_dir + '{} Final Report %H_%M.txt'.format(process))
    text = 'The following {} files where correctly uploaded to connect: {}' \
           '\n\nThe following {} files where generated correctly but could not be uploaded to connect: {}' \
           '\n\nThe following files where missing information, please contact triage after checking there ' \
           'where no previous request on the ISBNs.' \
           '\n\n{}: Missing Access Codes and URL (# missing codes): {}' \
           '\n\n{}: Missing Access Codes (# missing codes): {}' \
           '\n\n{}: Missing URL: {}' \
           '\n\n{}: Run out of Access Codes (# missing codes): {}'\
        .format(len(Report['OK']), Report['OK'], len(Report['Failed Import']), Report['Failed Import'],
                len(Report['Access Codes and URL']), Report['Access Codes and URL'], len(Report['Access Codes']),
                Report['Access Codes'], len(Report['URL']), Report['URL'], len(Report['Run out of codes']),
                Report['Run out of codes'])

    f = open(file_name, 'w')
    f.write(text)
    f.close()


def move_csv_file(Error, Check_file, School, Catalog, save_path, access_codes_file):
    if (not Error) & (not Check_file):
        # Make new directory to save file without errors
        new_dir = time.strftime('Access-Codes/%Y/%m/%d/Good/{}/{}/'.format(School, Catalog))
        os.makedirs(save_path + new_dir, exist_ok=True)
        os.rename(str(access_codes_file), save_path + new_dir + str(access_codes_file))
        print('File moved to {}'.format(new_dir))
    elif Error:
        # Make new directory to save file with errors
        new_dir = time.strftime('Access-Codes/%Y/%m/%d/Missing {}/{}/{}/'.format(Error, School, Catalog))
        os.makedirs(save_path + new_dir, exist_ok=True)
        os.rename(str(access_codes_file), save_path + new_dir + str(access_codes_file))
        print('File moved to {}'.format(new_dir))
    elif Check_file:
        # Make new directory to save file for manual check
        new_dir = time.strftime('Access-Codes/%Y/%m/%d/Ruby Error/{}/{}/'.format(School, Catalog))
        os.makedirs(save_path + new_dir, exist_ok=True)
        os.rename(str(access_codes_file), save_path + new_dir + str(access_codes_file))
        print('\nFile moved to {}'.format(new_dir))
