import subprocess
import pathlib
import pandas as pd

import Functions
import Chrome_navigator

def run_low_no_notification(Credentials, Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, Publishers,
                            Titles, URLs, Automatic_Verba_upload, process_todo):

    # Define paths and credentials
    API_Key = Credentials['API_Key']
    save_path = Credentials['csv_save_path']
    if Automatic_Verba_upload:
        Verba_Username = Credentials['Verba_Username']
        Verba_Password = Credentials['Verba_Password']
        driver = Chrome_navigator.verba_connect_login(my_username=Verba_Username, my_password=Verba_Password)

    # Report variable for storing information
    Report = {'OK': [], 'Failed Import': [], 'Missing Access Codes': [], 'Missing URL': [],
              'Missing Access Codes and URL': [], 'Run out of codes': [], 'eCampus Content Holding': [], 'Dismiss': []}

    # Change directory to Access-Codes where ruby
    Functions.ruby_directory(Credentials=Credentials)

    print('\n\nTotal cases: {}\n'.format(len(Billing_ISBNs)))
    # Start run for
    for Billing_ISBN, VBID, Total_quantity, School, Verba_School, Catalog, Publisher, Title, URL in zip(Billing_ISBNs,
                                                                                           VBIDs, quantities, Schools,
                                                                                           Verba_Schools, Catalogs,
                                                                                           Publishers, Titles, URLs):

        File_imported = False
        Available_codes = Chrome_navigator.check_available_codes(driver=driver, Verba_School=Verba_School,
                                                                 Catalog=Catalog, ISBN=Billing_ISBN)

        # eCampus Content Holding exception
        if Publisher == 'eCampus Content Holding':
            quantity = Total_quantity
            Missing_codes = Total_quantity
            if Available_codes:
                # Dismiss
                print('Already available codes for Publisher: {}\nDismiss request.'.format(Publisher))
                Error = 'Dismiss'
            else:
                print('No available codes for Publisher: {}\n.'.format(Publisher))
                Error = 'eCampus Content Holding'

        else:
            if Total_quantity <= Available_codes:
                quantity = Total_quantity
                Missing_codes = 0
                print('Already {} available codes for a {} codes request.\nDismiss'.format(Available_codes, quantity))
                Error = 'Dismiss'

            elif Total_quantity > Available_codes:
                quantity = Total_quantity - Available_codes
                print('Already {} available codes for a {} codes request.\n{} Codes needed.'.format(Available_codes,
                                                                                              Total_quantity,
                                                                                              quantity))
                # Run access codes ruby command
                ruby_command = 'ruby vst_acs_v1.rb {} {} {} {}'.format(VBID, Billing_ISBN, quantity, API_Key)
                print('\nRunning\n' + ruby_command)
                process = subprocess.run(ruby_command, check=True)

                # if ruby returned an error check file manually
                if process.returncode:
                    Ruby_run_error = True
                else:
                    Ruby_run_error = False

                # Read generated file (checking if there is only one access codes file)
                access_codes_files = list(pathlib.Path().glob('access_codes_{}*.csv'.format(VBID)))
                if len(access_codes_files) == 1:
                    access_codes_file = access_codes_files[0]
                    df = pd.read_csv(access_codes_file)
                else:
                    # If there are more access codes files STOP
                    input('\nMultiple acces_codes files found with same VBID. Please check your directory and delete'
                          'corresponding file\nPress Enter to continue')
                    access_codes_files = list(pathlib.Path().glob('access_codes_{}*.csv'.format(VBID)))
                    access_codes_file = access_codes_files[0]
                    df = pd.read_csv(access_codes_file)

                # Check if Access Code file is OK
                df, Error, Missing_codes = Functions.check_file(df=df, csv_file=access_codes_file, quantity=quantity,
                                                                URL=URL, Title=Title)

                if (not Error) and (not Ruby_run_error) and Automatic_Verba_upload:
                    # Open Verba Connect/School/Sttings and drop import
                    File_imported = Chrome_navigator.automatic_verba_upload(driver=driver, csv_file=access_codes_file,
                                                                            Verba_School=Verba_School, Catalog=Catalog)

                # Move access codes file to history folder
                Functions.move_csv_file(Error=Error, Check_file=Ruby_run_error, School=School, Catalog=Catalog,
                                        save_path=save_path, access_codes_file=access_codes_file)

        # Add if this title run successfully or not to report
        Report = Functions.append_to_report(Report=Report, File_imported=File_imported, Error=Error,
                                            Missing_codes=Missing_codes, School=School, Catalog=Catalog,
                                            Publisher=Publisher, Title=Title, Billing_ISBN=Billing_ISBN,
                                            quantity=quantity)

        # Write and save report as txt file
        Functions.write_report(Report, save_path, process=process_todo)
    Functions.write_final_report(Report, save_path, process=process_todo)
    print('\nProcess: {} Run successfully'.format(process_todo))

    return Report, driver


def run_code_reveal(Credentials, Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, Automatic_Verba_upload,
        process_todo):

    # Define paths and credentials
    API_Key = Credentials['API_Key']
    save_path = Credentials['csv_save_path']
    if Automatic_Verba_upload:
        Verba_Username = Credentials['Verba_Username']
        Verba_Password = Credentials['Verba_Password']
        driver = Chrome_navigator.verba_connect_login(my_username=Verba_Username, my_password=Verba_Password)

    # Report variable for storing information
    Report = {'OK': [], 'Failed Import': [], 'Missing Access Codes': [], 'Missing URL': [],
              'Missing Access Codes and URL': [], 'Run out of codes': [], 'eCampus Content Holding': [], 'Dismiss': []}

    # Change directory to Access-Codes where ruby
    Functions.ruby_directory(Credentials=Credentials)

    print('\n\nTotal cases: {}\n'.format(len(Billing_ISBNs)))
    # Start run for
    for Billing_ISBN, VBID, Total_quantity, School, Verba_School, Catalog in zip(Billing_ISBNs, VBIDs, quantities,
                                                                                 Schools, Verba_Schools, Catalogs):

        File_imported = False
        Available_codes = Chrome_navigator.check_available_codes(driver=driver, Verba_School=Verba_School,
                                                                 Catalog=Catalog, ISBN=Billing_ISBN)

        if Total_quantity < Available_codes:
            quantity = Total_quantity
            Missing_codes = 0
            print('Already {} available codes for a {} request.\nDismiss'.format(Available_codes, quantity))
            Error = 'Dismiss'

        if Total_quantity > Available_codes:
            quantity = Total_quantity - Available_codes + 10
            print('Already {} available codes for a {} codes request.\n{} Codes needed.'.format(Available_codes,
                                                                                                Total_quantity,
                                                                                                quantity))

            # Run access codes ruby command
            ruby_command = 'ruby vst_acs_v1.rb {} {} {} {}'.format(VBID, Billing_ISBN, quantity, API_Key)
            print('\nRunning\n' + ruby_command)
            process = subprocess.run(ruby_command, check=True)

            # if ruby returned an error check file manually
            if process.returncode:
                Ruby_run_error = True
            else:
                Ruby_run_error = False

            # Read generated file (checking if there is only one access codes file)
            access_codes_files = list(pathlib.Path().glob('access_codes_{}*.csv'.format(VBID)))
            if len(access_codes_files) == 1:
                access_codes_file = access_codes_files[0]
                df = pd.read_csv(access_codes_file)
            else:
                # If there are more access codes files STOP
                input('\nMultiple acces_codes files found with same VBID. Please check your directory and delete'
                      'corresponding file\nPress Enter to continue')
                access_codes_file = access_codes_files[0]
                df = pd.read_csv(access_codes_file)

            # Check if Access Code file is OK
            df, Error, Missing_codes = Functions.check_file(df=df, csv_file=access_codes_file, quantity=quantity,
                                                            URL=float('nan'), Title='')

            if (not Error) and (not Ruby_run_error) and Automatic_Verba_upload:
                # Open Verba Connect/School/Sttings and drop import
                File_imported = Chrome_navigator.automatic_verba_upload(driver=driver, csv_file=access_codes_file,
                                                                        Verba_School=Verba_School, Catalog=Catalog)

            # Move access codes file to history folder
            Functions.move_csv_file(Error=Error, Check_file=Ruby_run_error, School=School, Catalog=Catalog,
                                    save_path=save_path, access_codes_file=access_codes_file)

        # Add if this title run successfully or not to report
        Report = Functions.append_to_report(Report=Report, File_imported=File_imported, Error=Error,
                                            Missing_codes=Missing_codes, School=School, Catalog=Catalog, Publisher='',
                                            Title='', Billing_ISBN=Billing_ISBN, quantity=quantity)

        # Write and save report as txt file
        Functions.write_report(Report, save_path, process=process_todo)
    Functions.write_final_report(Report, save_path, process=process_todo)
    print('\nProcess: {} Run successfully'.format(process_todo))

    return Report, driver


def run_fake_code_reveal(Credentials, Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs,
                         Automatic_Verba_upload, driver, process_todo):

    # Define paths and credentials
    save_path = Credentials['csv_save_path']

    # Report variable for storing information
    Report = {'OK': [], 'Failed Import': [], 'Missing Access Codes': [], 'Missing URL': [],
              'Missing Access Codes and URL': [], 'Run out of codes': [], 'eCampus Content Holding': [], 'Dismiss': []}

    print('\n\nTotal cases: {}\n'.format(len(Billing_ISBNs)))
    # Start run for
    for Billing_ISBN, VBID, quantity, School, Verba_School, Catalog in zip(Billing_ISBNs, VBIDs, quantities, Schools,
                                                                           Verba_Schools, Catalogs):

        # Define data to put in Billing ISBN and Access Code columns
        Billing_ISBN_col = [Billing_ISBN] * quantity
        Access_code_col = ['No access code is needed for this content, to access navigate back to your LMS.'] * quantity

        # Make Dataframe
        access_codes_file = pd.DataFrame(columns=['Billing ISBN', 'Access Code', 'URL'])
        access_codes_file['Billing ISBN'] = Billing_ISBN_col
        access_codes_file['Access Code'] = Access_code_col

        # Save Dataframe to csv in current folder
        file_name = 'access_codes_{}.csv'.format(Billing_ISBN)
        access_codes_file.to_csv(file_name, index=False)

        File_imported = False
        if Automatic_Verba_upload:
            # Open Verba Connect/School/Sttings and drop import
            File_imported = Chrome_navigator.automatic_verba_upload(driver=driver, csv_file=file_name,
                                                                    Verba_School=Verba_School, Catalog=Catalog)

        Error = False
        Ruby_run_error = False
        Missing_codes = 0
        # Add if this title run successfully or not to report
        Report = Functions.append_to_report(Report=Report, File_imported=File_imported, Error=Error,
                                            Missing_codes=Missing_codes, School=School, Catalog=Catalog, Publisher='',
                                            Title='', Billing_ISBN=Billing_ISBN, quantity=quantity)

        # Move access codes file to history folder
        Functions.move_csv_file(Error=Error, Check_file=Ruby_run_error, School=School, Catalog=Catalog,
                                save_path=save_path, access_codes_file=file_name)

        # Write and save report as txt file
        Functions.write_report(Report=Report, save_path=save_path, process=process_todo)
    Functions.write_final_report(Report=Report, save_path=save_path, process=process_todo)
    print('\nProcess: {} Run successfully'.format(process_todo))