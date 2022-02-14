import Low_No_notification_setup
import DD_setup
import Paths_Credentials
import Functions
import Access_Codes_process

# Ask for desired process to run (Daily Delta or Low Notification)
process_todo = Functions.choose_process()

# Define Credentials for later use (API Key, BNED folder path, Verba credentials)
Credentials = Paths_Credentials.API_Paths()

# Define Verba username and password and save in credentials file
Credentials, Automatic_Verba_upload = Paths_Credentials.Verba_Credentials(Credentials)

if process_todo == 'Daily Delta':
    print('\nRunning Code reveal Cases\n')
    Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, run, DD = \
        DD_setup.run_code_reveal(Credentials=Credentials)
    if run:
        Report, driver = Access_Codes_process.run_code_reveal(Credentials=Credentials, Billing_ISBNs=Billing_ISBNs,
                                                              VBIDs=VBIDs, quantities=quantities, Schools=Schools,
                                                              Verba_Schools=Verba_Schools, Catalogs=Catalogs,
                                                              Automatic_Verba_upload=Automatic_Verba_upload,
                                                              process_todo=process_todo + ' Code Reveal')
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')

    print('\nRuning Fake Code Reveal Cases\n')
    Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, run = \
        DD_setup.run_fake_code_reveal(Credentials=Credentials, DD=DD)
    if run:
        Access_Codes_process.run_fake_code_reveal(Credentials=Credentials, Billing_ISBNs=Billing_ISBNs, VBIDs=VBIDs,
                                                  quantities=quantities, Schools=Schools, Verba_Schools=Verba_Schools,
                                                  Catalogs=Catalogs, Automatic_Verba_upload=Automatic_Verba_upload,
                                                  driver=driver, process_todo=process_todo + ' Fake Code Reveal')
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')

elif process_todo == 'Low Notification':
    Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, Publishers, Titles, URLs, run = \
        Low_No_notification_setup.run_low_notification_setup(Credentials=Credentials)
    if run:
        Access_Codes_process.run_low_no_notification(Credentials=Credentials, Billing_ISBNs=Billing_ISBNs, VBIDs=VBIDs,
                                                     quantities=quantities, Schools=Schools, Verba_Schools=Verba_Schools,
                                                     Catalogs=Catalogs, Publishers=Publishers, Titles=Titles, URLs=URLs,
                                                     Automatic_Verba_upload=Automatic_Verba_upload,
                                                     process_todo=process_todo)
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')

elif process_todo == 'No Notification':
    Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, URLs, run = \
        Low_No_notification_setup.run_no_notification_setup(Credentials=Credentials)
    if run:
        Access_Codes_process.run_low_no_notification(Credentials=Credentials, Billing_ISBNs=Billing_ISBNs, VBIDs=VBIDs,
                                                     quantities=quantities, Schools=Schools, Verba_Schools=Verba_Schools,
                                                     Catalogs=Catalogs, URLs=URLs,
                                                     Automatic_Verba_upload=Automatic_Verba_upload,
                                                     process_todo=process_todo)
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')
