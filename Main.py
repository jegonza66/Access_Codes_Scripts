import Setup
import Paths_Credentials
import Functions
import Access_Codes_process

# Ask for desired process to run (Daily Delta or Low Notification)
process_todo = Functions.choose_process()

# Define Credentials for later use (API Key, BNED folder path, Verba credentials)
Credentials = Paths_Credentials.API_Paths()

# Define Verba username and password and save in credentials file
Credentials = Paths_Credentials.Verba_Credentials(Credentials)

if process_todo == 'Daily Delta':
    # Run Code Reveal Cases
    print('\nRunning Code reveal Cases\n')
    Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, run, DD = \
        Setup.code_reveal(Credentials=Credentials)
    if run:
        Report, driver = Access_Codes_process.run_code_reveal(Credentials=Credentials, Billing_ISBNs=Billing_ISBNs,
                                                              VBIDs=VBIDs, quantities=quantities, Schools=Schools,
                                                              Verba_Schools=Verba_Schools, Catalogs=Catalogs,
                                                              process_todo=process_todo + ' Code Reveal')
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')

    # Run Fake Code Reveal Cases
    print('\nRuning Fake Code Reveal Cases\n')
    Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, run = \
        Setup.fake_code_reveal(Credentials=Credentials, DD=DD)
    if run:
        Report, driver = Access_Codes_process.run_fake_code_reveal(Credentials=Credentials, Billing_ISBNs=Billing_ISBNs,
                                                                   VBIDs=VBIDs, quantities=quantities, Schools=Schools,
                                                                   Verba_Schools=Verba_Schools,Catalogs=Catalogs,
                                                                   driver=driver,
                                                                   process_todo=process_todo + ' Fake Code Reveal')
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')


elif process_todo == 'Low Notification':
    Billing_ISBNs, VBIDs, quantities, Schools, Verba_Schools, Catalogs, Publishers, Titles, URLs, run = \
        Setup.low_notification(Credentials=Credentials)
    if run:
        Report, driver = Access_Codes_process.run_low_notification(Credentials=Credentials,
                                                                   Billing_ISBNs=Billing_ISBNs, VBIDs=VBIDs,
                                                                   quantities=quantities, Schools=Schools,
                                                                   Verba_Schools=Verba_Schools,Catalogs=Catalogs,
                                                                   Publishers=Publishers, Titles=Titles, URLs=URLs,
                                                                   process_todo=process_todo)
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')


elif process_todo == 'Special request':
    Billing_ISBNs, VBIDs, quantities, Verba_Schools, Catalogs, Publishers, run = \
        Setup.special_request()
    if run:
        Report, driver = Access_Codes_process.run_special_request(Credentials=Credentials,
                                                                   Billing_ISBNs=Billing_ISBNs, VBIDs=VBIDs,
                                                                   quantities=quantities,
                                                                   Verba_Schools=Verba_Schools, Catalogs=Catalogs,
                                                                   Publishers=Publishers, process_todo=process_todo)
    else:
        print('\nRun == False.\nLenght of ISBNs, quantities, URLs did not match. Please check excel files.')
