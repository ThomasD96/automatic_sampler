#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
from datetime import date
import os
import glob
import random
import datetime
import openpyxl
import xlsxwriter
import getpass


# In[2]:


username = getpass.getuser()


# In[3]:


def my_sampler ():
    
    # Restarts function again and again
    while True:
        datum = date.today()
        cwd = os.getcwd()
        path = 'C:\\Users\\'+username+'\\OneDrive - KPMG\\Desktop\\samples_irm'
        os.chdir(path)
        list_of_files_excel = glob.glob("*.xlsx")
        list_of_files_svn_sap = glob.glob('*')
        df_select_len_svn_sap = len(list_of_files_svn_sap)

        if df_select_len_svn_sap > 1:
            print ()
            print ()
            print ('##################################################################')
            print ('####################### AUTOMATIC SAMPLER ########################')
            print ('##################################################################')
            print ('')
            print ('Attention: More than one file detected in the ordner "samples_irm" on your Desktop!')
            print ('Please, delete the other files and press ENTER to restart the program.')
            print ('Merci!')
            timer_sleep = str(input())

        else:
            print ()
            print ()
            print ()
            print ('Welcome to the KPMG sampling-tool!')
            print ()
            print ('Do you want to sample SVN-Logs / SAP-outputs or normal samples?')
            print ('Type in either << svn >> / << sap >> or << normal >>.')
            format_input = str(input()).lower()
            print ('')
            
            ############################################################
            ####################### Sampling of SVN-data ###############
            ############################################################
            
            if format_input == 'svn':
                df_select = list_of_files_svn_sap[0]
                with open(df_select, 'r') as f:
                    svn_list = [line.strip() for line in f]
                df_1_original = pd.DataFrame(svn_list)
                df_1_original.columns = ['svn']
                df_1 = df_1_original.copy()
                df_1 = df_1.iloc[:, [0]]   
                str_match_input = '[r]+\d+\s+[|]'
                df_1 = df_1.dropna()
                df_1 = df_1[df_1.svn.str.match(str_match_input)]
                df_1_size = list(df_1.shape)[0]
                
            ###########################################################
            ###################### Sampling of SAP-txt-data ###########
            ###########################################################
            
            elif format_input == 'sap':
                list_of_files = list(list_of_files_svn_sap)
                sap_txt = '.txt'
                matching = [y for y in list_of_files if sap_txt in y]
                data_txt = matching[0]

                f = open(data_txt)
                mylines = f.readlines()
                f.close()
                
                df = pd.DataFrame(mylines) 
                df.columns = ['col']
                df['col'] = df['col'].astype(str)
                df = df[~df['col'].str.startswith('-')]
                df = df[df['col'].str.startswith('|')]
                
                # replace incorrect encodings
                df = df.replace('Ã„', 'Ä', regex=True)
                df = df.replace('Ã¤', 'ä', regex=True)
                df = df.replace('Ã–', 'Ö', regex=True)
                df = df.replace('Ã¶', 'ö', regex=True)
                df = df.replace('Ãœ', 'Ü', regex=True)
                df = df.replace('Ã¼', 'ü', regex=True)
                df.to_csv('test.csv', header = False)

                # get first row - will be futere columns
                future_cols = df.iloc[0]
                future_cols = list(future_cols)
                b = ' '.join([str(elem) for elem in future_cols]) 
                dlim = '|'
                c = b.split(dlim)
                k = [x.replace(' ', '') for x in c]

                df = pd.read_csv('test.csv')
                df.columns = ['dummy', 'sap']
                df = df['sap'].str.split('|', expand = True)
                df.columns = k
                
                # Remove empty-column named '' and \n
                df = df.drop(['\n', ''], axis=1, errors='ignore')
               
                df_1_original = df.copy()
                df_1 = df_1_original.copy()
                df_1_size = list(df_1.shape)[0]
                os.remove('test.csv')
            
            
            #############################################################
            ################ Sampling of normal Excel files #############
            #############################################################
            
            else:
                df_select = list_of_files_excel[0]
                print ('-----------------------------------------------')
                print ('Should the first n-rows be skipped? - Yes or No')
                skip_rows = str(input())
                skip_rows = skip_rows.lower()

                if skip_rows == 'yes':
                    print (' ')
                    print ('-----------------------------------------')
                    print (' ')
                    print ('Type in how many rows should be skipped: ')
                    skip_rows_input = int(input())
                    print (' ')
                    print ('-----------')
                    df_1_original = pd.read_excel(df_select)
                    df_1 = df_1_original.copy()
                    df_1 = df_1.iloc[skip_rows_input:]
                    df_1_size = list(df_1.shape)[0]
                else:
                    df_1_original = pd.read_excel(df_select)
                    df_1 = df_1_original.copy()
                    df_1_size = list(df_1.shape)[0]

            if df_1_size == 0:
                print (' ')
                print ('----------------------------------------------------------------------------------------')
                print ('Row-count is {0} - therefore select << 0 >>'.format(df_1_size))
                print ('----------------------------------------------------------------------------------------')
            elif df_1_size in range(1, 2):
                kam_control = 'annual-frequency'
                print (' ')
                print ('----------------------------------------------------------------------------------------')
                print ('Row-count is {0} - therefore select << 1 >> according to KAM (annual-frequency)'.format(df_1_size))
                print ('----------------------------------------------------------------------------------------')
            elif df_1_size in range (2, 5):
                kam_control = 'quarterly_frequency'
                print (' ')
                print ('----------------------------------------------------------------------------------------')
                print ('Row-count is {0} - therefore select << 2 >>according to KAM (quarterly-frequency)'.format (df_1_size))
                print ('----------------------------------------------------------------------------------------')
            elif df_1_size in range (5, 13):
                kam_control = 'monthly-frequency'
                print (' ')
                print ('----------------------------------------------------------------------------------------')
                print ('Row-count is {0} - therefore select << 2 >> according to KAM (monthly-frequency)'.format (df_1_size))
                print ('----------------------------------------------------------------------------------------')
            elif df_1_size in range (13, 53):
                kam_control = 'weekly-frequency'
                print (' ')
                print ('----------------------------------------------------------------------------------------')
                print ('Row-count is {0} - therefore select << 5 >> according to KAM (weekly-frequency)'.format (df_1_size))
                print ('----------------------------------------------------------------------------------------')
            elif df_1_size in range (53, 366):
                kam_control = 'daily-frequency'
                print (' ')
                print ('----------------------------------------------------------------------------------------')
                print ('Row-count is {0} - therefore select << 15 >> according to KAM (daily-frequency)'.format (df_1_size))
                print ('----------------------------------------------------------------------------------------')
            else:
                kam_control = 'reccuring manual control'
                print (' ')
                print ('----------------------------------------------------------------------------------------')
                print ('Row-count is {0} - therefore select << 25 >> according to KAM (reccuring manual control)'.format (df_1_size))
                print ('---------------------------------------------------------------------------------------')
                

             ## Input-Daten
            print (' ')
            print ('How many samples should be selected? Type in number:')
            sample_number = int(input())
            print (' ')
            print ('------------------------------------------')
            print ('Type in the name of the control-performer.')
            name = str(input())
            print (' ')
            print ('----------------------------------')
            print ('Type in the name of the GITC-name.')
            control_name = str(input())


            ## Excel-Layout for Sample
            row_count = str(df_1_size)
            sample_nr = str(sample_number)

            list_1 = ['KPMG Austria GmbH Wirtschaftsprüfungs- und Steuerberatungsgesellschaft', 'Porzellangasse 51', '1090 Wien', 
                      'Tel: +43 (1) 313 32-0', '', '', 'Procedures for random sampling in Excel/Python',
                      'Step 1: ', 
                      'Step 2: ', 
                      'Step 3:' , 
                      'Step 4:' 
                      '', 
                      '',
                      '',
                      'Control-Performer', 
                      'Control-nr. (GITC)', 
                      'Sample-Anzahl', 
                      'Date']


            string_input_list_2 = ['Calculation of total population {number_1} elements --> {number_2} according to KAM ({kam})']
            string_input_list_2 = [w.replace('{number_1}', row_count) for w in string_input_list_2]
            string_input_list_2 = [w.replace('{number_2}', sample_nr) for w in string_input_list_2]
            string_input_list_2 = [w.replace('{kam}', kam_control) for w in string_input_list_2]
            string_input_list_2 = ' '.join([str(elem) for elem in string_input_list_2]) 


            list_2 = ['', '', '', '', '', '', '', 
                      'Copy of the data file into the Desktop-folder “samples_irm”', 
                      'Start of the programme “kpmg_sampler.exe”', 
                      'Based on the internal Python pseudorandom number generator (Mersenne Twister), a sample is generated',
                      string_input_list_2, 
                      '',
                      '',
                      name, 
                      control_name, 
                      sample_number, 
                      datum]


            df_general = pd.DataFrame(list(zip(list_1, list_2)), columns =['Name', 'val'])
            sampling = df_1.sample(n = sample_number)
            sampling = sampling.dropna(axis = 1, how = 'all')
            sampling['index'] = np.arange(1, len(sampling)+1)
            sampling.set_index('index', inplace = True)

            ## Excel-Layout for Work-Paper
            list_3 = ['Prepared by:', 'Date', '', 'Legend', '✓', 'x', 'n/a', '', 'Conclusion on TOE:', '', '']
            list_4 = [name, datum, '', '', 'no exceptions noted', 'exceptions noted', 'not applicable', '', '', 'operating effectively', 'NOT OPERATING EFFECTIVELY']
            df_general_wb = pd.DataFrame(list(zip(list_3, list_4)), columns =['Name', 'val'])


            ##########################################################################

            ## Save Excel-Worksheets: Sample
            writer_1 = pd.ExcelWriter('sample_output.xlsx',  engine ='xlsxwriter')
            df_general.to_excel(writer_1, sheet_name ='KPMG_Information', startrow = 0, startcol = 0, header = False, index = False)
            df_1_original.to_excel(writer_1, sheet_name = 'pbc', startrow = 0, startcol = 0, header = True, index = False)
            sampling.to_excel(writer_1, sheet_name ='Sample-Output', startrow = 0, startcol = 0, header = True, index = True)

            ## Save Excel-Worksheet: Workbook        
            writer_2 = pd.ExcelWriter('workpaper.xlsx',  engine ='xlsxwriter')
            df_general_wb.to_excel(writer_2, sheet_name ='KPMG_Information', startrow = 0, startcol = 0, header = False, index = False)
            sampling_wb = sampling.copy()
            sampling_wb.index = np.arange(1, len(sampling_wb)+1)
            sampling_wb['AP'] =  None
            sampling_wb['Comment'] =  None

            sampling_wb.to_excel(writer_2, sheet_name ='KPMG_Information', startrow = 15, startcol = 2, header = False, index = True)
            workbook  = writer_2.book
            worksheet = writer_2.sheets['KPMG_Information']
            header_format = workbook.add_format({'bold': True, 'center_across': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})

            for col_num, value in enumerate(sampling_wb.columns.values):
                worksheet.write(14, col_num + 3, value, header_format)


            ## Close Excel
            writer_1.save()
            writer_2.save()


# In[ ]:


my_sampler()


# In[ ]:





# In[ ]:




