
from PyPDF2 import PdfFileReader, PdfFileWriter
from docx2pdf import convert
from func_repo import *
from tools import *

import pandas as pd
import pandoc
import fitz
import sys
import os

pd.options.display.max_columns = None  # display options for table
pd.options.display.width = None  # allows 'print table' to fill output screen
pd.options.mode.chained_assignment = None  # disables error caused by chained dataframe iteration


def make_it():
    # ----------------------------------------------------------------------------------------------------------------------
    # global variables
    insp_num = {'Elliott Rodgers': '110341',
                'Chris Ciappina': '120303',
                'Fabrizzio Simoni': '120304',
                'Parker Alvis': '120301',
                'Larry Rockefeller': '120291',
                'Lee Clark': '120065',
                'Ryan Bumpass': '120310',
                'Rob Campbell': '120302',
                'Tom Majkowski': '120166',
                'Brian Long': 'unknown'}  # inspector numbers
    name2sig = {'Elliott Rodgers': 'elliott_rodgers',
                'Chris Ciappina': 'chris_ciappina',
                'Fabrizzio Simoni': 'fabrizzio_simoni',
                'Parker Alvis': 'parker_alvis',
                'Larry Rockefeller': 'larry_rockefeller',
                'Lee Clark': 'lee_clark',
                'Ryan Bumpass': 'ryan_bumpass',
                'Rob Campbell': 'rob_campbell',
                'Tom Majkowski': 'tom_majkowski',
                'Brian Long': 'brian_long'}  # signature file name
    table_names = ['Table 1: Lead-Based Paint¹',
                   'Table 2: Deteriorated Lead-Based Paint¹',
                   'Table 3: Lead Containing Materials²',
                   'Table 4: Dust Wipe Sample Analysis',
                   'Table 5: Soil Sample Analysis',
                   'Table 6: Lead Hazard Control Options¹']
    proj_num = '220083.00'
    cwd = os.path.abspath(os.path.dirname(__file__))  # current working directory

    # loop through each app folder uploaded, if multiple uploaded at once
    for app in os.listdir(os.path.join(cwd, 'uploads')):  # loop through contents of "uploads" in root
        for paint in os.listdir(os.path.join(cwd, 'uploads', app)):  # get name of next directory for lead-based-paint docs
            if 'LBP' in paint:
                hold = paint
        bin_pat = os.path.join(cwd, 'uploads', app, hold, 'app_Data')  # bin_pat is the absolute path to "app_Data", the directory in each job folder that holds the working docs
        sched_sm = os.path.join(bin_pat, 'schedule')  # absolute path to the schedule directory
        sched_pat = os.path.join(sched_sm, os.listdir(sched_sm)[0])  # path to schedule.xlsx, a sheet that has information for each job done that week
        sched_hold = parse_excel(sched_pat)  # parse_excel cleans schedule.xlsx

        # create dataframe of info pertaining to each app number
        df = pd.DataFrame(sched_hold.loc[sched_hold['APP'] == app.split(' ')[0]])  # removes all rows that don't pertain to our working folders
        df = df.reset_index(drop=True)  # reset indices of df

        # call create_lra function on df to create report for working folders
        for index, row in df.iterrows():
            thero = row.to_numpy()  # convert df rows to arrays

            # commonly used paths and names
            app_data_pat = os.path.join(cwd, 'uploads', thero[0] + ' - ' + thero[5] + ' - ' + thero[6], thero[2] + '_LBP', 'app_Data')
            app_report_pat = os.path.join(cwd, 'finished_Docs', thero[0])

            if not os.path.exists(app_report_pat):  # create app folder in finished documents
                os.makedirs(app_report_pat)

            pdf_path1 = os.path.join(app_data_pat, 'lab_Results', os.listdir(os.path.join(app_data_pat, 'lab_Results'))[0])  # path to lab results

            gx = get_xrf(row.to_numpy())  # gx = raw xrf data

            # create separate xtab variables for creating tables as variables are modified for the different tables
            xtab = xrf_tables(gx, pdf_path1)
            xtab1 = xrf_tables(gx, pdf_path1)
            xtab2 = xrf_tables(gx, pdf_path1)

            # save xrf_clean.xlsx in app folder
            save_xrf_clean_xlsx(gx, thero)

            # create table 1: Lead Based Paint, save as table1_lbp.xlsx
            save_xrf_pos_xlsx(xtab, thero)

            xrf_clean_excel2pdf(gx, thero)  # save clean excel file as pdf, use beholden to get .xlsx path name

            create_photo_log_pdf(thero)

            create_lra(xtab1,  # dflis
                       thero,  # beholden
                       insp_num,  # from global variables
                       proj_num)  # from global variables

            create_lra_pdf(xtab1,  # dflis
                           thero,  # beholden
                           insp_num,  # from global variables
                           proj_num)  # from global variables

            create_lbpas(xtab2,  # dflis
                         thero,  # beholden
                         insp_num,  # from global variables
                         name2sig)  # from global variables

            wavelis = ['form_5.0_Page_1',
                       'form_5.0_Page_2',
                       'form_5.1']
            wavepath = []
            # get paths to these 3 forms
            for x in range(len(wavelis)):  # for x in range 3
                wavey = os.listdir(os.path.join(app_data_pat, wavelis[x]))  # wavey = form file name
                for y in wavey:  # if wavey has multiple files
                    if y[-4:] != '.pdf':  # if first file opened is not a pdf
                        img2pdf(os.path.join(app_data_pat, wavelis[x], y))  # then convert to pdf
                        wavepath.append(os.path.join(app_data_pat, wavelis[x], y[:-4] + '.pdf'))  # and append that pdf to the wavepath array
                        break  # end loop
                    else:
                        wavepath.append(os.path.join(app_data_pat, wavelis[x], y))  # if first file opened is pdf, appened to wavepath array
                        break  # end loop to avoid adding 2 files

            # get paths to floorplan, same method as wavepath forms
            for x in os.listdir(os.path.join(app_data_pat, 'floorplan')):
                if x[-4:] != '.pdf':
                    img2pdf(os.path.join(app_data_pat, 'floorplan', x))
                    floor_path = os.path.join(app_data_pat, 'floorplan', x[:-4] + '.pdf')
                    break
                else:
                    floor_path = os.path.join(app_data_pat, 'floorplan', x)
                    break

            # for some reason, the pdf with the lab results does not have the proper EOL documentation on the last page
            # for another unknown reason, splitting the last page from the main body of results, converting to png from a pixelmap and back to a pdf and recombining the pdfs solves this issue
            input1 = fitz.open(pdf_path1)  # input1 = lab_results.pdf
            page_end = input1.load_page(3)  # page_end is last page of results
            pix = page_end.get_pixmap()
            outputt = os.path.join(app_report_pat, 'res_end.png')
            pix.save(outputt)
            img1 = Image.open(outputt)
            img2 = img1.convert('RGB')
            img2.save(os.path.join(app_report_pat, 'res_end.pdf'))

            resmain_reader = PdfFileReader(pdf_path1)  # use pdffilereader to open and save the first 3 pages of the lab results as res_main.pdf
            resmain_writer = PdfFileWriter()  # pdffilewriter creates new pdf of the three pages
            for x in range(3):
                my_page = resmain_reader.getPage(x)
                resmain_writer.addPage(my_page)
            resmain_output = os.path.join(app_report_pat, 'res_main.pdf')  # path for save file
            with open(resmain_output, 'wb') as output:
                resmain_writer.write(output)  # save new pdf for first 3 pages of lab results as res_main.pdf

            merge_lis = [os.path.join(app_report_pat, row.to_numpy()[0] + '_LRA.pdf'),
                         os.path.join(cwd, 'reporting/LRA/attachments.pdf'),
                         os.path.join(cwd, 'reporting/LRA/floor_Plan.pdf'),
                         floor_path,
                         os.path.join(cwd, 'reporting/LRA/risk_Assessment.pdf'),
                         wavepath[0],
                         wavepath[1],
                         wavepath[2],
                         os.path.join(cwd, 'reporting/LRA/xrf_Photos.pdf'),
                         os.path.join(app_report_pat, thero[0] + '_xrf_clean.pdf'),
                         os.path.join(app_report_pat, thero[0] + '_photo_Log.pdf'),
                         os.path.join(cwd, 'reporting/LRA/lab_Results.pdf'),
                         os.path.join(app_report_pat, 'res_main.pdf'),
                         os.path.join(app_report_pat, 'res_end.pdf'),
                         os.path.join(cwd, 'reporting/LRA/method_all.pdf'),
                         os.path.join(cwd, 'reporting/LRA/lbpas.pdf'),
                         # os.path.join(app_report_pat, thero[0] + '_LBPAS.pdf'),
                         os.path.join(cwd, 'reporting/LRA/xrf_all.pdf'),
                         os.path.join(cwd, 'reporting/LRA/certs.pdf'),
                         os.path.join('reporting/Licensure/Lead', thero[1] + '.pdf'),
                         os.path.join(cwd, 'reporting/LRA/firm_license.pdf')]

            # merge all of the static and dynamic documents after everything has been created
            merge_pdfs(merge_lis, os.path.join(app_report_pat, thero[0] + '_LBP_Report_' + thero[11].strftime('%m%d%y') + '.pdf'))

            # if they exist, delete the res docs
            if os.path.exists(os.path.join(app_report_pat, 'res_end.png')):
                os.remove(os.path.join(app_report_pat, 'res_end.png'))
            if os.path.exists(os.path.join(app_report_pat, 'res_end.pdf')):
                os.remove(os.path.join(app_report_pat, 'res_end.pdf'))
            if os.path.exists(os.path.join(app_report_pat, 'res_main.pdf')):
                os.remove(os.path.join(app_report_pat, 'res_main.pdf'))

    print('make_it end')
