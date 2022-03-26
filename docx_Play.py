
from docx2pdf import convert
from PyPDF2 import PdfFileReader, PdfFileWriter
from func_repo import *
from tools import *

import pandas as pd
import fitz
import sys
import os

pd.options.display.max_columns = None  # display options for table
pd.options.display.width = None  # allows 'print table' to fill output screen
pd.options.mode.chained_assignment = None  # disables error caused by chained dataframe iteration


def make_it():
    # ----------------------------------------------------------------------------------------------------------------------
    # global variables
    insp_num = {'Chris Ciappina': '120303',
                'Fabrizzio Simoni': '120304',
                'Parker Alvis': '120301',
                'Larry Rockefeller': '120291',
                'Lee Clark': '120065',
                'Ryan Bumpass': '120310',
                'Rob Campbell': '120302',
                'Tom Majkowski': '120166',
                'Brian Long': 'unknown',
                'Kevin': 'unknown'}  # inspector numbers
    name2sig = {'Chris Ciappina': 'chris_ciappina',
                'Fabrizzio Simoni': 'fabrizzio_simoni',
                'Parker Alvis': 'parker_alvis',
                'Larry Rockefeller': 'larry_rockefeller',
                'Lee Clark': 'lee_clark',
                'Ryan Bumpass': 'ryan_bumpass',
                'Rob Campbell': 'rob_campbell',
                'Tom Majkowski': 'tom_majkowski',
                'Brian Long': 'brian_long',
                'Kevin': 'unknown'}  # signature file name
    proj_num = '210289.00'

    # ----------------------------------------------------------------------------------------------------------------------
    # input: schedule as excel file
    # output: variable df as dataframe containing pertinent information to the list of app numbers
    # ----------------------------------------------------------------------------------------------------------------------
    for app in os.listdir('uploads'):
        for paint in os.listdir('uploads/' + app):
            if 'LBP' in paint:
                hold = paint
        bin_pat = 'uploads/' + app + '/' + str(hold)
        sched_pat = bin_pat + '/schedule/' + str(os.listdir(bin_pat + '/schedule')[0])
        sched_hold = parse_excel(sched_pat)

        # create dataframe of info pertaining to each app number
        df = pd.DataFrame(sched_hold.loc[sched_hold['APP'] == app.split(' ')[0]])
        df = df.reset_index(drop=True)  # reset indices of df
        # ----------------------------------------------------------------------------------------------------------------------
        # input: df1 as dataframe containing pertinent information to the list of app numbers
        # output: LRA for each job folder present in folder "job_Folders/"
        # output: full xrf data as excel file and pdf
        # output: tables 1-5 as sheets in excel file
        # ----------------------------------------------------------------------------------------------------------------------

        # call create_lra function on df1 to create LRA's for all apps in dataframe
        for index, row in df.iterrows():
            thero = row.to_numpy()
            dash_str = thero[0] + ' - ' + thero[5] + ' - ' + thero[6]

            pdf_filename = os.listdir('uploads/' + str(dash_str) + '/' + str(thero[2]) + '_LBP/lab_Results/')
            pdf_path1 = 'uploads/' + str(dash_str) + '/' + str(thero[2]) + '_LBP/lab_Results/' + pdf_filename[0]
            gx = get_xrf(row.to_numpy())  # gx = raw xrf data
            xtab = xrf_tables(gx, pdf_path1)  # xtab =

            # save xrf_clean.xlsx in app folder
            save_xrf_clean_xlsx(gx,
                                thero)
            print('TEST_______________________________________________________________________')
            sys.stdout.flush()

            # create table 1: Lead Based Paint, save as table1_lbp.xlsx
            save_xrf_pos_xlsx(xtab,
                              thero)

            xrf_clean_excel2pdf(gx, thero)  # save clean excel file as pdf, use beholden to get .xlsx path name

            create_photo_log(thero)
            pat_photo_log = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_photo_Log.docx'
            convert(pat_photo_log)

            create_lra(xtab,  # dflis
                       thero,  # beholden
                       insp_num,  # from global variables
                       proj_num)  # from global variables

            path_lra = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LRA.docx'
            convert(path_lra)

            create_lbpas(xtab,  # dflis
                         thero,  # beholden
                         insp_num,  # from global variables
                         name2sig)  # from global variables

            path_lbpas = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LBPAS.docx'
            convert(path_lbpas)

            wavelis = ['form_5.0_Page_1',
                       'form_5.0_Page_2',
                       'form_5.1']
            wavepath = []
            for x in range(len(wavelis)):
                wavepath.append('uploads/' + dash_str + '/' + thero[2] + '_LBP/' + wavelis[x] + '/' + os.listdir('uploads/' + dash_str + '/' + thero[2] + '_LBP/' + wavelis[x])[0])
            for x in range(len(wavepath)):
                wavepath[x] = wavepath[x][:-4]

            for x in wavepath:
                img2pdf(x)

            floor_path = ('uploads/' + dash_str + '/' + thero[2] + '_LBP/floorplan/' + os.listdir('uploads/' + dash_str + '/' + thero[2] + '_LBP/floorplan')[0])[:-4]
            img2pdf(floor_path)

            input1 = fitz.open(pdf_path1)
            page_end = input1.load_page(3)
            pix = page_end.get_pixmap()
            outputt = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.png'
            pix.save(outputt)

            pdfpat = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.pdf'
            img1 = Image.open(outputt)
            img2 = img1.convert('RGB')
            img2.save(pdfpat)

            resmain_reader = PdfFileReader(pdf_path1)
            resmain_writer = PdfFileWriter()
            for x in range(3):
                my_page = resmain_reader.getPage(x)
                resmain_writer.addPage(my_page)
            resmain_output = 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_main.pdf'
            with open(resmain_output, 'wb') as output:
                resmain_writer.write(output)

            merge_lis = ['lead_Pit/LRA/finished_Docs/' + row.to_numpy()[0] + '/' + row.to_numpy()[0] + '_LRA.pdf',
                         'lead_Pit/reporting/LRA/attachments.pdf',
                         'lead_Pit/reporting/LRA/floor_Plan.pdf',
                         floor_path + '.pdf',
                         'lead_Pit/reporting/LRA/risk_Assessment.pdf',
                         wavepath[0] + '.pdf',
                         wavepath[1] + '.pdf',
                         wavepath[2] + '.pdf',
                         'lead_Pit/reporting/LRA/xrf_Photos.pdf',
                         'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_xrf_clean.pdf',
                         'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_photo_Log.pdf',

                         'lead_Pit/reporting/LRA/lab_Results.pdf',
                         'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_main.pdf',
                         'lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.pdf',
                         'lead_Pit/reporting/LRA/method_all.pdf',
                         'lead_Pit/reporting/LRA/lbpas.pdf',

                         'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LBPAS.pdf',
                         'lead_Pit/reporting/LRA/xrf_all.pdf',
                         'lead_Pit/reporting/LRA/certs.pdf',
                         'lead_Pit/reporting/Licensure/Lead/' + thero[1] + '.pdf',
                         'lead_Pit/reporting/LRA/firm_license.pdf',

                         'lead_Pit/reporting/LRA/rebuild.pdf']

            merge_pdfs(merge_lis, 'lead_Pit/LRA/finished_Docs/' + thero[0] + '/' + thero[0] + '_LBP_Report.pdf')

            os.remove('lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.png')
            os.remove('lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_end.pdf')
            os.remove('lead_Pit/LRA/finished_Docs/' + thero[0] + '/res_main.pdf')

