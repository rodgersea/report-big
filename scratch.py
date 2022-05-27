# input: elevation photos and xrf pos photos
# output: photo log as docx
def create_photo_log_docx(beholden):
    doc = docx.Document()  # create instance of a word document
    sections = doc.sections  # change the page margins
    for section in sections:  # set margins equal to 0 on all sides of doc container
        section.top_margin = Inches(0.25)
        section.bottom_margin = Inches(0)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # create elevation and xrf photo dic
    f_name = str(beholden[0]) + ' - ' + str(beholden[5] + ' - ' + str(beholden[6]))

    elev_lab_lis1 = []
    for x in range(4):
        elev_lab_lis1.append('Elevation ' + string.ascii_uppercase[x])
    arr_elev_pat1 = []
    for x in range(4):
        arr_elev_pat1.append(
            str('uploads/' + f_name + '/' + str(beholden[2]) + '_LBP/elevations/' + string.ascii_lowercase[x] + '.jpg'))

    xrf_lab_lis1 = []
    xrf_pos_pat_lis1 = os.listdir('uploads/' + f_name + '/' + str(beholden[2]) + '_LBP/xrf_Photos')
    xrf_pos_srt = []
    for x in range(len(xrf_pos_pat_lis1)):
        temp = re.findall(r'\d+', xrf_pos_pat_lis1[x])
        res = list(map(int, temp))[0]
        xrf_pos_srt.append(res)
    xrf_pos_srt_full = map(lambda y: 'Reading_' + str(y) + '_.jpg', sorted(xrf_pos_srt))
    xrf_pos_lab_full = map(lambda y: 'Reading ' + str(y), sorted(xrf_pos_srt))
    xrf_lab_lis1 = list(xrf_pos_srt_full)

    arr_xrf_pat = []
    for x in range(len(xrf_lab_lis1)):
        arr_xrf_pat.append('uploads/' + f_name + '/' + str(beholden[2]) + '_LBP/xrf_Photos/' + xrf_lab_lis1[x])

    lab_full1 = elev_lab_lis1 + list(xrf_pos_lab_full)
    pat_full = arr_elev_pat1 + arr_xrf_pat
    page_len = round(len(lab_full1) / 6)
    label_groups = [lab_full1[i:i + 6] for i in range(0, len(lab_full1), 6)]  # array of labels to the photos
    pat_groups1 = [pat_full[i:i + 6] for i in range(0, len(pat_full), 6)]  # array of paths to the photos

    table_arr = []
    header_arr = []
    brk_para_arr = []
    brk_run_arr = []
    # ------------------------------------------------------------------------------------------------------------------
    for x in range(page_len):
        # add header
        header_arr.append(doc.add_table(1, 2))
        header_widths = [7, 3]
        header_arr[x].alignment = WD_TABLE_ALIGNMENT.CENTER

        for i in range(1, 2):  # set table cell widths
            for cell in header_arr[x].columns[i].cells:
                cell.width = Inches(header_widths[i])
                cell.height = Inches(1.5)

        # left cell
        tab_head_left_cell = header_arr[x].cell(0, 0)
        par_head_left = tab_head_left_cell.paragraphs[0]
        run_head_left = par_head_left.add_run('Photo Log - ' + beholden[0])
        font_head = run_head_left.font
        font_head.size = Pt(14)
        run_head_left.font.color.rgb = RGBColor(0, 91, 184)
        par_head_left.alignment = 0

        # right cell
        tab_head_right_cell = header_arr[x].cell(0, 1)
        par_head_right = tab_head_right_cell.paragraphs[0]
        run_head_right = par_head_right.add_run()
        run_head_right.add_picture('lead_Pit/reporting/photo_Log/ei.jpg', width=Inches(2))
        par_head_right.alignment = 2
        par_head_right.add_run()

        # t_head formatting
        header_arr[x].cell(0, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        # ------------------------------------------------------------------------------------------------------------------

        lg = label_groups[x]
        pg = pat_groups1[x]
        table_arr.append(doc.add_table(len(lg) + 1, 2))
        table_arr[x].alignment = WD_TABLE_ALIGNMENT.CENTER
        table_arr[x].style = 'Table Grid'
        cell_arr = []
        par_arr = []
        run_arr = []
        for i in range(math.ceil(len(pg) / 2) * 2)[::2]:
            for j in range(2):
                try:
                    ipj = i + j
                    cell_arr.append(table_arr[x].cell(i, j))
                    cell_arr[ipj].width = Inches(3.5)
                    par_arr.append(cell_arr[ipj].paragraphs[0])
                    # par_arr[ipj].alignment = WD_ALIGN_PARAGRAPH.DISTRIBUTE
                    run_arr.append(par_arr[ipj].add_run())
                    run_arr[ipj].add_picture(pg[ipj], height=Inches(2.75))
                    last_p = doc.tables[-1].rows[-1].cells[-1].paragraphs[-1]
                    last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except:
                    pass

        cell_label_arr = []
        par_label_arr = []
        run_label_arr = []
        font_label_arr = []
        for i in range(math.ceil(len(pg) / 2) * 2)[1::2]:
            for j in range(2):
                try:
                    comb = i + j - 1
                    cell_label_arr.append(table_arr[x].cell(i, j))
                    par_label_arr.append(cell_label_arr[comb].paragraphs[0])
                    par_label_arr[comb].alignment = 1
                    run_label_arr.append(par_label_arr[comb].add_run(lg[comb]))
                    font_label_arr.append(run_label_arr[comb].font)
                    font_label_arr[comb].size = Pt(14)
                except:
                    pass
        skipper = 0
        for row in table_arr[x].rows:
            if skipper == 1:
                row.height = Pt(20)
                skipper = 0
            else:
                skipper = 1
        if page_len > 1 and x != (page_len - 1):
            brk_para_arr.append(doc.add_paragraph())
            brk_run_arr.append(brk_para_arr[x].add_run(''))

    # ------------------------------------------------------------------------------------------------------------------
    # SAVE DOCUMENT AS
    fl_pat = 'lead_Pit/LRA/finished_Docs'
    doc.save(str(fl_pat + '/' + beholden[0] + '/' + beholden[0] + '_photo_Log.docx'))
    # ------------------------------------------------------------------------------------------------------------------