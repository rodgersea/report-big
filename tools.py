
import pandas as pd
import numpy
from inspect import getframeinfo, stack
from PyPDF2 import PdfFileMerger, PdfFileReader
from PIL import Image

# TOOLS
# ----------------------------------------------------------------------------------------------------------------------
# vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv


# input: "paths" is passed as a list of pdf file names
# output: single pdf file saved to "output"
def merge_pdfs(paths, output):
    merged_object = PdfFileMerger()
    for pdff in paths:
        merged_object.append(PdfFileReader(pdff, strict=False), 'rb')
    merged_object.write(output)


# input: image path
# output: pdf with file name as "image path".pdf
def img2pdf(fdr_nm):
    imgpat = fdr_nm + '.jpg'
    pdfpat = fdr_nm + '.pdf'
    img1 = Image.open(imgpat)
    img2 = img1.convert('RGB')
    img2.save(pdfpat)
    del img2


# input: dictionary
# input: query
# output: key, value pair of matching entries in dictionary
def search_arr(dic, quer):  # this looks wrong
    for j in dic:
        if str(quer) in str(j):
            return j, dic[j]


# input: detail is the print string
# input: theobject is the object to be printed pretty
# output: detail, line number dispp is called, object type, object contents
def dispp(detail, theobject='nan'):
    print('___________________________________________________________________________________________________________')
    caller = getframeinfo(stack()[1][0])  # get line number where dispp is called
    print(detail,
          '\nline no:', caller.lineno,
          '\ntype:', ''.join([char for char in str(type(theobject)).split()[-1]][1:-2]))

    if isinstance(theobject, dict):  # print instructions for dictionary
        for key, value in theobject.items():
            print(key, '\n', value, '\n')
        print('')
        return
    if isinstance(theobject, str):  # print instructions for string
        print(theobject)
        print('^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^')
        return
    if isinstance(theobject, int):  # print instructions for string
        print(theobject)
        print('^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^')
        return
    if isinstance(theobject, list):  # print instructions for list
        print('len: ', len(theobject))
        for x in range(len(theobject)):
            print(theobject[x])
        print('^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^')
        return
    if isinstance(theobject, pd.DataFrame):
        with pd.option_context('display.max_rows', None, 'display.max_columns', None):
            print(theobject)
        print('^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^')
        return
    if isinstance(theobject, numpy.ndarray):
        print(theobject)
        print('^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^')
        return
    if isinstance(theobject, tuple):
        print(str(theobject))
        print('^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^')
        return


# input: list of values
# output: list of unique values
def unq_lis(listt):
    lis1 = []
    lis2 = ''
    for i in listt:
        if i not in lis1:
            lis1.append(i)
    for j in lis1:
        lis2 += str(j) + ', '
    return str(lis2[0:-2])

