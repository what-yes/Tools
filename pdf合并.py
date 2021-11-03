"""
A simple python script to merge all the pdf files in the directory where this script is located.

@author: Yuanyang Shao
"""

import PyPDF2
import os
import re


def main():
    # find all the pdf files in current directory.
    mypath = "C:\\Users\\Administrator\\Desktop\\a"
    pattern = r"\.pdf$"
    file_names_lst = [mypath + "\\" + f for f in os.listdir(mypath) if re.search(pattern, f, re.IGNORECASE) and not re.search(r'Merged.pdf', f)]

    # merge the file.
    opened_file = [open(file_name, 'rb') for file_name in file_names_lst]
    pdfFM = PyPDF2.PdfFileMerger()
    for file in opened_file:
        pdfFM.append(file)

    # output the file.
    with open(mypath + "\\Merged.pdf", 'wb') as write_out_file:
        pdfFM.write(write_out_file)

    # close all the input files.
    for file in opened_file:
        file.close()


if __name__ == '__main__':
    main()