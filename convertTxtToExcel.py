#!/mnt/lustre2/BI_Tools/tools/anaconda2/bin/python2.7

## Convert stat.txt to Excel
## Coding by Eunjung Kim
## Date: 2018. 02. 20

import pandas as pd
import glob, os
import openpyxl
import xlwt
from xlwt import Workbook, Formula, easyxf

os.chdir(".")
for stat in glob.glob("*_stat.txt"):            ###### ex.  stat = 1712UNHP-0001_stat.txt
        num_lines = sum(1 for line in open(stat))
#       print num_lines
        word = stat.split("_")                  ###### ex.  word = 1712UHNP-0001_stat
        fxls = word[0]+".xlsx"                  ###### ex.  fxls = 1712UNHP-0001_stat.xlsx

        ###### Read OrderID_stat.txt
        txt = pd.read_csv(stat, sep='\t')

        ###### Create OrderID_stat.xlsx
        writer = pd.ExcelWriter(fxls, engine='xlsxwriter')
        txt.to_excel(writer, sheet_name = 'Stat', index=False)

        ###### write worksheet
        workbook = writer.book
        worksheet1 = writer.sheets['Stat']

        ###### Format
        header_format = workbook.add_format({
        'bold': True,
        'valign': 'top',
        'fg_color': '#CCCCFF',
        'border': 1
         })

        for col_num, value in enumerate(txt.columns.values):
                worksheet1.write(0, col_num, value, header_format)

for md5 in glob.glob("*_download.txt"):         ###### ex.      md5     = 1712UNHP-0001_download.txt
        word2 = md5.split(".")                  ###### ex.      word2   = 1712UNHP-0001_downdload

        ###### Read ${orderID}_downdload.info
        txt2 = pd.read_csv(md5, sep='\t')

        ###### Create the second sheet in ${orderID}_stat.xlsx
        txt2.to_excel(writer, sheet_name = 'Download Info', index=False)

        ###### Write the worksheet
        workbook = writer.book
        worksheet2 = writer.sheets['Download Info']

        for col_num, value in enumerate(txt2.columns.values):
                worksheet2.write(0, col_num, value, header_format)


writer.save()

print stat, "is converted to", fxls
