import csv
import openpyxl
from openpyxl.styles import Border, Side
import pandas

#Read A1_Template and A1_Output

templatefile = openpyxl.load_workbook('/a/mitch.cs.fiu.edu./disk/mitch-b/dmis-research/Anchen/2019/a1/processing/2017FormA1.xlsx')

sheet = templatefile['Form A-1']
A1_output = pandas.read_excel('/a/mitch.cs.fiu.edu./disk/mitch-b/dmis-research/Anchen/2019/a1/processing/A1_output.xls')

with open('/a/mitch.cs.fiu.edu./disk/mitch-b/dmis-research/Anchen/2019/a1/processing/2017_zips_all.csv', newline='') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
    '''
    Select all columns and sort by zip column in ascending order
    Copy values from zip, county and fips columns to the previously open and empty Form A1 template.
    Paste in the corresponding columns.
    Close 2017_zips_all.csv without saving changes
    
    Open the file with the outputs and the added columns generated in step 1.
    Select column D (Const_type)
    Filter by each of the three construction types
    Copy the values from column S (Loss_Cost) into the corresponding column of the Form A1 template. Please note that the numbers of rows copied must match the number of rows in the template. This number is the number of valid zip codes
    Repeat the process for all the constructions types
    Save Form A1 template as a new file with a proper name identifying the date of the run and the form.
    '''
    #Prepare Data
    row_num = 0 
    zip_county_fips = []
    loss_cost = []
    A1_Cons = A1_output[['Cons_type']].values.tolist()
    A1_LossCost = A1_output[['Loss_Cost']].values.tolist()
    A1_ZipCode = A1_output[['ZipCode']].values.tolist()
   
    #Crop Zip data from zips_all
    for row in spamreader:
        if row_num != 0:
            rowread = row
            zip_county_fips.append((int(rowread[0]), rowread[3], int(rowread[4])))
        row_num = row_num + 1
    
    zip_county_fips.sort()

    #Crop Loss Data From A1_output
    for i in range(int(len(A1_Cons) / 3)):
        row_num = int(i * 3)    
        if (''.join(A1_Cons[row_num]) == 'Frame') and (''.join(A1_Cons[row_num + 1]) == 'Masonry') and (''.join(A1_Cons[row_num + 2]) == 'Manufactured'):
            loss_cost.append((A1_ZipCode[row_num][0], A1_LossCost[row_num][0], A1_LossCost[row_num + 1][0], A1_LossCost[row_num + 2][0]))

    loss_cost.sort()

    #Draw Border Template
    borderleft = Border(left=Side(border_style='medium',color='000000'),
                        right=Side(border_style='thin',color='000000'),
                        top=Side(border_style='thin',color='000000'),
                        bottom=Side(border_style='thin',color='000000'))

    borderright = Border(left=Side(border_style='thin',color='000000'),
                        right=Side(border_style='medium',color='000000'),
                        top=Side(border_style='thin',color='000000'),
                        bottom=Side(border_style='thin',color='000000'))

    borderleftbot = Border(left=Side(border_style='medium',color='000000'),
                           right=Side(border_style='thin',color='000000'),
                           top=Side(border_style='thin',color='000000'),
                           bottom=Side(border_style='medium',color='000000'))

    borderrightbot = Border(left=Side(border_style='thin',color='000000'),
                           right=Side(border_style='medium',color='000000'),
                           top=Side(border_style='thin',color='000000'),
                           bottom=Side(border_style='medium',color='000000'))

    borderbot = Border(left=Side(border_style='thin',color='000000'),
                       right=Side(border_style='thin',color='000000'),
                       top=Side(border_style='thin',color='000000'),
                       bottom=Side(border_style='medium',color='000000'))

    borderthin = Border(left=Side(border_style='thin',color='000000'),
                    right=Side(border_style='thin',color='000000'),
                    top=Side(border_style='thin',color='000000'),
                    bottom=Side(border_style='thin',color='000000'))

    #Fill out the form
    x = 9
    row = x
    for item in zip_county_fips:
        cell = 'A' + str(row)
        zipcode = str(item[0])
        if len(str(item[0])) < 5:
            for i in range(5 - len(str(item[0]))):
                zipcode = '0' + zipcode 
        sheet[cell] = zipcode
        sheet[cell].border = borderleft
        cell = 'B' + str(row)
        sheet[cell] = item[1]
        sheet[cell].border = borderthin
        cell = 'C' + str(row)
        sheet[cell] = item[2]
        sheet[cell].border = borderright
        row = row + 1
   
    row = x
    for item in loss_cost:
        cell = 'D' + str(row)
        sheet[cell] = item[1]
        sheet[cell].border = borderleft
        cell = 'E' + str(row)
        sheet[cell] = item[2]
        sheet[cell].border = borderthin
        cell = 'F' + str(row)
        sheet[cell] = item[3]
        sheet[cell].border = borderright
        row = row + 1

    #Draw Bottom Border
    row = row - 1
    sheet['A' + str(row)].border = borderleftbot
    sheet['B' + str(row)].border = borderbot
    sheet['C' + str(row)].border = borderrightbot
    sheet['D' + str(row)].border = borderleftbot
    sheet['E' + str(row)].border = borderbot
    sheet['F' + str(row)].border = borderrightbot

    #Save the form
    templatefile.save('/a/mitch.cs.fiu.edu./disk/mitch-b/dmis-research/Anchen/2019/a1/processing/2017FormA1_test.xlsx')
    
        
