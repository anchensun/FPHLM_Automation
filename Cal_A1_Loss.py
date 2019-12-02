import csv
import xlwt

with open('/a/mitch.cs.fiu.edu./disk/mitch-b/dmis-research/Anchen/2019/a1/processing/pr_unmatched/env/Results/A1_output.csv', newline='') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
    '''
    Total_Lms = LMs * Structure_Loss_Cost (Column F * Column J)
    Total_LMapp = LMapp * App_Loss_cost (Column G * Column K)
    Total_LMc = LMc * Contents_Loss_Cost (Column H * Column L)
    Total_Lmale = LMale * ALE_Loss_Cost (Column I * Column M)
    Total_Loss = Total_Lms + Total_LMapp + Total_LMc + Total_Lmale (Column N + Column O + Column P + Column Q)
    Loss_Cost = Total_Loss / LMs ((Column R / Column F) * 1000)
    '''
    row_num = 0
    row_write = []
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('A1_output')
    for row in spamreader:
        rowread = row
        row_num = row_num + 1
        if row_num == 1:
            rowread.append('Total_Lms')
            rowread.append('Total_LMapp')
            rowread.append('Total_LMc')
            rowread.append('Total_LMale')
            rowread.append('Total_Loss')
            rowread.append('Loss_Cost')
        else:
            Column_A = rowread[0]
            Column_B = int(rowread[1])
            Column_C = int(rowread[2])
            Column_D = rowread[3]
            Column_E = int(rowread[4])
            Column_F = int(rowread[5])
            Column_G = int(rowread[6])
            Column_H = int(rowread[7])
            Column_I = int(rowread[8])
            Column_J = float(rowread[9])
            Column_K = float(rowread[10])
            Column_L = float(rowread[11])
            Column_M = float(rowread[12])
            Column_N = (Column_F * Column_J)
            Column_O = (Column_G * Column_K)
            Column_P = (Column_H * Column_L)
            Column_Q = (Column_I * Column_M)
            Column_R = (Column_N + Column_O + Column_P + Column_Q)
            Column_S = (Column_R / Column_F) * 1000
            #row_write = [(Column_A, Column_B, Column_C, Column_D, Column_E, Column_F, Column_G, Column_H, Column_I, Column_J, Column_K, Column_L, Column_M, Column_N, Column_O, Column_P, Column_Q, Column_R, Column_S)]
            rowread.extend((Column_N, Column_O, Column_P, Column_Q, Column_R, Column_S))
        # Iterate over the data and write it out by row.
        col = 0
        for item in rowread:
            sheet.write((row_num - 1), col, rowread[col])
            col = col + 1
        #print(rowread)
    workbook.save('./A1_output.xls')
        
  
