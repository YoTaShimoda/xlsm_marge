import openpyxl

file = openpyxl.load_workbook('no4__2017-12-06_08-25-29.xlsm') #sheetファイル名を入れる
sheet = file['no4__2017-12-06_08-25-29'] #sheet名を入れprint(value)
second = sheet['D2'].value.second   #引くべき値を取得
ignore_value = 60-second+4   #無視すべき値を取得
start_cell = 'C'+str(ignore_value)
row_position = ignore_value
hr_value = 0
average = 0.000000
sum = 0

for i in range(0,1440):
    for a in range(0,59):
        row_position = int(row_position)+a
        cell_position = 'C'+str(row_position)
        hr_value = sheet[cell_position].value
        sum += hr_value
    average = round(sum/60)
    print(average)
    sum = 0




