import openpyxl
# https://pg-chain.com/python-excel-cell-write 参考

def return_write_cell(culumn,sheet,hour,minute):  #書き始めるcell決定用
    for i in range(5,1445):
        cell_name = culumn+str(i)
        data_value = sheet[cell_name].value
        write_hour = data_value.hour
        write_minute = data_value.minute
        if hour== write_hour and minute == write_minute:
            return('D'+str(i))
            break

file = openpyxl.load_workbook('no5__2017-12-06_08-25-28.xlsm') #　受け取りsheetファイル名を入れる
sheet = file['no5__2017-12-06_08-25-28'] # sheet名を入れprint(value)


file_name = 'marge_result.xlsm'
write_file = openpyxl.load_workbook('no5-1METS_20171206.xlsm')  #　書きこむファイル名を入力
write_sheet = write_file['no5-1METS_20171206'] #　書きこむsheet名を入力

# start position決定用
hour_minute_hour = sheet['D2'].value
second = hour_minute_hour.second   # 引くべき値を取得
ignore_value = 60-second+4   # 無視すべき値を取得

# 書きこむ場所決定用
hour = hour_minute_hour.hour
minute = hour_minute_hour.minute

# データ定義
start_cell = 'C'+str(ignore_value) #無視すべきあたいを抜いたスターするセルのいち
row_position = ignore_value
hr_value = 0
average = 0.000000
sum = 0
reset = False       #処理を終了するか否かの判定用

# 実際のメイン部分
for i in range(0,10):
    for a in range(0,59):
        row_position = int(row_position)+a
        cell_position = 'C'+str(row_position)
        hr_value = sheet[cell_position].value
        if hr_value == None:
            reset = True
            break
        sum += hr_value
    if reset == True:
        break
    average = round(sum/60)
    print(average)
    posi = return_write_cell('A',write_sheet,hour,minute)
    print(posi)
    write_sheet[posi] = average
    minute +=1
    if minute == 60:
        hour += 1
        minute = 0
    sum = 0
write_file.save(file_name)





