from openpyxl import Workbook, load_workbook

wb = load_workbook('excel3.xlsx')
ws = wb.active

ws.move_range('A3:E4', rows=2, cols= 2) #移動資料，起始、橫移、直移



wb.save('excel4.xlsx')