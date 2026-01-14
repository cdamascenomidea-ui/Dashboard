from openpyxl import load_workbook
wb = load_workbook('Dashboard_test.xlsx')
d = wb['Dashboard']
print('B3 formula:', d['B3'].value)
print('B4 formula:', d['B4'].value)
print('B5 formula:', d['B5'].value)
print('B6 formula:', d['B6'].value)
