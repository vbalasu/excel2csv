import os, sys
curdir = os.getcwd() # os.path.dirname(os.path.abspath(__file__))
if(len(sys.argv) < 2):
  print('Syntax: excel2csv <Excel file name>')
  sys.exit()
from win32com.client import Dispatch
xl = Dispatch("Excel.Application")
xl.Visible = False # otherwise excel is hidden
xl.DisplayAlerts = False
filename = os.path.abspath(sys.argv[1])
print('Opening file '+filename)
wb = xl.Workbooks.Open(filename)
for n in wb.Sheets:
  n.Select()
  f=os.path.abspath(curdir+'/'+n.Name+'.csv')
  print(f)
  wb.SaveAs(Filename=f, FileFormat=62, CreateBackup=False)
wb.Close()
xl.Quit()
