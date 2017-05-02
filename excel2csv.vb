Imports System
Public Module modmain
   Sub Main()
	 Dim excelfile As String, sheetname As String, topleft As String, csvfile As String
	 Console.WriteLine ("Path to Excel File?	[eg. c:\example.xlsx]")
	 excelfile = Console.Readline ()
	 Console.WriteLine (excelfile)
	 Console.WriteLine ("Name of sheet to export?	[eg. Sheet1]")
	 sheetname = Console.Readline ()
	 Console.Writeline (sheetname)
	 Console.WriteLine ("Top-left cell of range to export?	[eg. A1]")
	 topleft = Console.Readline ()
	 Console.Writeline (topleft)
	 Console.WriteLine ("Path to output CSV file?	[eg. c:\example.csv]")
	 csvfile = Console.Readline ()
	 Console.Writeline (csvfile)
	 Dim xl As Object
	 xl = CreateObject("Excel.Application")
	 ExportToCsv (xl, excelfile, sheetname, topleft, csvfile)
	 Console.Writeline ("Done")
   End Sub
   
   'Usage: ExportToCsv "c:\wamp\www\vb\inventory.xlsm", "Inventory List", "C3", "C:\wamp\www\vb\inventory.csv"
    Sub ExportToCsv(xl As Object, excelfile As String, sheetname As String, topleft As String, csvfile As String)
    '
    ' Macro1 Macro
    '

    '
    Console.Writeline ("1")
        xl.DisplayAlerts = False
    Console.Writeline ("2")
        xl.Workbooks.Open (excelfile)
    Console.Writeline ("3")
		Dim firstcell As Object
        firstcell = xl.Range("'" & sheetname & "'!" & topleft)      'eg. 'Inventory List'!C3
    Console.Writeline ("4")
		Dim lastcell As Object
		lastcell = firstcell.SpecialCells(11)	'xlCellTypeLastCell
    Console.Writeline ("5")
		xl.Range(firstcell, lastcell).Copy        'xlLastCell
    Console.Writeline ("7")
        xl.Workbooks.Add
    Console.Writeline ("8")
        xl.Selection.PasteSpecial (-4163, -4142, False, False)     'xlPasteValues=-4163, xlNone=-4142
    Console.Writeline ("9")
        xl.Application.CutCopyMode = False
    Console.Writeline ("10")
        xl.ActiveWorkbook.SaveAs (csvfile, 6, False)   'xlCSV=6
    Console.Writeline ("11")
        xl.ActiveWindow.Close
    Console.Writeline ("12")
        xl.Quit
    End Sub

End Module



