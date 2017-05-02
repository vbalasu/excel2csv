(
REM Input Excel file
echo c:\wamp\www\vb\inventory.xlsm
REM Sheet name
echo Inventory List
REM Top left cell
echo C3
REM Output CSV file
echo c:\wamp\www\vb\inventory.csv
)>excel2csv.script
type excel2csv.script | excel2csv.exe