Dim args, objExcel

Set objExcel = CreateObject("Excel.Application")
Set args = wscript.Arguments

objExcel.Workbooks.Open args(O) ' Capital letter "O"
objExcel.visible = True

objExcel.Run "Expiry"
objExcel.Run "Backup"
objExcel.ActiveWorkbook.Save

wscript.sleep 5000 'sleep for 5sec 'this code can be used to pause and check the terminal window activity

ObjExcel.Activeworkbook.Close(O) ' Capital letter "O"
ObjExcel.Quit