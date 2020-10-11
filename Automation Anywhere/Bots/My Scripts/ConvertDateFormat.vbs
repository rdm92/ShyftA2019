On Error resume next
Set Arg = WScript.Arguments
vFileSO = Arg(0)
vFolderCurrent = Arg(1)
vRow = Arg(2)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set File = FSO.CreateTextFile(vFolderCurrent & "\Date.txt",True)
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = getObject(vFileSO)
Set SOExcelSheet  = objExcel.ActiveSheet
Schld_date = SOExcelSheet.Cells(vRow, 4).Value
Schld_date = Right("00" & Month(Schld_date),2) & Right("00" & Day(Schld_date),2) & Right(Year(Schld_date),2)
Set objExcel = Nothing
File.Write Schld_date
File.Close
On error Goto 0