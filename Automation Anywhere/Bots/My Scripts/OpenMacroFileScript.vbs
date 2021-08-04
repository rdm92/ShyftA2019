Set Arg = WScript.Arguments
macroFilePath = WScript.Arguments.Item(0)
Dim xlApp, xlBook

Set xlApp = CreateObject("Excel.Application")
xlApp.DisplayAlerts = False
Set xlBook = xlApp.Workbooks.Open(macroFilePath, 0, True)
xlApp.Application.Visible = True
WScript.Sleep 5000

xlBook.Close true
xlApp.Quit
WScript.Sleep 5000
Set xlBook = Nothing
Set xlApp = Nothing

'WScript.Echo "Finished."
WScript.Quit