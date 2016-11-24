Dim qtApp ' Declare the Application object variable
Dim res
'Create the QTP Application object
Set qtApp = CreateObject("QuickTest.Application") 

'Make the QuickTest application visible
qtApp.Visible = True

'Make Sure about your script path  and script name in QC
'qtApp.Open "c:\CD Automation Project Backup\Run Full Regression\Regression Runner4", False
qtApp.Open "c:\CD Automation Project Backup\Run Full Regression\Regression Runner5", False
res = Cstr(Now())
res = Replace(res,":","_")
res = Replace(res, "/","_")
res = Replace(res, "\","_")
' Setting the result folder with TimeStamp
Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
qtResultsOpt.ResultsLocation = "C:\Test\" & res ' Set the 'results location

qtApp.Test.Run qtResultsOpt,true ' Run the test
'Close QTP
qtApp.quit
'Release Object
Set qtApp = Nothing


















'Kill_QTP_Process
Sub Kill_QTP_Process()
	strComputer = "."
	strProcessToKill = "QTPro.exe" 
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\"& strComputer & "\root\cimv2") 
	Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")
	count = 0
	For Each objProcess in colProcess
		objProcess.Terminate()
		count = count + 1
	Next 
End Sub
