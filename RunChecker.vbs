Dim diff,Date1,Date2,sTimeOut
Dim LockFilePath : LockFilePath = "C:\CD Automation Project Backup\Run Full Regression\WriteLock.txt"
Dim tempExcelFilePath : tempExcelFilePath =  "C:\CD Automation Project Backup\Run Full Regression\TempRunTimeExcel.xls"
Dim FSO:Set FSO = CreateObject("Scripting.FileSystemObject")
Dim qtApp  
Dim res
Dim sLogFilePath:sLogFilePath = "C:\CD Automation Project Backup\Run Full Regression\RunCheckerLog"&Replace(FormatDateTime(cDate(NOW),2),"/","_")&".txt"
Dim StrScriptName 
'msgbox sLogFilePath
Set qtApp = CreateObject("QuickTest.Application")  
' Default value if Timeout is not defined
sTimeOut = 90
While TRUE
	' Waits 5 minutes
	WScript.Sleep(300000)

	' Waits until excel file is released - according to the lock file 
	While FSO.FileExists(LockFilePath )
		Wscript.Sleep(1000)
	Wend
	' Getting the last executed script Start Time and current time (date1 and date 2 respectively)
	GetDatesFromExcel Date1, Date2 , sTimeOut 
	'Checks The Time Difference of the currently running script
	diff = Cint(DateDiff("n",Date1,Date2))

	' if time in minutes exeeded the retrieved timeout value for the script - stops the execution and delete the script relevant record from runfile
	On Error Resume NExt
' Change Script Timeout Here (minutes)
'	msgbox "Current TimeOut :" & sTimeOut
'sTimeOut  = 90
	If Cint(diff)>Cint(sTimeOut) Then
'		msgbox "Going to abort "
		WriteToLogFile "Failure : Stopping QTP Scenario - Timeout Exceeded For Scenario : " & CSTR(StrScriptName) & " ,Time passed since started: " & CSTR(diff) & ", Timeout Value is: " & CSTR(sTimeOut)
		WriteStopIndicatorToTempExecutionFile CSTR(StrScriptName)
		qtApp.Test.Stop  
		'closing QTP
		qtApp.quit:	WScript.Sleep(30000)
		' Restarting QTP
		qtApp.Launch :	WScript.Sleep(30000)
		res = Cstr(Now())
		res = Replace(res,":","_")
		res = Replace(res, "/","_")
		res = Replace(res, "\","_")
		'qtApp.Open "c:\CD Automation Project Backup\Run Full Regression\Regression Runner4", False
		qtApp.Open "c:\CD Automation Project Backup\Run Full Regression\Regression Runner5", False
		Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
		qtResultsOpt.ResultsLocation = "C:\Test\" & res ' Set the 'results location
		qtApp.Test.Run  qtResultsOpt,true
		Else
		WriteToLogFile "Info : Checking Scenario Timeout For Scenario :" & CSTR(StrScriptName) & " ,Time passed since started: " & CSTR(diff) & ", Timeout Value is: " & CSTR(sTimeOut)
	End If
	On Error GoTo 0
Wend


Sub GetDatesFromExcel(ByRef Date1,ByRef Date2,ByRef sTimeOut)
	 StrScriptName=""
	Dim i
	Dim Row_Count
			Set objExcel = CreateObject("Excel.Application")
			objExcel.WorkBooks.Open tempExcelFilePath
			Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
			Row_Count = objSheet.UsedRange.Rows.Count 
				
				'msgbox Cstr(Row_Count)
				For i=Row_Count To 1 step -1
					'msgbox "i is : "&CSTR(i)
					'msgbox objSheet.Cells(i, 2).Value
					'msgbox objSheet.Cells(i, 3).Value
					If  objSheet.Cells(i, 2).Value &""<>"" AND objSheet.Cells(i, 3).Value &""="" Then
						Date1 = objSheet.Cells(i, 2).Value
						Date2 = Now
						sTimeOut = objSheet.Cells(i, 5).Value
						StrScriptName = objSheet.Cells(i, 1).Value
						Exit For
					End If	
					'i = i -1
					
				Next
	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	Set objSheet = Nothing
	Set objExcel = Nothing	
End Sub

Sub ShowMessage(ByVal str)
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Popup str, 10
	Set objShell =Nothing
End Sub

'Writing to log file the given text
Sub WriteToLogFile(ByVal sText)
	Set LogFSO = CreateObject("Scripting.FileSystemObject")
	If NOT LogFSO.FileExists(sLogFilePath) Then
			Set objFile = LogFSO.CreateTextFile(sLogFilePath,TRUE)
		Else
		' For appending
			Set objFile = LogFSO.OpenTextFile(sLogFilePath,8,TRUE)
	End If
	objFile.WriteLine CSTR(NOW) & ":" & sText
	objFile.Close
End Sub
'Writing when scenario has been stopped to temp run time excel file
Sub WriteStopIndicatorToTempExecutionFile(ByVal sScriptName)
	Dim i
	Dim Row_Count
			Set objExcel = CreateObject("Excel.Application")
			objExcel.WorkBooks.Open tempExcelFilePath
			Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
			Row_Count = objSheet.UsedRange.Rows.Count 
				'msgbox Cstr(Row_Count)
				For i=Row_Count To 1 step -1
					'msgbox "i is : "&CSTR(i)
					'msgbox objSheet.Cells(i, 2).Value
					'msgbox objSheet.Cells(i, 3).Value
					If  objSheet.Cells(i, 1).Value = sScriptName Then
						objSheet.Cells(i, 3).Value = "Stopped"
						Exit For
					End If	
					'i = i -1
				Next
	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	Set objSheet = Nothing
	Set objExcel = Nothing	
End Sub