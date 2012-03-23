' ProcessKillLocal.vbs
' Sample VBScript to kill a program
' Author Guy Thomas http://computerperformance.co.uk/
' Version 2.7 - December 2010
' 22 March 2012: K Caulfield, Excelian Ltd. KillTrinity clone
' ------------------------ -------------------------------' 
Option Explicit
Dim objWMIService, objProcess, colProcess, colProperties, objFSO, objTextFile
Dim wshShell, wshNet
Dim strComputer, strProcessKill, strProcessesToKill
Dim strNameOfUser,strUserDomain
Dim strProcessName, strProcessOwner, strProcessHandle, strProcessSessionID, strProcessPID
Dim strLogonDomainUser, strNextLine
Dim arrTrinityProcesses, iProcessCount
Dim iProcesses
Const ForReading = 1

WSCript.Echo "-------------------------------------------"
WSCript.Echo "Kill Trinity Script - Standard Bank London."
WSCript.Echo "-------------------------------------------"

' ----------------------------------------------------
' Get List of Processes to Kill from KillProcess.txt
'-----------------------------------------------------

On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Not objFSO.FileExists("KillProcesses.txt") Then
	WSCript.Echo "-------------------------------------------"
	WScript.Echo "Input file " & "KillProcesses.txt" & " NOT FOUND."
	WSCript.Echo "-------------------------------------------"
	WScript.Quit(2)
End If

Set objTextFile = objFSO.OpenTextFile _
    ("KillProcesses.txt", ForReading)

If objTextFile.AtEndOfStream Then
	WSCript.Echo "-------------------------------------------"
	WScript.Echo "Input file " & "KillProcesses.txt" & " is EMPTY."
	WSCript.Echo "-------------------------------------------"
    WScript.Quit(2)
End If	

Do Until objTextFile.AtEndOfStream
	strNextLine = objTextFile.Readline
	arrTrinityProcesses = Split(strNextLine , ",")
	For iProcesses = 0 to Ubound(arrTrinityProcesses)
		Wscript.Echo "Process   : " & arrTrinityProcesses(iProcesses)
		strProcessesToKill = strProcessesToKill + "Name = '" & arrTrinityProcesses(iProcesses)
		strProcessesToKill = strProcessesToKill + "' OR "
	Next
Loop

objTextFile.Close

' ----------------------------------------------------
' Build SQL List of Processes, take off the last "OR"
'-----------------------------------------------------

strProcessesToKill = Left(strProcessesToKill,Len(strProcessesToKill)-4)
 
' Wscript.Echo strProcessesToKill

' ----------------------------------------------------
' Get session log details
'-----------------------------------------------------
Set wshNet = WScript.CreateObject("WScript.NetWork")

WSCript.Echo "-------------------------------------------"

strComputer = wshNet.ComputerName
WSCript.Echo "Server    : " & strComputer

strLogonDomainUser = UCase(wshNet.UserDomain & "\" & wshNet.UserName)
WSCript.Echo "Net User  : " & strLogonDomainUser

WSCript.Echo "-------------------------------------------"
					
' -----------------------------------------------------
' WMI Service Create
' -----------------------------------------------------

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _ 
& strComputer & "\root\cimv2") 

' -----------------------------------------------------
' Kill Trinity Processes for current user and session
' -----------------------------------------------------

'debug
WScript.Echo "Select * from Win32_Process Where " & strProcessesToKill

Set colProcess = objWMIService.ExecQuery _
("Select * from Win32_Process Where " & strProcessesToKill)

For Each objProcess in colProcess

	colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)

	strProcessName = objProcess.Name
	strProcessOwner = UCase(strUserDomain & "\" & strNameOfUser)
	strProcessHandle = objProcess.Handle
	strProcessPID = objProcess.ProcessId
	strProcessSessionID = objProcess.SessionID

	WSCript.Echo "-------------------------------------------"

	WSCript.Echo "Process   : " & strProcessName
	WSCript.Echo "Owner     : " & strProcessOwner
	WSCript.Echo "Handle/PID: " & strProcessHandle & "/" & strProcessPID
	WSCript.Echo "SessionID : " & strProcessSessionID

' -------------------------------------------------
' Inspect Process and determine whether to kill
' -------------------------------------------------

	If strProcessOwner = strLogonDomainUser Then
		WSCript.Echo "Kill?     : YES"
		objProcess.Terminate()
	Else
		WSCript.Echo "Kill?     : NO"
	End If

Next 

WSCript.Echo "-------------------------------------------"
WSCript.Echo "Kill Trinity Script ** END **"
WSCript.Echo "-------------------------------------------"


WScript.Quit 
' End of WMI KillTrinity Process