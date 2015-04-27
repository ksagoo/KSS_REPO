
'Usage "[ExecutablePath\]blSQLplus.exe /D:DOMAIN /S:[ScriptPath\]scriptname.sql /DB:<dbname> [/U:<db username> /P:<password>] [/debug]"
Dim DebugMode
Dim wsh
Set wsh = CreateObject("Wscript.shell")
Dim fso
Set fso = CreateObject("Scripting.filesystemobject")
Set cNamedArguments = WScript.Arguments.Named
InputFile = "C:\tmp\ora_sqllog.log"	'get sql logfile
'
' Check for help request
If WScript.Arguments.Named.Exists("?")   Or _
  WScript.Arguments.count < 3            Or _
  WScript.Arguments.Named.Exists("help")Then
  Call DisplayHelp()
  wscript.quit (1)
End If

If WScript.Arguments.Named.Exists("debug") Then
DebugMode = True
Dim DebugFile
DebugFile = wsh.ExpandEnvironmentStrings("%WINDIR%") & "\TEMP\ExecuteSQLdebug.log"
End If

'Set Target Script to run and test for existance
DB=Wscript.Arguments.Named("DB")
DB_SQSCRIPT=Wscript.Arguments.Named("S")

If Not fso.FileExists(DB_SQSCRIPT) Then 
If DebugMode = True Then appendToDebugLog
    sResult 1,DB_SQSCRIPT & " not found. Cannot proceed"
End If    

DOMAIN=UCase(Wscript.Arguments.Named("D"))


'KSS 26/06/09 : Check If Db User and Password passed and use Default Produser if not

If (WScript.Arguments.Named.Exists("U") And Not WScript.Arguments.Named.Exists("P")) Then
   sResult 1, "Both DB Username and Password must be supplied. Cannot Proceed"
ElseIf (WScript.Arguments.Named.Exists("P") And Not WScript.Arguments.Named.Exists("U")) Then
   sResult 1, "Both DB Username and Password must be supplied. Cannot Proceed"
ElseIf (WScript.Arguments.Named.Exists("P") And WScript.Arguments.Named.Exists("U")) Then
  USER=Trim(Wscript.Arguments.Named("U"))
  PASS=Trim(Wscript.Arguments.Named("P"))
  If USER ="" Or PASS ="" Then 
    sResult 1, "Db username or password passed is null. Cannot proceed" 
  End If  
Else
  USER="PRODUSER"
  
  Select Case DOMAIN
  Case "MNSUK"
  	PASS="london"
  Case "MNSUKDEV"
  	PASS="password"
  Case "MNSUKCATE"
  	PASS="password"
  End select
End IF

sqlPlus = wsh.regread("HKEY_LOCAL_MACHINE\Software\ORACLE\Oracle_Home") & "\bin\sqlplus.exe"

If Not fso.FileExists(DB_SQSCRIPT) Then _
	sResult 1,DB_SQSCRIPT & " not found. Cannot proceed"


'clear old SQL Log file
If fso.FileExists(InputFile) Then fso.DeleteFile(InputFile) 

'Run the .SQL script and create a SQL log file
strExec = "echo exit | " &sqlPlus & " " & USER & "/" & PASS & "@" & DB & " @" & DB_SQSCRIPT & " > c:\tmp\ora_sqllog.log"		

'write batch file
dim batchfile
set BatchFile = fso.OpenTextFile("C:\tmp\oraBatch.cmd",8,2)
Batchfile.writeline("@echo off")
Batchfile.writeline(strExec)

BatchFile.close

strExec = "cmd /c C:\tmp\orabatch.cmd"

Err.Clear

rc= wsh.Run (strExec,0,1)

fso.deletefile "C:\tmp\orabatch.cmd",1

If DebugMode=True Then appendToDebugLog

'Read SQL log

If Not fso.FileExists(InputFile) Then 
	sResult 99,"No Log found, results indeterminate for " & DB_SQSCRIPT
else
	Set TS = fso.OpenTextFile(InputFile)
	Do While Not TS.AtEndOfStream
	strTest = ts.ReadLine
	qRESULT=GetMatch (strTest)
	
	If qRESULT=1 Then sResult 1,"ORACLE SQL Script Failure Detected in " & DB_SQSCRIPT & vbCrLf & strTest
	Loop
End If

TS.Close

fso.DeleteFile inputfile,1
sResult 0, DB_SQSCRIPT & " script ran Successfully"

'------------------------------------------------------
Sub appendToDebugLog()
'------------------------------------------------------
Dim dbFile
Dim sqlLog
Set dbFile = fso.OpenTextFile(DebugFile,8,2)

dbFile.WriteLine(String(80,"-"))
dbFile.WriteLine("blSQLplus.exe Debug Session for " & DB_SQSCRIPT & vbTab & " on " & Now())
dbFile.WriteLine(String(80,"-"))
dbFile.WriteLine ""

If Not fso.FileExists("C:\tmp\ora_sqllog.log") Then
	dbFile.WriteLine "FATAL: Output log C:\tmp\ora_sqllog.log was not found."
	dbFile.WriteLine ""
	dbFile.WriteLine ""
	dbFile.Close
Else
	Set sqlLog = fso.OpenTextFile("C:\tmp\ora_sqllog.log",1)
	
	Do While Not sqlLog.AtEndOfStream
		Thisline = sqlLog.ReadLine
		dbFile.WriteLine(Thisline)
	Loop
	
	dbFile.WriteLine ""
	dbFile.WriteLine ""
	sqlLog.Close
	dbFile.Close

End If

End Sub

'------------------------------------------------------------------------------
' Function Check string for regular expressions
'------------------------------------------------------------------------------

Function GetMatch(target)
On Error resume next
'define regular expressions to test for
Dim RegExp,RegERROR
Set RegERROR = New RegExp

RegERROR.Pattern = "ERROR"
RegERROR.IgnoreCase = True

Set ERROR = RegERROR.Execute(strTest)

For Each match In ERROR
If match.FirstIndex>0 Then R1=1
Next

GetMatch=R1

End Function
'------------------------------------------------------------------------------
' Sub to Send StdOUT messages
'------------------------------------------------------------------------------
sub sResult(eeNum ,eeText )
Set stdOut=WScript.stdout
stdOut.WriteLine "blSQLplus.exe: " & eeText
stdOut.WriteLine "Return code was " & eeNum
WScript.Quit(eeNum)
End Sub

'------------------------------------------------------------------------------
' Sub to Display help
'------------------------------------------------------------------------------
Sub DisplayHelp()
   sResult 1, VbCrLf & "Usage:"  & VbCrLf &_
   VbCrLf &  "[ExecutablePath\]blSQLplus.exe /D:DOMAIN /S:[ScriptPath\]scriptname.sql /DB:<dbname> [/U:<db username> /P:<password>] [/debug]" & VbCrLf &_
   VbCrLf &  "Where DOMAIN is one of the following (MNSUK / MNSUKDEV / MNSUKCATE)" &_
   VbCrLf &  "Where db username /password are optional - Defaults for Produser taken if not supplied"
 End sub
