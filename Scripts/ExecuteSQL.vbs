On Error Resume Next
'Usage ExecuteSQL.exe /s:<path\filename.sql> [/c:<cluster name>] [/debug]
Dim DebugMode
Dim UTILITY,SCRIPT,SERVERNAME
Dim wsh
Set wsh = CreateObject("Wscript.shell")
Dim fso
Set fso = CreateObject("Scripting.filesystemobject")
Set cNamedArguments = WScript.Arguments.Named
InputFile = "C:\tmp\RunSQL.log"	'get sql logfile
'
' Check for help request
If WScript.Arguments.Named.Exists("?")   Or _
  Not WScript.Arguments.Named.Exists("S") Or _
   WScript.Arguments.Named.Exists("help")Then
  Call DisplayHelp()
  wscript.quit (1)
End If

If WScript.Arguments.Named.Exists("debug") Then
DebugMode = True
Dim DebugFile
DebugFile = wsh.ExpandEnvironmentStrings("%WINDIR%") & "\TEMP\ExecuteSQLdebug.log"
End If


'Set Target Script and Utility to run and test for existence
SCRIPT=Wscript.Arguments.Named("S")
UTILITY="c:\Program Files\Microsoft SQL Server\80\Tools\Binn\osql.exe"
If Not fso.FileExists(SCRIPT) Then
If DebugMode = True Then appendToDebugLog
	sResult 1,SCRIPT & " not found. Cannot proceed"
End If

	
If Not fso.FileExists(UTILITY) Then _
	sResult 1,UTILITY & " not found. Cannot proceed"
	

	
Set objWMISvc = GetObject( "winmgmts:\\.\root\cimv2" )
Set colItems = objWMISvc.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
For Each objItem in colItems
    If objItem.PartOfDomain Then
        HOSTNAME=objItem.name
        Select Case UCase(Mid(HOSTNAME,11,1))
			Case "D"
				DOMAIN = "MNSUKDEV"
			Case "C"
				DOMAIN = "MNSUKCATE"
			Case "P"
				DOMAIN = "MNSUK"
			Case Else
					sResult 99,"Failed to define DOMAIN for " & objitem.name _
					& ". Is this a proper server name?" _
					& " 11th Char of name must be C,D or P. "
		End Select
    Else
        sResult 99,objitem.name &" is not a Domain Member"
    End If
Next	

	
'DOMAIN=UCase(Wscript.Arguments.Named("D"))   ''Retired option but switch retained for legacy packages
USER=DOMAIN & "\Y0129080"
select Case DOMAIN	
	Case "MNSUK"
		PASS="YFZj4Gq8"
	Case "MNSUKDEV"
		PASS="mAK32haN1"
	Case "MNSUKCATE"
		PASS="YeRPu12G"
End select


'Define Environment 
Set objWMISvc = GetObject( "winmgmts:\\.\root\cimv2" )
Set colItems = objWMISvc.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
For Each objItem in colItems
If objItem.PartOfDomain Then
    HOSTNAME=objItem.name
Else
    sResult 99,objitem.name &" is not part of a Domain"
End If
Next

'If Cluster name specified on command line then use it instead
'of the hostname for the SQL Server name
If WScript.Arguments.Named.Exists("C") Then
	SERVERNAME=Wscript.Arguments.Named("C")
Else
	SERVERNAME=HOSTNAME
End If

'clear old SQL Log file
If fso.FileExists(InputFile) Then fso.DeleteFile(InputFile) 

'check for spaces in script name
If InStr(SCRIPT," ") > 0 Then
SCRIPT = Chr(34) & SCRIPT & Chr(34)
End If


'Run the .SQL script and create a SQL log file
strExec = "c:\tmp\psexec.exe -accepteula \\" & HOSTNAME _
		& " -u " & USER _
		& " -p " & PASS _
		& " " & Chr(34) & UTILITY _
		& Chr(34) & " -S " & SERVERNAME & " -i " & SCRIPT _
		& " -E -o c:\tmp\RunSQL.log"
		
Err.Clear

rc = wsh.Run (strExec,0,1)

If DebugMode=True Then appendToDebugLog

'Check for successful run
If Not rc=0 Then 
	Err.Raise rc
	sResult rc,"ExecuteSQL: failed to Launch Script " & SCRIPT & vbCrLf & Err.Description
End If

'Consider system drive variable here
'Read SQL log
If Not fso.FileExists(InputFile) Then 
	sResult 99,"No Log found, results indeterminate for " & SCRIPT
else
	Set TS = fso.OpenTextFile(InputFile)
	Do While Not TS.AtEndOfStream
		strTest = ts.ReadLine
		qRESULT=GetMatch (strTest)
		If qRESULT=3 Then
		TS.Close
		sResult 1,"SQL Script Failure Detected in " & SCRIPT & vbCrLf & strTest
		End if
	Loop
End If

TS.Close

sResult 0, SCRIPT & " script ran Successfully"



'------------------------------------------------------
Sub appendToDebugLog()
'------------------------------------------------------
Dim dbFile
Dim sqlLog
Set dbFile = fso.OpenTextFile(DebugFile,8,2)

dbFile.WriteLine(String(80,"-"))
dbFile.WriteLine("ExecuteSQL Debug Session for " & SCRIPT & vbTab & " on " & Now())
dbFile.WriteLine(String(80,"-"))
dbFile.WriteLine ""

If Not fso.FileExists("c:\tmp\RunSQL.log") Then
	dbFile.WriteLine "FATAL: Output log c:\tmp\RunSQL.log was not found."
	dbFile.WriteLine ""
	dbFile.WriteLine ""
	dbFile.Close
Else
	Set sqlLog = fso.OpenTextFile("c:\tmp\RunSQL.log",1)
	
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



'------------------------------------------------------
' Function Check string for regular expressions
'------------------------------------------------------

Function GetMatch(target)
'define regular expressions to test for
Dim RegExp,RegMSG,RegLVL,RegSTATE
Set RegMSG = New RegExp
Set RegLVL = New RegExp
Set RegSTAT = New RegExp
RegMSG.Pattern = "Msg "
RegMSG.IgnoreCase = True
RegLVL.Pattern = "Level "
RegLVL.IgnoreCase = True
RegSTAT.Pattern = "State "
RegSTAT.IgnoreCase = True


Set MSG = RegMSG.Execute(strTest)
Set LVL = RegLVL.Execute(strTest)
Set STAT = RegSTAT.Execute(strTest)

For Each match In MSG
	If match.FirstIndex>0 Then R1=1
Next
For Each match In LVL
	If match.FirstIndex>0 Then R2=1
Next
For Each match In STAT
	If match.FirstIndex>0 Then R3=1
Next
GetMatch=R1+R2+R3

End Function
'------------------------------------------------------
' Sub to Send StdOUT messages
'------------------------------------------------------
sub sResult(eeNum ,eeText )
Set stdOut=WScript.stdout
stdOut.WriteLine eeText
stdOut.WriteLine "Return code was " & eeNum

If fso.FileExists("c:\tmp\RunSQL.log") Then
	fso.DeleteFile "c:\tmp\RunSQL.log",1
End If


WScript.Quit(eeNum)

End Sub

'------------------------------------------------------
' Sub to Display help
'------------------------------------------------------
Sub DisplayHelp()
   sResult 1, VbCrLf & "Usage:"  & VbCrLf &_
   VbCrLf &  "c:\tmp\ExecuteSQL.exe /S:<[ScriptPath\]scriptname.sql> [/C:<cluster name>] [/debug]" & VbCrLf &_
   VbCrLf &  "Where DOMAIN is one of the following (MNSUK / MNSUKDEV / MNSUKCATE)"
End sub

