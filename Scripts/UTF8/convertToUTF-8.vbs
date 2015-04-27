'---------------------------------------------------------------------------------
' convertToUTF-8.vbs <InpuFile>
' Checks for existence of UNICODE format in file and converts to UTF-8 on requestls

'---------------------------------------------------------------------------------

Const YES = 6
Const NO = 7

Dim wsh
Set wsh = CreateObject("Wscript.shell")
Dim fso
Set fso = CreateObject( "Scripting.FileSystemObject" )

Dim strFileIn, strFileOut, strExec, boIsUnicode, intButton 

Dim args
Set args = WScript.Arguments

If Not args.Count = 1 Then
  MsgBox "Usage: convertToUTF-8.vbs <filename>"
  WScript.Quit(1)
End If

strFileIn = args(0)

boIsUnicode = IsUnicodeFile(strFileIn)

If boIsUnicode Then
   intButton = wsh.Popup("File contains Non-UTF8 Formats. Would you like to convert?",, "UTF-8 Conversion:", 4 + 32)
   
   If intButton = YES Then
       strFileOut = fso.GetBaseName( strFileIn ) & ".bak"
       
       'Make Copy of File
       strExec = "CMD /C COPY /Y " & strFileIn & " " & strFileOut
       rc = wsh.Run(strExec,0,1)
       
       If Not rc=0 Then 
	       MsgBox "Failed to make a copy of: " & strFileIn & " to " & strFileOut & " With error: " & Err.Description
	       WScript.Quit(1) 
       End If
       
       'Convert copy and overwrite original
       strExec = "CMD /C TYPE " & strFileOut & " > " & strFileIn
       rc = wsh.Run(strExec,0,1)
       
       If Not rc=0 Then 
	       MsgBox "Failed to execute TYPE Command on : " & strFileOut & " to " & strFileIn & " With error: " & Err.Description 
	       WScript.Quit(1)
       End If      
        
   End If    
Else
   MsgBox "File Is already UTF-8 Format"
End If      
   
'---------------------------------------------------------------------------------
Function IsUnicodeFile(filename) 
'---------------------------------------------------------------------------------
  Dim ts, char1, char2 
  
  Set ts = fso.opentextfile(filename) 
  
  On Error Resume Next 
  
  IsUnicodeFile = False 
  
  char1 =ts.read(1) 
  char2 =ts.read(1) 
  ts.close 
  
  If Err.Number = 0 Then 
    If asc(char1) = 255 and asc(char2) = 254 Then 
      IsUnicodeFile = True 
    End If 
  End If 
End Function 
