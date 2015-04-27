Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim CaseSense
CaseSense = False
Dim rCount

Dim regEx
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = true
Dim stdErr, StdOut
Set stdErr = WScript.StdErr
Set StdOut = WScript.StdOut

Dim args
Set args = WScript.Arguments

If Not (args.Count = 4  Or args.Count = 3) Then
StdErr.WriteLine "Usage: stringReplace.exe <filename> <search_string> <replace_string> [-preserveCase]"
WScript.Quit(1)
End If


filename = args(0)
If args.Count = 4 then
If LCase(args(3)) = "-preservecase" Then
caseSense = True
End If
End If


If Not fso.FileExists(filename) Then
stdErr.WriteLine "stringReplace: Error, cannot find file " & args(0)
WScript.Quit(53)
End If


If CaseSense = true then
oldString = args(1)
newString = args(2)
Else
oldString = UCase(args(1))
newString = UCase(args(2))
End if


Set OutPut = fso.OpenTextFile(filename & "$$",8,2)
Set Input = fso.OpenTextFile(filename,1)

Do While Not Input.AtEndOfStream
CurLine = Input.ReadLine
If CaseSense = true Then
	If InStr(curLine,Oldstring) > 0 Then
	curline = Replace(curline,oldstring,newstring)
	rCount = rCount +1
	End if
Else

	If InStr(UCase(curLine),oldstring) > 0 Then
		regEx.Pattern = oldstring
		Set matches = regEx.Execute(curline)
		For Each match In matches
			curline = Replace(curline,match.value,newstring)
			rCount = rCount +1
		next
	End If
End if
output.WriteLine curline

Loop

output.Close
input.Close


If rCount > 0 Then
fso.DeleteFile filename,1
fso.MoveFile filename & "$$", filename
StdOut.WriteLine "stringReplace.exe completed with " & rCount & " substitutions"
Else
fso.DeleteFile filename & "$$", 1
StdOut.WriteLine "stringReplace.exe completed with ZERO replacements - check your parameters"
End If
