Imports System.Security
Imports System.IO


Module Module1

    Dim WinDir As String = Nothing
    Dim Userid As String = Nothing
    Dim SecStr As New SecureString()
    Dim HostName As String = Nothing
    Dim SQLversion As String = Nothing
    Dim DebugMode As Boolean = False
    Dim Domain As String = Nothing
    Dim inputfile As String = Nothing
    Dim debugFile As String = Nothing
    Dim sqlCmd_out As String = Nothing
    Dim SQLcommand As String = Nothing
    Dim SQLargs As String = Nothing
    Dim SQLerr As Boolean = False


    Sub Main()
        If Not AppAuthorised() Then
            Console.WriteLine("blSQLcmd is not authorised to run")
            Environment.Exit(5)
        End If
        Main(Environment.GetCommandLineArgs())
    End Sub


    Private Sub Main(ByVal Args() As String)

        'Get Windows Directory
        WinDir = System.Environment.ExpandEnvironmentVariables("%WINDIR%")
        If WinDir = "" Then
            'something really bad has happended
            Console.WriteLine("blSQLcmd PANIC! failed to enumerate Windir. Exiting")
            System.Environment.Exit(1)
        End If


        If Not Args Is Nothing Then
            Select Case Args.Length
                Case 1
                    Console.WriteLine("No Arguments! Usage: blsqlcmd <DOMAIN> <scriptfile.sql> [/debug]")
                    System.Environment.Exit(1)
                Case 3
                    Exit Select
                Case 4
                    If UCase(Args(3)) = "/DEBUG" Then
                        DebugMode = True
                        debugFile = WinDir & "\TEMP\blSQL_debg.txt"
                        sqlCmd_out = WinDir & "\Temp\sqlcmd_out.txt"
                        If File.Exists(sqlCmd_out) Then
                            File.Delete(sqlCmd_out)
                        End If
                    End If
                Case Else
                    Console.WriteLine("Syntax error! Usage: blsqlcmd <DOMAIN> <scriptfile.sql> [/debug]")
            End Select
        End If


        'try getting the current machine name  for command line
        HostName = System.Environment.MachineName
        If HostName = "" Then
            Console.WriteLine("PANIC! - could not get hostname from environment")
            System.Environment.Exit(1)
        End If

        Domain = UCase(Args(1))
        inputfile = UCase(Args(2))

        If Not File.Exists(inputfile) Then
            Console.WriteLine("")
            Console.WriteLine("blSQLcmd.exe: ERROR, input file does not exist: " & inputfile)
            Console.WriteLine("")
            System.Environment.Exit(1)
        End If



        '=======================================================
        '-----SET Service Account Userid & password
        Select Case UCase(Domain)
            Case "MNSUKDEV"
                Userid = "Y0129080"
            Case "MNSUKCATE"
                Userid = "Y0129080"
            Case "MNSUK"
                Userid = "Y0129080"
            Case Else
                Console.WriteLine("")
                Console.WriteLine("BlSQLcmd.exe: ERROR, The domain must be MNSUKDEV, MNSUKCATE or MNSUK")
                Console.WriteLine("")
                Environment.Exit(1)
        End Select

        '========================================================
        setSecStr(UCase(Domain))
        '========================================================

        'detectSQL version
        If File.Exists("C:\Program Files\Microsoft SQL Server\90\Tools\binn\SQLCMD.EXE") Then
            SQLcommand = """C:\Program Files\Microsoft SQL Server\90\Tools\binn\SQLCMD.EXE"""
            SQLversion = "9"
        Else
            If File.Exists("C:\Program Files\Microsoft SQL Server\80\Tools\Binn\osql.exe") Then
                SQLcommand = """C:\Program Files\Microsoft SQL Server\80\Tools\Binn\osql.exe"""
                SQLversion = "8"
            End If
        End If

        WriteDebugHeader()

        '---------------------- WARM UP COMMAND --------------------------
        SQLargs = "-r -S " & HostName & " -E -Q " & Chr(34) & "SELECT @@VERSION" & Chr(34)
        Dim Preproc As New System.Diagnostics.ProcessStartInfo
        With Preproc
            .FileName = SQLcommand
            .UserName = Userid
            .Password = SecStr
            .Domain = UCase(Domain)
            .UseShellExecute = False
            .CreateNoWindow = True
            .Arguments = SQLargs
            .RedirectStandardError = True
        End With

        Try
            

            Dim psPre As System.Diagnostics.Process = Process.Start(Preproc)
            dbg("Preprocessing ProcessID = " & psPre.Id)
            Dim t1 As Integer = 0
            'loop until processid does not exist
            Do Until psPre.HasExited
                System.Threading.Thread.Sleep(1000)
                t1 = t1 + 1
            Loop


            Dim SE As StreamReader = psPre.StandardError
            If Not SE.ReadToEnd = "" Then
                SQLerr = True
                dbg("------------Redirected Standard Error --------------")
                dbg(SE.ReadToEnd())
                dbg("----------------------------------------------------")
            End If

            dbg("Preprocessing has completed after " & t1 & " seconds with return code " & psPre.ExitCode)
            psPre = Nothing
            SE = Nothing

        Catch ex As Exception
            'dbg("exception caught in psPre runtime. " & ex.Message)
            Console.WriteLine("exception caught in psPre runtime. " & ex.Message)
            Environment.Exit(1)
        End Try

        Preproc = Nothing
        '--------------------------------------------------------------


        '----------------------- MAIN SQL WORK HERE ------------------
        If DebugMode = True Then
            SQLargs = "-r -S " & HostName & " -i " & Chr(34) & inputfile & Chr(34) & " -o " & sqlCmd_out & " -E"
        Else
            SQLargs = "-r -S " & HostName & " -i " & Chr(34) & inputfile & Chr(34) & " -E"
        End If

        Dim sqlproc As New System.Diagnostics.ProcessStartInfo
        With sqlproc
            .FileName = SQLcommand
            .UserName = Userid
            .Password = SecStr
            .Domain = UCase(Domain)
            .UseShellExecute = False
            .CreateNoWindow = True
            .Arguments = SQLargs
            .RedirectStandardError = True
            .RedirectStandardOutput = True
        End With

        Try

            Dim ps As System.Diagnostics.Process = Process.Start(sqlproc)
            Dim Myproc As Integer = ps.Id
            dbg("MAIN SQL EXECUTION BEGINS HERE")
            dbg("Command utility: " & Replace(sqlproc.FileName, Chr(34), ""))
            dbg("Arguments:       " & Replace(sqlproc.Arguments, " -o " & sqlCmd_out, ""))
            dbg("User context:    " & sqlproc.Domain & "\" & sqlproc.UserName)
            dbg("ProcessID = " & ps.Id)


            'loop until processid does not exist
            Do Until ps.HasExited
                dbg("Waiting 2 more seconds for process " & ps.Id & " to end")
                System.Threading.Thread.Sleep(2000)
            Loop
            dbg("Process exited with rc=" & ps.ExitCode)
            dbg("")

           

            Dim SO As StreamReader = ps.StandardOutput
            Console.WriteLine(SO.ReadToEnd())
            Dim SE As StreamReader = ps.StandardError
            Dim ErrorCondition As String = Nothing
            ErrorCondition = SE.ReadToEnd()

            If Not ErrorCondition = "" Then
                SQLerr = True
                dbg("Output to stdErr: The following content was directed to the Standard Error")
                dbg(ErrorCondition)
                Console.WriteLine(ErrorCondition)
                dbg("******************** Output from sqlcmd follows ********************")
                appendOutputToDebug()
                dbg("********************************************************************")
                SecStr.Clear()
                Environment.Exit(1)
            Else
                dbg("******************** Output from sqlcmd follows ********************")
                appendOutputToDebug()
                dbg("********************************************************************")
                dbg("")

            End If

            ps = Nothing
            SO = Nothing
            SE = Nothing

        Catch ex As Exception
            dbg("exception caught in ps runtime. " & ex.Message)
            Console.WriteLine("exception caught in ps runtime. " & ex.Message)
            Exit Sub
        End Try

        sqlproc = Nothing

        '-------------------------------------------------------------
        SecStr.Clear()
        dbg("Processing complete!")




    End Sub


    Private Sub setSecStr(ByVal DomainName As String)

        Dim ptstring As String = ""
        Select Case UCase(DomainName)
            Case "MNSUKDEV"
                ptstring = Decode("H1a2KmNh3AQ")
            Case "MNSUKCATE"
                ptstring = Decode("@G1Pe2uRYH")
            Case "MNSUK"
                ptstring = Decode("@8GjFq4ZYH")
            Case Else
                Console.WriteLine("blSQLcmd: PANIC! Sub SetSecStr cound not setup secure password becuase domain was '" & DomainName & "'")
                System.Environment.Exit(1)
        End Select

        Dim ptArr() As Char = ptstring.ToCharArray
        For i As Integer = 0 To UBound(ptArr)
            SecStr.AppendChar(ptArr(i))

        Next

        ptstring = ".=.0.0.=."
        ptstring = ""

    End Sub

    Private Function AppAuthorised() As Boolean
        AppAuthorised = False
        Dim oNet As Object = Nothing
        oNet = CreateObject("Wscript.Network")

        If UCase(oNet.Username) = "BLADELOGICRSCD" Then
            AppAuthorised = True
        End If

        oNet = Nothing

    End Function


    Private Sub dbg(ByVal LogString As String)

        If Not DebugMode Then Exit Sub
        Dim tsOut As StreamWriter = Nothing
        tsOut = File.AppendText(debugFile)
        tsOut.WriteLine(Now() & vbTab & LogString)
        tsOut.Close()
        tsOut = Nothing

    End Sub


    Private Function Decode(ByVal sInput As String) As String

        Dim sTemp As String = ""
        Dim sEnd As String = ""
        Dim sStart As String = ""
        Dim iLoop As Integer
        Dim iLen As Integer
        Dim iMiddle As Integer
        Dim iRemainder As Integer
        sInput = Mid(sInput, 2, sInput.Length - 2)
        iLen = sInput.Length

        iRemainder = iLen Mod 2
        iMiddle = iLen \ 2

        For iLoop = iMiddle + iRemainder To 1 Step -1
            If iRemainder = 0 Then
                sTemp += Microsoft.VisualBasic.Mid(sInput, iLoop + iMiddle, 1)
            End If
            sTemp += Microsoft.VisualBasic.Mid(sInput, iLoop, 1)
            If iRemainder = 1 And iLoop <> 1 Then
                sTemp += Microsoft.VisualBasic.Mid(sInput, iLoop + iMiddle, 1)
            End If
        Next

        Return sTemp
    End Function


    Sub WriteDebugHeader()
        If Not DebugMode Then Exit Sub
        Dim tsOut As StreamWriter = Nothing
        tsOut = File.AppendText(debugFile)
        tsOut.WriteLine("")
        tsOut.WriteLine("--------------------------------------------------------------------------------------------")
        tsOut.WriteLine("--- blSQLcmd.exe debug for " & inputfile)
        tsOut.WriteLine("--------------------------------------------------------------------------------------------")
        tsOut.WriteLine("Sqlversion=" & SQLversion & ".0")
        tsOut.WriteLine("")
        tsOut.Close()
        tsOut = Nothing

    End Sub

    Sub appendOutputToDebug()

        If Not DebugMode Then Exit Sub
        Try

            If Not File.Exists(sqlCmd_out) Then

                dbg("WARNING: no sqlcmd output file was detected")

            End If

            Dim fs As FileStream = File.OpenRead(sqlCmd_out)
            Dim sr As New StreamReader(fs)
            Dim content As String = vbCrLf & sr.ReadToEnd & vbCrLf
            fs.Close()

            File.AppendAllText(debugFile, content)
            File.Delete(sqlCmd_out)
        Catch ex As Exception

            Console.WriteLine("Exception caught in appendOutputToDebug():" & ex.Message)

        End Try


    End Sub

End Module

