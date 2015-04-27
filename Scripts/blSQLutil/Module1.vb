Imports System.Security
Imports System.IO


Module Module1

    Dim WinDir As String = Nothing
    Dim Userid As String = Nothing
    Dim SecStr As New SecureString()
    Dim HostName As String = Nothing
    Dim DebugMode As Boolean = False
    Dim Domain As String = Nothing
    Dim cmdExe As String = Nothing
    Dim cmdArgs As String = Nothing
    Dim ProcError As Boolean



    Sub Main()
        'If Not AppAuthorised() Then
        'Console.WriteLine("blSQLutil is not authorised to run")
        'Environment.Exit(5)
        'End If
        Main(Environment.GetCommandLineArgs())
    End Sub


    Private Sub Main(ByVal Args() As String)

        If Not Args Is Nothing Then
            Select Case Args.Length
                Case 1
                    Console.WriteLine("No Arguments! Usage: blsqlutil <DOMAIN> <Executable> <arguments> [/debug]")
                    System.Environment.Exit(1)
                Case 4
                    Exit Select
                Case 5
                    If UCase(Args(4)) = "/DEBUG" Then
                        DebugMode = True
                    End If
                Case Else
                    Console.WriteLine("Syntax error! Usage: blsqlutil <DOMAIN> <Executable> <arguments> [/debug]")
                    System.Environment.Exit(1)
            End Select
        End If

        'Get Windows Directory
        WinDir = System.Environment.ExpandEnvironmentVariables("%WINDIR%")
        If WinDir = "" Then
            'something bad has happended
            Console.WriteLine("blSQLcmd PANIC! failed to enumerate Windir. Exiting")
            System.Environment.Exit(1)
        End If

        dbg(" -----------------------------blSQLutil.exe ----------------------------------")
        dbg(" Debug log started on " & Now())
        dbg(" Argument list:")

        For i As Integer = 0 To UBound(Args)
            dbg("Args(" & i & ")=" & Args(i))
        Next


        Domain = UCase(Args(1))
        cmdExe = Args(2)
        cmdArgs = Args(3)

       
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
                Console.WriteLine("Error: The domain must be MNSUKDEV, MNSUKCATE or MNSUK")
                dbg("Error: The domain must be MNSUKDEV, MNSUKCATE or MNSUK")
                Environment.Exit(1)
        End Select

        '========================================================
        setSecStr(UCase(Domain))
        dbg("Returned to main execution thread")
        '========================================================


        If Not File.Exists(cmdExe) Then
            dbg("==================================================")
            dbg("Could not find a file called " & cmdExe)
            dbg("Are you assuming that it is in the PATH?")
            dbg("===================================================")
        End If


        '----------------------- MAIN Processing HERE ------------------

        Dim sqlproc As New System.Diagnostics.ProcessStartInfo
        With sqlproc
            .FileName = cmdExe
            .UserName = Userid
            .Password = SecStr
            .Domain = UCase(Domain)
            .UseShellExecute = False
            .CreateNoWindow = True
            .Arguments = cmdArgs
            .RedirectStandardError = True
            .RedirectStandardOutput = True
        End With

        dbg("lauching process with " & sqlproc.UserName)

        Try
            Dim ps As System.Diagnostics.Process = Process.Start(sqlproc)
            Dim Myproc As Integer = ps.Id
            dbg("Process spawned as PID = " & ps.Id)

            'loop until processid does not exist
            Do Until ps.HasExited
                dbg("Waiting 2 more seconds for process " & ps.Id & " to end")
                System.Threading.Thread.Sleep(2000)
            Loop

            Dim SO As StreamReader = ps.StandardOutput
            Console.WriteLine(SO.ReadToEnd())
            Dim SE As StreamReader = ps.StandardError
            Dim ErrorCondition As String = Nothing
            ErrorCondition = SE.ReadToEnd()
            If Not ErrorCondition = "" Then
                ProcError = True
                dbg("------------Redirected Standard Error --------------")
                dbg(ErrorCondition)
                Console.WriteLine(ErrorCondition)
                dbg("----------------------------------------------------")
            End If

            '-------------------> ERROR CONTROL HERE
            If ProcError Then
                dbg("Messages to STDerr were detected. Exiting")
                dbg("----------------------------------------------------" & vbCrLf & vbLf)
                SecStr.Clear()
                Environment.Exit(1)
            Else
                dbg("Process exited with rc=" & ps.ExitCode)
            End If

            ps = Nothing
            SO = Nothing
            SE = Nothing

        Catch ex As Exception
            dbg("exception caught in ps runtime. " & ex.Message)
            Console.WriteLine("exception caught in ps runtime. " & ex.Message)
            dbg("-------------------------------------------------------------------")
            dbg("")
            Exit Sub
        End Try

        sqlproc = Nothing

        '-------------------------------------------------------------
        'clear password from memory
        SecStr.Clear()
        dbg("Processing complete")
        dbg("-------------------------------------------------------------------" & vbCrLf & vbLf)
    End Sub


    Private Sub setSecStr(ByVal DomainName As String)

        Dim ptstring As String = ""
        Select Case UCase(DomainName)
            Case "MNSUKDEV"
                ptstring = Decode("H1a2KmNh3AQ")
            Case "MNSUKCATE"
                ptstring = Decode("@uP.e#R%YH")
            Case "MNSUK"
                ptstring = Decode("@8GjFq4ZYH")
            Case Else
                dbg("PANIC! Sub SetSecStr cound not setup secure password becuase domain was '" & DomainName & "'")
                System.Environment.Exit(1)
        End Select

        Dim ptArr() As Char = ptstring.ToCharArray
        For i As Integer = 0 To UBound(ptArr)
            SecStr.AppendChar(ptArr(i))

        Next

        ptstring = ".=.0.0.=."
        ptString = ""

    End Sub

    Private Function AppAuthorised() As Boolean
        AppAuthorised = False
        Dim oNet As Object = Nothing
        oNet = CreateObject("Wscript.Network")
        dbg(oNet.Username)

        If UCase(oNet.Username) = "BLADELOGICRSCD" Then
            AppAuthorised = True
        End If

        oNet = Nothing

    End Function


    Private Sub dbg(ByVal LogString As String)

        If Not DebugMode Then Exit Sub
        Dim tsOut As StreamWriter = Nothing
        tsOut = File.AppendText(WinDir & "\TEMP\blUtil_debg.txt")
        tsOut.WriteLine(Now() & vbTab & LogString)
        tsOut.Close()
    End Sub


    Private Function Decode(ByVal sInput As String) As String
        dbg("Decoding password...")
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


End Module

