'SendTicketsi Transfer batch scale tickets to Scoular (interactive)
'Uses Dart PowerTCP SSH and SFTP for .NET 4.10
'Control-S calls configuration application
'SFTP files must be sent to the IFS on the IBM iseries

Imports System.IO
Imports ClassLibrary1
Imports Dart.Ssh
Imports System.Net.Sockets

Public Class SendTickets
    'Dim ssh1 As Dart.Ssh.Optimizations
    Dim sftp As New Dart.Ssh.Sftp
    Dim Ftp1 As New Dart.Ftp.Ftp
    Dim dbase As String = "test" ' production or test system"
    Dim splib As String
    Dim result As Boolean
    Dim loc As String
    Dim sadd As String
    Dim dpath As String
    Dim sid As String
    Dim spath As String
    Dim pkey As String
    Dim pass As String
    Dim uname As String
    Dim pkey2 As String
    Dim lpath As String
    Dim ppath As String
    Dim fpath As String
    Dim logfile As String
    Dim rcnt As Integer
    Dim yesno() As String = {"Y", "y", "N", "n"}
    Const sfile As String = "tickets.txt"
    Const nfile As String = "tickets.dat"
    Const iprefix As String = "/P"
    Const isuffix As String = "600I.txt"
    Const str_dir As String = "/qsys.lib/PCSPRDLIB.lib/"  'prod
    Const str_temp As String = "tickets.tmp"
    Dim srcfile As String, dstfile As String, nsrcfile As String
    Dim abort As Boolean = False
    Dim stest As Boolean = True
    Dim logwriter As StreamWriter


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Dim RetValues As Utils.Appvalues = Utils.GetAllSettings() 'get values from app.cfg
        SetValues(RetValues)
        'test for values
        'if no log folder fire error
        If Not Testsystem(lpath) Then abort = True
        Me.Text = Me.Text & " for location " & loc
        If Not abort Then
            Checkfolders() 'create folders if needed
        End If
        If Not abort Then
            Createlog() 'create log file
        End If
        If Not abort Then
            Createfilename
        End If
    End Sub
    Private Sub Main_KeyDown(ByVal Sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        NewMethod(e)
    End Sub

    Private Shared Sub NewMethod(e As KeyEventArgs)
        If (e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Control) Then
            Dim f As New Appconfigmaint
            f.Show()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Button1.Enabled Then Me.Button1.Enabled = False
        If Me.Button2.Enabled Then Me.Button2.Enabled = False
        logwriter.WriteLine(Date.Now & " File upload process initiated")
        'test for values
        Dim retval As Boolean
        retval = TestConnection()
        If retval Then
            If Checksource() Then
                'send file
                Sendfile()
                RunCommand

            Else
                'error abort
                abort = True
                CloseWriter()
                DisposeWriter()
                Exit Sub
            End If
        End If
        logwriter.WriteLine(GetDateTime() & " Process Completed----------")
        CloseWriter()
        DisposeWriter()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim int_result As Integer
        abort = True
        int_result = MsgBox("FTP transfer aborted," & vbCrLf & "Do you wish to archive and delete files?", MsgBoxStyle.YesNo, "Application terminated")
        If int_result = 6 Then  '6=yes archive files
            If Checksource() Then
                Rename_file() 'rename the file
                ArchiveCancel()
            End If
        Else
            Try
                logwriter.WriteLine(GetDateTime() & " File archive canceled by user")
            Catch ex As Exception
            End Try
        End If
        CloseWriter()
        DisposeWriter()
        Me.Close()
        Me.Dispose()
    End Sub
    Private Sub SetValues(retvalues As Utils.Appvalues)
        'set local variables to config values
        loc = retvalues.Location
        sadd = retvalues.FTP_Address
        dpath = retvalues.Destination_Path
        sid = retvalues.Source_Device_Address
        spath = retvalues.Source_Path
        pkey = retvalues.Password
        pass = retvalues.PassPhrase
        uname = retvalues.Username
        Me.lpath = spath & "log\" 'log file path
        Me.ppath = spath & "sent\" 'sent files path
        Me.fpath = spath & "fail\" 'failed files path
        Try
            sftp.Connection.RemoteEndPoint.HostNameOrAddress = sadd
            sftp.Connection.RemoteEndPoint.Port = 22
            sftp.Optimizations.BlockSize = 2056
        Catch ex As Exception
        End Try

    End Sub
    Private Sub Checkfolders()
        'check for log file folder and create if necessary
        If Not My.Computer.FileSystem.DirectoryExists(lpath) Then
            Try
                My.Computer.FileSystem.CreateDirectory(lpath)
            Catch ex As Exception
                If ex.Message = "Path/File access error." Then
                    'ignore
                Else
                    MsgBox("Folder" & lpath & " not created procedure aborted ", MsgBoxStyle.Critical, "Folder Error")
                    abort = True
                    Exit Sub
                End If
            End Try
        End If
        'check for folder to capture sent files
        If Not My.Computer.FileSystem.DirectoryExists(ppath) Then
            Try
                My.Computer.FileSystem.CreateDirectory(ppath)
            Catch ex As Exception
                If ex.Message = "Path/File access error." Then
                    'ignore
                Else
                    MsgBox("Folder" & ppath & " not created procedure aborted ", MsgBoxStyle.Critical, "Folder Error")
                    abort = True
                    Exit Sub
                End If
            End Try
        End If
        'check for folder to capture failed files
        If Not My.Computer.FileSystem.DirectoryExists(fpath) Then
            Try
                My.Computer.FileSystem.CreateDirectory(fpath)
            Catch ex As Exception
                If ex.Message = "Path/File access error." Then
                    'ignore
                Else
                    MsgBox("Folder" & fpath & " not created procedure aborted ", MsgBoxStyle.Critical, "Folder Error")
                    abort = True
                    Exit Sub
                End If
            End Try
        End If
    End Sub
    Private Sub Createlog()
        'create log file
        Dim dtmCurrentDate As Date = DateTime.Now
        logfile = String.Format("{0:yy}", dtmCurrentDate) & String.Format("{0:MM}", dtmCurrentDate) & ".txt"
        Try
            logwriter = My.Computer.FileSystem.OpenTextFileWriter(lpath & logfile, True)
        Catch ex As IOException
            MsgBox("Log file in use by another application, job cancelled", MsgBoxStyle.Critical, "Log file error")
            abort = True
            Exit Sub
        End Try
        Try
            If My.Computer.FileSystem.GetFileInfo(lpath & logfile).Length = 0 Then
                logwriter.WriteLine(GetDateTime() & " log file created")
            End If
        Catch ex As Exception
            MsgBox("Log file create error, job aborted", MsgBoxStyle.Critical, "Create log file")
            abort = True
            Exit Sub
        End Try
        logwriter.WriteLine(GetDateTime() & " File Transfer Initiated----------")
    End Sub
    Private Sub Createfilename()
        srcfile = spath & sfile 'source file on pc tickets.txt
        nsrcfile = spath & nfile 'renamed file ticket.dat
        dstfile = dpath & iprefix & loc & isuffix 'destination file on AS/400
    End Sub
    Private Sub Sendfile()
        Rename_file()
        If Not abort Then
            Format_file(nsrcfile) 'reformat file for upload
        End If
        If Not abort Then
            If sftp.Connection.AuthenticationComplete Then
                sftp.Optimizations.UseAppend = True
                'sftp.Put(nsrcfile, dstfile, CopyMode.Append)
                Dim r As CopyResult = sftp.Put(nsrcfile, dstfile, CopyMode.Append)
                If r.Percentage = 100 Then
                    Archivefile(r)
                    sftp.Close()

                Else
                    ArchiveFail(r.Status.ToString)
                End If
            End If
        End If

    End Sub
    Private Sub Capture_file()
        ' if renamed file exists capture to failed folder
        Dim aname As String = Fun_arcfil(GetDateTime)
        My.Computer.FileSystem.CopyFile(spath & nfile, aname)
        logwriter.WriteLine(GetDateTime() & " File " & nfile & " archived to: " & aname)
    End Sub
    Sub Rename_file()
        Try 'check for .dat file
            If My.Computer.FileSystem.FileExists(nsrcfile) Then
                Capture_file()
                My.Computer.FileSystem.DeleteFile(nsrcfile)
            End If
        Catch ex As Exception
            logwriter.WriteLine(GetDateTime() & " Previous ticket file delete failed, Process aborted")
            abort = True
            Exit Sub
        End Try
        Try
            My.Computer.FileSystem.RenameFile(srcfile, nfile)
        Catch ex As Exception
            logwriter.WriteLine(GetDateTime() & " Ticket file rename failed, Process aborted")
            abort = True
            Exit Sub
        End Try
    End Sub
    Public Sub Format_file(ByVal fname As String)
        'validate records in ticket file
        'copy valid ticket to temp file rename to .dat file
        Dim dtmdatetime As Date = Date.Now
        Dim r As String
        Dim d As String
        Dim oname As String
        oname = spath & str_temp 'temp ticket fileo
        Try
            My.Computer.FileSystem.DeleteFile(oname)
        Catch ex As Exception
        End Try
        'open streamreader and streamwriter
        Dim filein As New StreamReader(New FileStream(fname, FileMode.Open, FileAccess.Read))
        Dim fileo As New StreamWriter(New FileStream(oname, FileMode.Create, FileAccess.Write))

        Do While filein.Peek <> -1
            r = filein.ReadLine
            If r <> Nothing Then
                rcnt += 1
                d = r.Substring(2, 2) & "/" & r.Substring(4, 2) & "/" & r.Substring(0, 2) 'reformat date
                If IsDate(d) Then 'valid ticket date
                    If r.Substring(6, 6) > "000000" Then 'valid ticket #
                        'if pos 75-76 are blank and pos 77=0 and pos 481 =Y or n and pos 482 is blank remap the data.
                        Dim idx As Integer = Array.BinarySearch(yesno, r.Substring(480, 1))
                        If r.Substring(76, 1) = "0" And r.Substring(74, 2) = "  " And r.Substring(481, 1) = " " And idx > 0 Then  'adjust data elements for invalid field
                            r = r.Substring(0, 76) & " " & r.Substring(76, r.Length - 77)
                            Call Format_error(r)
                        End If
                        fileo.WriteLine(r)
                    End If
                End If
            End If
        Loop
        filein.Close()
        fileo.Close()
        filein = Nothing
        fileo = Nothing
        'rename temp file to ticket file
        Try
            My.Computer.FileSystem.DeleteFile(nsrcfile)
        Catch ex1 As FileNotFoundException
        Catch ex2 As Exception
            MsgBox("Delete of ticket file " & nsrcfile & " has failed." & vbCrLf &
                "Contact IT for assistance", MsgBoxStyle.Critical, "File Delete failure")
            logwriter.WriteLine(dtmdatetime & " File delete failed" & nsrcfile)
        End Try
        Try
            My.Computer.FileSystem.RenameFile(oname, nfile)
        Catch ex As Exception
            MsgBox("Rename of ticket file " & oname & " has failed." & vbCrLf &
            "Contact IT for assistance", MsgBoxStyle.Critical, "File Rename failure")
            logwriter.WriteLine(dtmdatetime & " File rename failed" & oname)
        End Try
    End Sub
    Private Sub Format_error(ByVal rr As String)
        Try
            logwriter.WriteLine(GetDateTime() & " Record format error, ticket " & rr.Substring(6, 6) & " reformatted record value: " & rr.Substring(0, 77))
            MessageBox.Show("Record format error ticket# " & rr.Substring(6, 6), "Record Format Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub Archivefile(r) 'move files to archive folders
        Dim fsize As Integer
        fsize = My.Computer.FileSystem.GetFileInfo(nsrcfile).Length
        If fsize > 0 Then
            'copy the file to archive folder
            Try
                My.Computer.FileSystem.CopyFile(spath & nfile, ppath & "\h_" & String.Format("{0:yyMMddHHmmss}", GetDateTime) & ".bac")
                logwriter.WriteLine(GetDateTime() & " File transfer complete " & fsize & " bytes transferred archived as " & ppath & "\h_" & String.Format("{0:yyMMddHHmmss}", GetDateTime) & ".bac")
                MessageBox.Show("File transfer complete " & rcnt & " Tickets Transferred as " & vbCrLf & r.Count & " bytes of data", "Ticlet File Transfer", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                logwriter.WriteLine(DateTime.Now & " File archive process failed!")
                MsgBox("File archive process failed! " & vbCrLf & "Please contact IT for assistance" _
                , MsgBoxStyle.Critical, "File archive failure")
                Exit Sub
            End Try
            'Delete old file 
            Fun_dltfil(nsrcfile)
        End If
    End Sub
    Private Sub ArchiveFail(str_status)
        Dim result As String
        MsgBox("FTP transfer failed," & str_status & vbCrLf & "File will be archived to: " & Fun_arcfil(GetDateTime), MsgBoxStyle.Critical, "File Transfer Error")
        My.Computer.FileSystem.CopyFile(spath & nfile, Fun_arcfil(GetDateTime))
        MsgBox("FTP transfer aborted," & vbCrLf & "File will be archived to: " & Fun_arcfil(GetDateTime) & vbCrLf & "Click OK to continue", MsgBoxStyle.Critical, "Transfer Aborted")
        logwriter.WriteLine(GetDateTime() & " File " & Fun_arcfil(GetDateTime) & " archived")
        result = Fun_dltfil(nsrcfile) 'delete file
        abort = True
    End Sub
    Private Sub ArchiveCancel()
        Try
            My.Computer.FileSystem.CopyFile(spath & nfile, fpath & "\h_" & String.Format("{0:yyMMddHHmmss}", DateTime.Now) & ".bac")
            MsgBox("FTP transfer aborted," & vbCrLf & "File will be archived to: " & Fun_arcfil(DateTime.Now) & vbCrLf & "Click OK to continue", MsgBoxStyle.Critical, "Transfer Aborted")
            logwriter.WriteLine(GetDateTime() & " File " & Fun_arcfil(GetDateTime) & " archived")
            result = Fun_dltfil(spath & nfile) 'delete file
        Catch ex As Exception
            MsgBox("File archive process failed! " & vbCrLf & "Please contact IT for assistance" _
            , MsgBoxStyle.Critical, "File archive failure")
            logwriter.WriteLine(GetDateTime() & " File archive failed")
        End Try
    End Sub
    Sub CloseWriter()
        Try
            logwriter.Close()
        Catch ex As Exception
        End Try
    End Sub
    Sub DisposeWriter()
        Try
            logwriter.Dispose()
        Catch ex As Exception
        End Try
    End Sub
    Sub RunCommand()
        'call command on IBM i to move file from IFS
        Dim r As Response
        Ftp1.Session.RemoteEndPoint.HostNameOrAddress = "devlpmnt"
        Ftp1.Session.Username = "COMPUWAY"
        Ftp1.Session.Password = "BATCH$07"
        Ftp1.Timeout = 30000 '10 second timeout
        Ftp1.Session.AsciiExtensions.Enabled = False
        'Connect and send command.
        Try
            Ftp1.Connect()
            'iresult = ftp1.Invoke(FtpCommand.System) 'login to ftp server
        Catch e1 As Dart.Ftp.ProtocolException
            logwriter.WriteLine(DateTime.Now & " Failed to connect to FTP server " & fun_strmsg(e1.Message))
            MsgBox("ProtocolException, job aborted " & vbCrLf & fun_strmsg(e1.Message), MsgBoxStyle.Critical, "FTP failure")
            abort = True
            Exit Sub
        Catch e2 As Dart.Ftp.DataException
            logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & fun_strmsg(e2.Message))
            MsgBox("DataException, " & vbCrLf & fun_strmsg(e2.Message), MsgBoxStyle.Critical, "FTP failure")
            abort = True
            Exit Sub
        Catch e3 As Dart.Ftp.ZStreamException
            logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & fun_strmsg(e3.Message))
            MsgBox("ZStreamException, " & vbCrLf & fun_strmsg(e3.Message), MsgBoxStyle.Critical, "FTP failure")
            abort = True
            Exit Sub
        Catch e4 As System.Net.Sockets.SocketException
            logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & "Socket Error " & e4.SocketErrorCode)
            MsgBox("Connect to FTP server failed,  Sockets Failure" & vbCrLf & " Socket Error " & e4.SocketErrorCode & vbCrLf, MsgBoxStyle.Critical, "FTP socket failure")
            abort = True
            Exit Sub
        Catch e5 As Exception
            logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & "Exception Error " & e5.Message)
            MsgBox("Connect to FTP server failed,  Unknown Failure" & vbCrLf & " Exception Error " & e5.Message & vbCrLf, MsgBoxStyle.Critical, "FTP socket failure")
            abort = True
            Exit Sub
        End Try
        If Not abort Then
            Try
                Ftp1.Authenticate(Ftp1.Session.Username, Ftp1.Session.Password)
            Catch e1 As Dart.Ftp.ProtocolException
                logwriter.WriteLine(DateTime.Now & " Failed to connect to FTP server " & fun_strmsg(e1.Message))
                MsgBox("ProtocolException, job aborted " & vbCrLf & fun_strmsg(e1.Message), MsgBoxStyle.Critical, "FTP failure")
                abort = True
                Exit Sub
            Catch e2 As Dart.Ftp.DataException
                logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & fun_strmsg(e2.Message))
                MsgBox("DataException, " & vbCrLf & fun_strmsg(e2.Message), MsgBoxStyle.Critical, "FTP failure")
                abort = True
                Exit Sub
            Catch e3 As Dart.Ftp.ZStreamException
                logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & fun_strmsg(e3.Message))
                MsgBox("ZStreamException, " & vbCrLf & fun_strmsg(e3.Message), MsgBoxStyle.Critical, "FTP failure")
                abort = True
                Exit Sub
            Catch e4 As System.Net.Sockets.SocketException
                logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & "Socket Error " & e4.SocketErrorCode)
                MsgBox("Connect to FTP server failed,  Sockets Failure" & vbCrLf & " Socket Error " & e4.SocketErrorCode & vbCrLf, MsgBoxStyle.Critical, "FTP socket failure")
                abort = True
                Exit Sub
            Catch e5 As Exception
                logwriter.WriteLine(Date.Now & " Failed to connect to FTP server, " & "Exception Error " & e5.Message)
                MsgBox("Connect to FTP server failed,  Unknown Failure" & vbCrLf & " Exception Error " & e5.Message & vbCrLf, MsgBoxStyle.Critical, "FTP socket failure")
                abort = True
                Exit Sub
            End Try
        End If
        ' Run remote command on host
        If Ftp1.Connected Then
            Dim cmd As String = "rcmd call PGM(DDUDDY/GNCTICKETS) PARM('" & loc & "')"
            'cmd = "rcmd call PGM(TSCPROGRAM/GNCTICKETS) PARM('" & loc & "')"
            r = Ftp1.Send(cmd)
            If r.Code <> 250 Then
                logwriter.WriteLine(Date.Now & " Failed to run CLP command, " & "Exception Error " & r.Text)
                MsgBox("Failed to run CLP command on host" & vbCrLf & " Exception Error " & r.Text, MsgBoxStyle.Critical, "FTP remote command failure")
                abort = True
            Else
                logwriter.WriteLine(Date.Now & " CLP command successful")
                'delete file on server
                Ftp1.Delete(dstfile)
                logwriter.WriteLine(Date.Now & dstfile & "file deleted successfully")
            End If
            Ftp1.Close()
        End If
    End Sub

    Private Function TestConnection() As Boolean
        Try
            sftp.Connect(timeout:=30000)
        Catch i As IOException
            MessageBox.Show("Cannot connect to: " & sadd & vbCrLf & i.Message, "Connection Failure", MessageBoxButtons.OK)
            logwriter.WriteLine(GetDateTime() & " Connection Error IOException")
            Return False
        Catch t As TimeoutException
            MessageBox.Show("Connection Timeout to: " & sadd & vbCrLf & t.Message, "Timeoout Error", MessageBoxButtons.OK)
            logwriter.WriteLine(GetDateTime() & " Connection error TimeoutException")
            Return False
        Catch s As System.Exception
            MessageBox.Show("Cannot connect to: " & sadd & vbCrLf & s.Message & vbCrLf & s.InnerException.InnerException.InnerException.ToString, "Connection Failure", MessageBoxButtons.OK)
            logwriter.WriteLine(GetDateTime() & " Connection error system exception")
            Return False
        End Try
        logwriter.WriteLine(GetDateTime() & " Connection to host successful")
        Try
            sftp.Authenticate(uname, pkey)
        Catch a As System.Security.Authentication.AuthenticationException
            MessageBox.Show("Authentication Error user: " & uname & vbCrLf & a.Message & vbCrLf & a.InnerException.ToString, "Authentication Error", MessageBoxButtons.OK)
            logwriter.WriteLine(GetDateTime() & " Authentication Error")
            Return False
        End Try
        logwriter.WriteLine(GetDateTime() & " Authentication to host successful")
        Return True
    End Function
    Private Function GetDateTime() As Date
        Return DateTime.Now
    End Function
    Private Function Testsystem(lpath) As Boolean
        If Not My.Computer.FileSystem.DirectoryExists(lpath) Then
            MsgBox("System values Not created", MsgBoxStyle.Critical, "System Error")
            stest = False
            Me.Button1.Enabled = False
            Return False
        End If
        Return True
    End Function


    Private Function Checksource() As Boolean
        If Not My.Computer.FileSystem.FileExists(srcfile) Then
            logwriter.WriteLine(GetDateTime() & " Ticket file Not found, transfer cancelled")
            MsgBox("Source file " & srcfile & " cannot be found, transfer cancelled", MsgBoxStyle.Critical, "File Not found")
            Return False
        End If
        Return True
    End Function

    Function Fun_arcfil(ByVal datetim As Date) As String
        Fun_arcfil = fpath & "h_" & String.Format("{0:yyMMddHHmmss}", datetim) & ".bac"
    End Function
    Function Fun_dltfil(ByVal filename As String)
        Dim dtmdatetime As Date = Date.Now
        Try
            My.Computer.FileSystem.DeleteFile(filename)
            logwriter.WriteLine(dtmdatetime & " File " & filename & " deleted")
        Catch ex As Exception
            MsgBox("Delete of source file " & filename & " has failed." & vbCrLf &
            "Contact IT for assistance", MsgBoxStyle.Critical, "File Delete failure")
            logwriter.WriteLine(dtmdatetime & " File delete failed")
        End Try
        Fun_dltfil = True
    End Function
    Function Fun_Strmsg(ByVal msg As String) As String
        'substring message for response portion & remove line feed
        Dim s As String = msg
        Dim dtmdatetime As Date = Date.Now
        Try
            Fun_Strmsg = s.Substring(Strings.InStr(s, "Response:") - 1)
            Fun_Strmsg = Fun_Strmsg.Replace(vbLf, "")
        Catch ex As Exception
            MsgBox("Error Occurred on Command Call " & ex.Message & vbCrLf &
            "Contact IT for assistance", MsgBoxStyle.Critical, "Command call Failure Please verify")
            logwriter.WriteLine(dtmdatetime & " Quote Command Failure")
        End Try
        Return 0
    End Function
End Class
