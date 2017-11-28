'SendTicketsi Transfer batch scale tickets to Scoular (interactive)
'Uses Dart PowerTCP SSH and SFTP for .NET 4.10
'Control-S calls configuration application
Imports System.IO
Imports ClassLibrary1
Public Class SendTickets
    Dim ssh1 As Dart.Ssh.Optimizations
    Dim sftp As New Dart.Ssh.Ssh
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
    Const sfile As String = "tickets.txt"
    Const nfile As String = "tickets.dat"
    Const iprefix As String = "/P"
    Const isuffix As String = "600I.PCP600I"
    Const str_dir As String = "PCSPRDLIB"   'prod
    Dim srcfile As String, dstfile As String, nsrcfile As String
    Dim abort As Boolean = False
    Dim stest As Boolean = True
    Dim logwriter As StreamWriter


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Dim RetValues As Utils.Appvalues = Utils.GetAllSettings()
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            logwriter.WriteLine(GetDateTime() & " Process Cancelled----------")
        Catch ex As Exception
        End Try

        Try
            logwriter.Close()
        Catch ex As Exception
        End Try
        Me.Dispose()
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
        'test for values
        If Not Testsystem(lpath) Then Exit Sub
        'Cleanup old ftp files
        Dim retval As Boolean
        retval = TestConnection()
        logwriter.WriteLine(GetDateTime() & " Process Completed----------")
        logwriter.Close()
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
            ssh1.BlockSize = 2056
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
        srcfile = spath & sfile 'source file on pc
        nsrcfile = spath & nfile
        dstfile = dpath & iprefix & loc & isuffix 'destination file on AS/400
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
            MsgBox("System values not created", MsgBoxStyle.Critical, "System Error")
            stest = False
            Me.Button1.Enabled = False
            Return False
        End If
        Return True
    End Function
End Class
