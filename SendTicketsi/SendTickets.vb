'SendTicketsi Transfer batch scale tickets to Scoular (interactive)
'Uses Dart PowerTCP FTP for .NET 4.7
Imports System.IO
Imports ClassLibrary1
Public Class SendTickets
    Dim dbase As String = "test" ' production or test system"
    Dim splib As String
    Dim result As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Dim RetValues As Utils.Appvalues = Utils.GetAllSettings()
        SetValues(RetValues)

        While result = False
            result = TestConnection(RetValues)
        End While
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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

    End Sub
    Private Sub SetValues(retvalues As Utils.Appvalues)
        Dim ftp As New Dart.Ftp.Ftp
        Dim sftp As New Dart.Ssh.Ssh
        ftp.Session.RemoteEndPoint.HostNameOrAddress = retvalues.FTP_Address
        ftp.Session.Username = "compuway"
        ftp.Session.Password = "batch01"
        sftp.Connection.RemoteEndPoint.HostNameOrAddress = retvalues.FTP_Address
        sftp.Connect()
        sftp.Authenticate("user", retvalues.PassPhrase, retvalues.Private_Key)
    End Sub
    Private Function TestConnection(retvalues As Utils.Appvalues) As Boolean

        Return False
    End Function

End Class
