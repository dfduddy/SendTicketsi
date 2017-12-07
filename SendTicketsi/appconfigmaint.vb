Imports System.Configuration
Imports ClassLibrary1


Public Class Appconfigmaint

    Private Sub Appconfigmaint_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim settings As New Utils.Appvalues
        settings = Utils.GetAllSettings
        Me.TextBox1.Text = settings.Location
        Me.TextBox2.Text = settings.FTP_Address
        Me.TextBox3.Text = settings.Destination_Path
        Me.TextBox4.Text = settings.Source_Device_Address
        Me.TextBox5.Text = settings.Source_Path
        Me.TextBox6.Text = settings.Password
        Me.TextBox7.Text = settings.PassPhrase
        Me.TextBox8.Text = settings.Username

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If EditValues = True Then
            Dim plist As New List(Of Utils.Appvalue)
            plist.Add(New Utils.Appvalue With {.Key = "Location", .Value = Me.TextBox1.Text})
            plist.Add(New Utils.Appvalue With {.Key = "FTP_Address", .Value = Me.TextBox2.Text})
            plist.Add(New Utils.Appvalue With {.Key = "Destination_Path", .Value = Me.TextBox3.Text})
            plist.Add(New Utils.Appvalue With {.Key = "Source_Device_Address", .Value = Me.TextBox4.Text})
            plist.Add(New Utils.Appvalue With {.Key = "Source_Path", .Value = Me.TextBox5.Text})
            plist.Add(New Utils.Appvalue With {.Key = "Password", .Value = Me.TextBox6.Text})
            plist.Add(New Utils.Appvalue With {.Key = "PassPhrase", .Value = Me.TextBox7.Text})
            plist.Add(New Utils.Appvalue With {.Key = "Username", .Value = Me.TextBox8.Text})
            Dim result As Boolean = Utils.AddUpdateAllAppSettings(plist)
            Dim lpath As String = Me.TextBox5.Text & "log" 'log file path
            Checkfolder(lpath)
            Me.Dispose()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub


    Private Sub Checkfolder(lpath)
        'check for log file folder and create if necessary
        If Not My.Computer.FileSystem.DirectoryExists(lpath) Then
            Try
                My.Computer.FileSystem.CreateDirectory(lpath)
            Catch ex As Exception
                If ex.Message = "Path/File access error." Then
                    'ignore
                Else
                    MessageBox.Show(Me, "Error Creating log file", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Sub
                End If
            End Try
        End If
    End Sub
    Public Function EditValues() As Boolean
        If TextBox1.Text.Trim.Length <= 0 Then
            MessageBox.Show(Me, "Location cannot be blank", "Missing Input", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox1.Focus()
            Return False
        End If
        If TextBox2.Text.Trim.Length <= 0 Then
            MessageBox.Show(Me, "Host Address cannot be blank", "Missing Input", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox2.Focus()
            Return False
        End If
        If TextBox3.Text.Trim.Length <= 0 Then
            MessageBox.Show(Me, "Destination Path cannot be blank", "Missing Input", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox3.Focus()
            Return False
        End If
        If TextBox4.Text.Trim.Length <= 0 Then
            MessageBox.Show(Me, "Source Device Address cannot be blank", "Missing Input", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox4.Focus()
            Return False
        End If
        If TextBox5.Text.Trim.Length <= 0 Then
            MessageBox.Show(Me, "Source Path cannot be blank", "Missing Input", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox5.Focus()
            Return False
        End If
        If TextBox6.Text.Trim.Length <= 0 Then
            MessageBox.Show(Me, "Password cannot be blank", "Missing Input", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox6.Focus()
            Return False
        End If
        If TextBox8.Text.Trim.Length <= 0 Then
            MessageBox.Show(Me, "Username cannot be blank", "Missing Input", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBox8.Focus()
            Return False
        End If
        Return True
    End Function
End Class