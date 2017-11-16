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
        Me.TextBox6.Text = settings.Private_Key
        Me.TextBox7.Text = settings.PassPhrase
        Me.TextBox8.Text = settings.username

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim plist As New List(Of Utils.Appvalue)
        plist.Add(New Utils.Appvalue With {.Key = "Location", .Value = Me.TextBox1.Text})
        plist.Add(New Utils.Appvalue With {.Key = "FTP_Address", .Value = Me.TextBox2.Text})
        plist.Add(New Utils.Appvalue With {.Key = "Destination_Path", .Value = Me.TextBox3.Text})
        plist.Add(New Utils.Appvalue With {.Key = "Source_Device_Address", .Value = Me.TextBox4.Text})
        plist.Add(New Utils.Appvalue With {.Key = "Source_Path", .Value = Me.TextBox5.Text})
        plist.Add(New Utils.Appvalue With {.Key = "Private_Key", .Value = Me.TextBox6.Text})
        plist.Add(New Utils.Appvalue With {.Key = "PassPhrase", .Value = Me.TextBox7.Text})
        plist.Add(New Utils.Appvalue With {.Key = "Username", .Value = Me.TextBox8.Text})
        Dim result As Boolean = Utils.AddUpdateAllAppSettings(plist)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class