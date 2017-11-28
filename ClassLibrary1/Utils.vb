Imports System.Configuration

Public Class Utils
    Public Shared Function GetAllSettings() As Appvalues
        'get setting from app.config
        Dim appval As New Appvalues
        Try
            Dim appsettings = System.Configuration.ConfigurationManager.AppSettings

            If appsettings.Count = 0 Then
                Return appval
            Else
                For Each key As String In appsettings.AllKeys
                    Select Case key
                        Case = "Location"
                            appval.Location = appsettings(key)
                        Case = "FTP_Address"
                            appval.FTP_Address = appsettings(key)
                        Case = "Destination_Path"
                            appval.Destination_Path = appsettings(key)
                        Case = "Source_Path"
                            appval.Source_Path = appsettings(key)
                        Case = "Source_Device_Address"
                            appval.Source_Device_Address = appsettings(key)
                        Case = "PassPhrase"
                            appval.PassPhrase = appsettings(key)
                        Case = "Password"
                            appval.Password = appsettings(key)
                        Case = "Username"
                            appval.Username = appsettings(key)
                    End Select

                Next
            End If
        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error reading app settings")
        End Try
        Return appval
    End Function
    Public Shared Function GetSetting(Key As String) As String
        Try
            Dim appsettings = System.Configuration.ConfigurationManager.AppSettings
            Dim result As String = appsettings(Key)
            If IsNothing(result) Then
                result = "not found"
            End If
            Return result
        Catch ex As Exception
        End Try
        Return "no keys found"
    End Function
    Public Shared Function AddUpdateAllAppSettings(plist As List(Of Appvalue)) As Boolean
        Try
            Dim configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            Dim settings = configFile.AppSettings.Settings
            For Each p As Appvalue In plist
                If IsNothing(p.Key) Then
                    settings.Add(p.Key, p.Value)
                Else
                    settings(p.Key).Value = p.Value
                End If
            Next
            configFile.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name)
        Catch ex As Exception

        End Try
        Return False
    End Function
    Public Class Appvalues
        Public Property Location As String
        Public Property FTP_Address As String
        Public Property Destination_Path As String
        Public Property Source_Path As String
        Public Property Source_Device_Address As String
        Public Property PassPhrase As String
        Public Property Password As String
        Public Property Username As String
    End Class
    Public Class Appvalue
        Public Property Key As String
        Public Property Value As String
    End Class
End Class
