Imports IBM.Data.DB2.iSeries
Public Class Class1
    Public Shared Function GetConnection(ByVal strDatabase As String) As iDB2Connection
        Return New iDB2Connection(GetConnectionString(strDatabase))
    End Function
    Public Shared Function GetConnectionString(ByVal strDatabase As String) As String
        Dim strString As String = Nothing
        Select Case strDatabase
            Case Is = "test"
                strString = "datasource=testing;userid=msolap;password=star@84;"
            Case Is = "prod"
                strString = "datasource=s10b1350;userid=msolap;password=star@84;"
        End Select
        Return strString
    End Function
    Public Class Appvalues2
        Public Property Location As String
        Public Property FtpAaddress As String
    End Class

End Class
