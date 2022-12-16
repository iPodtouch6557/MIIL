Imports System.IO

Class Application
    Private Sub Application_Startup(ByVal sender As Object, ByVal e As StartupEventArgs) Handles Me.Startup
        Try
            If Not Directory.Exists(Path & "MIIL\") Then Directory.CreateDirectory(Path & "MIIL\")
            If File.Exists(Path & "MIIL\Latest.log") Then File.Create(Path & "MIIL\Latest.log").Dispose()
            Log("[Start] 程序版本：" & VersionName & " (" & VersionCode & ")")
        Catch ex As Exception

        End Try
    End Sub
End Class
