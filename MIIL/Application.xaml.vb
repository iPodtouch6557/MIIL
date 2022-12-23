Imports System.IO
Imports System.Windows.Threading

Class Application
    Private Sub Application_Startup(ByVal sender As Object, ByVal e As StartupEventArgs) Handles Me.Startup
        Try
            If Not Directory.Exists(Path & "MIIL") Then Directory.CreateDirectory(Path & "MIIL")
            If Not File.Exists(Path & "MIIL\Log.txt") Then File.Create(Path & "MIIL\Log.txt").Dispose()
            Log("[Start] 程序版本：" & VersionType & " " & VersionName & " (" & VersionCode & ")")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Application_DispatcherUnhandledException(ByVal sender As Object, ByVal e As DispatcherUnhandledExceptionEventArgs) Handles Me.DispatcherUnhandledException
        ExShow(e.Exception, "出现未经处理的异常", Errorlevel.Feedback)
    End Sub
    Private Sub Application_Exit(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Exit
        Log("[System] 程序已退出")
    End Sub
End Class
