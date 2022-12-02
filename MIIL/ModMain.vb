Public Module ModMain

#Region "声明"
    Public Const VersionName As String = "1.0.0"
    Public Const VersionCode As Integer = 0
    ''' <summary>
    ''' 应用程序是否为快照版。
    ''' </summary>
    Public Const IsSnapshot As Boolean = False
    ''' <summary>
    ''' 应用程序是否为调试模式。
    ''' </summary>
    Public Const ModeDebug As Boolean = True
    ''' <summary>
    ''' 程序的启动路径，以“\”结尾。
    ''' </summary>
    Public Path As String = AppDomain.CurrentDomain.SetupInformation.ApplicationBase
    ''' <summary>
    ''' 包含程序名的完整启动路径。
    ''' </summary>
    Public FullPath As String = Path & AppDomain.CurrentDomain.SetupInformation.ApplicationName
#End Region

End Module
