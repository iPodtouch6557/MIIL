Imports System.IO

Public Module ModBase

#Region "声明"
    ''' <summary>
    ''' 应用程序的版本名称，格式为“1.0.0”。
    ''' </summary>
    Public Const VersionName As String = "1.0.0"
    ''' <summary>
    ''' 应用程序的版本编号。
    ''' </summary>
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

#Region "文本"
    ''' <summary>
    ''' 填充字符串长度。
    ''' </summary>
    ''' <param name="str">要填充的字符串。</param>
    ''' <param name="code">用于填充的字符。</param>
    ''' <param name="length">要填充的长度。</param>
    ''' <returns></returns>
    Public Function FillLength(ByVal str As String, ByVal code As String, ByVal length As Byte)
        If (Len(str) > length) Then Return Mid(str, 1, length)
        Return Mid(str.Replace(" ", "-").PadRight(length), str.Length + 1).Replace(" ", code) & str
    End Function
#End Region

#Region "文件"
    ''' <summary>
    ''' 获取文件或URL所在文件夹的路径。不包含路径将会抛出异常。
    ''' </summary>
    ''' <param name="FilePath">文件路径或URL。</param>
    ''' <returns>文件或URL所在文件夹的路径，以“\”结尾。</returns>
    Public Function GetFileDirectory(ByRef FilePath As String)
        If FilePath.EndsWith("\") Or FilePath.EndsWith("/") Then
            '已经为目录
            Return FilePath
        ElseIf FilePath.Contains(":\") Or FilePath.Contains(":/") Then
            '非简写的文件路径或URL
            Return Left(FilePath, FilePath.LastIndexOfAny({"\", "/"}) + 1)
        Else
            '简写的文件路径，已弃用
            'Dim RealPath As String
            'RealPath = Path & "MIIL\" & FilePath
            'Return Left(RealPath, RealPath.LastIndexOfAny({"\", "/"}) + 1)

            '不包含路径
            Throw New Exception("不包含路径：" & FilePath)
        End If
    End Function

    '注册表I/O
    ''' <summary>
    ''' 读取注册表。
    ''' </summary>
    ''' <param name="Key">键。</param>
    ''' <param name="DefaultValue">出现错误时返回的默认值。</param>
    ''' <returns></returns>
    Public Function ReadReg(ByVal Key As String, Optional ByVal DefaultValue As String = "") As String
        Try
            Dim ParentKey As Microsoft.Win32.RegistryKey
            ParentKey = My.Computer.Registry.CurrentUser
            Dim SoftKey As Microsoft.Win32.RegistryKey
            SoftKey = ParentKey.OpenSubKey("Software\MIIL", True)
            If SoftKey Is Nothing Then
                Return DefaultValue
            Else
                Dim readValue As New Text.StringBuilder
                readValue.AppendLine(SoftKey.GetValue(Key))
                Dim value = readValue.ToString.Replace(vbCrLf, "") '去除莫名的回车
                Return If(value = "", DefaultValue, value) '错误则返回默认值
            End If
        Catch ex As Exception
            ExShow(ex, "读取注册表时出现异常: " & Key)
            Return DefaultValue
        End Try
    End Function
#End Region

#Region "系统"
    ''' <summary>
    ''' 获取格式类似于“11:08:52.037”的当前时间的字符串。
    ''' </summary>
    ''' <returns>当前的时间。</returns>
    Public Function GetTime() As String
        Return Date.Now.ToLongTimeString & "." & FillLength(Date.Now.Millisecond, "0", 3)
    End Function
#End Region

    Private LogLock As New Object
    ''' <summary>
    ''' 输出Log(未完成)。
    ''' </summary>
    ''' <param name="Message">要输出的文本。</param>
    ''' <param name="Errlevel">错误等级。</param>
    Public Sub Log(Message As String, Optional ByVal Errlevel As Errorlevel = Errorlevel.Slient)
        Try
            Dim OutputText = "[" & GetTime() & "] " & Message
            Debug.WriteLine(OutputText)
            If Not File.Exists(Path & "MIIL\Latest.log") Then
                File.Create(Path & "MIIL\Latest.log").Dispose()
            End If
            '阻止线程并行
            SyncLock LogLock
                Using Writter As New StreamWriter(Path & "MIIL\Latest.log", True)
                    Writter.WriteLine(OutputText)
                    Writter.Close()
                End Using
            End SyncLock
            Select Case Errlevel

            End Select
        Catch
        End Try
    End Sub
End Module
