Imports System.IO
Imports Microsoft.VisualBasic.Logging
Imports MS.Internal.Text.TextInterface

Public Module Modules

#Region "主要方法"
    Private LogLock As New Object
    Public Sub Log(Message As String)
        Try
            Dim OutputText = "[" & GetTime() & "] " & Message
            Debug.WriteLine(OutputText)
            If Not File.Exists(Path & "PCL\log.txt") Then
                File.Create(Path & "PCL\log.txt").Dispose()
            End If
            SyncLock LogLock
                Using Writter As New StreamWriter(Path & "PCL\log.txt", True)
                    Writter.WriteLine(OutputText)
                    Writter.Close()
                End Using
            End SyncLock
        Catch
        End Try
    End Sub
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
            'Log("读取注册表时出错：" & ex)
        End Try
    End Function
#End Region

#Region "错误与反馈"
    ''' <summary>
    ''' 错误等级，范围为0至5之间的正整数。
    ''' </summary>
    Public Enum Errorlevel As Integer
        ''' <summary>
        ''' 仅输出Log。
        ''' </summary>
        Slient = 0
        ''' <summary>
        ''' 仅通过提示方式通知调试用户。
        ''' </summary>
        DebugOnly = 1
        ''' <summary>
        ''' 通过提示方式通知全部用户。
        ''' </summary>
        Hint = 2
        ''' <summary>
        ''' 通过弹窗方式通知全部用户。
        ''' </summary>
        MsgBox = 3
        ''' <summary>
        ''' 通过弹窗方式通知全部用户，并要求反馈。
        ''' </summary>
        Feedback = 4
        ''' <summary>
        ''' 尽可能多的弹出提示，要求反馈，然后终止程序。
        ''' </summary>
        Barrier = 5
    End Enum
    ''' <summary>
    ''' 处理错误。
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <param name="Description"></param>
    ''' <param name="Errlevel"></param>
    Public Sub ExShow(ex As Exception, Description As String, Optional Errlevel As Integer = Errorlevel.DebugOnly)
        Select Case Errlevel
            Case Errorlevel.Slient
                Log("[Error] " & Description)
        End Select
    End Sub
#End Region

#Region "系统"
    ''' <summary>
    ''' 获取格式类似于“11:08:52.037”的当前时间的字符串。
    ''' </summary>
    ''' <returns></returns>
    Public Function GetTime() As String
        Return Date.Now.ToLongTimeString & "." & FillLength(Date.Now.Millisecond, "0", 3)
    End Function
#End Region
End Module
