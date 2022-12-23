Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Public Module ModMain

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
    ''' 应用程序的版本类型(Release Snapshot Non-Stable)
    ''' </summary>
    Public Const VersionType As String = "Non-Stable"
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

    '窗体
    Public FrmMain As FormMain
#End Region

#Region "枚举"
    ''' <summary>
    ''' 程序返回值。
    ''' </summary>
    Public Enum ExitCode As Integer
        ''' <summary>
        ''' 由用户主动关闭。
        ''' </summary>
        Success = 0
        ''' <summary>
        ''' 由程序正常关闭。
        ''' </summary>
        Normal = 1
        Fail = 2
    End Enum
#End Region

#Region "文本"
    ''' <summary>
    ''' 填充字符串长度。
    ''' </summary>
    ''' <param name="str">要填充的字符串。</param>
    ''' <param name="code">用于填充的字符。</param>
    ''' <param name="length">要填充的长度。</param>
    ''' <returns></returns>
    Public Function FillLength(ByVal str As String, ByVal code As String, ByVal length As Byte) As String
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
    Public Function GetFileDirectory(ByRef FilePath As String) As String
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
                readValue.AppendLine(CStr(SoftKey.GetValue(Key)))
                Dim value = readValue.ToString.Replace(vbCrLf, "") '去除回车
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

#Region "错误与反馈"
    ''' <summary>
    ''' 错误等级，范围为0与5之间的整数。
    ''' </summary>
    Public Enum Errorlevel As Integer
        ''' <summary>
        ''' 仅输出Log。
        ''' </summary>
        Slient = 0
        ''' <summary>
        ''' 仅通过弹窗方式通知调试用户。
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
    ''' 处理错误(未完成)。
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <param name="Description">对错误的描述，例如“下载文件时出错”。</param>
    ''' <param name="Errlevel">错误等级。</param>
    Public Sub ExShow(ex As Exception, Description As String, Optional Errlevel As Errorlevel = Errorlevel.DebugOnly)
        Select Case Errlevel
            Case Errorlevel.Slient
                Log("[Error] " & Description)
                Log(GetStringFromException(ex))
            Case Errorlevel.DebugOnly
                Log("[Error] " & Description)
                Dim StrEx As String = GetStringFromException(ex)
                Log(StrEx)
                If ModeDebug Then
                    MsgBox(Description & vbCrLf & StrEx)
                End If
        End Select
    End Sub
    ''' <summary>
    ''' Exception转String。
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <param name="IsLong">是否添加堆栈信息，默认为True。</param>
    ''' <returns>转化后的String。</returns>
    Public Function GetStringFromException(ByVal ex As Exception, Optional ByVal IsLong As Boolean = True) As String
        '构造错误信息
        GetStringFromException = ex.Message.Replace(vbCrLf, "")
        If Not IsNothing(ex.InnerException) Then
            GetStringFromException = GetStringFromException & vbCrLf & "内部错误: " & ex.InnerException.Message.Replace(vbCrLf, "")
            If Not IsNothing(ex.InnerException.InnerException) Then
                GetStringFromException = GetStringFromException & vbCrLf & "内部错误: " & ex.InnerException.Message.Replace(vbCrLf, "")
            End If
        End If

        '添加堆栈
        If IsLong Then
            GetStringFromException = GetStringFromException & vbCrLf
            For Each Stk As String In ex.StackTrace.Split(vbCrLf)
                If Stk.Contains("MIIL.") Then '仅显示以MIIL为根命名空间的堆栈
                    GetStringFromException = GetStringFromException & vbCrLf & Stk
                End If
            Next
        End If

        '清理双回车
        GetStringFromException = GetStringFromException.Replace(vbCrLf & vbCrLf, vbCrLf).Replace(vbCrLf & vbCrLf, vbCrLf)
    End Function
#End Region

#Region "加密"
    'Base64
    ''' <summary>
    ''' 对字符串进行Base64编码。
    ''' </summary>
    ''' <param name="str">要编码的字符串。</param>
    ''' <returns>编码后的字符串。</returns>
    Public Function Base64Encode(str As String) As String
        Return Convert.ToBase64String(Encoding.ASCII.GetBytes(str))
    End Function
    ''' <summary>
    ''' 对字符串进行Base64解码。
    ''' </summary>
    ''' <param name="b64">要解码的字符串。</param>
    ''' <returns>解码后的字符串。</returns>
    Public Function Base64Decode(b64 As String) As String
        Return Encoding.ASCII.GetString(Convert.FromBase64String(b64))
    End Function

    'AES
    ''' <summary>
    ''' 加密文本为AES。
    ''' </summary>
    ''' <param name="Source">源文本。</param>
    ''' <param name="Key">密钥。</param>
    ''' <returns>加密后的字符串。</returns>
    Public Function EncriptStr(ByVal Source As String, ByVal Key As String) As String
        Dim encripter As Aes = Aes.Create("AES")
        '设置密钥
        Dim keyBytes() As Byte = (New MD5CryptoServiceProvider).ComputeHash(Encoding.Unicode.GetBytes(Key))

        encripter.BlockSize = keyBytes.Length * 8
        encripter.Key = keyBytes
        encripter.IV = keyBytes
        encripter.Mode = CipherMode.CBC
        encripter.Padding = PaddingMode.PKCS7
        Dim cripter As ICryptoTransform = encripter.CreateEncryptor()
        Dim inBuff As Byte() = Encoding.Unicode.GetBytes(Source)
        Return Convert.ToBase64String(cripter.TransformFinalBlock(inBuff, 0, inBuff.Length))
    End Function
    ''' <summary>
    ''' 解密AES编码的字符串。
    ''' </summary>
    ''' <param name="EncodedStr">AES加密后的密文。</param>
    ''' <param name="Key">加密时使用的密钥。</param>
    ''' <returns>解密后的字符串。</returns>
    Public Function DecriptStr(ByVal EncodedStr As String, ByVal Key As String) As String
        Dim decripter As Aes = Aes.Create("AES")
        '设置密钥
        Dim keyBytes() As Byte = (New MD5CryptoServiceProvider).ComputeHash(Encoding.Unicode.GetBytes(Key))
        decripter.BlockSize = keyBytes.Length * 8

        decripter.Key = keyBytes
        decripter.IV = keyBytes
        decripter.Mode = CipherMode.CBC
        decripter.Padding = PaddingMode.PKCS7
        Dim cripter As ICryptoTransform = decripter.CreateDecryptor()
        Dim inBuff As Byte() = Convert.FromBase64String(EncodedStr)
        Return Encoding.Unicode.GetString(cripter.TransformFinalBlock(inBuff, 0, inBuff.Length))
    End Function

    'MD5
    ''' <summary>
    ''' 获取字符串MD5。发生异常会返回空字符串。
    ''' </summary>
    ''' <param name="str">要获取MD5的字符串。</param>
    ''' <returns>成功则返回MD5；发生异常时返回空字符串。</returns>
    Public Function GetMD5(ByVal str As String) As String
        Try
            'PCL那里抄的啊，自己改了改而已XD
            Dim result As String = ""
            Dim hashvalue As Byte() = CType(CryptoConfig.CreateFromName("MD5"), HashAlgorithm).ComputeHash(Encoding.ASCII.GetBytes(str))
            Dim i As Double
            For i = 0 To hashvalue.Length - 1
                result += FillLength(Hex(hashvalue(i)).ToLower, "0", 2)
            Next
            Return result
        Catch ex As Exception
            ExShow(ex, "获取字符串MD5时出错")
            Return ""
        End Try
    End Function
#End Region

    ''' <summary>
    ''' 以正常方式结束程序。
    ''' </summary>
    ''' <param name="Code">程序需要返回的值。</param>
    Public Sub ExitNormally(Code As ExitCode)
        Try
            Log("[System] 收到关闭指令，返回值: " & Code)
            My.Application.Shutdown(Code)
        Catch ex As Exception
            ExShow(ex, "正常关闭程序失败", Errorlevel.Barrier)
        End Try
    End Sub
    Public Sub EndForcely()
        On Error Resume Next
        My.Application.Shutdown()
        Process.GetCurrentProcess.Kill()
        End
    End Sub

    Private LogLock As New Object
    ''' <summary>
    ''' 输出Log。
    ''' </summary>
    ''' <param name="Message">要输出的文本。</param>
    Public Sub Log(Message As String)
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
        Catch
        End Try
    End Sub
End Module
