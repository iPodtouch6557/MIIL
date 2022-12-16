Public Module ModMain

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
        Crash = 5
    End Enum
    ''' <summary>
    ''' 处理错误(未完成)。
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <param name="Description"></param>
    ''' <param name="Errlevel"></param>
    Public Sub ExShow(ex As Exception, Description As String, Optional Errlevel As Errorlevel = Errorlevel.DebugOnly)
        Select Case Errlevel
            Case Errorlevel.Slient
                Log("[Error] " & Description)
            Case Errorlevel.DebugOnly
                Log("[Error] " & Description)
        End Select
    End Sub
    ''' <summary>
    ''' Exception转String(未完成)。
    ''' </summary>
    ''' <param name="ex">发生的Exception。</param>
    ''' <returns>转化后的String。</returns>
    Public Function GetStringFromException(ByVal ex As Exception) As String
        GetStringFromException = ex.Message
    End Function
#End Region

End Module
