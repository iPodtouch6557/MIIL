Public Class FormMain

#Region "窗体事件"
    '从任务栏关闭
    Private Sub FormMain_Closing(ByVal sender As Object, ByVal e As ComponentModel.CancelEventArgs) Handles Me.Closing
        EndForcely()
    End Sub
    '窗口拖拽
    Private Sub WindowTitle_MouseLeftButtonDown() Handles WindowTitle.MouseLeftButtonDown
        DragMove()
    End Sub
#End Region

#Region "顶部选择栏 | 右侧"
    Private BtnMin_
    Private Sub BtnWindowTitleMinisize_MouseLeftButtonDown(ByVal sender As Object, ByVal e As MouseButtonEventArgs) Handles BtnWindowTitleMinisize.MouseLeftButtonDown

    End Sub
#End Region

End Class
