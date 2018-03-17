Public Class Form1
    Private SndKeys As New Class1
    'https://msdn.microsoft.com/en-us/library/windows/desktop/dd375731(v=vs.85).aspx
    Public Const WM_KEYDOWN = &H100
    Public Const VK_F1 = &H70
    Private Delegate Sub voidDelegate()
    Dim _t As Threading.Thread
    Private IsAbortedManully As Boolean '检测 SndKeys的线程 是否人为终止的
    Private IsButton1Focused As Boolean '检测 Button1 是否有焦点
    'Private Const 
    'https://stackoverflow.com/questions/13727172/vb-net-keydown-event-on-whole-form BUG: 无法执行
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        'If TextBox1.Text <> "" Then 检测按键的 m.Msg 值
        '    Stop
        'End If
        If m.Msg = WM_KEYDOWN Then ' BUG: 无法执行
            Stop
            Select Case m.WParam
                Case VK_F1
                    Stop
                    SndKeys.Start()
                Case Keys.F2
                    SndKeys.AbortAll()
                Case Keys.A
                    TextBox1.Text &= "good!"
            End Select
        End If
        MyBase.WndProc(m)
    End Sub
    ''这个只有在窗口处于焦点才有效
    'Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
    '    Select Case keyData
    '        Case Keys.F2
    '            ' Do something 
    '            TextBox1.Text = "good!"
    '        Case Keys.F3
    '            ' Do more

    '        Case Keys.Escape
    '            ' Crap

    '        Case Else
    '            Return MyBase.ProcessCmdKey(msg, keyData)

    '    End Select

    '    Return True
    'End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        IsAbortedManully = False '初始化，为将来是否人为终止 线程 做铺垫
        _t.Sleep(3000) '开火前准备
        SndKeys.Start()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SndKeys.AbortAll()
        IsAbortedManully = True
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IsAbortedManully = True
        _t = New Threading.Thread(Sub()
                                      Do
                                          If SndKeys.IsThreadAlive Then '如果线程已经启动
                                              Me.Invoke(New voidDelegate(Sub()
                                                                             Me.Text = "Started!"
                                                                             TextBox1.Visible = False '保护
                                                                             CheckBox1.Enabled = False
                                                                         End Sub))
                                              If IsButton1Focused Then
                                                  SndKeys.AbortAll()
                                                  IsAbortedManully = True
                                              End If
                                          ElseIf CheckBox1.Checked And (Not IsAbortedManully) Then '如果不是人为终止的 才能重新开始 'BUG 这一步无法执行
                                                  SndKeys.Start()
                                              Else
                                                  Me.Invoke(New voidDelegate(Sub()
                                                                             Me.Text = "Stopped!"
                                                                             TextBox1.Visible = True '取消保护
                                                                             CheckBox1.Enabled = True
                                                                         End Sub))
                                          End If
                                          Me.Invoke(New voidDelegate(Sub()
                                                                         If Me.Focused Then
                                                                             Label3.Text = "GotFocus"
                                                                             'Else
                                                                             '    Label3.Text = "LostFocus"
                                                                         End If
                                                                     End Sub))
                                      Loop
                                  End Sub)
        _t.Start()
    End Sub
    '当窗口关闭时
    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        SndKeys.AbortAll()
        If _t.IsAlive Then
            _t.Abort()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        SndKeys.Interval = Val(TextBox1.Text)
        Label2.Text = Val(TextBox1.Text) '同步显示在label2上
    End Sub

    Private Sub Button1_GotFocus(sender As Object, e As System.EventArgs) Handles Button1.GotFocus
        IsButton1Focused = True
    End Sub

    Private Sub Button1_LostFocus(sender As Object, e As System.EventArgs) Handles Button1.LostFocus
        IsButton1Focused = False
    End Sub
End Class
