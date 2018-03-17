Public Class Class1
    Private _t As System.Threading.Thread '多线程容器
    Private _txtPath As String = "Lines.txt" '将要发出的话 的 txt 的地址
    Private _f1 As System.IO.StreamReader
    'Private _p As Process '意义不明，可能只是用来启动程序
    Public IsThreadAlive As Boolean '方便外部访问 线程是否还活着
    Public Interval As Int16 = 0 '发送信息的冷却时间,默认是0秒一个

    Public Sub Start()
        If Not System.IO.File.Exists(_txtPath) Then
            Dim f As New System.IO.FileStream(_txtPath, IO.FileMode.OpenOrCreate) '创建一个（如果没有txt）
            f.Close()
            MessageBox.Show("请到文件根目录寻找Lines.txt以回车为分行标志写你要发出的内容" & vbCrLf & "made by Banic", "Fresh Start")
            Exit Sub
        End If
        _f1 = New System.IO.StreamReader(_txtPath, System.Text.Encoding.Default)
        _t = New Threading.Thread(Sub()
                                      '_p = New Process
                                      '_p.StartInfo.FileName = "Notepad.exe"
                                      '_p.Start() '之后才sendkey
                                      SndKeys()
                                      AbortAll()
                                  End Sub)
        _t.Start()
        IsThreadAlive = True
    End Sub

    Public Sub AbortAll()
        IsThreadAlive = False
        Try
            If _t.IsAlive Then
                _t.Abort()
            End If
        Finally
        End Try
        '_p.Close()
    End Sub

    Private Sub SndKeys()
        Dim line As String = GetLine()
        'Threading.Thread.Sleep(3000) '准备开火 '准备阶段交给Form1
        While Not _f1.EndOfStream '如果没有读完
            Threading.Thread.Sleep(Interval)
            SendKeys.SendWait(line)
            SendKeys.SendWait("{ENTER}")
            Do
                line = GetLine()
            Loop Until line <> "" Or _f1.EndOfStream
        End While
    End Sub

    Private Function GetLine() As String
        Dim line As String = _f1.ReadLine
        Return line
    End Function
End Class
