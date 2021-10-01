Imports System.Net
Friend Class CSubSystemListener
    Inherits System.Net.Sockets.TcpListener

#Region "變數宣告"
    Private m_ListenThread As System.Threading.Thread   ' 等待接收Client連入的執行緒
    Private m_Dispatcher As CSubSystemDispatcher
#End Region
    'Private temp As IPEndPoint
    Public Sub New(ByVal AddressOrHostName As String, ByVal Port As Integer, ByVal Dispatcher As CSubSystemDispatcher) 'ByVal OnConnect As ConnectDel)

        '依照個別IP做Listening 20121212 MT 
        MyBase.New(IPAddress.Parse(AddressOrHostName), Port)

        Me.m_Dispatcher = Dispatcher
    End Sub

    ' 開始聆聽
    Public Sub StartListen()
        If Not MyBase.Active Then
            Try
                MyBase.Start()      ' Start Listener
            Catch ex As System.Exception
                Throw New Exception("啟動SubSystemDispatcher的TcpListener失敗。錯誤原因：" & ex.Message)
            End Try

            ' 啟動Listen的執行緒
            Me.m_ListenThread = New System.Threading.Thread(AddressOf ListenProcess)
            Me.m_ListenThread.Name = "SubSystemDispatcherListener"
            Me.m_ListenThread.SetApartmentState(Threading.ApartmentState.STA)
            Me.m_ListenThread.Start()
        End If
    End Sub

    ' 停止聆聽
    Public Sub StopListen()
        If MyBase.Active Then
            Try
                MyBase.Stop()       ' Stop Listener
            Catch ex As System.Exception
                Debug.WriteLine("關閉SubSystemDispatcher的TcpListener失敗。錯誤原因：" & ex.Message)
            End Try
        End If
    End Sub

    ' 接收連入的Client
    Private Sub ListenProcess()
        Try
            While MyBase.Active
                Me.m_Dispatcher.Connect(MyBase.AcceptTcpClient())
            End While
        Catch ex As System.Net.Sockets.SocketException  ' TcpListener.Stop 所導致的例外
            '中止操作被 WSACancelBlockingCall 呼叫打斷。
        Catch ex As System.Exception        ' 其他的例外
            Debug.WriteLine("ListenProcess => " & ex.Message)
        End Try
    End Sub
End Class


