Public Class CControlManager

#Region "變數宣告"
    Private m_ClientManager As ArrayList
    Private m_WaitResponseEventManager As ArrayList
    Private m_Timeout As Integer        ' 送出REQUEST後, 等待多久後去接收回應
    Private m_IsTimeout As Boolean      ' REQUEST是否超過時間沒有收到回應
    Private m_WaitAllResponseEvent As System.Threading.AutoResetEvent   ' 等待所有的Client回傳結果
    Private m_WaitResponseEvents() As System.Threading.AutoResetEvent
    Private m_WaitResponseThread As System.Threading.Thread
#End Region

#Region "事件宣告"
    Public Event RemoteDisconnect()                             ' 遠端產生TcpClient.Close
    Public Event ReceiveOccurError(ByVal ErrMessage As String)  ' 接收的Process發生例外錯誤, 必須將所有的Clients中斷(Close)
#End Region

    Public Sub New()
        Me.m_ClientManager = New ArrayList
        Me.m_WaitResponseEventManager = New ArrayList
        Me.m_WaitAllResponseEvent = New System.Threading.AutoResetEvent(False)
    End Sub

    ' 建立連線
    Public Sub CreateClient(ByVal AddressOrHostName As String, ByVal Port As Integer, ByVal Id As Integer, ByVal ClientName As String)
        Me.CreateClient(AddressOrHostName, Port, Id, ClientName, "")
    End Sub

    ' 建立連線
    Public Sub CreateClient(ByVal AddressOrHostName As String, ByVal Port As Integer, ByVal Id As Integer, ByVal ClientName As String, ByVal LogPath As String)
        Dim ControlClient As CControlClient

        Try
            ControlClient = New CControlClient(AddressOrHostName, Port, Id, ClientName, Me, LogPath)
            Me.m_ClientManager.Add(ControlClient)
            Me.m_WaitResponseEventManager.Add(ControlClient.WaitResponseEvent)

        Catch ex As Exception
            Try
                Me.Disconnect()
            Catch ex2 As Exception
                Throw New Exception(ClientName & "關閉連線失敗，錯誤原因：" & ex2.Message)
            End Try
            Throw New Exception(ClientName & "連線失敗，錯誤原因：" & ex.Message)
        End Try
    End Sub

    Friend Sub RemoteClient(ByVal Client As CControlClient)
        SyncLock Me
            Me.m_ClientManager.Remove(Client)
        End SyncLock
    End Sub

    ' 將所有的Client都中斷連線
    Public Sub Disconnect()
        Dim i As Integer

        Try
            Me.m_WaitResponseEventManager.Clear()
            For i = Me.m_ClientManager.Count - 1 To 0 Step -1
                CType(Me.m_ClientManager.Item(i), CControlClient).Close()
            Next i
            Me.m_ClientManager.Clear()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub Reset()
        Dim i As Integer

        Try
            For i = 0 To Me.m_ClientManager.Count - 1
                CType(Me.m_ClientManager.Item(i), CControlClient).Reset()
            Next i
        Catch ex As Exception
            Throw New Exception("ClearRequest發生錯誤，錯誤原因：" & ex.Message)
        End Try
    End Sub

    Public Sub PrepareRequest(ByVal ClientId As Integer, ByVal Command As String, Optional ByVal Param1 As String = "", Optional ByVal Param2 As String = "", Optional ByVal Param3 As String = "", Optional ByVal Param4 As String = "", Optional ByVal Param5 As String = "", Optional ByVal Param6 As String = "", Optional ByVal Param7 As String = "", Optional ByVal Param8 As String = "", Optional ByVal Param9 As String = "")
        Dim i As Integer
        Dim ControlClient As CControlClient
        Dim Request As CSetRequest

        ' 是否有Client可以設定資料
        If Me.m_ClientManager.Count = 0 Then Throw New Exception("尚未建立連線")

        Try
            For i = 0 To Me.m_ClientManager.Count - 1
                ControlClient = CType(Me.m_ClientManager.Item(i), CControlClient)
                If ControlClient.Id = ClientId Then
                    Request = New CSetRequest
                    With Request
                        .SetCommand = Command
                        .SetParam1 = Param1
                        .SetParam2 = Param2
                        .SetParam3 = Param3
                        .SetParam4 = Param4
                        .SetParam5 = Param5
                        .SetParam6 = Param6
                        .SetParam7 = Param7
                        .SetParam8 = Param8
                        .SetParam9 = Param9
                    End With
                    ControlClient.Request = Request
                    Exit For
                End If
            Next i
        Catch ex As Exception
            Throw New Exception("PrepareRequest發生錯誤，錯誤原因：" & ex.Message)
        End Try
    End Sub

    Public Sub PrepareAllRequest(ByVal Command As String, Optional ByVal Param1 As String = "", Optional ByVal Param2 As String = "", Optional ByVal Param3 As String = "", Optional ByVal Param4 As String = "", Optional ByVal Param5 As String = "", Optional ByVal Param6 As String = "", Optional ByVal Param7 As String = "", Optional ByVal Param8 As String = "", Optional ByVal Param9 As String = "")
        Dim i As Integer
        Dim ControlClient As CControlClient
        Dim Request As CSetRequest

        ' 是否有Client可以設定資料
        If Me.m_ClientManager.Count = 0 Then Throw New Exception("尚未建立連線")

        Try
            For i = 0 To Me.m_ClientManager.Count - 1
                ControlClient = CType(Me.m_ClientManager.Item(i), CControlClient)
                Request = New CSetRequest
                With Request
                    .SetCommand = Command
                    .SetParam1 = Param1
                    .SetParam2 = Param2
                    .SetParam3 = Param3
                    .SetParam4 = Param4
                    .SetParam5 = Param5
                    .SetParam6 = Param6
                    .SetParam7 = Param7
                    .SetParam8 = Param8
                    .SetParam9 = Param9
                End With
                ControlClient.Request = Request
            Next i
        Catch ex As Exception
            Throw New Exception("PrePareAllRequest發生錯誤，錯誤原因：" & ex.Message)
        End Try
    End Sub

    ' 傳送命令
    Public Function SendRequest(ByVal Timeout As Integer) As CResponseResult
        Dim i As Integer
        Dim ControlClient As CControlClient
        Dim SetResponseResult As CSetResponseResult
        Dim Response As CResponse

        Me.m_Timeout = Timeout
        Me.m_WaitAllResponseEvent.Reset()

        ' 是否有Client可以送資料
        If Me.m_ClientManager.Count = 0 Then Throw New Exception("尚未建立連線")
        ' 判斷是否有設定要送出的資料
        For i = 0 To Me.m_ClientManager.Count - 1
            If CType(Me.m_ClientManager.Item(i), CControlClient).Request IsNot Nothing Then Exit For
            If i = Me.m_ClientManager.Count - 1 Then Throw New Exception("沒有設定Request")
        Next i

        Try
            ' 產生WaitHandle Array
            ReDim m_WaitResponseEvents(Me.m_WaitResponseEventManager.Count - 1)
            For i = 0 To Me.m_WaitResponseEventManager.Count - 1
                Me.m_WaitResponseEvents(i) = CType(Me.m_WaitResponseEventManager.Item(i), System.Threading.AutoResetEvent)
                Me.m_WaitResponseEvents(i).Reset()
            Next i

            ' 送出資料
            For i = 0 To Me.m_ClientManager.Count - 1
                ControlClient = CType(Me.m_ClientManager.Item(i), CControlClient)
                If ControlClient.Request IsNot Nothing Then
                    ControlClient.CanReceive = True
                    ControlClient.SendCommand()
                Else
                    ControlClient.WaitResponseEvent.Set()
                End If
            Next i

            ' 等待接收資料
            Me.m_WaitResponseThread = New System.Threading.Thread(AddressOf Me.WaitResponseProcess)
            Me.m_WaitResponseThread.Name = "WaitResponseProcess"
            Me.m_WaitResponseThread.IsBackground = True
            Me.m_WaitResponseThread.Start()
            Me.m_WaitAllResponseEvent.WaitOne()

            SetResponseResult = New CSetResponseResult

            If Me.m_IsTimeout Then
                SetResponseResult.SetCommResult = eCommResult.TIMEOUT
                SetResponseResult.SetErrMessage = "Timeout"
            Else
                ' 傳回RESPONSE資料
                SetResponseResult.SetCommResult = eCommResult.OK
                For i = 0 To Me.m_ClientManager.Count - 1
                    ControlClient = CType(Me.m_ClientManager.Item(i), CControlClient)
                    Response = ControlClient.Response
                    If Response IsNot Nothing Then
                        If Response.Result = eResponseResult.ERR Then
                            SetResponseResult.SetCommResult = eCommResult.ERR
                            SetResponseResult.SetErrMessage = Response.Param1
                            SetResponseResult.SetErrCode = Response.Param2
                            SetResponseResult.AddResponse(Response)
                        ElseIf Response.Result = eResponseResult.OK Then
                            SetResponseResult.AddResponse(Response)
                        Else
                            SetResponseResult.SetCommResult = eCommResult.ERR
                            SetResponseResult.SetErrMessage = "Error response format"
                        End If
                    End If
                Next i
            End If
            Return SetResponseResult
        Catch ex As Exception
            Throw New Exception("SendRequest發生錯誤，錯誤原因：" & ex.Message)
        End Try
    End Function

    ' 接收資料, 等全部的Client都收到回應之後, 再引發對應的事件
    Private Sub WaitResponseProcess()
        Me.m_IsTimeout = False  ' Reset
        Me.m_IsTimeout = Not System.Threading.WaitHandle.WaitAll(Me.m_WaitResponseEvents, Me.m_Timeout, False)
        Me.m_WaitAllResponseEvent.Set()
    End Sub

    Friend Sub OnRemoteDisconnect()
        SyncLock Me
            Me.Disconnect()
            RaiseEvent RemoteDisconnect()
        End SyncLock
    End Sub

    Friend Sub OnReceiveOccurError(ByVal ErrMessage As String)
        SyncLock Me
            RaiseEvent ReceiveOccurError(ErrMessage)
        End SyncLock
    End Sub
End Class