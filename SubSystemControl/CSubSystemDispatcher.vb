Public Class CSubSystemDispatcher

#Region "變數宣告"
    Private m_SubSystemListener As CSubSystemListener
    Private m_Client As System.Net.Sockets.TcpClient         ' 接受的Client
    Private m_ReceiveThread As System.Threading.Thread       ' 接收資料的執行緒
    'Private m_RaiseEventThread As System.Threading.Thread
    Private m_Request As CRequest
    'Private m_Wait_Event As System.Threading.AutoResetEvent  '觸發RemoteControl Event
    'Private m_Exit As Boolean = False                        '停止RemoteControl非同步Thread
    Private m_RemoteControlThread As System.Threading.Thread    'RemoteControl Thread

    Private m_LogPath As String            ' 2008.03.29 add
    Private m_LogRecorder As CLogRecorder  ' 2008.03.29 add
#End Region

#Region "事件宣告"
    Public Event RemoteConnectComing()                          ' 遠端連入
    Public Event RemoteControl(ByVal Request As CRequest)       ' 接收遠端的命令控制
    Public Event RemoteDisconnect()                            ' 遠端產生TcpClient.Close
    Public Event ReceiveOccurError(ByVal ErrMessage As String)  ' 接收的Process發生例外錯誤, 必須將TcpClient中斷(Close)
#End Region


    Public Sub New()
        Me.New("")
    End Sub
    ' overloads
    Public Sub New(ByVal LogPath As String)
        Me.m_LogPath = LogPath
    End Sub

    ' 建立Listener
    Public Sub CreateListener(ByVal AddressOrHostName As String, ByVal Port As Integer)

        Try
            If Me.m_SubSystemListener Is Nothing Then
                Me.m_SubSystemListener = New CSubSystemListener(AddressOrHostName, Port, Me)
                ' 2008.03.29 add
                If Not Me.m_LogPath = "" Then
                    If Not System.IO.Directory.Exists(Me.m_LogPath) Then System.IO.Directory.CreateDirectory(Me.m_LogPath) ' 檢查System Log的檔案路徑是否存在，不存在創建一個新的
                    Me.m_LogRecorder = New CLogRecorder
                    Me.m_LogRecorder.Open(Me.m_LogPath, AddressOrHostName & "_" & Port.ToString, Now)
                    Me.m_LogRecorder.WriteLog(Now, ",---------- Start Log!! ----------")
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        'Else

        ' Throw New Exception("CreateListener只允許被執行一次！")
        ' End If
    End Sub

    ' 開始聆聽
    Public Sub StartListen()
        Try
            System.Threading.Thread.Sleep(100)
            If Not Me.m_SubSystemListener Is Nothing Then Me.m_SubSystemListener.StartListen()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' 停止聆聽
    Public Sub StopListen()
        Try
            'Me.m_Exit = True
            'Me.m_Wait_Event.Set()
            System.Threading.Thread.Sleep(100)
            If Me.m_Client IsNot Nothing Then
                If Me.m_Client.Connected Then
                    Me.m_Client.Close()
                End If
            End If
            If Not Me.m_SubSystemListener Is Nothing Then Me.m_SubSystemListener.StopListen()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' 關閉TcpClient
    Public Sub Disconnect()
        Try
            If Not Me.m_LogPath = "" Then
                Me.m_LogRecorder.WriteLog(Now, ",---------- Stop Log!! ----------")
                Me.m_LogRecorder.Close()
            End If
            If Me.m_Client IsNot Nothing AndAlso Me.m_Client.Connected Then Me.m_Client.Close()
        Catch ex As Exception
            Throw New Exception("關閉SubSystemDispatcher通訊傳送端失敗。錯誤原因：" & ex.Message)
        End Try
    End Sub

    ' 是否接受目前連入的Client
    Friend Sub Connect(ByVal RemoteClient As System.Net.Sockets.TcpClient)
        Try

            If Not Me.m_ReceiveThread Is Nothing Then
                Me.m_ReceiveThread = Nothing
            End If
            'Me.StopListen()             ' 關閉Listener
            If Not Me.m_Client Is Nothing AndAlso Me.m_Client.Connected Then
                Me.m_Client.Close()
            End If
            System.Threading.Thread.Sleep(500)
            Me.m_Client = RemoteClient  ' 接受
            ' 建立接收資料執行緒

            Me.m_ReceiveThread = New System.Threading.Thread(AddressOf ReceiveProcess)
            Me.m_ReceiveThread.Name = "SubSystemDispatcherReceiver"
            Me.m_ReceiveThread.SetApartmentState(Threading.ApartmentState.STA)
            Me.m_ReceiveThread.Start()
            RaiseEvent RemoteConnectComing()
        Catch ex As Exception
            Debug.WriteLine("Connect => " & ex.Message)
        End Try
    End Sub

    ' 接收資料
    Private Sub ReceiveProcess()
        Dim ReceiveBuffer(16777216) As Byte      ' 接收資料的Buffer, 2^24
        Dim ReceiveLength As Integer            ' 實際收到的資料長度
        Dim RestartListener As Boolean
        'Dim Request As CRequest

        Try
            RestartListener = True
            Try
                Do
                    ReceiveLength = Me.m_Client.GetStream.Read(ReceiveBuffer, 0, ReceiveBuffer.Length)  ' 同步式接收資料
                    Try
                        If ReceiveLength > 0 Then
                            'RaiseEvent RemoteControl(Me.ParseReceiveCommand(System.Text.Encoding.UTF8.GetString(ReceiveBuffer, 0, ReceiveLength)))  '引發事件
                            ' 引發非同步式RemoteControl事件執行緒
                            Me.m_Request = Me.ParseReceiveCommand(System.Text.Encoding.UTF8.GetString(ReceiveBuffer, 0, ReceiveLength))
                            'Me.m_Wait_Event.Set()
                            'Me.m_RaiseEventThread = New System.Threading.Thread(AddressOf DoEventProcess)
                            'Me.m_RaiseEventThread.Name = "SubSystemRemotControl"
                            'Me.m_RaiseEventThread.Start()
                            Me.m_remotecontrolThread = New System.Threading.Thread(AddressOf DoEventProcess)
                            Me.m_RemoteControlThread.Name = "SubSystemRemotControl"
                            Me.m_RemoteControlThread.SetApartmentState(Threading.ApartmentState.STA)
                            Me.m_RemoteControlThread.Start()
                        End If
                    Catch ex As Exception
                        ' 處理RemoteControl的例外
                        Me.SendCommand(Me.GetErrorResponse(ex.Message)) ' 回覆錯誤訊息
                    End Try
                Loop While ReceiveLength > 0
                If ReceiveLength = 0 Then

                    RaiseEvent RemoteDisconnect() ' 對方斷線
                End If

            Catch ex As System.IO.IOException   ' 本地斷線, 發生TcpClient.Close
                RestartListener = False         ' 不用重新啟動Listener
            Catch ex As Exception               ' 發生其他意外
                ' 處理SendCommand與RemoteDisconnect的例外
                Debug.WriteLine("Inner ReceiveProcess => " & ex.Message)
                If Me.m_Client.Connected Then Me.m_Client.Close()
                RaiseEvent ReceiveOccurError(ex.Message)    'Raise Event To Notify The Upper Form
            End Try
        Catch ex As Exception
            ' 處理Client.Close與ReceiveOccurError的例外
            Debug.WriteLine("Outer ReceiveProcess => " & ex.Message)
        Finally
            'Me.m_Client.Close()
            Me.m_Client = Nothing
            If RestartListener Then Me.StartListen() ' 重新啟動Listener
        End Try
    End Sub

    '引發RemoteControl事件
    Private Sub DoEventProcess()
        'While Not Me.m_Exit
        'Me.m_Wait_Event.WaitOne()
        'If Not Me.m_Exit Then RaiseEvent RemoteControl(Me.m_Request)
        'End While
        RaiseEvent RemoteControl(Me.m_Request)
    End Sub

    ' 解析REQUEST-XML資料
    Private Function ParseReceiveCommand(ByVal ReceiveCommand As String) As CRequest
        Dim XmlDoc As System.Xml.XmlDocument
        Dim SetRequest As CSetRequest

        Try
            SetRequest = New CSetRequest
            XmlDoc = New System.Xml.XmlDocument
            XmlDoc.LoadXml(ReceiveCommand)
            With SetRequest
                .SetCommand = XmlDoc.GetElementsByTagName("Command").Item(0).InnerText.ToUpper
                .SetParam1 = XmlDoc.GetElementsByTagName("Param1").Item(0).InnerText
                .SetParam2 = XmlDoc.GetElementsByTagName("Param2").Item(0).InnerText
                .SetParam3 = XmlDoc.GetElementsByTagName("Param3").Item(0).InnerText
                .SetParam4 = XmlDoc.GetElementsByTagName("Param4").Item(0).InnerText
                .SetParam5 = XmlDoc.GetElementsByTagName("Param5").Item(0).InnerText
                .SetParam6 = XmlDoc.GetElementsByTagName("Param6").Item(0).InnerText
                .SetParam7 = XmlDoc.GetElementsByTagName("Param7").Item(0).InnerText
                .SetParam8 = XmlDoc.GetElementsByTagName("Param8").Item(0).InnerText
                .SetParam9 = XmlDoc.GetElementsByTagName("Param9").Item(0).InnerText
            End With
            '2008.03.29 add
            If Not Me.m_LogPath = "" Then
                Me.m_LogRecorder.WriteLog(Now, ",[Receive = " & SetRequest.Command & "," & SetRequest.Param1 & "," & SetRequest.Param2 & "," & SetRequest.Param3 & "," & SetRequest.Param4 & "," & SetRequest.Param5 & "," & SetRequest.Param6 & "," & SetRequest.Param7 & "," & SetRequest.Param8 & "," & SetRequest.Param9 & "]")
            End If
            Return SetRequest
        Catch ex As Exception
            Debug.WriteLine("ParseReceiveCommand => " & ex.Message)
            Throw ex
        End Try
    End Function

    ' 產生ERROR RESPONSE-XML資料
    Private Function GetErrorResponse(ByVal ErrMessage As String) As System.Xml.XmlDocument
        Dim Response As CSetResponse

        Try
            Response = New CSetResponse
            With Response
                .SetResult = eResponseResult.ERR
                .SetParam1 = ErrMessage
            End With
            Return Response.ResponseXml()
        Catch ex As Exception
            Debug.WriteLine("GetErrorResponse => " & ex.Message)
            Throw ex
        End Try
    End Function

    ' 傳送資料
    Private Sub SendCommand(ByVal XmlDoc As System.Xml.XmlDocument)
        Dim SendCommands() As Byte

        Try
            SendCommands = System.Text.Encoding.UTF8.GetBytes(XmlDoc.OuterXml)
            Me.m_Client.GetStream.Write(SendCommands, 0, SendCommands.GetLength(0))
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' 回傳執行結果
    Public Sub ReturnResponse(ByVal Result As eResponseResult, Optional ByVal Param1 As String = "", Optional ByVal Param2 As String = "", Optional ByVal Param3 As String = "", Optional ByVal Param4 As String = "", Optional ByVal Param5 As String = "", Optional ByVal Param6 As String = "", Optional ByVal Param7 As String = "", Optional ByVal Param8 As String = "", Optional ByVal Param9 As String = "")
        Dim SetResponse As CSetResponse

        Try
            'If Not Me.m_Client.Connected Then Exit Sub ' 避免逾時斷線之後, 在送出執行結果時發生錯誤
            If Me.m_Client Is Nothing Then
                Exit Sub ' 避免逾時斷線之後, 在送出執行結果時發生錯誤
            Else
                If Not Me.m_Client.Connected Then Exit Sub ' 避免逾時斷線之後, 在送出執行結果時發生錯誤
            End If

            SetResponse = New CSetResponse
            With SetResponse
                .SetResult = Result
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
            '2008.03.29 add
            If Not Me.m_LogPath = "" Then
                Me.m_LogRecorder.WriteLog(Now, ",[Response = " & SetResponse.Result.ToString & "," & SetResponse.Param1 & "," & SetResponse.Param2 & "," & SetResponse.Param3 & "," & SetResponse.Param4 & "," & SetResponse.Param5 & "," & SetResponse.Param6 & "," & SetResponse.Param7 & "," & SetResponse.Param8 & "," & SetResponse.Param9 & "]")
            End If
            Me.SendCommand(SetResponse.ResponseXml())  ' 送出執行結果

        Catch ex As Exception
            Throw New Exception("傳送Response到Controller失敗。")
        End Try
    End Sub

End Class


