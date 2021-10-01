Friend Class CControlClient

#Region "變數宣告"
    Private m_Id As Integer
    Private m_WaitResponseEvent As System.Threading.AutoResetEvent
    Private m_Client As System.Net.Sockets.TcpClient
    Private m_ReceiveThread As System.Threading.Thread
    Private m_ClientName As String
    Private m_ControlManager As CControlManager
    Private m_CanReceive As Boolean     ' True to allow to receive the remote response; otherwise, false
    Private m_Request As CRequest
    Private m_Response As CResponse

    Private m_LogPath As String            ' 2008.02.18 add
    Private m_LogRecorder As CLogRecorder  ' 2008.02.18 add
#End Region

#Region "存取變數"
    Public ReadOnly Property Id() As Integer
        Get
            Return Me.m_Id
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return Me.m_ClientName
        End Get
    End Property


    Public ReadOnly Property WaitResponseEvent() As System.Threading.AutoResetEvent
        Get
            Return Me.m_WaitResponseEvent
        End Get
    End Property

    Public WriteOnly Property CanReceive() As Boolean
        Set(ByVal value As Boolean)
            Me.m_CanReceive = value
        End Set
    End Property

    Public Property Request() As CRequest
        Get
            Return Me.m_Request
        End Get
        Set(ByVal value As CRequest)
            Me.m_Request = value
        End Set
    End Property

    Public Property Response() As CResponse
        Get
            Return Me.m_Response
        End Get
        Set(ByVal value As CResponse)
            Me.m_Response = value
        End Set
    End Property
#End Region

    Public Sub New(ByVal AddressOrHostName As String, ByVal Port As Integer, ByVal Id As Integer, ByVal ClientName As String, ByVal ControlManager As CControlManager)
        Me.New(AddressOrHostName, Port, Id, ClientName, ControlManager, "")
    End Sub
    ' overloads
    Public Sub New(ByVal AddressOrHostName As String, ByVal Port As Integer, ByVal Id As Integer, ByVal ClientName As String, ByVal ControlManager As CControlManager, ByVal LogPath As String)
        Try
            Me.m_Id = Id
            Me.m_ClientName = ClientName
            Me.m_LogPath = LogPath

            Me.m_Client = New System.Net.Sockets.TcpClient(AddressOrHostName, Port)
            Me.m_WaitResponseEvent = New System.Threading.AutoResetEvent(False)
            Me.m_ControlManager = ControlManager

            ' 建立接收資料執行緒
            Me.m_ReceiveThread = New System.Threading.Thread(AddressOf ReceiveProcess)
            Me.m_ReceiveThread.Name = ClientName
            Me.m_ReceiveThread.Start()

            ' 2008.02.18 add
            If Not Me.m_LogPath = "" Then
                If Not Me.m_LogRecorder Is Nothing Then
                    Me.m_LogRecorder.WriteLog(Now, ",---------- Stop Log!! ----------")
                    Me.m_LogRecorder.Close()
                End If
                If Not System.IO.Directory.Exists(Me.m_LogPath) Then System.IO.Directory.CreateDirectory(Me.m_LogPath) ' 檢查System Log的檔案路徑是否存在，不存在創建一個新的
                Me.m_LogRecorder = New CLogRecorder
                Me.m_LogRecorder.Open(Me.m_LogPath, Me.m_ClientName, Now)
                Me.m_LogRecorder.WriteLog(Now, ",---------- Start Log!! ----------")
            End If
        Catch ex As System.Net.Sockets.SocketException
            ' 無法連線，因為目標電腦拒絕連線。
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub Reset()
        Me.m_CanReceive = False
        Me.m_Request = Nothing
        Me.m_Response = Nothing
    End Sub

    Public Sub Close()
        If Not Me.m_LogPath = "" Then
            If Not Me.m_LogRecorder Is Nothing Then
                Me.m_LogRecorder.WriteLog(Now, ",---------- Stop Log!! ----------")
                Me.m_LogRecorder.Close()
            End If
        End If
        If Me.m_Client.Connected Then Me.m_Client.Close()
    End Sub

    '接收資料
    Private Sub ReceiveProcess()
        Dim ReceiveBuffer(4048576) As Byte         ' 接收資料的Buffer, 1K
        Dim ReceiveLength As Long            ' 實際收到的資料長度

        Try
            Try
                Do
                    ReceiveLength = Me.m_Client.GetStream.Read(ReceiveBuffer, 0, ReceiveBuffer.Length)  ' 同步式接收資料
                    Try
                        If ReceiveLength > 0 AndAlso Me.m_CanReceive Then
                            Me.m_Response = Me.ParseReceiveCommand(System.Text.Encoding.UTF8.GetString(ReceiveBuffer, 0, ReceiveLength))
                            Me.m_CanReceive = False
                            Me.m_WaitResponseEvent.Set()
                        End If
                    Catch ex As Exception
                        ' 處理ParseReceiveCommand的例外
                        Me.m_Response = Me.GetErrorResponse(ex.Message)
                        Me.m_CanReceive = False
                        Me.m_WaitResponseEvent.Set()
                    End Try
                Loop While ReceiveLength > 0
                If ReceiveLength = 0 Then Me.m_ControlManager.OnRemoteDisconnect() ' 對方斷線, Sub System的程式直接被關掉或是發生錯誤
            Catch ex As System.IO.IOException   ' 本地斷線, 發生TcpClient.Close
            Catch ex As Exception               ' 發生其他意外
                ' 處理GetErrorResponse與OnRemoteDisconnect的例外
                Debug.WriteLine("Inner ReceiveProcess => " & ex.Message)
                If Me.m_Client.Connected Then Me.m_Client.Close()
                Me.m_ControlManager.OnReceiveOccurError(ex.Message)
            End Try
        Catch ex As Exception
            ' 處理Client.Close與OnReceiveOccurError的例外
            Debug.WriteLine("Outer ReceiveProcess => " & ex.Message)
        Finally
            Me.m_ControlManager.RemoteClient(Me)
        End Try
    End Sub

    ' 產生ERROR RESPONSE
    Private Function GetErrorResponse(ByVal ErrMessage As String) As CResponse
        Dim SetResponse As CSetResponse

        Try
            SetResponse = New CSetResponse
            With SetResponse
                .SetResult = eResponseResult.ERR
                .SetParam1 = "接收Response時發生錯誤，錯誤原因：" & ErrMessage
            End With
            Return SetResponse
        Catch ex As Exception
            Debug.WriteLine("GetErrorResponse => " & ex.Message)
            Throw ex
        End Try
    End Function

    ' 傳送資料
    Public Sub SendCommand()
        Dim SendCommands() As Byte

        Try
            SendCommands = System.Text.Encoding.UTF8.GetBytes(Me.Request.RequestXml.OuterXml)
            Me.m_Client.GetStream.Write(SendCommands, 0, SendCommands.GetLength(0))
            If Not Me.m_LogPath = "" Then
                Me.m_LogRecorder.WriteLog(Now, ",[Request = " & Me.m_Request.Command & "," & Me.m_Request.Param1 & "," & Me.m_Request.Param2 & "," & Me.m_Request.Param3 & "," & Me.m_Request.Param4 & "," & Me.m_Request.Param5 & "," & Me.m_Request.Param6 & "," & Me.m_Request.Param7 & "," & Me.m_Request.Param8 & "," & Me.m_Request.Param9 & "]")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function ParseReceiveCommand(ByVal ReceiveCommand As String) As CResponse
        Dim XmlDoc As System.Xml.XmlDocument
        Dim SetResponse As CSetResponse
        Dim ResponseResult As String

        Try
            SetResponse = New CSetResponse
            XmlDoc = New System.Xml.XmlDocument
            XmlDoc.LoadXml(ReceiveCommand)
            With SetResponse
                ResponseResult = XmlDoc.GetElementsByTagName("Result").Item(0).InnerText
                Select Case ResponseResult
                    Case eResponseResult.OK.ToString
                        .SetResult = eResponseResult.OK
                        'Case eResponseResult.NG.ToString
                        '    .SetResult = eResponseResult.NG
                    Case eResponseResult.ERR.ToString
                        .SetResult = eResponseResult.ERR
                End Select
                '.SetNgOrErrMessage = XmlDoc.GetElementsByTagName("NgOrErrMessage").Item(0).InnerText
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
            ' 2008.02.18 add
            'Me.m_LogRecorder.WriteLog(Now, Now & ",[Request = " & Me.m_Request.Command & "," & Me.m_Request.Param1 & "," & Me.m_Request.Param2 & "," & Me.m_Request.Param3 & "," & Me.m_Request.Param4 & "," & Me.m_Request.Param5 & "," & Me.m_Request.Param6 & "," & Me.m_Request.Param7 & "," & Me.m_Request.Param8 & "," & Me.m_Request.Param9 & "] , [Response = " & SetResponse.Result.ToString & "," & SetResponse.Param1 & "," & SetResponse.Param2 & "," & SetResponse.Param3 & "," & SetResponse.Param4 & "," & SetResponse.Param5 & "," & SetResponse.Param6 & "," & SetResponse.Param7 & "," & SetResponse.Param8 & "," & SetResponse.Param9 & "]")
            If Not Me.m_LogPath = "" Then
                Me.m_LogRecorder.WriteLog(Now, ",[Response = " & SetResponse.Result.ToString & "," & SetResponse.Param1 & "," & SetResponse.Param2 & "," & SetResponse.Param3 & "," & SetResponse.Param4 & "," & SetResponse.Param5 & "," & SetResponse.Param6 & "," & SetResponse.Param7 & "," & SetResponse.Param8 & "," & SetResponse.Param9 & "]")
            End If
            Return SetResponse
        Catch ex As Exception
            Debug.WriteLine("ParseReceiveCommand => " & ex.Message)
            Throw ex
        End Try
    End Function
End Class
