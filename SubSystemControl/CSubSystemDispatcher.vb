Public Class CSubSystemDispatcher

#Region "�ܼƫŧi"
    Private m_SubSystemListener As CSubSystemListener
    Private m_Client As System.Net.Sockets.TcpClient         ' ������Client
    Private m_ReceiveThread As System.Threading.Thread       ' ������ƪ������
    'Private m_RaiseEventThread As System.Threading.Thread
    Private m_Request As CRequest
    'Private m_Wait_Event As System.Threading.AutoResetEvent  'Ĳ�oRemoteControl Event
    'Private m_Exit As Boolean = False                        '����RemoteControl�D�P�BThread
    Private m_RemoteControlThread As System.Threading.Thread    'RemoteControl Thread

    Private m_LogPath As String            ' 2008.03.29 add
    Private m_LogRecorder As CLogRecorder  ' 2008.03.29 add
#End Region

#Region "�ƥ�ŧi"
    Public Event RemoteConnectComing()                          ' ���ݳs�J
    Public Event RemoteControl(ByVal Request As CRequest)       ' �������ݪ��R�O����
    Public Event RemoteDisconnect()                            ' ���ݲ���TcpClient.Close
    Public Event ReceiveOccurError(ByVal ErrMessage As String)  ' ������Process�o�ͨҥ~���~, �����NTcpClient���_(Close)
#End Region


    Public Sub New()
        Me.New("")
    End Sub
    ' overloads
    Public Sub New(ByVal LogPath As String)
        Me.m_LogPath = LogPath
    End Sub

    ' �إ�Listener
    Public Sub CreateListener(ByVal AddressOrHostName As String, ByVal Port As Integer)

        Try
            If Me.m_SubSystemListener Is Nothing Then
                Me.m_SubSystemListener = New CSubSystemListener(AddressOrHostName, Port, Me)
                ' 2008.03.29 add
                If Not Me.m_LogPath = "" Then
                    If Not System.IO.Directory.Exists(Me.m_LogPath) Then System.IO.Directory.CreateDirectory(Me.m_LogPath) ' �ˬdSystem Log���ɮ׸��|�O�_�s�b�A���s�b�Ыؤ@�ӷs��
                    Me.m_LogRecorder = New CLogRecorder
                    Me.m_LogRecorder.Open(Me.m_LogPath, AddressOrHostName & "_" & Port.ToString, Now)
                    Me.m_LogRecorder.WriteLog(Now, ",---------- Start Log!! ----------")
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        'Else

        ' Throw New Exception("CreateListener�u���\�Q����@���I")
        ' End If
    End Sub

    ' �}�l��ť
    Public Sub StartListen()
        Try
            System.Threading.Thread.Sleep(100)
            If Not Me.m_SubSystemListener Is Nothing Then Me.m_SubSystemListener.StartListen()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' �����ť
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

    ' ����TcpClient
    Public Sub Disconnect()
        Try
            If Not Me.m_LogPath = "" Then
                Me.m_LogRecorder.WriteLog(Now, ",---------- Stop Log!! ----------")
                Me.m_LogRecorder.Close()
            End If
            If Me.m_Client IsNot Nothing AndAlso Me.m_Client.Connected Then Me.m_Client.Close()
        Catch ex As Exception
            Throw New Exception("����SubSystemDispatcher�q�T�ǰe�ݥ��ѡC���~��]�G" & ex.Message)
        End Try
    End Sub

    ' �O�_�����ثe�s�J��Client
    Friend Sub Connect(ByVal RemoteClient As System.Net.Sockets.TcpClient)
        Try

            If Not Me.m_ReceiveThread Is Nothing Then
                Me.m_ReceiveThread = Nothing
            End If
            'Me.StopListen()             ' ����Listener
            If Not Me.m_Client Is Nothing AndAlso Me.m_Client.Connected Then
                Me.m_Client.Close()
            End If
            System.Threading.Thread.Sleep(500)
            Me.m_Client = RemoteClient  ' ����
            ' �إ߱�����ư����

            Me.m_ReceiveThread = New System.Threading.Thread(AddressOf ReceiveProcess)
            Me.m_ReceiveThread.Name = "SubSystemDispatcherReceiver"
            Me.m_ReceiveThread.SetApartmentState(Threading.ApartmentState.STA)
            Me.m_ReceiveThread.Start()
            RaiseEvent RemoteConnectComing()
        Catch ex As Exception
            Debug.WriteLine("Connect => " & ex.Message)
        End Try
    End Sub

    ' �������
    Private Sub ReceiveProcess()
        Dim ReceiveBuffer(16777216) As Byte      ' ������ƪ�Buffer, 2^24
        Dim ReceiveLength As Integer            ' ��ڦ��쪺��ƪ���
        Dim RestartListener As Boolean
        'Dim Request As CRequest

        Try
            RestartListener = True
            Try
                Do
                    ReceiveLength = Me.m_Client.GetStream.Read(ReceiveBuffer, 0, ReceiveBuffer.Length)  ' �P�B���������
                    Try
                        If ReceiveLength > 0 Then
                            'RaiseEvent RemoteControl(Me.ParseReceiveCommand(System.Text.Encoding.UTF8.GetString(ReceiveBuffer, 0, ReceiveLength)))  '�޵o�ƥ�
                            ' �޵o�D�P�B��RemoteControl�ƥ�����
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
                        ' �B�zRemoteControl���ҥ~
                        Me.SendCommand(Me.GetErrorResponse(ex.Message)) ' �^�п��~�T��
                    End Try
                Loop While ReceiveLength > 0
                If ReceiveLength = 0 Then

                    RaiseEvent RemoteDisconnect() ' ����_�u
                End If

            Catch ex As System.IO.IOException   ' ���a�_�u, �o��TcpClient.Close
                RestartListener = False         ' ���έ��s�Ұ�Listener
            Catch ex As Exception               ' �o�ͨ�L�N�~
                ' �B�zSendCommand�PRemoteDisconnect���ҥ~
                Debug.WriteLine("Inner ReceiveProcess => " & ex.Message)
                If Me.m_Client.Connected Then Me.m_Client.Close()
                RaiseEvent ReceiveOccurError(ex.Message)    'Raise Event To Notify The Upper Form
            End Try
        Catch ex As Exception
            ' �B�zClient.Close�PReceiveOccurError���ҥ~
            Debug.WriteLine("Outer ReceiveProcess => " & ex.Message)
        Finally
            'Me.m_Client.Close()
            Me.m_Client = Nothing
            If RestartListener Then Me.StartListen() ' ���s�Ұ�Listener
        End Try
    End Sub

    '�޵oRemoteControl�ƥ�
    Private Sub DoEventProcess()
        'While Not Me.m_Exit
        'Me.m_Wait_Event.WaitOne()
        'If Not Me.m_Exit Then RaiseEvent RemoteControl(Me.m_Request)
        'End While
        RaiseEvent RemoteControl(Me.m_Request)
    End Sub

    ' �ѪRREQUEST-XML���
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

    ' ����ERROR RESPONSE-XML���
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

    ' �ǰe���
    Private Sub SendCommand(ByVal XmlDoc As System.Xml.XmlDocument)
        Dim SendCommands() As Byte

        Try
            SendCommands = System.Text.Encoding.UTF8.GetBytes(XmlDoc.OuterXml)
            Me.m_Client.GetStream.Write(SendCommands, 0, SendCommands.GetLength(0))
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ' �^�ǰ��浲�G
    Public Sub ReturnResponse(ByVal Result As eResponseResult, Optional ByVal Param1 As String = "", Optional ByVal Param2 As String = "", Optional ByVal Param3 As String = "", Optional ByVal Param4 As String = "", Optional ByVal Param5 As String = "", Optional ByVal Param6 As String = "", Optional ByVal Param7 As String = "", Optional ByVal Param8 As String = "", Optional ByVal Param9 As String = "")
        Dim SetResponse As CSetResponse

        Try
            'If Not Me.m_Client.Connected Then Exit Sub ' �קK�O���_�u����, �b�e�X���浲�G�ɵo�Ϳ��~
            If Me.m_Client Is Nothing Then
                Exit Sub ' �קK�O���_�u����, �b�e�X���浲�G�ɵo�Ϳ��~
            Else
                If Not Me.m_Client.Connected Then Exit Sub ' �קK�O���_�u����, �b�e�X���浲�G�ɵo�Ϳ��~
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
            Me.SendCommand(SetResponse.ResponseXml())  ' �e�X���浲�G

        Catch ex As Exception
            Throw New Exception("�ǰeResponse��Controller���ѡC")
        End Try
    End Sub

End Class


