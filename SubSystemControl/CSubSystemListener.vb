Imports System.Net
Friend Class CSubSystemListener
    Inherits System.Net.Sockets.TcpListener

#Region "�ܼƫŧi"
    Private m_ListenThread As System.Threading.Thread   ' ���ݱ���Client�s�J�������
    Private m_Dispatcher As CSubSystemDispatcher
#End Region
    'Private temp As IPEndPoint
    Public Sub New(ByVal AddressOrHostName As String, ByVal Port As Integer, ByVal Dispatcher As CSubSystemDispatcher) 'ByVal OnConnect As ConnectDel)

        '�̷ӭӧOIP��Listening 20121212 MT 
        MyBase.New(IPAddress.Parse(AddressOrHostName), Port)

        Me.m_Dispatcher = Dispatcher
    End Sub

    ' �}�l��ť
    Public Sub StartListen()
        If Not MyBase.Active Then
            Try
                MyBase.Start()      ' Start Listener
            Catch ex As System.Exception
                Throw New Exception("�Ұ�SubSystemDispatcher��TcpListener���ѡC���~��]�G" & ex.Message)
            End Try

            ' �Ұ�Listen�������
            Me.m_ListenThread = New System.Threading.Thread(AddressOf ListenProcess)
            Me.m_ListenThread.Name = "SubSystemDispatcherListener"
            Me.m_ListenThread.SetApartmentState(Threading.ApartmentState.STA)
            Me.m_ListenThread.Start()
        End If
    End Sub

    ' �����ť
    Public Sub StopListen()
        If MyBase.Active Then
            Try
                MyBase.Stop()       ' Stop Listener
            Catch ex As System.Exception
                Debug.WriteLine("����SubSystemDispatcher��TcpListener���ѡC���~��]�G" & ex.Message)
            End Try
        End If
    End Sub

    ' �����s�J��Client
    Private Sub ListenProcess()
        Try
            While MyBase.Active
                Me.m_Dispatcher.Connect(MyBase.AcceptTcpClient())
            End While
        Catch ex As System.Net.Sockets.SocketException  ' TcpListener.Stop �ҾɭP���ҥ~
            '����ާ@�Q WSACancelBlockingCall �I�s���_�C
        Catch ex As System.Exception        ' ��L���ҥ~
            Debug.WriteLine("ListenProcess => " & ex.Message)
        End Try
    End Sub
End Class


