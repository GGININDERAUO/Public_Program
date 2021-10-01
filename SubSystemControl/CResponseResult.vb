Public Class CResponseResult

#Region "�ܼƫŧi"
    Protected m_CommResult As eCommResult       ' Result Status
    'Protected m_NgOrErrMessage As String
    Protected m_ErrMessage As String
    Protected m_ErrorCode As String
    Protected m_Responses() As CResponse
#End Region

    Public ReadOnly Property CommResult() As eCommResult
        Get
            Return Me.m_CommResult
        End Get
    End Property

    'Public ReadOnly Property NgOrErrMessage() As String
    '    Get
    '        Return Me.m_NgOrErrMessage
    '    End Get
    'End Property

    Public ReadOnly Property ErrMessage() As String
        Get
            Return Me.m_ErrMessage
        End Get
    End Property

    Public ReadOnly Property ErrCode() As String
        Get
            Return Me.m_ErrorCode
        End Get
    End Property

    Public ReadOnly Property Responses() As CResponse()
        Get
            Return Me.m_Responses
        End Get
    End Property
End Class
