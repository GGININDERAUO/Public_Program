Friend Class CSetResponseResult
    Inherits CResponseResult

    Public Sub New()
        ReDim Me.m_Responses(0)
        Me.m_Responses(0) = Nothing
    End Sub

    Public WriteOnly Property SetCommResult() As eCommResult
        Set(ByVal value As eCommResult)
            Me.m_CommResult = value
        End Set
    End Property

    'Public WriteOnly Property SetNgOrErrMessage() As String
    '    Set(ByVal value As String)
    '        Me.m_NgOrErrMessage = value
    '    End Set
    'End Property

    Public WriteOnly Property SetErrMessage() As String
        Set(ByVal value As String)
            Me.m_ErrMessage = value
        End Set
    End Property

    Public WriteOnly Property SetErrCode() As String
        Set(ByVal value As String)
            Me.m_ErrorCode = value
        End Set
    End Property

    Public Sub AddResponse(ByVal Response As CResponse)
        If Me.m_Responses.GetLength(0) = 1 AndAlso Me.m_Responses(0) Is Nothing Then
            Me.m_Responses(0) = Response
        Else
            ReDim Preserve Me.m_Responses(Me.m_Responses.GetLength(0))
            Me.m_Responses(Me.m_Responses.GetLength(0) - 1) = Response
        End If
    End Sub
End Class
