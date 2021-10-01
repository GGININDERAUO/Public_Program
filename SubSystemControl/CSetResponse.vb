Friend Class CSetResponse
    Inherits CResponse

    Public Sub New()
        Me.m_Result = eResponseResult.OK
        'Me.m_NgOrErrMessage = ""
        Me.m_Param1 = ""
        Me.m_Param2 = ""
        Me.m_Param3 = ""
        Me.m_Param4 = ""
        Me.m_Param5 = ""
        Me.m_Param6 = ""
        Me.m_Param7 = ""
        Me.m_Param8 = ""
        Me.m_Param9 = ""
    End Sub

    Public WriteOnly Property SetResult() As eResponseResult
        Set(ByVal value As eResponseResult)
            Me.m_Result = value
        End Set
    End Property

    'Public WriteOnly Property SetNgOrErrMessage() As String
    '    Set(ByVal value As String)
    '        Me.m_NgOrErrMessage = value
    '    End Set
    'End Property

    Public WriteOnly Property SetParam1() As String
        Set(ByVal value As String)
            Me.m_Param1 = value
        End Set
    End Property

    Public WriteOnly Property SetParam2() As String
        Set(ByVal value As String)
            Me.m_Param2 = value
        End Set
    End Property

    Public WriteOnly Property SetParam3() As String
        Set(ByVal value As String)
            Me.m_Param3 = value
        End Set
    End Property

    Public WriteOnly Property SetParam4() As String
        Set(ByVal value As String)
            Me.m_Param4 = value
        End Set
    End Property

    Public WriteOnly Property SetParam5() As String
        Set(ByVal value As String)
            Me.m_Param5 = value
        End Set
    End Property

    Public WriteOnly Property SetParam6() As String
        Set(ByVal value As String)
            Me.m_Param6 = value
        End Set
    End Property

    Public WriteOnly Property SetParam7() As String
        Set(ByVal value As String)
            Me.m_Param7 = value
        End Set
    End Property

    Public WriteOnly Property SetParam8() As String
        Set(ByVal value As String)
            Me.m_Param8 = value
        End Set
    End Property

    Public WriteOnly Property SetParam9() As String
        Set(ByVal value As String)
            Me.m_Param9 = value
        End Set
    End Property
End Class