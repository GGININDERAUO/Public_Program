Friend Class CSetRequest
    Inherits CRequest

    Public Sub New()
        Me.m_Command = ""
        'Me.m_PanelId = ""
        'Me.m_Pattern = ""
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

    Public WriteOnly Property SetCommand() As String
        Set(ByVal value As String)
            Me.m_Command = value
        End Set
    End Property

    'Public WriteOnly Property SetPanelId() As String
    '    Set(ByVal value As String)
    '        Me.m_PanelId = value
    '    End Set
    'End Property

    'Public WriteOnly Property SetPattern() As String
    '    Set(ByVal value As String)
    '        Me.m_Pattern = value
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