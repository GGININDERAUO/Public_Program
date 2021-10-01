Public Class CResponse

#Region "�ܼƫŧi"
    Protected m_Result As eResponseResult
    'Protected m_NgOrErrMessage As String
    Protected m_Param1 As String
    Protected m_Param2 As String
    Protected m_Param3 As String
    Protected m_Param4 As String
    Protected m_Param5 As String
    Protected m_Param6 As String
    Protected m_Param7 As String
    Protected m_Param8 As String
    Protected m_Param9 As String
#End Region

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

    Public ReadOnly Property Result() As eResponseResult
        Get
            Return Me.m_Result
        End Get
    End Property

    'Public ReadOnly Property NgOrErrMessage() As String
    '    Get
    '        Return Me.m_NgOrErrMessage
    '    End Get
    'End Property

    Public ReadOnly Property Param1() As String
        Get
            Return Me.m_Param1
        End Get
    End Property

    Public ReadOnly Property Param2() As String
        Get
            Return Me.m_Param2
        End Get
    End Property

    Public ReadOnly Property Param3() As String
        Get
            Return Me.m_Param3
        End Get
    End Property

    Public ReadOnly Property Param4() As String
        Get
            Return Me.m_Param4
        End Get
    End Property

    Public ReadOnly Property Param5() As String
        Get
            Return Me.m_Param5
        End Get
    End Property

    Public ReadOnly Property Param6() As String
        Get
            Return Me.m_Param6
        End Get
    End Property

    Public ReadOnly Property Param7() As String
        Get
            Return Me.m_Param7
        End Get
    End Property

    Public ReadOnly Property Param8() As String
        Get
            Return Me.m_Param8
        End Get
    End Property

    Public ReadOnly Property Param9() As String
        Get
            Return Me.m_Param9
        End Get
    End Property

    Public Function ResponseXml() As System.Xml.XmlDocument
        Dim XmlDoc As System.Xml.XmlDocument
        Dim XmlElement As System.Xml.XmlElement

        XmlDoc = New System.Xml.XmlDocument
        ' Declaration
        XmlDoc.AppendChild(XmlDoc.CreateXmlDeclaration("1.0", "UTF-8", String.Empty))
        ' Root Node
        XmlDoc.AppendChild(XmlDoc.CreateElement("Response"))
        ' Result
        XmlElement = XmlDoc.CreateElement("Result")
        XmlElement.InnerText = Me.m_Result.ToString
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' NgOrErrMessage
        'XmlElement = XmlDoc.CreateElement("NgOrErrMessage")
        'XmlElement.InnerText = Me.m_NgOrErrMessage
        'XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param1
        XmlElement = XmlDoc.CreateElement("Param1")
        XmlElement.InnerText = Me.m_Param1
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param2
        XmlElement = XmlDoc.CreateElement("Param2")
        XmlElement.InnerText = Me.m_Param2
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param3
        XmlElement = XmlDoc.CreateElement("Param3")
        XmlElement.InnerText = Me.m_Param3
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param4
        XmlElement = XmlDoc.CreateElement("Param4")
        XmlElement.InnerText = Me.m_Param4
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param5
        XmlElement = XmlDoc.CreateElement("Param5")
        XmlElement.InnerText = Me.m_Param5
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param6
        XmlElement = XmlDoc.CreateElement("Param6")
        XmlElement.InnerText = Me.m_Param6
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param7
        XmlElement = XmlDoc.CreateElement("Param7")
        XmlElement.InnerText = Me.m_Param7
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param8
        XmlElement = XmlDoc.CreateElement("Param8")
        XmlElement.InnerText = Me.m_Param8
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        ' Param9
        XmlElement = XmlDoc.CreateElement("Param9")
        XmlElement.InnerText = Me.m_Param9
        XmlDoc.DocumentElement.AppendChild(XmlElement)
        Return XmlDoc
    End Function
End Class
