Friend Class CLogRecorder
#Region "�ܼƫŧi"
    Private m_FileStream As System.IO.FileStream
    Private m_Writer As System.IO.StreamWriter

    Private m_Filepath As String
    Private m_Filename As String
    Private m_LogDate As Date
#End Region

    Public Sub Open(ByVal filepath As String, ByVal prefixFilename As String, ByVal logDate As Date)
        Me.m_Filepath = filepath
        Me.m_Filename = prefixFilename
        Me.m_LogDate = logDate

        Me.m_FileStream = New System.IO.FileStream(filepath & "\" & prefixFilename & "_" & ConvertLogDate(logDate) & ".log", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.Read)
        Me.m_Writer = New System.IO.StreamWriter(Me.m_FileStream, System.Text.Encoding.Default)
        Me.m_Writer.AutoFlush = True
        Me.m_FileStream.Seek(0, IO.SeekOrigin.End)
    End Sub

    Private Function ConvertLogDate(ByVal logDate As Date)
        Return Format(logDate, "yyyyMMdd")
    End Function

    Public Sub Close()
        Me.m_FileStream.Close()
    End Sub

    Public Function ReadLog() As ArrayList
        Dim ArrList As ArrayList
        Dim Reader As System.IO.StreamReader

        ArrList = New ArrayList
        Me.m_FileStream.Seek(0, IO.SeekOrigin.Begin)
        Reader = New System.IO.StreamReader(Me.m_FileStream, System.Text.Encoding.Default)
        While Not Reader.EndOfStream
            ArrList.Add(Reader.ReadLine)
        End While
        Me.m_FileStream.Seek(0, IO.SeekOrigin.End)
        Return ArrList
    End Function

    Public Sub WriteLog(ByVal logDate As Date, ByVal log As String)
        Dim logDate1 As Date = Format(Me.m_LogDate, "yyyy/MM/dd")
        Dim logDate2 As Date = Format(logDate, "yyyy/MM/dd")

        Try
            If logDate1 <> logDate2 Then
                Me.Close()
                Me.Open(Me.m_Filepath, Me.m_Filename, logDate)
            End If
            SyncLock Me.m_Writer
                Me.m_Writer.WriteLine(Format(logDate, "yyyy/MM/dd HH:mm:ss ") & log)
            End SyncLock
        Catch ex As Exception
            Debug.WriteLine("Miss!")
        End Try

    End Sub

End Class
