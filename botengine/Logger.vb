Class Logger

    Dim strm As IO.StreamWriter = Nothing

    Sub New(ByVal filename As String)
        strm = New IO.StreamWriter(filename, True)
    End Sub

    Sub WriteLog(ByVal str As String)
        If strm IsNot Nothing Then
            strm.WriteLine(Now + " " + str)
            strm.Flush()
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If strm IsNot Nothing Then
            strm.Close()
        End If
        MyBase.Finalize()
    End Sub
End Class