Module TimetableModule

    Structure BusInfo
        Dim Leaving As TimeSpan 'GetBusInfoの時に格納される
        Dim BusNumber As Integer
        Dim Destination As String
    End Structure

    Dim Timetable As New Dictionary(Of String, SortedDictionary(Of TimeSpan, BusInfo))

    Sub LoadTimetable(ByVal dbfilepath As String)
        Dim xdoc = XDocument.Load(dbfilepath)
        Dim root = xdoc.Element("BusDataSet")

        Timetable.Add("weekday", New SortedDictionary(Of TimeSpan, BusInfo))
        Timetable.Add("saturday", New SortedDictionary(Of TimeSpan, BusInfo))
        Timetable.Add("sunday", New SortedDictionary(Of TimeSpan, BusInfo))

        For Each tt In root.Elements("timetable")
            Dim dayofweek As String = tt.Attribute("days_of_the_week")
            For Each route In tt.Elements("route")
                Dim number As Integer = route.Attribute("number")
                For Each dest In route.Elements("destination")
                    Dim destination As String = dest.Attribute("name")

                    Dim bi As New BusInfo
                    bi.BusNumber = number
                    bi.Destination = destination
                    For Each t In dest.Elements("time")
                        Timetable(dayofweek).Add(TimeSpan.Parse(t.Value), bi)
                    Next
                Next
            Next
        Next

    End Sub

    Function GetBusInfo(ByVal dt As DateTime, ByVal max As Integer) As List(Of BusInfo)
        Dim buslist As New List(Of BusInfo)
        Dim listname As String
        If dt.DayOfWeek = DayOfWeek.Saturday Then
            listname = "saturday"
        ElseIf dt.DayOfWeek = DayOfWeek.Sunday Then
            listname = "sunday"
        Else
            listname = "weekday"
        End If

        For Each bus In Timetable(listname)
            If bus.Key >= dt.TimeOfDay Then ' AndAlso bus.Key <= (dt.TimeOfDay + New TimeSpan(0, margin, 0)) Then
                Dim b As BusInfo = bus.Value
                b.Leaving = bus.Key
                buslist.Add(b)
            End If
            If buslist.Count >= max Then
                Exit For
            End If
        Next

        Return buslist
    End Function

End Module
