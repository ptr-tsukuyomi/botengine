Module Task
    ' 時間指定されたタスクと即時に実行するタスクを扱う

    Class IntervalTask
        Public begintime As DateTime
        Public act As Action
        Public interval As TimeSpan
        Public _nexttime As DateTime
        Public endtime As DateTime
    End Class

    Dim q As New Queue
    Dim tlist As New List(Of IntervalTask)

    Sub AddTask(ByVal act As Action)
        SyncLock q.SyncRoot
            q.Enqueue(act)
        End SyncLock
    End Sub

    Sub AddIntervalTask(ByVal ttask As IntervalTask)
        ttask._nexttime = ttask.begintime
        While Now > ttask._nexttime
            ttask._nexttime += ttask.interval
        End While
        tlist.Add(ttask)
    End Sub

    Sub DoTask()
        For Each tt In tlist
            If tt._nexttime.Year = Now.Year AndAlso
                tt._nexttime.Month = Now.Month AndAlso
                tt._nexttime.Day = Now.Day AndAlso
                Math.Truncate(tt._nexttime.TimeOfDay.TotalSeconds) = Math.Truncate(Now.TimeOfDay.TotalSeconds) Then
                AddTask(tt.act)
                tt._nexttime += tt.interval
            End If
        Next
        If q.Count <> 0 Then
            DirectCast(q.Dequeue(), Action)()
        End If
    End Sub

End Module
