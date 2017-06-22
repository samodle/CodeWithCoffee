Public Class DemoData
    Public Function CreateDowntimeData(lineIndex As Integer, ByVal startTime As Date, ByVal endTime As Date) As DowntimeDataset
        Dim d = New DowntimeDataset(AllProdLines(lineIndex), getDowntimeDataset(lineIndex, startTime, endTime))
        Return d
    End Function


    Private Function getDowntimeDataset(ByVal lineIndex As Integer, ByVal startTime As Date, ByVal endTime As Date) As List(Of DowntimeEvent)
        Dim ret As New List(Of DowntimeEvent)

        Dim i As Integer = 0
        Dim x As Date = startTime, y As Date, ut As Integer, dt As Integer

        ut = GetRandom(200, 900)
        dt = GetRandom(55, 600)
        y = x.AddSeconds(dt)
        ret.Add(New DowntimeEvent(AllProdLines(lineIndex), x, y, dt, ut))

        While x < startTime And ret.Count < 3000
            ut = GetRandom(15000, 95000)
            dt = GetRandom(10000, 75000)

            x = y.AddSeconds(ut)
            y = x.AddSeconds(dt)
            ret.Add(New DowntimeEvent(AllProdLines(lineIndex), x, y, dt, ut))
        End While

        While x < startTime And ret.Count < 6000
            ut = GetRandom(6900, 65000)
            dt = GetRandom(5900, 55000)

            x = y.AddSeconds(ut)
            y = x.AddSeconds(dt)
            ret.Add(New DowntimeEvent(AllProdLines(lineIndex), x, y, dt, ut))
        End While

        While x < endTime And ret.Count < 9000
            ut = GetRandom(400, 900)
            dt = GetRandom(50, 600)

            x = y.AddSeconds(ut)
            y = x.AddSeconds(dt)
            ret.Add(New DowntimeEvent(AllProdLines(lineIndex), x, y, dt, ut))
        End While

        Return ret
    End Function
    ' 
    ' Private Generator As System.Random = New System.Random()
    Private Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        ' by making Generator static, we preserve the same instance '
        ' (i.e., do not create new instances with the same seed over and over) '
        ' between calls '
        Static Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function
End Class
