Namespace prstoryAggFcns
    Public Class StandardDeviationFunction
        Inherits Telerik.Windows.Data.AggregateFunction(Of RateLossEvent, Double)
        Public Sub New()
            Me.AggregationExpression = Function(items) StdDev(items)
        End Sub

        Private Function StdDev(source As IEnumerable(Of RateLossEvent)) As Double
            Dim itemCount = source.Count()
            If itemCount > 1 Then
                Dim values = source.[Select](Function(i) i.RatePCT)
                Dim average = values.Average()
                Dim sum = values.Sum(Function(v) Math.Pow(v - average, 2))
                Return Math.Sqrt(sum / (itemCount - 1))
            End If

            Return 0
        End Function
    End Class

    Public Class CustomAverageFunction
        Inherits Telerik.Windows.Data.AggregateFunction(Of RateLossEvent, Double)
        Public Sub New()
            Me.AggregationExpression = Function(items) StdDev(items)
        End Sub

        Private Function StdDev(source As IEnumerable(Of RateLossEvent)) As Double
            Dim itemCount = source.Count()
            If itemCount > 1 Then
                Dim values = source.[Select](Function(i) i.RatePCT)
                Dim average = values.Average()
                Return Math.Round(average, 1)
            End If

            Return 0
        End Function
    End Class
End Namespace