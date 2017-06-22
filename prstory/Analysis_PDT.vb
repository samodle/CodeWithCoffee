Public Class PDTeventReport

#Region "Properties & Variables"
    'raw data
    Private _sourceEvent As DTevent

    Private _RawMeanVal As Double
    Private _RawstdDev As Double
    Private _AdjMeanVal As Double
    Private _AdjStdDev As Double
    Private _MaxDTValue As Double
    Private _MinDTValue As Double

    Private _rawDowntimes As New list(Of Double)
    ' Private _firstUptimes As New list(Of Double)

    'beer me statistics (we'll save these for later ... )
    Private _SurvivalPcts As New list(Of Double)
    Private _AverageSurvival As Double

    'properties...
    Public ReadOnly Property Name As String
        Get
            Return _sourceevent.name
        End Get
    End Property
    Public ReadOnly Property Mean As Double
        Get
            Return _rawmeanval
        End Get
    End Property
    Public ReadOnly Property MaxDT As Double
        Get
            Return _maxdtvalue
        End Get
    End Property
    Public ReadOnly Property MinDT As Double
        Get
            Return _mindtvalue
        End Get
    End Property
    Public ReadOnly Property OneDevAbove As Double
        Get
            Return _RawMeanVal + _RawstdDev
        End Get
    End Property
		public readonly property OneDevBelow as double
        Get
            Return _RawMeanVal - _RawstdDev
        End Get
    End Property
	#End region

#Region "Construction"
    Public Sub New(tgtEvent As DTevent, SourceData As List(Of DowntimeEvent))
        Dim i As Integer
        _sourceEvent = tgtEvent

        _MaxDTValue = 0
        _MinDTValue = 10000
        'populate dt list
        For i = 0 To SourceData.Count - 2 'all but that last one
            If tgtEvent.Name.Equals(SourceData(i).Tier2) Then
                With SourceData(i)
                    _rawDowntimes.Add(.DT)
                    '   _firstuptimes.add(sourcedata(i + 1).ut)
                    If .DT > _MaxDTValue Then _MaxDTValue = .DT
                    If .DT < _MinDTValue Then _MinDTValue = .DT
                End With
            End If
        Next
        If tgtEvent.Equals(SourceData(SourceData.Count - 1)) Then _rawDowntimes.Add(SourceData(SourceData.Count - 1).DT) 'dont forget that last guy!
        'if _mindtvalue = 10000 then _mindtvalue = 0 'not sure if we need this? -sro
        'do some math
        getRawMeanAndStdDev()
    End Sub
#End Region

    Private Sub getRawMeanAndStdDev()
        Dim Squares As New List(Of Double)
        Dim filteredList As New List(Of Double)
        Dim SquareAvg As Double, tmpMeanDist As Double
        'get the raw average
        _rawmeanval = _rawdowntimes.Average()
        'use it to find dat std dev
        For Each value As Double In _rawdowntimes
            Squares.Add(Math.Pow(value - _rawmeanval, 2))
        Next
        SquareAvg = Squares.Average
        _rawstddev = Math.Sqrt(SquareAvg)

        'LUKE! We're gonna have company!
        Squares.Clear()
        For i = 0 To _rawdowntimes.Count - 1
            tmpMeanDist = (_rawdowntimes(i) - _rawmeanval) / _rawstddev
            If (tmpMeanDist > 3) Then 'Or (DailyStops(i) < 11 And tmpMeanDist > 3) Then
                'Uh, everythings under control. Situation normal.
            Else
                filteredList.Add(_rawdowntimes(i))
            End If
        Next
        'are you not entertained? (we're gonna find it again)
        _adjmeanval = filteredList.Average
        For Each value As Double In filteredList
            Squares.Add(Math.Pow(value - _adjmeanval, 2))
        Next
        SquareAvg = Squares.Average
        _adjstddev = Math.Sqrt(SquareAvg)
    End Sub


End Class

Public Module pdtHTMLexport

    Public Sub HTML_exportPDTcandlesticks(rawData As List(Of PDTeventReport), PDTindex As Integer)
        Dim fsT As Object
        Dim fileName As String

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        fileName = SERVER_FOLDER_PATH & "PDTanalysis.html"


        'lets write some HTML!
        fsT.writetext("<html>" & vbCrLf)
        fsT.writetext("<head>" & vbCrLf)

        fsT.writetext("<script type=" & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "https://www.google.com/jsapi" & Chr(34) & "></script>" & vbCrLf)
        fsT.writetext("<script type=" & Chr(34) & "text/javascript" & Chr(34) & ">" & vbCrLf)
        fsT.writetext("google.load(" & Chr(34) & "visualization" & Chr(34) & ", " & Chr(34) & "1" & Chr(34) & ", {packages:[" & Chr(34) & "corechart" & Chr(34) & "]});" & vbCrLf)
        fsT.writetext("google.setOnLoadCallback(drawChart);" & vbCrLf)
        fsT.writetext("function drawChart() {" & vbCrLf)
        fsT.writetext("var data = google.visualization.arrayToDataTable([" & vbCrLf)

        For i As Integer = 0 To rawData.count - 1
            With rawData(i)
                fsT.writetext("['" & .Name & "'," & Math.Round(.MinDT, 1) & "," & Math.Round(.OneDevBelow, 1) & "," & Math.Round(.OneDevAbove, 1) & "," & Math.Round(.MaxDT, 1) & "]," & vbCrLf)
            End With
        Next

        ' ['PDT 1', 20, 28, 38, 45],
        ' ['PDT 2', 31, 38, 55, 66],
        ' ['PDT 3', 50, 55, 77, 80],
        ' ['PDT 4', 77, 77, 66, 50],
        ' ['PDT 5', 68, 66, 22, 15]
        '// Treat first row as data as well.
        fsT.writetext("], true);" & vbCrLf)

        fsT.writetext("var options = {" & vbCrLf)
        fsT.writetext("legend:'none'," & vbCrLf)
        '//  colors:['#e0440e', '#e6693e', '#ec8f6e', '#f3b49f', '#f6c7b6']" & vbCrLf)
        fsT.writetext("candlestick: {" & vbCrLf)
        fsT.writetext("fallingColor: { strokeWidth: 0, fill: '#58ACFA' }," & vbCrLf) '// red
        fsT.writetext("risingColor: { strokeWidth: 0, fill: '#58ACFA' }" & vbCrLf)   '// green
        fsT.writetext("}" & vbCrLf)
        fsT.writetext("};" & vbCrLf)

        fsT.writetext("var chart = new google.visualization.CandlestickChart(document.getElementById('chart_div'));" & vbCrLf)

        fsT.writetext("chart.draw(data, options);" & vbCrLf)
        fsT.writetext("}" & vbCrLf)
        fsT.writetext("</script>" & vbCrLf)
        fsT.writetext("</head>" & vbCrLf)
        fsT.writetext("<body>" & vbCrLf)
        fsT.writetext("<div id=" & Chr(34) & "chart_div" & Chr(34) & " style=" & Chr(34) & "width: 100%; height: 100%;" & Chr(34) & "></div>" & vbCrLf)
        fsT.writetext("</body>" & vbCrLf)
        fsT.writetext("</html>" & vbCrLf)

        'fin
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing
    End Sub

End Module



