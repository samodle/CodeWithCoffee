Module HTML_am
    Public Sub HTML_exportBubbleREDEP() 'parentLine As productionLine)
        Dim fsT As Object
        Dim fileName As String
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        fileName = SERVER_FOLDER_PATH & "redep" & ".html"

        fsT.writetext("<html>" & vbCrLf)
        fsT.writetext("<head>" & vbCrLf)

        fsT.writetext("<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & ">" & vbCrLf)
        fsT.writetext("<title>Dependency Analysis</title>" & vbCrLf)
        fsT.writetext("<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "style.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf)
        fsT.writetext("<script src=" & Chr(34) & "amcharts.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        fsT.writetext("<script src=" & Chr(34) & "xy.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        fsT.writetext("<script>" & vbCrLf)
        fsT.writetext("var chart;" & vbCrLf)
        fsT.writetext("var chartData = [" & vbCrLf)
        fsT.writetext("{" & vbCrLf)
        fsT.writetext("" & Chr(34) & "x" & Chr(34) & ": 1," & vbCrLf)
        fsT.writetext("" & Chr(34) & "y" & Chr(34) & ": 14," & vbCrLf)
        fsT.writetext("" & Chr(34) & "value" & Chr(34) & ": 59" & vbCrLf)
        fsT.writetext("}," & vbCrLf)
        fsT.writetext("" & Chr(34) & "x" & Chr(34) & ": 17," & vbCrLf)
        fsT.writetext("" & Chr(34) & "y" & Chr(34) & ": 6," & vbCrLf)
        fsT.writetext("" & Chr(34) & "value" & Chr(34) & ": 35" & vbCrLf)
        fsT.writetext("}" & vbCrLf)
        fsT.writetext("];" & vbCrLf)

        fsT.writetext("AmCharts.ready(function () {" & vbCrLf)
        '// XY Chart
        fsT.writetext("chart = new AmCharts.AmXYChart();" & vbCrLf)
        fsT.writetext("chart.dataProvider = chartData;" & vbCrLf)
        fsT.writetext("chart.startDuration = 1.5;" & vbCrLf)

        '// AXES
        '// X
        fsT.writetext("var xAxis = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.writetext("xAxis.labelFunction = formatValue;" & vbCrLf)
        fsT.writetext("xAxis.title = " & Chr(34) & "Post Dependency" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("xAxis.position = " & Chr(34) & "bottom" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("xAxis.autoGridCount = true;" & vbCrLf)
        fsT.writetext("chart.addValueAxis(xAxis);" & vbCrLf)

        '// Y
        fsT.writetext("var yAxis = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.writetext("yAxis.labelFunction = formatValue;" & vbCrLf)
        fsT.writetext("yAxis.title = " & Chr(34) & "Pre Dependency" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("yAxis.position = " & Chr(34) & "left" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("yAxis.autoGridCount = true;" & vbCrLf)
        fsT.writetext("chart.addValueAxis(yAxis);" & vbCrLf)

        ' // GRAPH
        fsT.writetext("var graph = new AmCharts.AmGraph();" & vbCrLf)
        fsT.writetext("graph.valueField = " & Chr(34) & "value" & Chr(34) & ";" & vbCrLf) '// valueField responsible for the size of a bullet
        fsT.writetext("graph.xField = " & Chr(34) & "x" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("graph.yField = " & Chr(34) & "y" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("graph.lineAlpha = 0;" & vbCrLf)
        fsT.writetext("graph.bullet = " & Chr(34) & "bubble" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("graph.balloonText = " & Chr(34) & "Pre SPD:<b>[[x]]</b> Post SPD:<b>[[y]]</b><br>Dependency:<b>[[value]]%</b>" & Chr(34) & vbCrLf)
        fsT.writetext("chart.addGraph(graph);" & vbCrLf)
        '// WRITE
        fsT.writetext("chart.write(" & Chr(34) & "chartdiv" & Chr(34) & ");" & vbCrLf)
        fsT.writetext("});" & vbCrLf)


        fsT.writetext("function formatValue(value, formattedValue, valueAxis){" & vbCrLf)
        fsT.writetext("if(value < 10){" & vbCrLf)
        fsT.writetext("return " & Chr(34) & "little" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("}" & vbCrLf)
        fsT.writetext("else if(value = 10){" & vbCrLf)
        fsT.writetext("return " & Chr(34) & "so-so" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("}" & vbCrLf)
        fsT.writetext("else if (value == 100){" & vbCrLf)
        fsT.writetext("return " & Chr(34) & "a lot!" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("}" & vbCrLf)
        fsT.writetext("else{" & vbCrLf)
        fsT.writetext("return " & Chr(34) & "" & Chr(34) & ";" & vbCrLf)
        fsT.writetext("}" & vbCrLf)
        fsT.writetext("};" & vbCrLf)
        fsT.writetext("</script>" & vbCrLf)
        fsT.writetext("</head>" & vbCrLf)

        fsT.writetext("<body>" & vbCrLf)
        fsT.writetext("<div id=" & Chr(34) & "chartdiv" & Chr(34) & " style=" & Chr(34) & "width: 100%; height: 400px;" & Chr(34) & "></div>" & vbCrLf)
        fsT.writetext("</body>" & vbCrLf)

        fsT.writetext("</html>" & vbCrLf)


        'fin
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing
    End Sub
End Module
