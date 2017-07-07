Imports System.Globalization
Module HTML_Multiline

    Public Sub CreateHTMLMultiLine_Summary(ByVal paramObj As Object, Optional IsTeamAnalysis As Boolean = False)

        Dim fsT As Object
        Dim fileName As String
        Dim us As New CultureInfo("en-US")

        Dim PRList As List(Of Double) = paramObj(0)
        Dim SPDList As List(Of Double) = paramObj(1)
        Dim MTBFList As List(Of Double) = paramObj(2)
        Dim UPDTList As List(Of Double) = paramObj(3)
        Dim PDTList As List(Of Double) = paramObj(4)
        Dim LineNamesList As List(Of String) = paramObj(5)
        Dim DatesList_starttime As List(Of String) = paramObj(6)
        Dim DatesList_endtime As List(Of String) = paramObj(7)
        Dim isalldatessame As Boolean = paramObj(8)
        Dim MSUList As List(Of Double) = paramObj(9)
        Dim CasesList As List(Of Double) = paramObj(10)
        Dim RateLossList As List(Of Double) = paramObj(11)
        Dim AdjustedUnitsList As List(Of Double) = paramObj(12)
        Dim MTTRList As List(Of Double) = paramObj(13)


        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        fileName = SERVER_FOLDER_PATH & "MultilineSummary.html"

        fsT.WriteText("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-'//W3C'//DTD HTML 4.01'//EN" & Chr(34) & " " & Chr(34) & "http:'//www.w3.org/TR/html4/strict.dtd" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<html>" & vbCrLf)
        fsT.WriteText("<head>" & vbCrLf)
        fsT.WriteText("<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "style.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "amcharts.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "serial.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        fsT.WriteText("<script>" & vbCrLf)

        fsT.WriteText("var chart;" & vbCrLf)


        fsT.WriteText("var titlesofchart = [{" & vbCrLf)
        If IsTeamAnalysis = False Then
            fsT.Writetext("'text': 'Overall line results'" & vbCrLf)
        Else
            fsT.Writetext("'text': 'Teams results side by side'" & vbCrLf)
        End If
        fsT.Writetext("}, {" & vbCrLf)
        fsT.Writetext("'text': ' ' ," & vbCrLf)
        fsT.Writetext(" 'bold': false" & vbCrLf)
        fsT.Writetext("}];" & vbCrLf)


        fsT.WriteText("var chartData = [" & vbCrLf)
        'THIS IS WHERE THE RAW DATA GOES
        '          {
        '      "fault": "a1",
        '      "dt": 23.5,
        '      "stops": 18.1
        '  },

        For eventIncrementer As Integer = 0 To PRList.Count - 1
            fsT.WriteText("{" & Chr(34) & "PR" & Chr(34) & ": " & Math.Round(100 * PRList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "stops per day" & Chr(34) & ": " & Math.Round(SPDList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            If isalldatessame = True Then
                fsT.WriteText(Chr(34) & "linename" & Chr(34) & ": " & Chr(34) & LineNamesList(eventIncrementer) & Chr(34) & "," & vbCrLf)
            Else
                fsT.WriteText(Chr(34) & "linename" & Chr(34) & ": " & Chr(34) & LineNamesList(eventIncrementer) & " \n " & DatesList_starttime(eventIncrementer) & " \n " & DatesList_endtime(eventIncrementer) & Chr(34) & "," & vbCrLf)
            End If
            fsT.WriteText(Chr(34) & "updt" & Chr(34) & ": " & Math.Round(100 * UPDTList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "pdt" & Chr(34) & ": " & Math.Round(100 * PDTList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                fsT.WriteText(Chr(34) & "msu" & Chr(34) & ": " & Math.Round(MSUList(eventIncrementer) / 1000, 2).ToString("######0.00", us) & "," & vbCrLf)
            Else
                fsT.WriteText(Chr(34) & "msu" & Chr(34) & ": " & Math.Round(MSUList(eventIncrementer)).ToString("######0.0", us) & "," & vbCrLf)
            End If
            fsT.WriteText(Chr(34) & "cases" & Chr(34) & ": " & Math.Round(CasesList(eventIncrementer)).ToString("######0.0", us) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "rateloss" & Chr(34) & ": " & Math.Round(100 * RateLossList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)

            ' If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            '     fsT.WriteText(Chr(34) & "adjustedunits" & Chr(34) & ": " & Math.Round(adjustedunitsList(eventIncrementer) / 1000, 2).ToString("######0.00", us) & "," & vbCrLf)
            ' Else
            fsT.WriteText(Chr(34) & "adjustedunits" & Chr(34) & ": " & Math.Round(AdjustedUnitsList(eventIncrementer)).ToString("######0.0", us) & "," & vbCrLf)
            '  End If
            fsT.WriteText(Chr(34) & "mttr" & Chr(34) & ": " & Math.Round(MTTRList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "mtbf" & Chr(34) & ": " & Math.Round(MTBFList(eventIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            If eventIncrementer = PRList.Count - 1 Then
                fsT.WriteText("}" & vbCrLf)
            Else
                fsT.WriteText("}," & vbCrLf)
            End If
        Next
        '||||||||||||||||||||||||||||||||

        fsT.WriteText("];" & vbCrLf)


        fsT.WriteText("AmCharts.ready(function () {" & vbCrLf)
        '// SERIAL CHART
        fsT.WriteText("chart = new AmCharts.AmSerialChart()" & vbCrLf)

        fsT.WriteText("chart.dataProvider = chartData; " & vbCrLf)
        fsT.WriteText("chart.categoryField = " & Chr(34) & "linename" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.startDuration = 1; " & vbCrLf)
        fsT.WriteText("chart.titles = titlesofchart; " & vbCrLf)
        fsT.WriteText("chart.export; " & vbCrLf)


        '// AXES
        '// category
        fsT.WriteText("var categoryAxis = chart.categoryAxis; " & vbCrLf)
        fsT.WriteText("categoryAxis.gridPosition = " & Chr(34) & "start" & Chr(34) & "; " & vbCrLf)
        If PRList.Count - 1 > 6 Then
            fsT.WriteText("categoryAxis.labelRotation = 30;" & vbCrLf)
        Else
            fsT.WriteText("categoryAxis.labelRotation = 0;" & vbCrLf)
        End If


        '// value  PR
        fsT.WriteText("var valueAxis = new AmCharts.ValueAxis(); " & vbCrLf)
        fsT.WriteText("valueAxis.axisColor = " & Chr(34) & "#2C99C3" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("valueAxis.gridAlpha = 0; " & vbCrLf)
        fsT.WriteText("valueAxis.axisThickness = 2; " & vbCrLf)
        fsT.WriteText("valueAxis.minimum = 0; " & vbCrLf)
        fsT.WriteText("chart.addValueAxis(valueAxis); " & vbCrLf)

        '// second value axis (on the right) spd
        fsT.WriteText("var valueAxis2 = new AmCharts.ValueAxis(); " & vbCrLf)
        fsT.WriteText("valueAxis2.position = " & Chr(34) & "right" & Chr(34) & "; " & vbCrLf) '// this line makes the axis to appear on the right
        fsT.WriteText("valueAxis2.axisColor = " & Chr(34) & "#fcd202" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("valueAxis2.gridAlpha = 0; " & vbCrLf)
        fsT.WriteText("valueAxis.unit = " & Chr(34) & "%" & Chr(34) & ";" & vbCrLf)
        fsT.WriteText("valueAxis2.axisThickness = 2; " & vbCrLf)
        fsT.WriteText("valueAxis2.minimum = 0; " & vbCrLf)
        fsT.WriteText("chart.addValueAxis(valueAxis2); " & vbCrLf)

        '// third value axis (on the right) mtbf
        fsT.WriteText("var valueAxis3 = new AmCharts.ValueAxis(); " & vbCrLf)
        fsT.WriteText("valueAxis3.position = " & Chr(34) & "right" & Chr(34) & "; " & vbCrLf) '// this line makes the axis to appear on the right
        fsT.writetext("valueAxis3.offset = 30; " & vbCrLf)
        fsT.WriteText("valueAxis3.axisColor = " & Chr(34) & "#FF8C00" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("valueAxis3.gridAlpha = 0; " & vbCrLf)
        fsT.WriteText("valueAxis3.axisThickness = 2; " & vbCrLf)
        fsT.WriteText("valueAxis3.minimum = 0; " & vbCrLf)
        fsT.WriteText("chart.addValueAxis(valueAxis3); " & vbCrLf)

        '// fourth value axis (on the right) mtbf
        fsT.WriteText("var valueAxis4 = new AmCharts.ValueAxis(); " & vbCrLf)
        fsT.WriteText("valueAxis4.position = " & Chr(34) & "right" & Chr(34) & "; " & vbCrLf) '// this line makes the axis to appear on the right
        fsT.writetext("valueAxis4.offset = 30; " & vbCrLf)
        fsT.WriteText("valueAxis4.axisColor = " & Chr(34) & "#CD5C5C" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("valueAxis4.gridAlpha = 0; " & vbCrLf)
        fsT.WriteText("valueAxis4.axisThickness = 2; " & vbCrLf)
        fsT.WriteText("valueAxis4.minimum = 0; " & vbCrLf)
        fsT.WriteText("chart.addValueAxis(valueAxis4); " & vbCrLf)

        '// GRAPHS
        '// column graph PR
        fsT.WriteText("var graph1 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph1.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            fsT.WriteText(" graph1.title = " & Chr(34) & "PR" & Chr(34) & "; " & vbCrLf)
        Else
            fsT.WriteText(" graph1.title = " & Chr(34) & "Availability" & Chr(34) & "; " & vbCrLf)
        End If

        fsT.WriteText("graph1.lineColor = " & Chr(34) & "#33cccc" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.valueField = " & Chr(34) & "PR" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.lineAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph1.fillAlphas = 1; " & vbCrLf)
        fsT.WriteText("graph1.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.labelText = " & Chr(34) & "[[value]]%" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b>[[title]] : </b> [[PR]] % </span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph1); " & vbCrLf)

        '// line 2nd spd
        fsT.WriteText("var graph2 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph2.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.title = " & Chr(34) & "Stops per day" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.lineColor = " & Chr(34) & "#ff6699" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.valueField = " & Chr(34) & "stops per day" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.lineThickness = 0; " & vbCrLf)
        fsT.WriteText("graph2.lineAlpha = 0; " & vbCrLf)
        fsT.WriteText("graph2.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.bulletBorderThickness = 3; " & vbCrLf)
        fsT.WriteText("graph2.bulletBorderColor = " & Chr(34) & "#fcd202" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.bulletBorderAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph2.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.valueAxis = valueAxis2; " & vbCrLf)
        fsT.WriteText("graph2.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b>[[value]] stops per day </b></span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph2);" & vbCrLf)

        ' line 3rd mtbf
        fsT.WriteText("var graph3 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph3.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.title = " & Chr(34) & "MTBF (min)" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.lineColor = " & Chr(34) & "#33ccff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.valueField = " & Chr(34) & "mtbf" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.lineThickness = 0; " & vbCrLf)
        fsT.WriteText("graph3.lineAlpha = 0; " & vbCrLf)
        fsT.WriteText("graph3.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.bulletBorderThickness = 3; " & vbCrLf)
        fsT.WriteText("graph3.bulletBorderColor = " & Chr(34) & "#FF8C00" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.bulletBorderAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph3.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.valueAxis = valueAxis3; " & vbCrLf)
        fsT.WriteText("graph3.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.balloonText = " & Chr(34) & "<span style='font-size:13px;'>[[title]]:<b>[[value]]</b></span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph3);" & vbCrLf)

        ' line 3rd MTTR
        fsT.WriteText("var graph13 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph13.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.title = " & Chr(34) & "MTTR (min)" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.lineColor = " & Chr(34) & "#cc66ff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.valueField = " & Chr(34) & "mttr" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.lineThickness = 0; " & vbCrLf)
        fsT.WriteText("graph13.lineAlpha = 0; " & vbCrLf)
        fsT.WriteText("graph13.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.bulletBorderThickness = 3; " & vbCrLf)
        fsT.WriteText("graph13.bulletBorderColor = " & Chr(34) & "#FF8C00" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.bulletBorderAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph13.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.valueAxis = valueAxis3; " & vbCrLf)
        fsT.WriteText("graph13.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph13.balloonText = " & Chr(34) & "<span style='font-size:13px;'>[[title]]:<b>[[value]]</b></span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph13);" & vbCrLf)
          fsT.WriteText("chart.hideGraph(graph13);" & vbCrLf)


        '// column graph UPDT
        fsT.WriteText("var graph4 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph4.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph4.title = " & Chr(34) & "UPDT%" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph4.lineColor = " & Chr(34) & "#ff6666" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph4.valueField = " & Chr(34) & "updt" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph4.lineAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph4.fillAlphas = 1; " & vbCrLf)
        fsT.WriteText("graph4.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph4.labelText = " & Chr(34) & "[[value]]%" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph4.valueAxis = valueAxis; " & vbCrLf)
        fsT.WriteText("graph4.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b> Unplanned DT: </b> [[updt]] %  </span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph4); " & vbCrLf)

        '// column graph PDT
        fsT.WriteText("var graph5 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph5.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText(" graph5.title = " & Chr(34) & "PDT%" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph5.lineColor = " & Chr(34) & "#ffcc66" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph5.valueField = " & Chr(34) & "pdt" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph5.lineAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph5.fillAlphas = 1; " & vbCrLf)
        fsT.WriteText("graph5.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph5.labelText = " & Chr(34) & "[[value]]%" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph5.valueAxis = valueAxis; " & vbCrLf)
        fsT.WriteText("graph5.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b> Planned DT: </b> [[pdt]] %  </span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph5); " & vbCrLf)

        '// column graph MSU
        fsT.WriteText("var graph6 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph6.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            fsT.WriteText(" graph6.title = " & Chr(34) & "Production (MSU)" & Chr(34) & "; " & vbCrLf)
        Else
            fsT.WriteText(" graph6.title = " & Chr(34) & "Sched Time (min)" & Chr(34) & "; " & vbCrLf)
        End If

        fsT.WriteText("graph6.lineColor = " & Chr(34) & "#b3d9ff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph6.valueField = " & Chr(34) & "msu" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph6.lineAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph6.fillAlphas = 1; " & vbCrLf)
        fsT.WriteText("graph6.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph6.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph6.valueAxis = valueAxis4; " & vbCrLf)
        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            fsT.WriteText("graph6.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b> Production: </b> [[msu]] msu  </span>" & Chr(34) & "; " & vbCrLf)
        Else
            fsT.WriteText("graph6.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b> Sched. Time: </b> [[msu]] min  </span>" & Chr(34) & "; " & vbCrLf)
        End If
        fsT.WriteText("chart.addGraph(graph6); " & vbCrLf)

        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            '// column graph Cases
            fsT.WriteText("var graph7 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph7.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText(" graph7.title = " & Chr(34) & "Production (actual cases)" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph7.lineColor = " & Chr(34) & "#0000ff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph7.valueField = " & Chr(34) & "cases" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph7.lineAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph7.fillAlphas = 1; " & vbCrLf)
            fsT.WriteText("graph7.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph7.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph7.valueAxis = valueAxis4; " & vbCrLf)
            fsT.WriteText("graph7.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b> Production: </b> [[cases]] cases  </span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph7); " & vbCrLf)
            If IsTeamAnalysis = False Then fsT.WriteText("chart.hideGraph(graph7);" & vbCrLf)

            fsT.WriteText("var graph8 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph8.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText(" graph8.title = " & Chr(34) & "Rate & Quality Losses" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph8.lineColor = " & Chr(34) & "#ffb3b3" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph8.valueField = " & Chr(34) & "rateloss" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph8.lineAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph8.fillAlphas = 1; " & vbCrLf)
            fsT.WriteText("graph8.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph8.labelText = " & Chr(34) & "[[value]]%" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph8.valueAxis = valueAxis; " & vbCrLf)
            fsT.WriteText("graph8.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b> Rate/Scrap Loss: </b> [[rateloss]] %  </span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph8); " & vbCrLf)
            fsT.WriteText("chart.hideGraph(graph8);" & vbCrLf)

            fsT.WriteText("var graph9 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph9.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText(" graph9.title = " & Chr(34) & "Adjusted Units" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph9.lineColor = " & Chr(34) & "#a556f0" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph9.valueField = " & Chr(34) & "adjustedunits" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph9.lineAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph9.fillAlphas = 1; " & vbCrLf)
            fsT.WriteText("graph9.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph9.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph9.valueAxis = valueAxis4; " & vbCrLf)
            fsT.WriteText("graph9.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b> Adjusted Units: </b> [[adjustedunits]]  </span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph9); " & vbCrLf)
            fsT.WriteText("chart.hideGraph(graph9);" & vbCrLf)


        End If

        '// LEGEND
        fsT.WriteText("var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.WriteText("legend.useGraphSettings = true;" & vbCrLf)
        fsT.WriteText("chart.addLegend(legend);" & vbCrLf)
        fsT.WriteText("chart.hideGraph(graph3);" & vbCrLf)
        fsT.WriteText("chart.hideGraph(graph6);" & vbCrLf)



        '// WRITE
        fsT.WriteText("chart.write(" & Chr(34) & "chartdiv" & Chr(34) & ");" & vbCrLf)
        fsT.WriteText("});" & vbCrLf)
        fsT.WriteText("</script>" & vbCrLf)
        fsT.WriteText("</head>" & vbCrLf)

        fsT.WriteText("<body>" & vbCrLf)
        fsT.WriteText("<div id=" & Chr(34) & "chartdiv" & Chr(34) & " style=" & Chr(34) & "width:100%; height:430px;" & Chr(34) & "></div>" & vbCrLf)








        'wrap it up
        fsT.WriteText("</body>" & vbCrLf)
        fsT.WriteText("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception

        End Try
        fsT = Nothing
    End Sub
    Public Sub CreateHTMLMultiLine_Bylossarea(ByVal paramObj As Object)

        Dim fsT As Object
        Dim fileName As String
        Dim us As New CultureInfo("en-US")

        Dim DTpctList As List(Of Double) = paramObj(0)
        Dim SPDList As List(Of Double) = paramObj(1)
        Dim MTBFList As List(Of Double) = paramObj(2)
        Dim StopsList As List(Of Double) = paramObj(3)
        Dim LineNamesList As List(Of String) = paramObj(4)
        Dim MTTRlist As List(Of Double) = paramObj(5)
        Dim selectedfailuremodename As String = paramObj(6)
        Dim DatesList_starttime As List(Of String) = paramObj(7)
        Dim DatesList_endtime As List(Of String) = paramObj(8)
        Dim isalldatessame As Boolean = paramObj(9)
        Dim ShowStops_insteadofSPD As Boolean = False


        If selectedfailuremodename = "CO" Or InStr(selectedfailuremodename, "Changeover", vbTextCompare) > 0 Or InStr(selectedfailuremodename, "change", vbTextCompare) Or InStr(selectedfailuremodename, "C/O", vbTextCompare) > 0 Then
            ShowStops_insteadofSPD = True
        End If

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "MultilineByLossArea.html"




        fsT.WriteText("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-'//W3C'//DTD HTML 4.01'//EN" & Chr(34) & " " & Chr(34) & "http:'//www.w3.org/TR/html4/strict.dtd" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<html>" & vbCrLf)
        fsT.WriteText("<head>" & vbCrLf)
        fsT.WriteText("<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "style.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "amcharts.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "serial.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        'fsT.WriteText("<script src=" & Chr(34) & "export.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        'fsT.WriteText("<link  type=" & Chr(34) & "text/css" & Chr(34) & " href=" & Chr(34) & "export.css" & Chr(34) & " rel=" & Chr(34) & "stylesheet" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<script>" & vbCrLf)

        fsT.WriteText("var chart;" & vbCrLf)

        fsT.WriteText("var titlesofchart = [{" & vbCrLf)
        fsT.Writetext("'text': 'Loss Comparison'" & vbCrLf)
        fsT.Writetext("}, {" & vbCrLf)
        fsT.Writetext("'text': '" & selectedfailuremodename & "' ," & vbCrLf)
        fsT.Writetext(" 'bold': false" & vbCrLf)
        fsT.Writetext("}];" & vbCrLf)


        fsT.WriteText("var chartData = [" & vbCrLf)
        'THIS IS WHERE THE RAW DATA GOES
        '          {
        '      "fault": "a1",
        '      "dt": 23.5,
        '      "stops": 18.1
        '  },

        For eventIncrementer As Integer = 0 To DTpctList.Count - 1
            fsT.WriteText("{" & Chr(34) & "DTpct" & Chr(34) & ": " & DTpctList(eventIncrementer).ToString("######0.0", us) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "stops per day" & Chr(34) & ": " & Math.Round(SPDList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "stops" & Chr(34) & ": " & Math.Round(StopsList(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            If isalldatessame = True Then
                fsT.WriteText(Chr(34) & "linename" & Chr(34) & ": " & Chr(34) & LineNamesList(eventIncrementer) & Chr(34) & "," & vbCrLf)
            Else
                fsT.WriteText(Chr(34) & "linename" & Chr(34) & ": " & Chr(34) & LineNamesList(eventIncrementer) & " \n " & DatesList_starttime(eventIncrementer) & " \n " & DatesList_endtime(eventIncrementer) & Chr(34) & "," & vbCrLf)
            End If


            fsT.WriteText(Chr(34) & "mttr" & Chr(34) & ": " & Math.Round(MTTRlist(eventIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "mtbf" & Chr(34) & ": " & Math.Round(MTBFList(eventIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            If eventIncrementer = DTpctList.Count - 1 Then
                fsT.WriteText("}" & vbCrLf)
            Else
                fsT.WriteText("}," & vbCrLf)
            End If
        Next
        '||||||||||||||||||||||||||||||||

        fsT.WriteText("];" & vbCrLf)


        fsT.WriteText("AmCharts.ready(function () {" & vbCrLf)
        '// SERIAL CHART
        fsT.WriteText("chart = new AmCharts.AmSerialChart()" & vbCrLf)

        fsT.WriteText("chart.dataProvider = chartData; " & vbCrLf)
        fsT.WriteText("chart.categoryField = " & Chr(34) & "linename" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.titles = titlesofchart; " & vbCrLf)
        fsT.WriteText("chart.startDuration = 1; " & vbCrLf)
        fsT.WriteText("chart.export; " & vbCrLf)


        '// AXES
        '// category
        fsT.WriteText("var categoryAxis = chart.categoryAxis; " & vbCrLf)
        fsT.WriteText("categoryAxis.gridPosition = " & Chr(34) & "start" & Chr(34) & "; " & vbCrLf)
        If DTpctList.Count - 1 > 6 Then
            fsT.WriteText("categoryAxis.labelRotation = 30;" & vbCrLf)
        Else
            fsT.WriteText("categoryAxis.labelRotation = 0;" & vbCrLf)
        End If

        '// value
        fsT.WriteText("var valueAxis = new AmCharts.ValueAxis(); " & vbCrLf)
        fsT.WriteText("valueAxis.axisColor = " & Chr(34) & "#2C99C3" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("valueAxis.gridAlpha = 0; " & vbCrLf)
        fsT.WriteText("valueAxis.axisThickness = 2; " & vbCrLf)
        fsT.WriteText("valueAxis.minimum = 0; " & vbCrLf)
        fsT.WriteText("chart.addValueAxis(valueAxis); " & vbCrLf)

        '// second value axis (on the right)
        fsT.WriteText("var valueAxis2 = new AmCharts.ValueAxis(); " & vbCrLf)
        fsT.WriteText("valueAxis2.position = " & Chr(34) & "right" & Chr(34) & "; " & vbCrLf) '// this line makes the axis to appear on the right
        fsT.WriteText("valueAxis2.axisColor = " & Chr(34) & "#fcd202" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("valueAxis2.gridAlpha = 0; " & vbCrLf)
        fsT.WriteText("valueAxis.unit = " & Chr(34) & "%" & Chr(34) & ";" & vbCrLf)
        fsT.WriteText("valueAxis2.axisThickness = 2; " & vbCrLf)
        fsT.WriteText("valueAxis2.minimum = 0; " & vbCrLf)
        fsT.WriteText("chart.addValueAxis(valueAxis2); " & vbCrLf)

        '// third value axis (on the right)
        fsT.WriteText("var valueAxis3 = new AmCharts.ValueAxis(); " & vbCrLf)
        fsT.WriteText("valueAxis3.position = " & Chr(34) & "right" & Chr(34) & "; " & vbCrLf) '// this line makes the axis to appear on the right
        fsT.writetext("valueAxis3.offset = 30; " & vbCrLf)
        fsT.WriteText("valueAxis3.axisColor = " & Chr(34) & "#FF8C00" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("valueAxis3.gridAlpha = 0; " & vbCrLf)
        fsT.WriteText("valueAxis3.axisThickness = 2; " & vbCrLf)
        fsT.WriteText("valueAxis3.minimum = 0; " & vbCrLf)
        fsT.WriteText("chart.addValueAxis(valueAxis3); " & vbCrLf)

        '// GRAPHS
        '// column graph DT
        fsT.WriteText("var graph1 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph1.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText(" graph1.title = " & Chr(34) & "DT %" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.lineColor = " & Chr(34) & "#33cccc" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.valueField = " & Chr(34) & "DTpct" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.lineAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph1.fillAlphas = 1; " & vbCrLf)
        fsT.WriteText("graph1.labelText = " & Chr(34) & "[[value]]%" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph1.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b>[[DTpct]] % </b> </span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph1); " & vbCrLf)

        '// line 2nd spd
        fsT.WriteText("var graph2 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph2.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)

        If ShowStops_insteadofSPD = True Then
            fsT.WriteText("graph2.title = " & Chr(34) & "Events" & Chr(34) & "; " & vbCrLf)
        Else
            fsT.WriteText("graph2.title = " & Chr(34) & "Stops per day" & Chr(34) & "; " & vbCrLf)
        End If
        fsT.WriteText("graph2.lineColor = " & Chr(34) & "#ff6699" & Chr(34) & "; " & vbCrLf)
        If ShowStops_insteadofSPD = True Then
            fsT.WriteText("graph2.valueField = " & Chr(34) & "stops" & Chr(34) & "; " & vbCrLf)
        Else
            fsT.WriteText("graph2.valueField = " & Chr(34) & "stops per day" & Chr(34) & "; " & vbCrLf)
        End If
        fsT.WriteText("graph2.lineThickness = 0; " & vbCrLf)
        fsT.WriteText("graph2.lineAlpha = 0; " & vbCrLf)
        fsT.WriteText("graph2.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.bulletBorderThickness = 3; " & vbCrLf)
        fsT.WriteText("graph2.bulletBorderColor = " & Chr(34) & "#fcd202" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.bulletBorderAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph2.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph2.valueAxis = valueAxis2; " & vbCrLf)
        fsT.WriteText("graph2.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
        If ShowStops_insteadofSPD = True Then
            fsT.WriteText("graph2.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b>[[value]] Events </b></span>" & Chr(34) & "; " & vbCrLf)
        Else
            fsT.WriteText("graph2.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b>[[value]] stops per day </b></span>" & Chr(34) & "; " & vbCrLf)
        End If


        fsT.WriteText("chart.addGraph(graph2);" & vbCrLf)

        '// line 3rd mtbf
        fsT.WriteText("var graph3 = new AmCharts.AmGraph(); " & vbCrLf)
        fsT.WriteText("graph3.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.title = " & Chr(34) & "MTBF (min)" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.lineColor = " & Chr(34) & "#33ccff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.valueField = " & Chr(34) & "mtbf" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.lineThickness = 0; " & vbCrLf)
        fsT.WriteText("graph3.lineAlpha = 0; " & vbCrLf)
        fsT.WriteText("graph3.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.bulletBorderThickness = 3; " & vbCrLf)
        fsT.WriteText("graph3.bulletBorderColor = " & Chr(34) & "#FF8C00" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.bulletBorderAlpha = 1; " & vbCrLf)
        fsT.WriteText("graph3.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.valueAxis = valueAxis3; " & vbCrLf)
        fsT.WriteText("graph3.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("graph3.balloonText = " & Chr(34) & "<span style='font-size:13px;'>[[title]]:<b>[[value]]</b></span>" & Chr(34) & "; " & vbCrLf)
        fsT.WriteText("chart.addGraph(graph3);" & vbCrLf)


        '// line 4 mttr
        If ShowStops_insteadofSPD = True Then
            '// line 4 changeover MTTR
            fsT.WriteText("var graph4 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph4.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.title = " & Chr(34) & "MTTR (min)" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.lineColor = " & Chr(34) & "#cc66ff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.valueField = " & Chr(34) & "mttr" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.lineAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph4.fillAlphas = 1; " & vbCrLf)
            fsT.WriteText("graph4.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.valueAxis = valueAxis3; " & vbCrLf)
            fsT.WriteText("graph4.labelText = " & Chr(34) & "[[value]] min" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.balloonText = " & Chr(34) & "<span style='font-size:13px;'>[[title]]:<b>[[value]]</b></span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph4);" & vbCrLf)


        Else
            '// line 4 mttr
            fsT.WriteText("var graph4 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph4.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.title = " & Chr(34) & "MTTR (min)" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.lineColor = " & Chr(34) & "#cc66ff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.valueField = " & Chr(34) & "mttr" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.lineThickness = 0; " & vbCrLf)
            fsT.WriteText("graph4.lineAlpha = 0; " & vbCrLf)
            fsT.WriteText("graph4.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.bulletBorderThickness = 3; " & vbCrLf)
            fsT.WriteText("graph4.bulletBorderColor = " & Chr(34) & "#b3cccc" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.bulletBorderAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph4.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.valueAxis = valueAxis3; " & vbCrLf)
            fsT.WriteText("graph4.labelText = " & Chr(34) & "[[value]]" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.balloonText = " & Chr(34) & "<span style='font-size:13px;'>[[title]]:<b>[[value]]</b></span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph4);" & vbCrLf)

        End If




        '// LEGEND
        fsT.WriteText("var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.WriteText("legend.useGraphSettings = true;" & vbCrLf)
        fsT.WriteText("chart.addLegend(legend);" & vbCrLf)
        fsT.WriteText("chart.hideGraph(graph3);" & vbCrLf)
        If ShowStops_insteadofSPD = False Then  ' is not changeover
            fsT.WriteText("chart.hideGraph(graph4);" & vbCrLf)
        End If


        '// WRITE
        fsT.WriteText("chart.write(" & Chr(34) & "chartdiv" & Chr(34) & ");" & vbCrLf)
        fsT.WriteText("});" & vbCrLf)
        fsT.WriteText("</script>" & vbCrLf)
        fsT.WriteText("</head>" & vbCrLf)

        fsT.WriteText("<body>" & vbCrLf)
        fsT.WriteText("<div id=" & Chr(34) & "chartdiv" & Chr(34) & " style=" & Chr(34) & "width:100%; height:400px;" & Chr(34) & "></div>" & vbCrLf)








        'wrap it up
        fsT.WriteText("</body>" & vbCrLf)
        fsT.WriteText("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception

        End Try
        fsT = Nothing
    End Sub

    Public Sub CreateHTMLMultiline_RollupChart1_Pie(paramObj As Object)  'DTpct
        Dim fsT As Object
        Dim fileName As String
        Dim us As New CultureInfo("en-US")

        Dim LineNamesList As List(Of String) = paramObj(0)
        Dim DTpctList As List(Of Double) = paramObj(1)
        Dim failuremodename As String = paramObj(2)

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Multiline_RollupChart1.html"




        fsT.WriteText("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-'//W3C'//DTD HTML 4.01'//EN" & Chr(34) & " " & Chr(34) & "http:'//www.w3.org/TR/html4/strict.dtd" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<html>" & vbCrLf)
        fsT.WriteText("<head>" & vbCrLf)
        fsT.WriteText("<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "style.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "amcharts.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "pie.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        'fsT.WriteText("<script src=" & Chr(34) & "export.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        'fsT.WriteText("<link  type=" & Chr(34) & "text/css" & Chr(34) & " href=" & Chr(34) & "export.css" & Chr(34) & " rel=" & Chr(34) & "stylesheet" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<script>" & vbCrLf)

        fsT.WriteText("var chart;" & vbCrLf)

        fsT.WriteText("var titlesofchart = [{" & vbCrLf)
        fsT.Writetext("'text': 'Rolled up DT%'" & vbCrLf)
        fsT.Writetext("}, {" & vbCrLf)
        fsT.Writetext("'text': '" & failuremodename & " ' ," & vbCrLf)
        fsT.Writetext(" 'bold': false" & vbCrLf)
        fsT.Writetext("}];" & vbCrLf)


        fsT.WriteText("var chartData = [" & vbCrLf)
        'THIS IS WHERE THE RAW DATA GOES
        '          {
        '      "fault": "a1",
        '      "dt": 23.5,
        '      "stops": 18.1
        '  },

        For eventIncrementer As Integer = 0 To DTpctList.Count - 1
            fsT.WriteText("{" & Chr(34) & "linename" & Chr(34) & ": " & Chr(34) & LineNamesList(eventIncrementer) & Chr(34) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "dtpct" & Chr(34) & ": " & Math.Round(DTpctList(eventIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            If eventIncrementer = LineNamesList.Count - 1 Then
                fsT.WriteText("}" & vbCrLf)
            Else
                fsT.WriteText("}," & vbCrLf)
            End If
        Next
        '||||||||||||||||||||||||||||||||

        fsT.WriteText("];" & vbCrLf)


        fsT.WriteText("AmCharts.ready(function () {" & vbCrLf)
        '// PIE CHART
        fsT.WriteText("chart = new AmCharts.AmPieChart()" & vbCrLf)

        fsT.WriteText("chart.dataProvider = chartData; " & vbCrLf)
        fsT.WriteText("chart.titleField = 'linename'; " & vbCrLf)
        fsT.WriteText("chart.valueField = 'dtpct'; " & vbCrLf)
        fsT.WriteText("chart.outlineColor = '#FFFFFF'; " & vbCrLf)
        fsT.WriteText("chart.outlineAlpha = 0.8;" & vbCrLf)
        fsT.WriteText("chart.outlineThickness = 2;" & vbCrLf)
        fsT.WriteText("chart.labelText = '[[title]]: [[value]]%';" & vbCrLf)
        fsT.WriteText("chart.titles = titlesofchart; " & vbCrLf)
        fsT.WriteText("chart.startDuration = 1; " & vbCrLf)
        fsT.WriteText("chart.export; " & vbCrLf)

        '//Legend
        'fsT.WriteText("var legend = New AmCharts.AmLegend();" & vbCrLf)
        'fsT.WriteText("legend.align = 'center';" & vbCrLf)
        'fsT.WriteText("legend.markerType = 'circle';" & vbCrLf)
        'fsT.WriteText("chart.balloonText = '[[title]]<br><span style='font-size:14px'><b>[[value]]</b></span>';" & vbCrLf)
        'fsT.WriteText(" chart.addLegend(legend);" & vbCrLf)


        '// WRITE
        fsT.WriteText("chart.write(" & Chr(34) & "chartdiv" & Chr(34) & ");" & vbCrLf)
        fsT.WriteText("});" & vbCrLf)
        fsT.WriteText("</script>" & vbCrLf)
        fsT.WriteText("</head>" & vbCrLf)

        fsT.WriteText("<body>" & vbCrLf)
        fsT.WriteText("<div id=" & Chr(34) & "chartdiv" & Chr(34) & " style=" & Chr(34) & "width:100%; height:480px;" & Chr(34) & "></div>" & vbCrLf)



        'wrap it up
        fsT.WriteText("</body>" & vbCrLf)
        fsT.WriteText("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception

        End Try
        fsT = Nothing

    End Sub
    Public Sub CreateHTMLMultiline_RollupChart2_Pie(paramObj As Object)   'SPD
        Dim fsT As Object
        Dim fileName As String
        Dim us As New CultureInfo("en-US")
        Dim LineNamesList As List(Of String) = paramObj(0)
        Dim SPDList As List(Of Double) = paramObj(1)
        Dim failuremodename As String = paramObj(2)

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Multiline_RollupChart2.html"




        fsT.WriteText("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-'//W3C'//DTD HTML 4.01'//EN" & Chr(34) & " " & Chr(34) & "http:'//www.w3.org/TR/html4/strict.dtd" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<html>" & vbCrLf)
        fsT.WriteText("<head>" & vbCrLf)
        fsT.WriteText("<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "style.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "amcharts.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        fsT.WriteText("<script src=" & Chr(34) & "pie.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        'fsT.WriteText("<script src=" & Chr(34) & "export.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
        'fsT.WriteText("<link  type=" & Chr(34) & "text/css" & Chr(34) & " href=" & Chr(34) & "export.css" & Chr(34) & " rel=" & Chr(34) & "stylesheet" & Chr(34) & ">" & vbCrLf)
        fsT.WriteText("<script>" & vbCrLf)

        fsT.WriteText("var chart;" & vbCrLf)

        fsT.WriteText("var titlesofchart = [{" & vbCrLf)
        fsT.Writetext("'text': 'Rolled up stops per day'" & vbCrLf)
        fsT.Writetext("}, {" & vbCrLf)
        fsT.Writetext("'text': '" & failuremodename & " ' ," & vbCrLf)
        fsT.Writetext(" 'bold': false" & vbCrLf)
        fsT.Writetext("}];" & vbCrLf)


        fsT.WriteText("var chartData = [" & vbCrLf)
        'THIS IS WHERE THE RAW DATA GOES
        '          {
        '      "fault": "a1",
        '      "dt": 23.5,
        '      "stops": 18.1
        '  },

        For eventIncrementer As Integer = 0 To SPDList.Count - 1
            fsT.WriteText("{" & Chr(34) & "linename" & Chr(34) & ": " & Chr(34) & LineNamesList(eventIncrementer) & Chr(34) & "," & vbCrLf)
            fsT.WriteText(Chr(34) & "spd" & Chr(34) & ": " & Math.Round(SPDList(eventIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            If eventIncrementer = LineNamesList.Count - 1 Then
                fsT.WriteText("}" & vbCrLf)
            Else
                fsT.WriteText("}," & vbCrLf)
            End If
        Next
        '||||||||||||||||||||||||||||||||

        fsT.WriteText("];" & vbCrLf)


        fsT.WriteText("AmCharts.ready(function () {" & vbCrLf)
        '// PIE CHART
        fsT.WriteText("chart = new AmCharts.AmPieChart()" & vbCrLf)

        fsT.WriteText("chart.dataProvider = chartData; " & vbCrLf)
        fsT.WriteText("chart.titleField = 'linename'; " & vbCrLf)
        fsT.WriteText("chart.valueField = 'spd'; " & vbCrLf)
        fsT.WriteText("chart.outlineColor = '#FFFFFF'; " & vbCrLf)
        fsT.WriteText("chart.outlineAlpha = 0.8;" & vbCrLf)
        fsT.WriteText("chart.outlineThickness = 2;" & vbCrLf)
        fsT.WriteText("chart.labelText = '[[title]]: [[value]]';" & vbCrLf)
        fsT.WriteText("chart.titles = titlesofchart; " & vbCrLf)
        fsT.WriteText("chart.startDuration = 1; " & vbCrLf)
        fsT.WriteText("chart.export; " & vbCrLf)

        '//Legend
        'fsT.WriteText("var legend = New AmCharts.AmLegend();" & vbCrLf)
        'fsT.WriteText("legend.align = 'center';" & vbCrLf)
        'fsT.WriteText("legend.markerType = 'circle';" & vbCrLf)
        'fsT.WriteText("chart.balloonText = '[[title]]<br><span style='font-size:14px'><b>[[value]]</b></span>';" & vbCrLf)
        'fsT.WriteText(" chart.addLegend(legend);" & vbCrLf)


        '// WRITE
        fsT.WriteText("chart.write(" & Chr(34) & "chartdiv" & Chr(34) & ");" & vbCrLf)
        fsT.WriteText("});" & vbCrLf)
        fsT.WriteText("</script>" & vbCrLf)
        fsT.WriteText("</head>" & vbCrLf)

        fsT.WriteText("<body>" & vbCrLf)
        fsT.WriteText("<div id=" & Chr(34) & "chartdiv" & Chr(34) & " style=" & Chr(34) & "width:100%; height:480px;" & Chr(34) & "></div>" & vbCrLf)



        'wrap it up
        fsT.WriteText("</body>" & vbCrLf)
        fsT.WriteText("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception

        End Try
        fsT = Nothing


    End Sub
End Module
