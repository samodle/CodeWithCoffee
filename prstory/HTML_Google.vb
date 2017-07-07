Imports System.Globalization
Imports System.Net


Module HTML_Motion


    'AM Charts 
    'Daily PR Overall
    Public Sub exportMotion_PR_HTML_AMCHART(rawData As Motion_LinePRReport) ' As List(Of DTevent))
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & 0 & "D.html"
        ''''''''


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = true; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.dateFormats = [{" & vbCrLf)
        fsT.Writetext("                    period: 'fff'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'ss'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'mm'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'hh'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'DD'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'WW'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'MM'," & vbCrLf)
        fsT.Writetext("                    format: 'MMM'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'YYYY'," & vbCrLf)
        fsT.Writetext("                    format: 'YYYY'" & vbCrLf)
        fsT.Writetext("                }];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph1.title = 'Unplanned DT%';" & vbCrLf)
        fsT.Writetext("               graph1.valueField = 'UnplannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second graph" & vbCrLf)
        fsT.Writetext("               var graph2 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph2.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph2.title = 'Planned DT%';" & vbCrLf)
        fsT.Writetext("               graph2.valueField = 'PlannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph2.bullet = 'square';" & vbCrLf)
        fsT.Writetext("               graph2.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph2.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third graph" & vbCrLf)
        fsT.Writetext("               var graph3 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph3.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph3.valueField = 'PR';" & vbCrLf)
        If My.Settings.AdvancedSettings_isAvailabilityMode Then
            fsT.Writetext("               graph3.title = 'Availability%';" & vbCrLf)
        Else
            fsT.Writetext("               graph3.title = 'PR%';" & vbCrLf)
        End If


        fsT.Writetext("               graph3.bullet = 'triangleUp';" & vbCrLf)
        fsT.Writetext("               graph3.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph3.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 0 To rawData.DailyReports.Count - 1
            'fsT.Writetext("                       {date: new Date('" & rawData.getHTMLdataString_AMCHarts_DateObj(timeIncrementer) & "')" & "," & vbCrLf)
            'Format("10/24/2015 07:30:00 AM", "Short Date") & " " & Format("10/24/2015 07:30:00 AM", "Long Time")

            fsT.Writetext("                       {date: new Date('" & Format(rawData.getHTMLdataString_AMCHarts_DateObj(timeIncrementer), "MM dd yyyy") & "')" & "," & vbCrLf)
            fsT.Writetext("                       PlannedDowntime: " & (Math.Round(rawData.getHTMLdataString_AMCharts_PDT(timeIncrementer), 1)).ToString("######0.0", us) & "," & vbCrLf)
            fsT.Writetext("                       UnplannedDowntime: " & (Math.Round(rawData.getHTMLdataString_AMCharts_UPDT(timeIncrementer), 1)).ToString("######0.0", us) & "," & vbCrLf)
            fsT.Writetext("                       PR: " & (Math.Round(rawData.getHTMLdataString_AMCharts_PR(timeIncrementer), 1)).ToString("######0.0", us) & vbCrLf)
            If timeIncrementer <> rawData.DailyReports.Count - 1 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        'fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)


        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub

    'Daily DT & SPD for selected failuremode
    Public Sub exportMotion_PR_HTML_AMCHART_selectedfailuremode(rawData As MotionReport, isDT As Boolean, failuremodeno As Integer)
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        If isDT Then
            fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "D.html"
        Else
            fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "S.html"
        End If


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = true; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.dateFormats = [{" & vbCrLf)
        fsT.Writetext("                    period: 'fff'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'ss'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'mm'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'hh'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'DD'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'WW'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'MM'," & vbCrLf)
        fsT.Writetext("                    format: 'MMM'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'YYYY'," & vbCrLf)
        fsT.Writetext("                    format: 'YYYY'" & vbCrLf)
        fsT.Writetext("                }];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        If isDT Then
            fsT.Writetext("               graph1.title = 'Unplanned DT%';" & vbCrLf)
        Else
            fsT.Writetext("               graph1.title = 'Stops per Day';" & vbCrLf)
        End If

        fsT.Writetext("               graph1.valueField = 'UnplannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 0 To rawData.DailyReports.Count - 1
            fsT.Writetext("                       {date: new Date('" & Format(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_Date(timeIncrementer, False, failuremodeno), "MM dd yyyy") & "')" & "," & vbCrLf)
            If isDT Then
                fsT.Writetext("                       UnplannedDowntime: " & (Math.Round(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd(timeIncrementer, True, failuremodeno), 1)).ToString("######0.0", us) & "," & vbCrLf)
            Else
                fsT.Writetext("                       UnplannedDowntime: " & (Math.Round(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd(timeIncrementer, False, failuremodeno))).ToString("######0.0", us) & "," & vbCrLf)
            End If

            If timeIncrementer <> rawData.DailyReports.Count - 1 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        ' fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)




        Try
            'fin
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub


    'NEW Code to be taken to  C# AM CHarts SelectedFailuremode DT and SPD
    Public Sub exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(rawData As MotionReport, isDT As Boolean, failuremodeno As Integer)
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        If isDT Then
            fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "D_Monthly.html"
        Else
            fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "S_Monthly.html"
        End If


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        If isDT Then
            fsT.Writetext("               graph1.title = 'Unplanned DT%';" & vbCrLf)
        Else
            fsT.Writetext("               graph1.title = 'Stops per Day';" & vbCrLf)
        End If

        fsT.Writetext("               graph1.valueField = 'UnplannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 1 To 3
            fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_Date_Monthly(timeIncrementer, False, failuremodeno) & "'" & "," & vbCrLf)
            If isDT Then
                fsT.Writetext("                       UnplannedDowntime: " & Math.Round(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd_Monthly(timeIncrementer, True, failuremodeno), 1).ToString("######0.0", us) & "," & vbCrLf)
            Else
                fsT.Writetext("                       UnplannedDowntime: " & Math.Round(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd_Monthly(timeIncrementer, False, failuremodeno)).ToString("######0.0", us) & "," & vbCrLf)
            End If

            If timeIncrementer <> 3 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        ' fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)








        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try

        fsT = Nothing
    End Sub
    Public Sub exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(rawData As MotionReport, isDT As Boolean, failuremodeno As Integer)
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        If isDT Then
            fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "D_Weekly.html"
        Else
            fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "S_Weekly.html"
        End If


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        If isDT Then
            fsT.Writetext("               graph1.title = 'Unplanned DT%';" & vbCrLf)
        Else
            fsT.Writetext("               graph1.title = 'Stops per Day';" & vbCrLf)
        End If

        fsT.Writetext("               graph1.valueField = 'UnplannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)













        For timeIncrementer = ((rawData.DailyReports.Count - 1) Mod 7) To rawData.DailyReports.Count - 1 Step 1
            If timeIncrementer + 6 < 90 Then
                fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_Date_Weekly(timeIncrementer, False, failuremodeno, timeIncrementer + 6) & "'" & "," & vbCrLf)
                If isDT Then
                    fsT.Writetext("                       UnplannedDowntime: " & Math.Round(rawData.getHTMLdataString_AMCharts_DTpctorSPD_Weekly(timeIncrementer, True, failuremodeno, timeIncrementer + 6), 1).ToString("######0.0", us) & "," & vbCrLf)
                Else
                    fsT.Writetext("                       UnplannedDowntime: " & Math.Round(rawData.getHTMLdataString_AMCharts_DTpctorSPD_Weekly(timeIncrementer, False, failuremodeno, timeIncrementer + 6)).ToString("######0.0", us) & "," & vbCrLf)
                End If

                If timeIncrementer + 6 < 83 Then
                    fsT.Writetext("}," & vbCrLf)
                End If
                timeIncrementer = timeIncrementer + 6
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        ' fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)


        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub

    'NEW Code to be taken to  C# AM Charts Total Line PR
    Public Sub exportMotion_PR_HTML_AMCHART_Monthly(rawData As Motion_LinePRReport) ' As List(Of DTevent))
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & 0 & "D_Monthly.html"
        ''''''''


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        'fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)



        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph1.title = 'Unplanned DT%';" & vbCrLf)
        fsT.Writetext("               graph1.valueField = 'UnplannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second graph" & vbCrLf)
        fsT.Writetext("               var graph2 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph2.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph2.title = 'Planned DT%';" & vbCrLf)
        fsT.Writetext("               graph2.valueField = 'PlannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph2.bullet = 'square';" & vbCrLf)
        fsT.Writetext("               graph2.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph2.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third graph" & vbCrLf)
        fsT.Writetext("               var graph3 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph3.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph3.valueField = 'PR';" & vbCrLf)
        If My.Settings.AdvancedSettings_isAvailabilityMode Then
            fsT.Writetext("               graph3.title = 'Availability%';" & vbCrLf)
        Else
            fsT.Writetext("               graph3.title = 'PR%';" & vbCrLf)
        End If


        fsT.Writetext("               graph3.bullet = 'triangleUp';" & vbCrLf)
        fsT.Writetext("               graph3.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph3.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 1 To 3
            fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_AMCharts_Dateobj_Monthly(timeIncrementer) & "'," & vbCrLf)
            fsT.Writetext("                       PlannedDowntime: " & Math.Round(rawData.getHTMLdataString_AMCharts_PDT_Monthly(timeIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)
            fsT.Writetext("                       UnplannedDowntime: " & Math.Round(rawData.getHTMLdataString_AMCharts_UPDT_Monthly(timeIncrementer), 1).ToString("######0.0", us) & "," & vbCrLf)

            If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                fsT.Writetext("                       PR: " & Math.Round(rawData.getHTMLdataString_AMCharts_PR_Monthly(timeIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            Else
                fsT.Writetext("                       PR: " & Math.Round(100 - rawData.getHTMLdataString_AMCharts_PDT_Monthly(timeIncrementer) - rawData.getHTMLdataString_AMCharts_UPDT_Monthly(timeIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            End If



            If timeIncrementer <> 3 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        'fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try

        fsT = Nothing
    End Sub
    Public Sub exportMotion_PR_HTML_AMCHART_Weekly(rawData As Motion_LinePRReport) ' As List(Of DTevent))
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & 0 & "D_Weekly.html"
        ''''''''


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        'fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)



        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph1.title = 'Unplanned DT%';" & vbCrLf)
        fsT.Writetext("               graph1.valueField = 'UnplannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second graph" & vbCrLf)
        fsT.Writetext("               var graph2 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph2.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph2.title = 'Planned DT%';" & vbCrLf)
        fsT.Writetext("               graph2.valueField = 'PlannedDowntime';" & vbCrLf)
        fsT.Writetext("               graph2.bullet = 'square';" & vbCrLf)
        fsT.Writetext("               graph2.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph2.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third graph" & vbCrLf)
        fsT.Writetext("               var graph3 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph3.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph3.valueField = 'PR';" & vbCrLf)
        If My.Settings.AdvancedSettings_isAvailabilityMode Then
            fsT.Writetext("               graph3.title = 'Availability%';" & vbCrLf)
        Else
            fsT.Writetext("               graph3.title = 'PR%';" & vbCrLf)
        End If


        fsT.Writetext("               graph3.bullet = 'triangleUp';" & vbCrLf)
        fsT.Writetext("               graph3.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph3.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)

        For timeIncrementer = ((rawData.DailyReports.Count - 1) Mod 7) To rawData.DailyReports.Count - 1 Step 1
            If timeIncrementer + 6 < 90 Then
                fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_AMCharts_Dateobj_Weekly(timeIncrementer, timeIncrementer + 6) & "'," & vbCrLf)
                fsT.Writetext("                       PlannedDowntime: " & Math.Round(rawData.getHTMLdataString_AMCharts_PDT_Weekly(timeIncrementer, timeIncrementer + 6), 1).ToString("######0.0", us) & "," & vbCrLf)
                fsT.Writetext("                       UnplannedDowntime: " & Math.Round(rawData.getHTMLdataString_AMCharts_UPDT_Weekly(timeIncrementer, timeIncrementer + 6), 1).ToString("######0.0", us) & "," & vbCrLf)

                If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                    fsT.Writetext("                       PR: " & Math.Round(rawData.getHTMLdataString_AMCharts_PR_Weekly(timeIncrementer, timeIncrementer + 6), 1).ToString("######0.0", us) & vbCrLf)
                Else
                    fsT.Writetext("                       PR: " & Math.Round(100 - rawData.getHTMLdataString_AMCharts_UPDT_Weekly(timeIncrementer, timeIncrementer + 6) - rawData.getHTMLdataString_AMCharts_PDT_Weekly(timeIncrementer, timeIncrementer + 6), 1).ToString("######0.0", us) & vbCrLf)
                End If





                If timeIncrementer + 6 < 83 Then
                    fsT.Writetext("}," & vbCrLf)
                End If
                timeIncrementer = timeIncrementer + 6
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        'fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub



    'NEW Code to be taken to C# AM CHarts MTBF
    Public Sub exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF(rawData As MotionReport, isDT As Boolean, failuremodeno As Integer)
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "MTBF.html"




        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = true; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.dateFormats = [{" & vbCrLf)
        fsT.Writetext("                    period: 'fff'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'ss'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'mm'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'hh'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'DD'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'WW'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'MM'," & vbCrLf)
        fsT.Writetext("                    format: 'MMM'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'YYYY'," & vbCrLf)
        fsT.Writetext("                    format: 'YYYY'" & vbCrLf)
        fsT.Writetext("                }];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)

        fsT.Writetext("               graph1.title = 'MTBF (min)';" & vbCrLf)


        fsT.Writetext("               graph1.valueField = 'MTBF';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 0 To rawData.DailyReports.Count - 1
            fsT.Writetext("                       {date: new Date('" & Format(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_Date(timeIncrementer, False, failuremodeno), "MM dd yyyy") & "')" & "," & vbCrLf)

            fsT.Writetext("                       MTBF: " & Math.Round(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_MTBF(timeIncrementer, True, failuremodeno), 1).ToString("######0.0", us) & "," & vbCrLf)


            If timeIncrementer <> rawData.DailyReports.Count - 1 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        ' fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)






        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub
    Public Sub exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Monthly(rawData As MotionReport, isDT As Boolean, failuremodeno As Integer)
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "MTBF_Monthly.html"



        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)

        fsT.Writetext("               graph1.title = 'MTBF (min)';" & vbCrLf)



        fsT.Writetext("               graph1.valueField = 'MTBF';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 1 To 3
            fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_Date_Monthly(timeIncrementer, False, failuremodeno) & "'" & "," & vbCrLf)

            fsT.Writetext("                      MTBF: " & Math.Round(rawData.getHTMLdataString_AMCharts_MTBF_Monthly(timeIncrementer, True, failuremodeno), 1).ToString("######0.0", us) & "," & vbCrLf)


            If timeIncrementer <> 3 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        ' fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)








        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub
    Public Sub exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Weekly(rawData As MotionReport, isDT As Boolean, failuremodeno As Integer)
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & prStoryCard.Stops & "_" & failuremodeno & "MTBF_Weekly.html"



        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)

        fsT.Writetext("               graph1.title = 'MTBF (min)';" & vbCrLf)


        fsT.Writetext("               graph1.valueField = 'MTBF';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)


        For timeIncrementer = ((rawData.DailyReports.Count - 1) Mod 7) To rawData.DailyReports.Count - 1 Step 1
            If timeIncrementer + 6 < 90 Then
                fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_Date_Weekly(timeIncrementer, False, failuremodeno, timeIncrementer + 6) & "'" & "," & vbCrLf)

                fsT.Writetext("                       MTBF: " & Math.Round(rawData.getHTMLdataString_AMCharts_MTBF_Weekly(timeIncrementer, True, failuremodeno, timeIncrementer + 6), 1).ToString("######0.0", us) & "," & vbCrLf)

                If timeIncrementer + 6 < 83 Then
                    fsT.Writetext("}," & vbCrLf)
                End If
                timeIncrementer = timeIncrementer + 6
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        ' fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)








        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub


    'New Code to be taken to C# AM Charts Overall Line SPD and MTBF - Daily Monthly and Weekly
    Public Sub exportMotion_SPD_HTML_AMCHART(rawData As Motion_LinePRReport) ' As List(Of DTevent))
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & 0 & "S.html"
        ''''''''


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = true; // as our data is date-based, we set parseDates to true" & vbCrLf)
        fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.dateFormats = [{" & vbCrLf)
        fsT.Writetext("                    period: 'fff'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'ss'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN:SS'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'mm'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'hh'," & vbCrLf)
        fsT.Writetext("                    format: 'JJ:NN'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'DD'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'WW'," & vbCrLf)
        fsT.Writetext("                    format: 'DD'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'MM'," & vbCrLf)
        fsT.Writetext("                    format: 'MMM'" & vbCrLf)
        fsT.Writetext("                }, {" & vbCrLf)
        fsT.Writetext("                    period: 'YYYY'," & vbCrLf)
        fsT.Writetext("                    format: 'YYYY'" & vbCrLf)
        fsT.Writetext("                }];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph1.title = 'Stops per day';" & vbCrLf)
        fsT.Writetext("               graph1.valueField = 'SPD';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 0 To rawData.DailyReports.Count - 1
            fsT.Writetext("                       {date: new Date('" & Format(rawData.getHTMLdataString_AMCHarts_DateObj(timeIncrementer), "MM dd yyyy") & "')" & "," & vbCrLf)
            fsT.Writetext("                       SPD: " & Math.Round(rawData.getHTMLdataString_AMCharts_SPD(timeIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            If timeIncrementer <> rawData.DailyReports.Count - 1 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        'fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)


        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub
    Public Sub exportMotion_SPD_HTML_AMCHART_Monthly(rawData As Motion_LinePRReport) ' As List(Of DTevent))
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & 0 & "S_Monthly.html"
        ''''''''


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        'fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)



        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph1.title = 'Stops per day';" & vbCrLf)
        fsT.Writetext("               graph1.valueField = 'SPD';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)
        For timeIncrementer = 1 To 3
            fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_AMCharts_Dateobj_Monthly(timeIncrementer) & "'," & vbCrLf)
            fsT.Writetext("                       SPD: " & Math.Round(rawData.getHTMLdataString_AMCharts_SPD_Monthly(timeIncrementer), 1).ToString("######0.0", us) & vbCrLf)
            If timeIncrementer <> 3 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        'fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub
    Public Sub exportMotion_SPD_HTML_AMCHART_Weekly(rawData As Motion_LinePRReport) ' As List(Of DTevent))
        Dim fsT As Object
        Dim fileName As String
        Dim timeIncrementer As Integer
        Dim us As New CultureInfo("en-US")
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object


        fileName = SERVER_FOLDER_PATH & "Motion" & 0 & "S_Weekly.html"
        ''''''''


        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'date';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)
        fsT.Writetext("               categoryAxis.parseDates = false; // as our data is date-based, we set parseDates to true" & vbCrLf)
        'fsT.Writetext("               categoryAxis.minPeriod = 'DD'; // our data is daily, so we set minPeriod to DD" & vbCrLf)
        fsT.Writetext("               categoryAxis.minorGridEnabled = true;" & vbCrLf)
        fsT.Writetext("               categoryAxis.axisColor = '#DADADA';" & vbCrLf)
        fsT.Writetext("               categoryAxis.twoLineMode = true;" & vbCrLf)



        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
        fsT.Writetext("               graph1.title = 'Stops per day';" & vbCrLf)
        fsT.Writetext("               graph1.valueField = 'SPD';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)

        For timeIncrementer = ((rawData.DailyReports.Count - 1) Mod 7) To rawData.DailyReports.Count - 1 Step 1
            If timeIncrementer + 6 < 90 Then
                fsT.Writetext("                       {date: '" & rawData.getHTMLdataString_AMCharts_Dateobj_Weekly(timeIncrementer, timeIncrementer + 6) & "'," & vbCrLf)
                fsT.Writetext("                       SPD: " & Math.Round(rawData.getHTMLdataString_AMCharts_SPD_Weekly(timeIncrementer, timeIncrementer + 6), 1).ToString("######0.0", us) & vbCrLf)
                If timeIncrementer + 6 < 83 Then
                    fsT.Writetext("}," & vbCrLf)
                End If
                timeIncrementer = timeIncrementer + 6
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        'fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)

        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
        End Try
        fsT = Nothing
    End Sub

End Module

Module HTML
    Public Sub createSPCchart(bubblecount As Integer)

        Dim fsT As Object
        Dim fileName As String
        Dim FileToDelete As String
        ' Dim seriestype As String
        Dim charttype As String
        ' Dim formattype As String
        Dim legendtype As String
        '   Dim barwidth As String
        Dim cardstring As String
        '   Dim setcolumnsSPC As String
        Dim daycount As Integer
        Dim tmpEvent As inControlDTevent
        Dim MasterDataSet As inControlReport
        Dim actualSPD As Double
        Dim meanSPD As Double
        Dim SD1pos As Double
        Dim SD2pos As Double
        Dim SD3pos As Double

        cardstring = ""
        FileToDelete = SERVER_FOLDER_PATH & "SPC" & ".html"
        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete(FileToDelete)
        End If
        MasterDataSet = New inControlReport(AllProdLines(selectedindexofLine_temp), incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate) ', selectedRLcolumn)
        tmpEvent = MasterDataSet.inControlEvents(bubblecount)
        For daycount = 0 To MasterDataSet.SPCNumberOfDays - 1
            actualSPD = tmpEvent.DailyStops(daycount)
            meanSPD = tmpEvent.AdjMu_InvMTDF
            SD1pos = meanSPD + tmpEvent.AdjSigma_InvMTDF
            SD2pos = meanSPD + (2 * tmpEvent.AdjSigma_InvMTDF)
            SD3pos = meanSPD + (3 * tmpEvent.AdjSigma_InvMTDF)
            cardstring = cardstring & "['" & daycount & "'," & actualSPD & "," & meanSPD & "," & SD1pos & "," & SD2pos & "," & SD3pos & "],"

        Next
        'chartdatatable_line1 = "['Day','ActualSPD', 'MeanSPD', SD1pos, SD2pos, SD3,pos],"


        'seriestype = "seriesType: 'bars', colors: ['#008080', 'gray']};"
        charttype = "var chart = new google.visualization.LineChart(document.getElementById('chart_div'));"
        ' formattype = "formatter.format(data,1,2)"
        legendtype = "legend: { position: 'none'},"
        'barwidth = "bar: {groupWidth: '50'},"
        '    chartsize = "width: '900', height: '270',"
        '    chartlabel = "var options = {title:'SPC',"
        ' setcolumnsSPC = "view.setColumns([0, 1,{ calc: 'stringify', sourceColumn: 1, type:   'string', role:   'annotation' },2,3]);"

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        fileName = SERVER_FOLDER_PATH & "SPC" & ".html"
        'FIND DATA TO EXPORT

        fsT.writetext("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">" & vbCrLf)
        fsT.writetext("<html>" & vbCrLf)
        fsT.writetext("<head>" & vbCrLf)


        fsT.writetext(" <script type=" & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "https://www.google.com/jsapi" & Chr(34) & "></script>" & vbCrLf)
        fsT.writetext(" <script type=" & Chr(34) & "text/javascript" & Chr(34) & ">google.load('visualization', '1.1', {packages: ['corechart']});</script>" & vbCrLf)

        fsT.writetext(" <script type=" & Chr(34) & "text/javascript" & Chr(34) & ">" & vbCrLf)
        fsT.writetext("google.load('visualization', '1.1', {packages: ['line']});" & vbCrLf)
        fsT.writetext("google.setOnLoadCallback(drawChart);" & vbCrLf)
        fsT.writetext("    function drawChart()" & vbCrLf)
        fsT.writetext("   {" & vbCrLf)

        fsT.writetext("var data = new google.visualization.DataTable(); data.addColumn('string', 'Day'); data.addColumn('number', 'Actual Stops'); data.addColumn('number', 'Mean'); data.addColumn('number', 'Standard Deviation 1'); data.addColumn('number', 'Standard Deviation 2'); data.addColumn('number', 'Standard Deviation 3');" & vbCrLf)
        fsT.writetext("data.addRows([" & vbCrLf)
        fsT.writetext(cardstring & vbCrLf)
        fsT.writetext("]);" & vbCrLf)




        fsT.writetext("var options = {" & vbCrLf)

        fsT.writetext("hAxis: { gridlines: { count: 0 } }," & vbCrLf)
        fsT.writetext("vAxis: { gridlines: { count: 0 } }," & vbCrLf)
        fsT.writetext("width: 900, height: 270," & vbCrLf)
        fsT.writetext("legend: { position: 'none'}," & vbCrLf)
        fsT.writetext("'chartArea': {'width': '100%', 'height': '80%'}," & vbCrLf)
        fsT.writetext("};")

        fsT.writetext("var chart = new google.charts.Line(document.getElementById('linechart_material'));" & vbCrLf)
        fsT.writetext("chart.draw(data, options);" & vbCrLf)

        fsT.writetext("}" & vbCrLf)
        fsT.writetext("google.setOnLoadCallback(drawChart);" & vbCrLf)
        fsT.writetext("</script>" & vbCrLf)




        fsT.writetext("</head>" & vbCrLf)
        fsT.writetext("<body>" & vbCrLf)
        fsT.writetext("  <div id='linechart_material'></div>" & vbCrLf)

        'wrap it up
        fsT.WriteText("</body>" & vbCrLf)
        fsT.WriteText("</html>" & vbCrLf)

        'fin
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing

    End Sub
    Sub CreateSurvivalPlot_AMCHarts(survivaltabledata As Array, actualListsize_ofselectedfailuremodeList As Integer, selectedfailuremodelist As List(Of String))
        Dim fsT As Object
        Dim fileName As String
        Dim us As New CultureInfo("en-US")
        Dim FileToDelete As String
        Dim k As Integer
        Dim cardstring As String
        '   Dim setcolumnsSPC As String

        cardstring = ""
        fileName = SERVER_FOLDER_PATH & "weibull.html"
        FileToDelete = SERVER_FOLDER_PATH & "Weibull" & ".html"
        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete(FileToDelete)
        End If

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        fsT.Writetext("<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>" & vbCrLf)
        fsT.Writetext("<html>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <head>" & vbCrLf)
        fsT.Writetext("        <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf)
        fsT.Writetext("        <title>amCharts examples</title>" & vbCrLf)
        fsT.Writetext("        <link rel='stylesheet' href='style.css' type='text/css'>" & vbCrLf)
        fsT.Writetext("        <script src='amcharts.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("        <script src='serial.js' type='text/javascript'></script>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("        <script>" & vbCrLf)
        fsT.Writetext("           var chart;" & vbCrLf)
        fsT.Writetext("           var chartData = [];" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           AmCharts.ready(function () {" & vbCrLf)
        fsT.Writetext("               // generate some random data first" & vbCrLf)
        fsT.Writetext("               generateChartData();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SERIAL CHART" & vbCrLf)
        fsT.Writetext("               chart = new AmCharts.AmSerialChart();" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               chart.dataProvider = chartData;" & vbCrLf)
        fsT.Writetext("               chart.categoryField = 'uptime';" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // listen for 'dataUpdated' event (fired when chart is inited) and call zoomChart method when it happens" & vbCrLf)
        fsT.Writetext("               chart.addListener('dataUpdated', zoomChart);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // AXES" & vbCrLf)
        fsT.Writetext("               // category" & vbCrLf)
        fsT.Writetext("               var categoryAxis = chart.categoryAxis;" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // first value axis (on the left)" & vbCrLf)
        fsT.Writetext("               var valueAxis1 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisColor = '#FF6600';" & vbCrLf)
        fsT.Writetext("               valueAxis1.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               valueAxis1.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // second value axis (on the right)" & vbCrLf)
        fsT.Writetext("               var valueAxis2 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis2.position = 'right'; // this line makes the axis to appear on the right" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisColor = '#FCD202';" & vbCrLf)
        fsT.Writetext("               valueAxis2.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis2.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis2);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // third value axis (on the left, detached)" & vbCrLf)
        fsT.Writetext("               valueAxis3 = new AmCharts.ValueAxis();" & vbCrLf)
        fsT.Writetext("               valueAxis3.offset = 50; // this line makes the axis to appear detached from plot area" & vbCrLf)
        fsT.Writetext("               valueAxis3.gridAlpha = 0;" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisColor = '#B0DE09';" & vbCrLf)
        fsT.Writetext("               valueAxis3.axisThickness = 2;" & vbCrLf)
        fsT.Writetext("               chart.addValueAxis(valueAxis3);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // GRAPHS" & vbCrLf)
        fsT.Writetext("               // first graph" & vbCrLf)
        fsT.Writetext("               var graph1 = new AmCharts.AmGraph();" & vbCrLf)
        fsT.Writetext("               graph1.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)

        fsT.Writetext("               graph1.title = 'Total Line Survival Probability';" & vbCrLf)


        fsT.Writetext("               graph1.valueField = 'Survivability';" & vbCrLf)
        fsT.Writetext("               graph1.bullet = 'round';" & vbCrLf)
        fsT.Writetext("               graph1.hideBulletsCount = 30;" & vbCrLf)
        fsT.Writetext("               graph1.bulletBorderThickness = 1;" & vbCrLf)
        fsT.Writetext("               chart.addGraph(graph1);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        If selectedfailuremodelist.Count <> 0 Then
            fsT.Writetext("               // second graph" & vbCrLf)
            fsT.Writetext("               var graph2 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph2.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph2.title = '" & selectedfailuremodelist(0) & "';" & vbCrLf)
            fsT.Writetext("               graph2.valueField = 'SurvivabilityFM1';" & vbCrLf)
            fsT.Writetext("               graph2.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph2.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph2.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph2);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 1 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // third graph" & vbCrLf)
            fsT.Writetext("               var graph3 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph3.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph3.title = '" & selectedfailuremodelist(1) & "';" & vbCrLf)
            fsT.Writetext("               graph3.valueField = 'SurvivabilityFM2';" & vbCrLf)
            fsT.Writetext("               graph3.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph3.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph3.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph3);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 2 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // fourth graph" & vbCrLf)
            fsT.Writetext("               var graph4 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph4.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph4.title = '" & selectedfailuremodelist(2) & "';" & vbCrLf)
            fsT.Writetext("               graph4.valueField = 'SurvivabilityFM3';" & vbCrLf)
            fsT.Writetext("               graph4.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph4.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph4.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph4);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 3 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // fifth graph" & vbCrLf)
            fsT.Writetext("               var graph5 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph5.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph5.title = '" & selectedfailuremodelist(3) & "';" & vbCrLf)
            fsT.Writetext("               graph5.valueField = 'SurvivabilityFM4';" & vbCrLf)
            fsT.Writetext("               graph5.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph5.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph5.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph5);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 4 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // six graph" & vbCrLf)
            fsT.Writetext("               var graph6 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph6.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph6.title = '" & selectedfailuremodelist(4) & "';" & vbCrLf)
            fsT.Writetext("               graph6.valueField = 'SurvivabilityFM5';" & vbCrLf)
            fsT.Writetext("               graph6.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph6.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph6.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph6);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 5 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // seven graph" & vbCrLf)
            fsT.Writetext("               var graph7 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph7.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph7.title = '" & selectedfailuremodelist(5) & "';" & vbCrLf)
            fsT.Writetext("               graph7.valueField = 'SurvivabilityFM6';" & vbCrLf)
            fsT.Writetext("               graph7.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph7.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph7.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph7);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 6 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // eigth graph" & vbCrLf)
            fsT.Writetext("               var graph8 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph8.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph8.title = '" & selectedfailuremodelist(6) & "';" & vbCrLf)
            fsT.Writetext("               graph8.valueField = 'SurvivabilityFM7';" & vbCrLf)
            fsT.Writetext("               graph8.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph8.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph8.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph8);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 7 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // ninth graph" & vbCrLf)
            fsT.Writetext("               var graph9 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph9.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph9.title = '" & selectedfailuremodelist(7) & "';" & vbCrLf)
            fsT.Writetext("               graph9.valueField = 'SurvivabilityFM8';" & vbCrLf)
            fsT.Writetext("               graph9.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph9.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph9.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph9);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 8 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // ten graph" & vbCrLf)
            fsT.Writetext("               var graph10 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph10.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph10.title = '" & selectedfailuremodelist(8) & "';" & vbCrLf)
            fsT.Writetext("               graph10.valueField = 'SurvivabilityFM9';" & vbCrLf)
            fsT.Writetext("               graph10.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph10.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph10.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph10);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)
            If 9 = actualListsize_ofselectedfailuremodeList Then GoTo starthere
            fsT.Writetext("               // eleventh graph" & vbCrLf)
            fsT.Writetext("               var graph11 = new AmCharts.AmGraph();" & vbCrLf)
            fsT.Writetext("               graph11.valueAxis = valueAxis1; // we have to indicate which value axis should be used" & vbCrLf)
            fsT.Writetext("               graph11.title = '" & selectedfailuremodelist(9) & "';" & vbCrLf)
            fsT.Writetext("               graph11.valueField = 'SurvivabilityFM10';" & vbCrLf)
            fsT.Writetext("               graph11.bullet = 'round';" & vbCrLf)
            fsT.Writetext("               graph11.hideBulletsCount = 30;" & vbCrLf)
            fsT.Writetext("               graph11.bulletBorderThickness = 1;" & vbCrLf)
            fsT.Writetext("               chart.addGraph(graph11);" & vbCrLf)
            fsT.Writetext("" & vbCrLf)

            fsT.Writetext("" & vbCrLf)
        End If
starthere:
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // CURSOR" & vbCrLf)
        fsT.Writetext("               var chartCursor = new AmCharts.ChartCursor();" & vbCrLf)
        fsT.Writetext("               chartCursor.cursorAlpha = 0.1;" & vbCrLf)
        fsT.Writetext("               chartCursor.fullWidth = true;" & vbCrLf)
        fsT.Writetext("               chart.addChartCursor(chartCursor);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // SCROLLBAR" & vbCrLf)
        fsT.Writetext("               var chartScrollbar = new AmCharts.ChartScrollbar();" & vbCrLf)
        fsT.Writetext("               chart.addChartScrollbar(chartScrollbar);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // LEGEND" & vbCrLf)
        fsT.Writetext("               var legend = new AmCharts.AmLegend();" & vbCrLf)
        fsT.Writetext("               legend.marginLeft = 110;" & vbCrLf)
        fsT.Writetext("               legend.useGraphSettings = true;" & vbCrLf)
        fsT.Writetext("               chart.addLegend(legend);" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("               // WRITE" & vbCrLf)
        fsT.Writetext("               chart.write('chartdiv');" & vbCrLf)
        fsT.Writetext("           });" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // generate some random data, quite different range" & vbCrLf)
        fsT.Writetext("           function generateChartData() {" & vbCrLf)

        fsT.Writetext("" & vbCrLf)

        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("                   chartData.push(" & vbCrLf)

        Dim tempuptime As Double
        Dim tempsurv As Double
        Dim tempsurv_fm As Double

        tempuptime = survivaltabledata(0, 0)
        tempsurv = survivaltabledata(2, 0)


        fsT.Writetext("                       {uptime: " & tempuptime.ToString("######0.000", us) & "," & vbCrLf)
        fsT.Writetext("                       Survivability: " & tempsurv.ToString("######0.000", us) & "," & vbCrLf)
        If selectedfailuremodelist.Count <> 0 Then

            For k = 0 To actualListsize_ofselectedfailuremodeList - 1
                tempsurv_fm = survivaltabledata(k + 3, 0)
                fsT.Writetext("                       SurvivabilityFM" & k + 1 & ": " & tempsurv_fm.ToString("######0.000", us) & "" & vbCrLf)

                If k >= actualListsize_ofselectedfailuremodeList - 1 Then
                    '    Exit For
                Else
                    fsT.Writetext("," & vbCrLf)
                End If

            Next
        End If
        fsT.Writetext("}," & vbCrLf)

        Dim i As Integer

        For i = 1 To survivaltabledata.GetLength(1) - 1
            tempuptime = survivaltabledata(0, i)

            fsT.Writetext("                       {uptime: " & tempuptime.ToString("######0.000", us) & "," & vbCrLf)
            If InStr(ReturnNullifzero(survivaltabledata(2, i)).GetType.ToString, "ouble") > 0 Then
                tempsurv = survivaltabledata(2, i)
                fsT.Writetext("                       Survivability: " & tempsurv.ToString("######0.000", us) & "," & vbCrLf)
            Else
                fsT.Writetext("                       Survivability: " & ReturnNullifzero(survivaltabledata(2, i)) & "," & vbCrLf)
            End If

            If selectedfailuremodelist.Count <> 0 Then
                For k = 0 To actualListsize_ofselectedfailuremodeList - 1
                    If InStr(ReturnNullifzero(survivaltabledata(k + 3, i)).GetType.ToString, "ouble") > 0 Then
                        tempsurv_fm = survivaltabledata(k + 3, i)
                        fsT.Writetext("                       SurvivabilityFM" & k + 1 & ": " & tempsurv_fm.ToString("######0.000", us) & "" & vbCrLf)
                    Else
                        fsT.Writetext("                      SurvivabilityFM" & k + 1 & ": " & ReturnNullifzero(survivaltabledata(k + 3, i)) & "" & vbCrLf)
                    End If
                    If k >= actualListsize_ofselectedfailuremodeList - 1 Then
                        '  Exit For
                    Else
                        fsT.Writetext("," & vbCrLf)
                    End If
                Next
            End If
            'fsT.Writetext("}," & vbCrLf)

            If i <> survivaltabledata.GetLength(1) - 1 Then
                fsT.Writetext("}," & vbCrLf)
            End If
        Next
        fsT.Writetext("                   });" & vbCrLf)


        'fsT.Writetext("               }" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("           // this method is called when chart is first inited as we listen for 'dataUpdated' event" & vbCrLf)
        fsT.Writetext("           function zoomChart() {" & vbCrLf)
        fsT.Writetext("               // different zoom methods can be used - zoomToIndexes, zoomToDates, zoomToCategoryValues" & vbCrLf)
        'fsT.Writetext("               chart.zoomToIndexes(10, 20);" & vbCrLf)
        fsT.Writetext("           }" & vbCrLf)
        fsT.Writetext("        </script>" & vbCrLf)
        fsT.Writetext("    </head>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("    <body>" & vbCrLf)
        fsT.Writetext("        <div id='chartdiv' style='width: 100%; height: 400px;'></div>" & vbCrLf)
        fsT.Writetext("    </body>" & vbCrLf)
        fsT.Writetext("" & vbCrLf)
        fsT.Writetext("</html>" & vbCrLf)




        'fin
        Try
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        Catch ex As Exception
            MsgBox("Chart could not be loaded. Please contact das.l@pg.com to give feedback with detail report info such as - line name, time frame being analyzed etc.")
        End Try
        fsT = Nothing
    End Sub

End Module