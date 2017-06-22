Imports System.Globalization
Module HTML
    Public Sub createMotionChart(DailyLEDSreports As List(Of SummaryReport))
        Dim fsT As Object
        Dim fileName As String
        Dim FileToDelete As String
        'Dim tmpDTevent As DTevent
        Dim i As Integer
        Dim cardstring As String
        Dim prstoryMOTIONreport As prStoryMainPageReport
        Dim starttempdate As Date
        Dim endtempdate As Date


        If False Then
            cardstring = ""
            endtempdate = endtimeselected

            For i = 1 To 30

                endtempdate = endtempdate.AddDays(-1)
                starttempdate = endtempdate.AddDays(-1)
                prstoryMOTIONreport = New prStoryMainPageReport(selectedindexofLine_temp, starttempdate, endtempdate)
                'tmpDTevent = prStoryReport.getCardEventInfo(31, i)

                cardstring = cardstring & "['Line Performance'," & starttempdate & "," & prstoryMOTIONreport.PR & "," & prstoryMOTIONreport.StopsPerDay & ", 'Total'],"
            Next i


            FileToDelete = SERVER_FOLDER_PATH & "Motion" & ".html"
            If System.IO.File.Exists(FileToDelete) = True Then
                System.IO.File.Delete(FileToDelete)
            End If
            fsT = CreateObject("ADODB.Stream")
            fsT.Type = 2 'Specify stream type - we want To save text/string data.
            fsT.Charset = "utf-8" 'Specify charset For the source text data.
            fsT.Open() 'Open the stream And write binary data To the object

            fileName = SERVER_FOLDER_PATH & "Motion" & ".html"
            'FIND DATA TO EXPORT

            fsT.writetext("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">" & vbCrLf)
            fsT.writetext("<html>" & vbCrLf)
            fsT.writetext("<head>" & vbCrLf)


            fsT.writetext(" <script type=" & Chr(34) & "text/javascript" & Chr(34) & " src=" & Chr(34) & "https://www.google.com/jsapi" & Chr(34) & "></script>" & vbCrLf)
            fsT.writetext("<script type='text/javascript'>" & vbCrLf)
            fsT.writetext("google.load('visualization', '1', {packages: ['motionchart']});" & vbCrLf)

            fsT.writetext("google.setOnLoadCallback(drawChart);" & vbCrLf)
            fsT.writetext("    function drawChart()" & vbCrLf)
            fsT.writetext("   {" & vbCrLf)

            fsT.writetext("var data = new google.visualization.DataTable();" & vbCrLf)

            fsT.writetext("data.addColumn('string', 'Line');" & vbCrLf)
            fsT.writetext("data.addColumn('date', 'Date');" & vbCrLf)
            fsT.writetext("data.addColumn('number', 'PR Loss');" & vbCrLf)
            fsT.writetext("data.addColumn('number', 'Stops');" & vbCrLf)
            fsT.writetext("data.addColumn('string', 'Type');" & vbCrLf)


            fsT.writetext("data.addRows([" & vbCrLf)
            '     fsT.writetext(chartdatatable_line2 & vbCrLf)
            fsT.writetext("]);" & vbCrLf)

            fsT.writetext("var chart = new google.visualization.MotionChart(document.getElementById('chart_div'));" & vbCrLf)
            fsT.writetext("chart.draw(data, {width: 1200, height: 500});" & vbCrLf)
            fsT.writetext("}" & vbCrLf)

            fsT.writetext("</script>" & vbCrLf)

            fsT.writetext("</head>" & vbCrLf)
            fsT.writetext("<body>" & vbCrLf)
            fsT.writetext("<div id='chart_div'></div>" & vbCrLf)

            'wrap it up
            fsT.WriteText("</body>" & vbCrLf)
            fsT.WriteText("</html>" & vbCrLf)

            'fin
            fsT.SaveToFile(fileName, 2) 'Save binary data To disk
            fsT = Nothing
        End If
    End Sub

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

    Public Sub CreateD3LossChart_HTML(cardnumber As Integer)
        Dim fsT As Object
        Dim fileName As String

        Dim FileToDelete As String

        'Deletes existing file.
        FileToDelete = SERVER_FOLDER_PATH & "card" & cardnumber & ".html"
        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete(FileToDelete)
        End If

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        fileName = SERVER_FOLDER_PATH & "index.html" ' "card" & cardnumber & ".html"




        fsT.writeText("<!DOCTYPE html>" & vbCrLf)
        fsT.writeText("<meta charset='utf-8'>" & vbCrLf)
        fsT.writeText("<style>" & vbCrLf)

        fsT.writeText(".node {" & vbCrLf)
        fsT.writeText("cursor: pointer;" & vbCrLf)
        fsT.writeText("}" & vbCrLf)

        fsT.writeText(".node:hover {" & vbCrLf)
        fsT.writeText("stroke: #000;" & vbCrLf)
        fsT.writeText("stroke-width: 1.5px;" & vbCrLf)
        fsT.writeText("}" & vbCrLf)

        fsT.writeText(".node--leaf {" & vbCrLf)
        fsT.writeText("fill: white;" & vbCrLf)
        fsT.writeText("}" & vbCrLf)

        fsT.writeText(".label {" & vbCrLf)
        fsT.writeText("font: 11px 'Helvetica Neue', Helvetica, Arial, sans-serif;" & vbCrLf)
        fsT.writeText("text-anchor: middle;" & vbCrLf)
        fsT.writeText("text-shadow: 0 1px 0 #fff, 1px 0 0 #fff, -1px 0 0 #fff, 0 -1px 0 #fff;" & vbCrLf)
        fsT.writeText("}" & vbCrLf)

        fsT.writeText(".label," & vbCrLf)
        fsT.writeText(".node--root," & vbCrLf)
        fsT.writeText(".node--leaf {" & vbCrLf)
        fsT.writeText("pointer-events: none;" & vbCrLf)
        fsT.writeText("}" & vbCrLf)

        fsT.writeText("</style>" & vbCrLf)
        fsT.writeText("<body>" & vbCrLf)
        fsT.writeText("<script src='http://d3js.org/d3.v3.min.js'></script>" & vbCrLf)
        fsT.writeText("<script>" & vbCrLf)

        fsT.writeText("var margin = 20," & vbCrLf)
        fsT.writeText("diameter = 960;" & vbCrLf)

        fsT.writeText("var color = d3.scale.linear()" & vbCrLf)
        fsT.writeText(".domain([-1, 5])" & vbCrLf)
        fsT.writeText(".range(['hsl(152,80%,80%)', 'hsl(228,30%,40%)'])" & vbCrLf)
        fsT.writeText(".interpolate(d3.interpolateHcl);" & vbCrLf)

        fsT.writeText("var pack = d3.layout.pack()" & vbCrLf)
        fsT.writeText(".padding(2)" & vbCrLf)
        fsT.writeText(".size([diameter - margin, diameter - margin])" & vbCrLf)
        fsT.writeText(".value(function(d) { return d.size; })" & vbCrLf)

        fsT.writeText("var svg = d3.select('body').append('svg')" & vbCrLf)
        fsT.writeText(".attr('width', diameter)" & vbCrLf)
        fsT.writeText(".attr('height', diameter)" & vbCrLf)
        fsT.writeText(".append('g')" & vbCrLf)
        fsT.writeText(".attr('transform', 'translate(' + diameter / 2 + ',' + diameter / 2 + ')');" & vbCrLf)

        fsT.writeText("d3.json('flare.json', function(error, root) {" & vbCrLf)
        fsT.writeText("if (error) return console.error(error);" & vbCrLf)

        fsT.writeText("var focus = root," & vbCrLf)
        fsT.writeText("nodes = pack.nodes(root)," & vbCrLf)
        fsT.writeText("view;" & vbCrLf)

        fsT.writeText("var circle = svg.selectAll('circle')" & vbCrLf)
        fsT.writeText(".data(nodes)" & vbCrLf)
        fsT.writeText(".enter().append('circle')" & vbCrLf)
        fsT.writeText(".attr('class', function(d) { return d.parent ? d.children ? 'node' : 'node node--leaf' : 'node node--root'; })" & vbCrLf)
        fsT.writeText(".style('fill', function(d) { return d.children ? color(d.depth) : null; })" & vbCrLf)
        fsT.writeText(".on('click', function(d) { if (focus !== d) zoom(d), d3.event.stopPropagation(); });" & vbCrLf)

        fsT.writeText("var Text = svg.selectAll('text')" & vbCrLf)
        fsT.writeText(".data(nodes)" & vbCrLf)
        fsT.writeText(".enter().append('text')" & vbCrLf)
        fsT.writeText(".attr('class', 'label')" & vbCrLf)
        fsT.writeText(".style('fill-opacity', function(d) { return d.parent === root ? 1 : 0; })" & vbCrLf)
        fsT.writeText(".style('display', function(d) { return d.parent === root ? null : 'none'; })" & vbCrLf)
        fsT.writeText(".text(function(d) { return d.name; });" & vbCrLf)

        fsT.writeText("var node = svg.selectAll('circle,text');" & vbCrLf)

        fsT.writeText("d3.select('body')" & vbCrLf)
        fsT.writeText(".style('background', color(-1))" & vbCrLf)
        fsT.writeText(".on('click', function() { zoom(root); });" & vbCrLf)

        fsT.writeText("zoomTo([root.x, root.y, root.r * 2 + margin]);" & vbCrLf)

        fsT.writeText("function zoom(d) {" & vbCrLf)
        fsT.writeText("var focus0 = focus; focus = d;" & vbCrLf)

        fsT.writeText("var transition = d3.transition()" & vbCrLf)
        fsT.writeText(".duration(d3.event.altKey ? 7500 : 750)" & vbCrLf)
        fsT.writeText(".tween('zoom', function(d) {" & vbCrLf)
        fsT.writeText("var i = d3.interpolateZoom(view, [focus.x, focus.y, focus.r * 2 + margin]);" & vbCrLf)
        fsT.writeText("return function(t) { zoomTo(i(t)); };" & vbCrLf)
        fsT.writeText("});" & vbCrLf)

        fsT.writeText("transition.selectAll('text')" & vbCrLf)
        fsT.writeText(".filter(function(d) { return d.parent === focus || this.style.display === 'inline'; })" & vbCrLf)
        fsT.writeText(".style('fill-opacity', function(d) { return d.parent === focus ? 1 : 0; })" & vbCrLf)
        fsT.writeText(".each('start', function(d) { if (d.parent === focus) this.style.display = 'inline'; })" & vbCrLf)
        fsT.writeText(".each('end', function(d) { if (d.parent !== focus) this.style.display = 'none'; });" & vbCrLf)
        fsT.writeText("}" & vbCrLf)

        fsT.writeText("function zoomTo(v) {" & vbCrLf)
        fsT.writeText("var k = diameter / v[2]; view = v;" & vbCrLf)
        fsT.writeText("node.attr('transform', function(d) { return 'translate(' + (d.x - v[0]) * k + ', ' + (d.y - v[1]) * k + ')'; });" & vbCrLf)
        fsT.writeText("circle.attr('r', function(d) { return d.r * k; });" & vbCrLf)
        fsT.writeText("}" & vbCrLf)
        fsT.writeText("});" & vbCrLf)

        fsT.writeText("d3.select(self.frameElement).style('height', diameter + 'px');" & vbCrLf)

        fsT.writeText("</script>" & vbCrLf)
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing
    End Sub

    Sub CreateSurvivalPlot(survivaltabledata As String)

        Dim fsT As Object
        Dim fileName As String
        Dim FileToDelete As String

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




        fsT.writetext("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">" & vbCrLf)
        fsT.writetext("<html>" & vbCrLf)
        fsT.writetext("<head>" & vbCrLf)


        fsT.writetext("<script type='text/javascript' src=" & Chr(34) & "https://www.google.com/jsapi?autoload={'modules':[{'name':'visualization', 'version':'1','packages':['corechart']}]}" & Chr(34) & "></script>" & vbCrLf)
        fsT.writetext("<script type='text/javascript'>" & vbCrLf)
        fsT.writetext("google.setOnLoadCallback(drawChart);" & vbCrLf)
        fsT.writetext("function drawChart()" & vbCrLf)
        fsT.writetext("   {" & vbCrLf)
        fsT.writetext("var data = google.visualization.arrayToDataTable([" & vbCrLf)
        'fsT.writetext("['Time', 'Survivability']," & vbCrLf)
        fsT.writetext(survivaltabledata & vbCrLf)
        '''''''
        fsT.writetext("]);" & vbCrLf)

        fsT.writetext("var options = { curveType:  'function', legend: { position: 'bottom' }, width: 770, height: 370, 'chartArea': {'width': '70%', 'height': '70%'} , vAxis: { title: 'Uptime Survival Probability', viewWindow:{ max:1.0, min:0.0 } } , hAxis: { title: 'Time (minutes)'}  };" & vbCrLf)



        fsT.writetext("var chart = new google.visualization.LineChart(document.getElementById('curve_chart'));" & vbCrLf)
        fsT.writetext("chart.draw(data, options);}" & vbCrLf)

        fsT.writetext("</script>" & vbCrLf)




        fsT.writetext("</head>" & vbCrLf)
        fsT.writetext("<body>" & vbCrLf)
        fsT.writetext("<div id='curve_chart'></div>" & vbCrLf)

        'wrap it up
        fsT.WriteText("</body>" & vbCrLf)
        fsT.WriteText("</html>" & vbCrLf)

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
