Imports System.Collections.ObjectModel
Imports Telerik.Charting
Imports Telerik.Windows.Controls.ChartView
Imports System.Threading

Public Class Window_RateLossReport

#Region "Variables"
    Private updateGraphThread As Thread

    Private mybrushTabSelected As New SolidColorBrush(Color.FromRgb(37, 160, 218))
    Private mybrushTabNOTSelected As New SolidColorBrush(Color.FromRgb(202, 221, 228))

    Public eventListX As New List(Of RateLossEvent)
    Public EventList As New List(Of RateLossEvent)
    Public schedTime As Double
    Private rawData As Object(,)

    Private eventsON As Boolean = False
    Private KPInum As Integer = 0
    Public netRateLossTime As Double = 0
    Public isDone As Boolean = False

    'chart Dimiables
    Private ChartData As New List(Of Tuple(Of String, Double, Double, Integer))
    Private ChartData2 As New List(Of Tuple(Of String, Double, Double, Integer))

    Private ChartNames As New List(Of String)
    Private ChartDTpct As New List(Of Double)
    Private ChartDT As New List(Of Double)
    Private ChartStops As New List(Of String)
    Private ChartNames2 As New List(Of String)
    Private ChartDTpct2 As New List(Of Double)
    Private ChartDT2 As New List(Of Double)
    Private ChartStops2 As New List(Of String)
    Private CurrentDowntimeField As Integer = DowntimeField.Reason1
    Private CurrentDowntimeField2 As Integer = DowntimeField.Tier1

    Private CurrentTier1Selection As String = ""

    Public FilteredList As List(Of String) = New List(Of String)



    Public ReadOnly Property ActiveDataCollection() As ObservableCollection(Of RateLossEvent)
        Get
            Return _ActiveDataCollection
        End Get
    End Property
    Dim _ActiveDataCollection As New ObservableCollection(Of RateLossEvent)()

#End Region

    Public Sub NewKPISelected2()
        Dim y As New List(Of Tuple(Of String, Double, Double, Integer))

        If KPIcomboBox2.SelectedItem = "DT %" Then
            KPInum = 0
        Else
            KPInum = 1
        End If

        populateChartValues()
        updateGraph()
        populateChartValues2(CurrentTier1Selection)
        updateGraph2()

    End Sub

    Public Sub NewKPISelected3()
        Dim x As String = KPIcomboBox3.SelectedItem

        If x = "Mapping" Then
            CurrentDowntimeField2 = DowntimeField.Tier1
        ElseIf x = "Reason 1" Then
            CurrentDowntimeField2 = DowntimeField.Reason1
        ElseIf x = "Reason 2" Then
            CurrentDowntimeField2 = DowntimeField.Reason2
        ElseIf x = "Reason 3" Then
            CurrentDowntimeField2 = DowntimeField.Reason3
        ElseIf x = "Reason 4" Then
            CurrentDowntimeField2 = DowntimeField.Reason4
        ElseIf x = "Team" Then
            CurrentDowntimeField2 = DowntimeField.Team
        ElseIf x = "% Rate" Then
            CurrentDowntimeField2 = DowntimeField.Tier2
        ElseIf x = "SKU" Then
            CurrentDowntimeField2 = DowntimeField.ProductGroup
        End If
        If ChartData.Count > 0 Then
            If CurrentTier1Selection = "" Then CurrentTier1Selection = ChartData(0).Item1
            populateChartValues2(CurrentTier1Selection)
            updateGraph2()
        End If
    End Sub

    Private Sub sortChartData()
        Dim y As New List(Of Tuple(Of String, Double, Double, Integer))

        If KPIcomboBox2.SelectedItem = "DT %" Then
            y = ChartData.OrderBy(Function(X) X.Item2).ToList
        Else
            y = ChartData.OrderBy(Function(X) X.Item4).ToList
        End If
        y.Reverse()
        ChartData = y


    End Sub

    Private Sub sortChartData2()
        Dim y As New List(Of Tuple(Of String, Double, Double, Integer))

        If KPIcomboBox2.SelectedItem = "DT %" Then
            y = ChartData2.OrderBy(Function(X) X.Item2).ToList
        Else
            y = ChartData2.OrderBy(Function(X) X.Item4).ToList
        End If
        y.Reverse()
        ChartData2 = y


    End Sub

    Public Sub NewKPISelected()


        Dim x As String = KPIcomboBox.SelectedItem

        If x = "Mapping" Then
            CurrentDowntimeField = DowntimeField.Tier1
        ElseIf x = "Reason 1" Then
            CurrentDowntimeField = DowntimeField.Reason1
        ElseIf x = "Reason 2" Then
            CurrentDowntimeField = DowntimeField.Reason2
        ElseIf x = "Reason 3" Then
            CurrentDowntimeField = DowntimeField.Reason3
        ElseIf x = "Reason 4" Then
            CurrentDowntimeField = DowntimeField.Reason4
        ElseIf x = "Team" Then
            CurrentDowntimeField = DowntimeField.Team
        ElseIf x = "% Rate" Then
            CurrentDowntimeField = DowntimeField.Tier2
        ElseIf x = "SKU" Then
            CurrentDowntimeField = DowntimeField.ProductGroup
        ElseIf x = "Fault" Then
            CurrentDowntimeField = DowntimeField.Fault

        End If
        '  CurrentTier1Selection = ""

        updateSecondMappingBox()
        populateChartValues()
        If ChartData.Count > 0 Then NewTier1Selected(ChartData(0).Item1)
        updateSecondMappingBox()
        ' populateChartValues2()
        updateGraph()
        ' updateGraph2()
    End Sub

    ' Dim eventsOn As Boolean = False
    Private Sub updateSecondMappingBox()
        Dim x As String
        KPIcomboBox3.Items.Clear()
        x = "Mapping"
        If x <> KPIcomboBox.SelectedItem Then KPIcomboBox3.Items.Add(x)
        x = "Reason 1"
        If x <> KPIcomboBox.SelectedItem Then KPIcomboBox3.Items.Add(x)
        x = "Reason 2"
        If x <> KPIcomboBox.SelectedItem Then KPIcomboBox3.Items.Add(x)
        x = "% Rate"
        If x <> KPIcomboBox.SelectedItem Then KPIcomboBox3.Items.Add(x)
        x = "Team"
        If x <> KPIcomboBox.SelectedItem Then KPIcomboBox3.Items.Add(x)
        x = "SKU"
        If x <> KPIcomboBox.SelectedItem Then KPIcomboBox3.Items.Add(x)
        x = "Fault"
        If x <> KPIcomboBox.SelectedItem Then KPIcomboBox3.Items.Add(x)

        KPIcomboBox3.SelectedIndex = 0
    End Sub


    Private Sub populateChartValues()
        ChartNames.Clear()
        ChartDT.Clear()
        ChartDTpct.Clear()
        ChartStops.Clear()
        ChartData.Clear()

        Dim tmpString As String, tmpIndex As Integer
        For i As Integer = 0 To EventList.Count - 1
            tmpString = EventList(i).getFieldFromInteger(CurrentDowntimeField)
            tmpIndex = ChartNames.IndexOf(tmpString)
            If tmpIndex = -1 Then
                ChartNames.Add(tmpString)
                ChartDT.Add(EventList(i).DT)
                ChartStops.Add(1)
            Else
                ChartDT(tmpIndex) += EventList(i).DT
                ChartStops(tmpIndex) += 1
            End If
        Next


        For j As Integer = 0 To ChartDT.Count - 1
            ChartDTpct.Add(Math.Round(ChartDT(j) * 100 / schedTime, 1))
        Next

        ' updateGraph()
        For k As Integer = 0 To ChartDT.Count - 1
            ChartData.Add(New Tuple(Of String, Double, Double, Integer)(ChartNames(k), ChartDTpct(k), ChartDT(k), ChartStops(k)))
        Next

        sortChartData()
    End Sub



    Private Sub populateChartValues2(targetString As String)
        Dim testString As String
        ChartNames.Clear()
        ChartDT.Clear()
        ChartDTpct.Clear()
        ChartStops.Clear()
        ChartData2.Clear()

        Dim tmpString As String, tmpIndex As Integer
        For i As Integer = 0 To EventList.Count - 1
            testString = EventList(i).getFieldFromInteger(CurrentDowntimeField)
            If testString = targetString Then
                tmpString = EventList(i).getFieldFromInteger(CurrentDowntimeField2)
                tmpIndex = ChartNames.IndexOf(tmpString)
                If tmpIndex = -1 Then
                    ChartNames.Add(tmpString)
                    ChartDT.Add(EventList(i).DT)
                    ChartStops.Add(1)
                Else
                    ChartDT(tmpIndex) += EventList(i).DT
                    ChartStops(tmpIndex) += 1
                End If
            End If
        Next


        For j As Integer = 0 To ChartDT.Count - 1
            ChartDTpct.Add(Math.Round(ChartDT(j) * 100 / schedTime, 1))
        Next

        ' updateGraph()
        For k As Integer = 0 To ChartDT.Count - 1
            ChartData2.Add(New Tuple(Of String, Double, Double, Integer)(ChartNames(k), ChartDTpct(k), ChartDT(k), ChartStops(k)))
        Next

        sortChartData2()
    End Sub


    Private Sub updateGraph()
        Dim blankDataTemplate = New DataTemplate("")
        ParetoChartView.Series.Clear()
        '  ParetoChartView.Palette = Trends_defaultChartColors()

            dim xyz = New LinearAxis()
        xyz.minimum = 0
        ParetoChartView.VerticalAxis = xyz

        'Dim secondaryVAxis = New LinearAxis()
        'secondaryVAxis.HorizontalLocation = AxisHorizontalLocation.Right

        'find axis titles
        '  Dim AxisTitle1 As String = "DT%"
        '   Dim AxisTitle2 As String = ""

        Dim primarySeries As CategoricalSeries = New BarSeries()
        '   Dim secondarySeries As CategoricalSeries = New LineSeries()

        '   secondarySeries.VerticalAxis = secondaryVAxis

        For i As Integer = 0 To ChartData.Count - 1
            Dim tmpPoint As CategoricalDataPoint = New CategoricalDataPoint()
            If KPInum = 0 Then
                tmpPoint.Value = ChartData(i).Item2
                tmpPoint.Label = ChartData(i).Item1 & " " & ChartData(i).Item2 & "% DT"
                ParetoChartView.VerticalAxis.Title = "DT%"
            Else
                tmpPoint.Value = ChartData(i).Item4
                tmpPoint.Label = ChartData(i).Item1 & " " & ChartData(i).Item4 & " Events"
                ParetoChartView.VerticalAxis.Title = "Events"
            End If
            tmpPoint.Category = ChartData(i).Item1

            primarySeries.DataPoints.Add(tmpPoint)


        Next

        'add the data for right time period
        '   String labelIntroString = "MultiLine " + getStringForEnum_Metric(ListofSelectedKPI_LineTrends[metricInc]) + ": "

        '        If (LineTrends_analysistimeperiod == 1) Then ' daily Then

        '        For (Int() i = 0 i < intermediate.Trends_Line_MasterDataList_Daily_RollUp[metricIndex].Count i++)

        '        Double value = intermediate.Trends_Line_MasterDataList_Daily_RollUp[metricIndex][i]
        '       newSeries.DataPoints.Add(New CategoricalDataPoint  Value = value, Category = intermediate.Multi_AllSystemReports_Daily[0][i].startTime.ToString("MM/dd"), Label = labelIntroString + Math.Round(value, 1) )


        'add the correct point template
        '    If (LineTrends_isLineGraph) Then

        'String hexColor = Color_HexFromPaletteEntry(Trends_defaultChartColors(), ParetoChartView.Series.Count)
        'newSeries.PointTemplate = Telerik_getLinePoint("#" + hexColor)



        'wrap it up
        primarySeries.TrackBallInfoTemplate = blankDataTemplate
        ParetoChartView.Series.Add(primarySeries)

        'wrap it up
        '    secondarySeries.TrackBallInfoTemplate = blankDataTemplate
        '    ParetoChartView.Series.Add(secondarySeries)

        ParetoChartView.HorizontalAxis.LabelInterval = 1
        ParetoChartView.HorizontalAxis.LabelFitMode = AxisLabelFitMode.MultiLine
    End Sub


    Private Sub updateGraph2()
        Dim blankDataTemplate = New DataTemplate("")
        ParetoChartView2.Series.Clear()
        '  ParetoChartView.Palette = Trends_defaultChartColors()
        dim xyz = New LinearAxis()
        xyz.minimum = 0
        ParetoChartView2.VerticalAxis = xyz

        'Dim secondaryVAxis = New LinearAxis()
        'secondaryVAxis.HorizontalLocation = AxisHorizontalLocation.Right

        'find axis titles
        '  Dim AxisTitle1 As String = "DT%"
        '   Dim AxisTitle2 As String = ""

        Dim primarySeries As CategoricalSeries = New BarSeries()
        '   Dim secondarySeries As CategoricalSeries = New LineSeries()

        '   secondarySeries.VerticalAxis = secondaryVAxis

        For i As Integer = 0 To ChartData2.Count - 1
            Dim tmpPoint As CategoricalDataPoint = New CategoricalDataPoint()
            If KPInum = 0 Then
                tmpPoint.Value = ChartData2(i).Item2
                tmpPoint.Label = ChartData2(i).Item1 & " " & ChartData2(i).Item2 & "% DT"
                ParetoChartView2.VerticalAxis.Title = "DT%"
            Else
                tmpPoint.Value = ChartData2(i).Item4
                tmpPoint.Label = ChartData2(i).Item1 & " " & ChartData2(i).Item4 & " Events"
                ParetoChartView2.VerticalAxis.Title = "Events"
            End If
            tmpPoint.Category = ChartData2(i).Item1

            primarySeries.DataPoints.Add(tmpPoint)


        Next

        'add the data for right time period
        '   String labelIntroString = "MultiLine " + getStringForEnum_Metric(ListofSelectedKPI_LineTrends[metricInc]) + ": "

        '        If (LineTrends_analysistimeperiod == 1) Then ' daily Then

        '        For (Int() i = 0 i < intermediate.Trends_Line_MasterDataList_Daily_RollUp[metricIndex].Count i++)

        '        Double value = intermediate.Trends_Line_MasterDataList_Daily_RollUp[metricIndex][i]
        '       newSeries.DataPoints.Add(New CategoricalDataPoint  Value = value, Category = intermediate.Multi_AllSystemReports_Daily[0][i].startTime.ToString("MM/dd"), Label = labelIntroString + Math.Round(value, 1) )


        'add the correct point template
        '    If (LineTrends_isLineGraph) Then

        'String hexColor = Color_HexFromPaletteEntry(Trends_defaultChartColors(), ParetoChartView.Series.Count)
        'newSeries.PointTemplate = Telerik_getLinePoint("#" + hexColor)



        'wrap it up
        primarySeries.TrackBallInfoTemplate = blankDataTemplate
        ParetoChartView2.Series.Add(primarySeries)

        'wrap it up
        '    secondarySeries.TrackBallInfoTemplate = blankDataTemplate
        '    ParetoChartView.Series.Add(secondarySeries)

        ParetoChartView2.HorizontalAxis.LabelInterval = 1
        ParetoChartView2.HorizontalAxis.LabelFitMode = AxisLabelFitMode.MultiLine
    End Sub

    Public DoIFilter As Boolean = False
    Public inclusionList As List(Of String) = New List(Of String)
    Private FilterField As DowntimeField

    Public Sub New(lineName As String, schedTime As Double, startTime As Date, endTime As Date, rawData As Object(,), isFilterOn As Boolean, inclusionList As List(Of String), filterFIeld As DowntimeField, INDEX as integer)

        ' This call is required by the designer.
        InitializeComponent()
        lineindex = index
        ' Add any initialization after the InitializeComponent() call.
        TitleLabel.Content = lineName & " Rate Loss"
        Me.schedTime = schedTime
        Me.rawData = rawData

        Dim BusyContent As New List(Of String)
        Dim rnd = New Random()

        DoIFilter = isFilterOn
        Me.inclusionList = inclusionList
        Me.FilterField = filterFIeld

        BusyContent.Add("Programming Flux Capacitor...")
        BusyContent.Add("Warming Hyperdrive...")
        BusyContent.Add("Spinning up the hamster...")
        BusyContent.Add("Shovelling coal into the server...")
        BusyContent.Add("Gremlins frantically finding data...")
        BusyContent.Add("Waiting for Godot...")
        BusyContent.Add("Replacing the vacuum tubes...")
        BusyContent.Add("Determining Universal Physical Constants...")



        Dim randomLineName = BusyContent(rnd.Next(0, BusyContent.Count))
        Title = lineName + " - Rate Loss"
        MainDateLabel.Content = Format(startTime, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(endtimeselected, "MMMM dd, yyyy HH:mm").ToString & vbNewLine
        BusyIndicator.BusyContent = randomLineName

        KPIcomboBox.Items.Add("Mapping")
        KPIcomboBox.Items.Add("Reason 1")
        KPIcomboBox.Items.Add("Reason 2")
        ''   KPIcomboBox.Items.Add("Reason 3")
        '  KPIcomboBox.Items.Add("Reason 4")
        KPIcomboBox.Items.Add("% Rate")
        KPIcomboBox.Items.Add("Team")
        KPIcomboBox.Items.Add("SKU")
        KPIcomboBox.Items.Add("Fault")

        KPIcomboBox3.Items.Add("Mapping")
        KPIcomboBox3.Items.Add("Reason 1")
        KPIcomboBox3.Items.Add("Reason 2")
        ''   KPIcomboBox.Items.Add("Reason 3")
        '  KPIcomboBox.Items.Add("Reason 4")
        KPIcomboBox3.Items.Add("% Rate")
        KPIcomboBox3.Items.Add("Team")
        KPIcomboBox3.Items.Add("SKU")


        KPIcomboBox2.Items.Add("DT %")
        KPIcomboBox2.Items.Add("Events")



        Dim bw = New ComponentModel.BackgroundWorker()
        '  bw.WorkerReportsProgress = True
        bw.WorkerSupportsCancellation = True
        AddHandler bw.DoWork, AddressOf bw_DoWork
        '  AddHandler bw.ProgressChanged, AddressOf bw_ProgressChanged
        AddHandler bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted

        bw.RunWorkerAsync()
    End Sub

#Region "Data Table"
    Private Sub bw_DoWork()
        parseRateData(rawData)
        isDone = True
        ' populateChartValues()
        ' System.Threading.Thread.Sleep(2000)
    End Sub
    Private Sub bw_RunWorkerCompleted()

        ' MessageBox.Show("Here we are!")

        '  populateChartValues()

        ParetoTabClicked()

        BusyIndicator.IsBusy = False

        '  updateGraphThread = New Thread(AddressOf populateChartValues)
        '  updateGraphThread.SetApartmentState(ApartmentState.STA)
        '  updateGraphThread.Start()

        KPIcomboBox.SelectedItem = "Reason 1"
        KPIcomboBox2.SelectedItem = "DT %"
    End Sub
    private lineIndex as integer
    Public Sub parseRateData(rawData As Object(,))
        Dim masterEvent As New RateLossEvent(rawData, lineIndex)

        eventListX = masterEvent.getAllEventsFromMaster()
        ' EventList = masterEvent.getAllEventsFromMaster()

        For i As Integer = 0 To eventListX.Count - 1
            If DoIFilter Then
                If FilterField = DowntimeField.Team Then
                    If (inclusionList.IndexOf(eventListX(i).Team) > -1) Then
                        EventList.Add(eventListX(i))
                        netRateLossTime += eventListX(i).DT

                        _ActiveDataCollection.Add(eventListX(i))
                    End If
                Else 'sku
                    If (inclusionList.IndexOf(eventListX(i).ProductGroup) > -1) Then
                        EventList.Add(eventListX(i))
                        netRateLossTime += eventListX(i).DT
                        _ActiveDataCollection.Add(eventListX(i))
                    End If
                End If
            Else
                EventList.Add(eventListX(i))
                netRateLossTime += eventListX(i).DT
                _ActiveDataCollection.Add(eventListX(i))
            End If
        Next

        ' Dim iv As Integer = 1
        'iv += 1

    End Sub
#End Region

#Region "UI"
    Private Sub BarChartSelectionBehavior_SelectionChanged(sender As Object, e As ChartSelectionChangedEventArgs)
        If e.AddedPoints.Count > 0 Then

            ' Dim barSeries = DirectCast(Me.ParetoChartView.Series(0), BarSeries)
            Dim x As String = e.AddedPoints(0).Label
            '  Me.UpdateAll(barSeries.DataPoints)
            Dim targetIndex As Integer = -1
            Dim targetString As String = ""
            TitleLabel2.Content = "No Selection"

            For i = 0 To ChartData.Count - 1
                If x.Contains(ChartData(i).Item1) Then
                    targetIndex = i
                    targetString = ChartData(i).Item1
                    i = ChartData.Count
                End If
            Next

            If targetIndex > -1 Then
                NewTier1Selected(targetString)
            End If

            '   TitleLabel2.Content = "No Selection"
        End If
    End Sub

    Private Sub NewTier1Selected(selectedString As String)
        TitleLabel2.Content = selectedString
        populateChartValues2(selectedString)
        updateGraph2()
    End Sub

    Private Sub ChartTrackBallBehavior_InfoUpdated(sender As Object, e As TrackBallInfoEventArgs)

        Dim tmpString As String = ""
        For Each info As DataPointInfo In e.Context.DataPointInfos

            '// info.DisplayHeader = "Custom data point header"
            tmpString += info.DataPoint.Label + Environment.NewLine
        Next

        e.Header = tmpString
    End Sub

    Private Sub ParetoTabClicked()
        ParetoTab.Background = mybrushTabSelected
        RawDataTab.Background = mybrushTabNOTSelected
        RawDataGridView.Visibility = Visibility.Hidden
        ParetoChartView.Visibility = Visibility.Visible
        KPIcomboBox.Visibility = Visibility.Visible
        KPIcomboBox2.Visibility = Visibility.Visible
        ParetoChartView2.Visibility = Visibility.Visible
        KPIcomboBox3.Visibility = Visibility.Visible
        TitleLabel2.Visibility = Visibility.Visible
    End Sub
    Private Sub RawDataTabClicked()
        ParetoTab.Background = mybrushTabNOTSelected
        RawDataTab.Background = mybrushTabSelected
        RawDataGridView.Visibility = Visibility.Visible
        ParetoChartView.Visibility = Visibility.Hidden
        KPIcomboBox.Visibility = Visibility.Hidden
        KPIcomboBox2.Visibility = Visibility.Hidden
        ParetoChartView2.Visibility = Visibility.Hidden
        KPIcomboBox3.Visibility = Visibility.Hidden
        TitleLabel2.Visibility = Visibility.Hidden

    End Sub


    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        sender.Opacity = 0.7
    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        sender.Opacity = 1.0
    End Sub

    Sub BackButtonClicked()

        Me.Close()
        'My.Settings.Save()
        '  Dim mainprstorywindow As New WindowMain_prstory
        ' Me.Owner.Visibility = Windows.Visibility.Visible


        bargraphreportwindow_Open = False
    End Sub
#End Region


End Class
