Imports System.ComponentModel
Imports System.Threading
Imports System.Collections.ObjectModel
Imports Telerik.Windows.Controls.ChartView
Imports Telerik.Charting

Public Module RandomNumberGen
    Public random As Random = New Random()

    Public Function GetRandomDouble(minimum As Double, maximum As Double) As Double
        Return random.NextDouble() * (maximum - minimum) + minimum
    End Function

    Public Function GetRandomInt(minimum As Double, maximum As Double) As Integer
        Return (random.NextDouble() * (maximum - minimum) + minimum)
    End Function
End Module


Public Class LineDisplayObject
    Public myName As String

    Public Property Name As String

        Get
            Return myName
        End Get
        Set(value As String)
            myName = value
        End Set
    End Property

    Public Sub New(name As String)
        Me.myName = name
    End Sub
End Class


Public Class Window_Multiline2
    Private PotentialLineIndeces As New List(Of Integer)

    Dim _LineList As New ObservableCollection(Of LineDisplayObject)()



    'Dim _LossTreeList_selectedlinetemp As New ArrayList

    Public ReadOnly Property LineList() As ObservableCollection(Of LineDisplayObject)
        Get
            Return _LineList
        End Get
    End Property

#Region "Constructor"
    Public Sub New(lineIndex As Integer, startTime As Date, endTime As Date)

        ' This call is required by the designer.
        InitializeComponent()
        MainDateLabel.Content = Format(startTime, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(endtimeselected, "MMMM dd, yyyy HH:mm").ToString & vbNewLine

        PotentialLineIndeces = GetCompatiblyMappedLines(lineIndex)
        ToggleLineCheckbox()

        dtgreenbox.Visibility = Visibility.Visible
        stopsgreenbox.Visibility = Visibility.Hidden
        mtbfgreenbox.Visibility = Visibility.Hidden

        LaunchAddLine()

        'stopsframe.Visibility = Visibility.Visible


        Tier1comboBox.Items.Add("Roll-up")
        Tier1comboBox.Items.Add("Benchmark")


        Tier2comboBox.Items.Add("Line")
        Tier2comboBox.Items.Add("Failure Mode")
        Tier2comboBox.SelectedItem = "Failure Mode"

        Tier1comboBox.SelectedIndex = 1

        areEventsOn = True
    End Sub
#End Region

#Region "Charts"

    Private Sub populateAllCharts()
        PopulateTier1Chart()
        PopulateTier2Chart()
        PopulateTier3Chart()
    End Sub
    Private Sub PopulateTier1Chart()
        'demo setup
        Dim numberOfLines As Integer = RandomNumberGen.GetRandomInt(9, 30)

        '   List<string> fakeLossNames = new List<string>(new string[] { "Filling Section", "SkyNet Activated", "Waning Gibbous", "Waxing Gibbous", "Siesta", "CILs", "Training", "Forktruck Race", "Materials", "Catching A Bird", "Failure Mode X", "Fault 587", "Fault 8932", "Electrical", "Making", "Engineering", "Utilities", "Case Packaging", "Major Breakdown", "Startup/Shutdown", "Maintenance" })

        'make the chart
        Dim blankDataTemplate = New DataTemplate("")
        Tier1Chart.Series.Clear()
        Tier1Chart.Palette = getLineColors()

        Tier1Chart.VerticalAxis = New LinearAxis()
        If stopsgreenbox.Visibility = Visibility.Visible Then
            Tier1Chart.VerticalAxis.Title = "Stops"
        ElseIf mtbfgreenbox.Visibility = Visibility.Visible Then
            Tier1Chart.VerticalAxis.Title = "MTBF (min)"
        Else
            Tier1Chart.VerticalAxis.Title = "Availability (%)"
        End If


        Dim newSeries As CategoricalSeries = New BarSeries()
        Dim newSeries1 As CategoricalSeries = New BarSeries()
        Dim newSeries2 As CategoricalSeries = New BarSeries()



        If Not isTier1RollUp Then
            newSeries.CombineMode = ChartSeriesCombineMode.Stack
            newSeries1.CombineMode = ChartSeriesCombineMode.Stack
            newSeries2.CombineMode = ChartSeriesCombineMode.Stack

            For j = 0 To numberOfLines
                Dim rndName As String = "Packing Line " & j
                Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)
                Dim x = New CategoricalDataPoint()
                x.Value = value
                x.Category = rndName
                x.Label = rndName & ": " & Math.Round(value, 1) & "%"

                newSeries.DataPoints.Add(x)
                ' newSeries.DataPoints.Add(New CategoricalDataPoint { Value = value, Category = rndName, Label = rndName + ": " + Math.Round(value, 1) + "%" })
            Next j
            For j = 0 To numberOfLines
                Dim rndName As String = "Packing Line " & j
                Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)
                Dim x = New CategoricalDataPoint()
                x.Value = value
                x.Category = rndName
                x.Label = rndName & ": " & Math.Round(value, 1) & "%"

                newSeries2.DataPoints.Add(x)
                ' newSeries.DataPoints.Add(New CategoricalDataPoint { Value = value, Category = rndName, Label = rndName + ": " + Math.Round(value, 1) + "%" })
            Next j
            For j = 0 To numberOfLines
                Dim rndName As String = "Packing Line " & j
                Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)
                Dim x = New CategoricalDataPoint()
                x.Value = value
                x.Category = rndName
                x.Label = rndName & ": " & Math.Round(value, 1) & "%"

                newSeries1.DataPoints.Add(x)
                ' newSeries.DataPoints.Add(New CategoricalDataPoint { Value = value, Category = rndName, Label = rndName + ": " + Math.Round(value, 1) + "%" })
            Next j

        Else
            For ix = 0 To 2
                Dim rndName As String = "All Selected Lines"
                Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)
                Dim x = New CategoricalDataPoint()
                x.Value = value
                x.Category = rndName
                x.Label = rndName & ": " & Math.Round(value, 1) & "%"

                newSeries.DataPoints.Add(x)
            Next ix

        End If

        newSeries.TrackBallInfoTemplate = blankDataTemplate
        Tier1Chart.Series.Add(newSeries)

        newSeries1.TrackBallInfoTemplate = blankDataTemplate
        Tier1Chart.Series.Add(newSeries1)

        newSeries2.TrackBallInfoTemplate = blankDataTemplate
        Tier1Chart.Series.Add(newSeries2)

        ' tier1chart.HorizontalAxis.LabelInterval = 6
        Tier1Chart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.MultiLine

    End Sub
    Private Sub PopulateTier2Chart()
        Dim numberOfLines As Integer = RandomNumberGen.GetRandomInt(5, 13)

        Dim blankDataTemplate = New DataTemplate("")
        Tier2Chart.Series.Clear()
        Tier2Chart.Palette = getLineColors()

        Tier2Chart.VerticalAxis = New LinearAxis()
        If stopsgreenbox.Visibility = Visibility.Visible Then
            Tier2Chart.VerticalAxis.Title = "Stops"
        ElseIf mtbfgreenbox.Visibility = Visibility.Visible Then
            Tier2Chart.VerticalAxis.Title = "MTBF (min)"
        Else
            Tier2Chart.VerticalAxis.Title = "Availability (%)"
        End If

        Dim newSeries As CategoricalSeries = New BarSeries()
        For j = 0 To numberOfLines
            Dim rndName As String
            If isTier2ByLine Then
                rndName = "Packing Line " & j
            Else
                rndName = "Failure Mode " & j
            End If
            Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)
            Dim x = New CategoricalDataPoint()
            x.Value = value
            x.Category = rndName
            x.Label = rndName & ": " & Math.Round(value, 1) & "%"

            newSeries.DataPoints.Add(x)
        Next j


        newSeries.TrackBallInfoTemplate = blankDataTemplate
        Tier2Chart.Series.Add(newSeries)

        Tier2Chart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.MultiLine

    End Sub

    Private Sub PopulateTier3Chart()
        Dim numberOfLines As Integer = RandomNumberGen.GetRandomInt(4, 14)

        Dim blankDataTemplate = New DataTemplate("")
        Tier3Chart.Series.Clear()
        Tier3Chart.Palette = getLineColors()

        Tier3Chart.VerticalAxis = New LinearAxis()
        If stopsgreenbox.Visibility = Visibility.Visible Then
            Tier3Chart.VerticalAxis.Title = "Stops"
        ElseIf mtbfgreenbox.Visibility = Visibility.Visible Then
            Tier3Chart.VerticalAxis.Title = "MTBF (min)"
        Else
            Tier3Chart.VerticalAxis.Title = "Availability (%)"
        End If

        Dim newSeries As CategoricalSeries = New BarSeries()
        For j = 0 To numberOfLines
            Dim rndName As String = "Packing Line " & j
            Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)
            Dim x = New CategoricalDataPoint()
            x.Value = value
            x.Category = rndName
            x.Label = rndName & ": " & Math.Round(value, 1) & "%"

            newSeries.DataPoints.Add(x)
        Next j


        newSeries.TrackBallInfoTemplate = blankDataTemplate
        Tier3Chart.Series.Add(newSeries)

        Tier3Chart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.MultiLine

    End Sub

    Private Sub ChartTrackBallBehavior_InfoUpdated(sender As Object, e As TrackBallInfoEventArgs)

        Dim tmpString As String = ""
        For Each info As DataPointInfo In e.Context.DataPointInfos
            ' info.DisplayHeader = "Custom data point header"
            tmpString += info.DataPoint.Label + Environment.NewLine
        Next
        e.Header = tmpString
    End Sub


    Public Function getLineColors() As ChartPalette

        Dim tmp As ChartPalette = New ChartPalette()
        addPaletteEntry(tmp, 50, 205, 240)
        addPaletteEntry(tmp, 254, 118, 58)
        addPaletteEntry(tmp, 153, 192, 73)
        addPaletteEntry(tmp, 1, 149, 159)
        addPaletteEntry(tmp, 115, 127, 65)
        addPaletteEntry(tmp, 119, 199, 198)
        addPaletteEntry(tmp, 189, 171, 210)
        addPaletteEntry(tmp, 76, 74, 75)
        addPaletteEntry(tmp, 255, 175, 2)
        addPaletteEntry(tmp, 150, 76, 143)
        addPaletteEntry(tmp, 18, 135, 170)
        Return tmp
    End Function
    Private Sub addPaletteEntry(ByRef palette As ChartPalette, R As Byte, G As Byte, B As Byte)

        Dim tmp = New PaletteEntry()
        tmp.Fill = New SolidColorBrush(Color.FromRgb(R, G, B))
        tmp.Stroke = New SolidColorBrush(Color.FromRgb(R, G, B))
        palette.GlobalEntries.Add(tmp)
    End Sub
#End Region

    Private Sub DoAnalyze(sender As Object, e As MouseEventArgs)
        Dim HeaderString As String = ""
        Dim tmpList As New List(Of String)

        For i = 0 To lineListView.SelectedItems.Count - 1
            tmpList.Add(lineListView.SelectedItems(i).Name)
            HeaderString = HeaderString & "  " & lineListView.SelectedItems(i).Name
        Next i

        ChartHeaderLabel.Content = HeaderString
        CloseAddLine()
    End Sub

#Region "UI Stuff"
    Dim isTier1RollUp As Boolean = False
    Dim isTier2ByLine As Boolean = False
    Private Sub Tier1ComboChanged(sender As Object, e As SelectionChangedEventArgs)
        If Tier1comboBox.SelectedItem = "Roll-up" Then
            isTier1RollUp = True
            Tier2comboBox.Visibility = Visibility.Visible
        Else
            isTier1RollUp = False
            areEventsOn = False
            Tier2comboBox.SelectedItem = "Failure Mode"
            areEventsOn = True
            Tier2comboBox.Visibility = Visibility.Hidden
        End If
        populateAllCharts()
    End Sub

    Dim areEventsOn As Boolean = False
    Private Sub Tier2ComboChanged(sender As Object, e As SelectionChangedEventArgs)
        If areEventsOn Then
            If Tier2comboBox.SelectedItem = "Line" Then
                isTier2ByLine = True
            Else
                isTier2ByLine = False
            End If
            PopulateTier2Chart()
            PopulateTier3Chart()
        End If
    End Sub

    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        sender.opacity = 0.8

    End Sub
    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        sender.opacity = 1.0

    End Sub
    Private Sub Specialmousemove(sender As Object, e As MouseEventArgs)
        sender.opacity = 0.8
        sender.background = mybrushNOTESblue
        sender.foreground = mybrushlanguagewhite

    End Sub
    Private Sub Specialmouseleave(sender As Object, e As MouseEventArgs)
        sender.opacity = 1.0
        sender.background = mybrushlanguagewhite
        sender.foreground = mybrushdefaultfontgray
    End Sub
    Private Sub BackButtonClicked(sender As Object, e As MouseButtonEventArgs)
        Me.Close()
    End Sub

    Private Sub LaunchAddLine()
        SplashCanvas.Visibility = Visibility.Visible
    End Sub
    Private Sub CloseAddLine()
        SplashCanvas.Visibility = Visibility.Hidden
    End Sub

    Public Sub ToggleLineCheckbox()
        _LineList.Clear()
        If checkBox.IsChecked Then
            For i = 0 To PotentialLineIndeces.Count - 1
                _LineList.Add(New LineDisplayObject(AllProdLines(PotentialLineIndeces(i)).ToString))
            Next
        Else
            For i = 0 To AllProdLines.Count - 1
                _LineList.Add(New LineDisplayObject(AllProdLines(i).ToString))
            Next
        End If
    End Sub


#End Region

    Private Function GetCompatiblyMappedLines(LineIndex As Integer) As List(Of Integer)
        Dim tmpList As New List(Of Integer)

        For i = 0 To AllProdLines.Count - 1
            If AllProdLines(i).prStoryMapping = AllProdLines(LineIndex).prStoryMapping Then tmpList.Add(i)
        Next

        Return tmpList
    End Function


    Private Sub FrameMouseMove(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Hand

        sender.Opacity = 0.8


        If sender Is DTpercentframe Then
            DTpercentframe.Opacity = 0.8
        End If

        If sender Is stopsframe Then
            stopsframe.Opacity = 0.8
        End If

        If sender Is mtbfframe Then
            mtbfframe.Opacity = 0.8
        End If



    End Sub
    Private Sub Frameclick(sender As Object, e As MouseButtonEventArgs)
        If sender Is stopsframe Then
            dtgreenbox.Visibility = Visibility.Hidden
            stopsgreenbox.Visibility = Visibility.Visible
            mtbfgreenbox.Visibility = Visibility.Hidden
        ElseIf sender Is DTpercentframe Then
            dtgreenbox.Visibility = Visibility.Visible
            stopsgreenbox.Visibility = Visibility.Hidden
            mtbfgreenbox.Visibility = Visibility.Hidden
        ElseIf sender Is mtbfframe Then
            dtgreenbox.Visibility = Visibility.Hidden
            stopsgreenbox.Visibility = Visibility.Hidden
            mtbfgreenbox.Visibility = Visibility.Visible
        End If

        populateAllCharts()

    End Sub


    Private Sub FrameMouseLeave(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Arrow
        sender.Opacity = 1.0
        If sender Is DTpercentframe Then
            DTpercentframe.Opacity = 1.0
        ElseIf sender Is stopsframe Then
            stopsframe.Opacity = 1.0
        ElseIf sender Is mtbfframe Then
            mtbfframe.Opacity = 1.0
        End If
    End Sub




End Class
