
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes
Imports Telerik.Charting
Imports Telerik.Windows.Controls.ChartView

'
'  	Only show the following columns: %Avail Loss, MTBF, MTTR, STOPS PER DAY –
'a.	expandable if they click somewhere to show Total Events, total DT 
'b.	For now eliminate the difference between Stops Based and Events Based – just show Stops 
'i.	Some other toggle which will show Events mode?   Or perhaps we can just set it by business. 
'3.	Graph the losses -  have some column where you can choose which ones to add to the graph (so they don’t all need to be in there) 
'


Namespace PRSTORY_ULTIMATE
    Public NotInheritable Class RandomNumberGen
        Private Sub New()
        End Sub
        Public Shared random As New Random()

        Public Shared Function GetRandomDouble(minimum As Double, maximum As Double) As Double
            Return random.NextDouble() * (maximum - minimum) + minimum
        End Function
        Public Shared Function GetRandomInt(minimum As Double, maximum As Double) As Integer
            Return CInt(random.NextDouble() * (maximum - minimum) + minimum)
        End Function
    End Class

    ''' <summary>
    ''' Interaction logic for Window_LossAllocation.xaml
    ''' </summary>
    Public Class Window_LossAllocation
#Region "Properties"
        '   Public ActiveDataCollection As ObservableCollection(Of LossAllocationDisplayEvent) = New ObservableCollection(Of LossAllocationDisplayEvent)


        Dim ActiveDataCollection As New ObservableCollection(Of LossAllocationDisplayEvent)()
        Public ReadOnly Property ActiveDataCollectionX() As ObservableCollection(Of LossAllocationDisplayEvent)
            Get
                Return ActiveDataCollection
            End Get
        End Property
        Dim demoMode As Boolean = False
#End Region

        'constructor for demo values
        Public Sub New()
            InitializeComponent()

            initializeDemoData()

            populateLineChart()
            populateBarChart()

            prclicked()
        End Sub

        'constructor w/ series passed in
        Public Sub New(series_Names As List(Of String), series_X As List(Of List(Of Double)), series_Y As List(Of List(Of Double)))

            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.

            Dim numberOfLines As Integer = series_Names.Count
            Dim numberOfTimePeriods As Integer = series_X(0).Count

            'make the chart
            Dim blankDataTemplate = New DataTemplate("")
            Loss_LineChart.Series.Clear()
            Loss_LineChart.Palette = getLineColors()

            Loss_LineChart.VerticalAxis = New LinearAxis()
            Loss_LineChart.VerticalAxis.Title = "MTBF / MTTR (min)"


            For i As Integer = 0 To numberOfLines - 1

                Dim rndName As String = series_Names(i)
                'please note that this code, while clever, is silly and needs to be removed
                Dim newSeries As CategoricalSeries = New LineSeries()

                For j As Integer = 0 To numberOfTimePeriods - 1
                    Dim value As Double = series_Y(i)(j)

                    Dim x = New CategoricalDataPoint
                    x.Value = value
                    x.Category = series_X(i)(j).ToString()
                    x.Label = (rndName & Convert.ToString(": ")) & Math.Round(value, 1) & "min"

                    newSeries.DataPoints.Add(x)

                Next

                newSeries.TrackBallInfoTemplate = blankDataTemplate
                Loss_LineChart.Series.Add(newSeries)
            Next

            Loss_LineChart.HorizontalAxis.LabelInterval = 6
            Loss_LineChart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.None

            mtbfclicked()
        End Sub



        Private Sub initializeDemoData()
            For i As Integer = 0 To 16
                ActiveDataCollection.Add(New LossAllocationDisplayEvent())
            Next
        End Sub


        Private Sub stopclicked()
            stopsbutton.Opacity = 1.0
            prbutton.Opacity = 0.2
            mtbfbutton.Opacity = 0.2

            If demoMode Then
                populateLineChart()
                populateBarChart()
            End If
        End Sub

        Private Sub prclicked()
            stopsbutton.Opacity = 0.2
            prbutton.Opacity = 1.0
            mtbfbutton.Opacity = 0.2

            If demoMode Then
                populateLineChart()
                populateBarChart()
            End If
        End Sub

        Private Sub mtbfclicked()
            stopsbutton.Opacity = 0.2
            prbutton.Opacity = 0.2
            mtbfbutton.Opacity = 1.0

            If demoMode Then
                populateLineChart()
                populateBarChart()
            End If
        End Sub

        Private Sub ToggleChartType()
            If (Loss_LineChart.Visibility = Visibility.Visible) Then
                Loss_LineChart.Visibility = Visibility.Hidden
                Loss_BarChart.Visibility = Visibility.Visible
            Else
                Loss_LineChart.Visibility = Visibility.Visible
                Loss_BarChart.Visibility = Visibility.Hidden
            End If
        End Sub

        Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
            'sender.Opacity = 0.7
        End Sub

        Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
            ' sender.Opacity = 1.0
        End Sub


#Region "Charting"

        Private Sub populateBarChart()
            'demo setup
            Dim numberOfLines As Integer = RandomNumberGen.GetRandomInt(3, 8)

            Dim fakeLossNames As New List(Of String)(New String() {"Filling Section", "SkyNet Activated", "Waning Gibbous", "Waxing Gibbous", "Siesta", "CILs",
                "Training", "Forktruck Race", "Materials", "Catching A Bird", "Failure Mode X", "Fault 587",
                "Fault 8932", "Electrical", "Making", "Engineering", "Utilities", "Case Packaging",
                "Major Breakdown", "Startup/Shutdown", "Maintenance"})

            'make the chart
            Dim blankDataTemplate = New DataTemplate("")
            Loss_BarChart.Series.Clear()
            Loss_BarChart.Palette = getLineColors()

            Loss_BarChart.VerticalAxis = New LinearAxis()
            Loss_BarChart.VerticalAxis.Title = "Availability (%)"

            Dim newSeries As CategoricalSeries = New BarSeries()



            For j As Integer = 0 To numberOfLines - 1
                Dim rndName As String = fakeLossNames.OrderBy(Function(s) Guid.NewGuid()).First()
                'please note that this code, while clever, is silly and needs to be removed
                Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)

                Dim x = New CategoricalDataPoint
                x.Value = value
                x.Category = rndName 'chartCategories(j).ToString("MM/dd")
                x.Label = (rndName & Convert.ToString(": ")) & Math.Round(value, 1) & "%"

                newSeries.DataPoints.Add(x)



                '   newSeries.DataPoints.Add(New CategoricalDataPoint() With {
                '   Key.Value = value,
                '   Key.Category = rndName,
                '   Key.Label = (rndName & Convert.ToString(": ")) & Math.Round(value, 1) & "%"
                '  })
            Next

            newSeries.TrackBallInfoTemplate = blankDataTemplate
            Loss_BarChart.Series.Add(newSeries)

            ' Loss_BarChart.HorizontalAxis.LabelInterval = 6;
            Loss_BarChart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.MultiLine

        End Sub


        Private Sub populateLineChart()
            'demo setup
            Dim numberOfLines As Integer = RandomNumberGen.GetRandomInt(1, 4)
            Dim numberOfTimePeriods As Integer = 60
            Dim fakeLossNames As New List(Of String)(New String() {"Filling Section", "SkyNet Activated", "Waning Gibbous", "Waxing Gibbous", "Siesta", "CILs",
                "Training", "Forktruck Race", "Materials", "Catching A Bird", "Failure Mode X", "Fault 587",
                "Fault 8932", "Electrical", "Making", "Engineering", "Utilities", "Case Packaging",
                "Major Breakdown", "Startup/Shutdown", "Maintenance"})
            Dim chartCategories As New List(Of DateTime)()
            Dim startDate As DateTime = DateTime.Now.AddDays(-90)
            For i As Integer = 0 To numberOfTimePeriods - 1
                chartCategories.Add(startDate.AddDays(i))
            Next

            'make the chart
            Dim blankDataTemplate = New DataTemplate("")
            Loss_LineChart.Series.Clear()
            Loss_LineChart.Palette = getLineColors()

            Loss_LineChart.VerticalAxis = New LinearAxis()
            Loss_LineChart.VerticalAxis.Title = "Availability (%)"


            For i As Integer = 0 To numberOfLines - 1

                Dim rndName As String = fakeLossNames.OrderBy(Function(s) Guid.NewGuid()).First()
                'please note that this code, while clever, is silly and needs to be removed
                Dim newSeries As CategoricalSeries = New LineSeries()

                For j As Integer = 0 To numberOfTimePeriods - 1
                    Dim value As Double = RandomNumberGen.GetRandomDouble(4, 70)

                    Dim x = New CategoricalDataPoint
                    x.Value = value
                    x.Category = chartCategories(j).ToString("MM/dd")
                    x.Label = (rndName & Convert.ToString(": ")) & Math.Round(value, 1) & "%"

                    newSeries.DataPoints.Add(x)

                    '  newSeries.DataPoints.Add(New CategoricalDataPoint() With {
                    '  Key.Value = value,
                    '  Key.Category = chartCategories(j).ToString("MM/dd"),
                    '  Key.Label = (rndName & Convert.ToString(": ")) & Math.Round(value, 1) & "%"
                    '  })
                Next

                newSeries.TrackBallInfoTemplate = blankDataTemplate
                Loss_LineChart.Series.Add(newSeries)
            Next

            Loss_LineChart.HorizontalAxis.LabelInterval = 6
            Loss_LineChart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.None

        End Sub

        Public Function getLineColors() As ChartPalette
            Dim tmp = New ChartPalette()
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


#Region "UI Events"

        Private Sub ToggleGrids(sender As Object, e As MouseEventArgs)
            If Loss_GridView.Visibility = Visibility.Visible Then
                Loss_GridView.Visibility = Visibility.Hidden
                Loss_GridView2.Visibility = Visibility.Visible
            Else
                Loss_GridView2.Visibility = Visibility.Hidden
                Loss_GridView.Visibility = Visibility.Visible
            End If
        End Sub

        Private Sub ChartTrackBallBehavior_InfoUpdated(sender As Object, e As TrackBallInfoEventArgs)
            Dim tmpString = ""
            For Each info As DataPointInfo In e.Context.DataPointInfos
                ' info.DisplayHeader = "Custom data point header";
                tmpString += info.DataPoint.Label & Environment.NewLine
            Next

            e.Header = tmpString
        End Sub

        Public Sub Gridview_SelectionChanged(sender As Object, e As Telerik.Windows.Controls.SelectionChangeEventArgs)
            If e.AddedItems.Count > 0 Then
                'make sure we have some items
                'find the event, the name, and the index in our collection
                Dim tmpEvent = DirectCast(e.AddedItems(0), LossAllocationDisplayEvent)
                Dim lossName = tmpEvent.Name
                Dim tmpIndex As Integer = -1
                For i As Integer = 0 To ActiveDataCollection.Count - 1
                    If ActiveDataCollection(i).Name = lossName Then
                        tmpIndex = i
                        Exit For
                    End If
                Next
                'do something about it!
                If tmpIndex > -1 Then
                    populateLineChart()
                    populateBarChart()
                Else
                    'somehow we selected something not in the bound list? this shouldn't happen?
                    MessageBox.Show(Convert.ToString("Unknown Loss Selected: ") & lossName)
                End If
            End If
        End Sub
#End Region

    End Class


    Public Class LossAllocationDisplayEvent
#Region "Properties"
        Public Property Name() As String
            Get
                Return m_Name
            End Get
            Set
                m_Name = Value
            End Set
        End Property
        Private m_Name As String

        Public Property Tier2() As String
            Get
                Return m_Tier2
            End Get
            Set
                m_Tier2 = Value
            End Set
        End Property
        Private m_Tier2 As String
        Public Property DT() As Double
            Get
                Return m_DT
            End Get
            Set
                m_DT = Value
            End Set
        End Property
        Private m_DT As Double
        Public Property DTpct() As String
            Get
                Return m_DTpct
            End Get
            Set
                m_DTpct = Value
            End Set
        End Property
        Private m_DTpct As String

        Public Stops As Double = 0

        Public ReadOnly Property MTBF() As Double
            Get
                Return If(Stops = 0, 0, Math.Round(UT / Stops, 1))
            End Get
        End Property
        Public ReadOnly Property MTTR() As Double
            Get
                Return If(Stops = 0, 0, Math.Round(DT / Stops, 1))
            End Get
        End Property
        Public ReadOnly Property SPD() As Double
            Get
                Return If(MTDF = 0, 0, Math.Round(1 / MTDF, 1))
            End Get
        End Property

        Private ReadOnly Property MTDF() As Double
            Get
                Return MTBF / 1440
            End Get
        End Property

        Public UT As Double = 0
#End Region

#Region "Construction"
        'random stuff for demos
        Public Sub New()
            Dim fakeLossNames As New List(Of String)(New String() {"Filling Section", "SkyNet Activated", "Waning Gibbous", "Waxing Gibbous", "Siesta", "CILs",
                "Training", "Forktruck Race", "Materials", "Catching A Bird", "Failure Mode X", "Fault 587",
                "Fault 8932", "Electrical", "Making", "Engineering", "Utilities", "Case Packaging",
                "Major Breakdown", "Startup/Shutdown", "Maintenance"})
            Dim rndName As String = fakeLossNames.OrderBy(Function(s) Guid.NewGuid()).First()
            'please note that this code, while clever, is silly and needs to be removed
            Dim rndName2 As String = fakeLossNames.OrderBy(Function(s) Guid.NewGuid()).First()
            'please note that this code, while clever, is silly and needs to be removed
            Dim rndStops = RandomNumberGen.GetRandomInt(2, 21)
            Dim rndUT = 999
            Dim rndDT = RandomNumberGen.GetRandomDouble(19, 85)
            Dim rndDTpct = 100 * rndDT / 1440

            Me.Name = rndName
            Me.Tier2 = rndName2
            Me.Stops = rndStops
            Me.UT = rndUT
            Me.DT = Math.Round(rndDT, 2)

            Me.DTpct = Math.Round(rndDTpct, 1) & "%"
        End Sub

        'actual constructor
        Public Sub New(Name As String, Stops As Double, UT As Double, DT As Double, DTpct As Double)
            Me.Name = Name
            Me.Stops = Stops
            Me.UT = UT
            Me.DT = DT

            Me.DTpct = Math.Round(DTpct, 1) & "%"
        End Sub
#End Region

    End Class
End Namespace

