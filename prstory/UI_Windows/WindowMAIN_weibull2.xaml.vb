
Imports System.Collections.ObjectModel
Imports Telerik.Windows.Controls.ChartView
Imports Telerik.Charting

Public Class WindowMAIN_weibull2
    Private prstoryReport_weibull As prStoryMainPageReport

    Public modeNames As New List(Of String)
    Public modeXVals As New List(Of List(Of Double))
    Public modeYVals As New List(Of List(Of Double))

    Public systemXVals As New List(Of Double)
    Public systemYVals As New List(Of Double)

    Private _ActiveDataCollection As New ObservableCollection(Of WeibullDisplayEvent)

    Private selectedNames As New List(Of String)
    Public Property ActiveDataCollection As ObservableCollection(Of WeibullDisplayEvent)
        Get
            Return _ActiveDataCollection
        End Get
        Set(value As ObservableCollection(Of WeibullDisplayEvent))
            _ActiveDataCollection = value
        End Set
    End Property



    Public Sub New(ByVal storyReport As prStoryMainPageReport)
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        prstoryReport_weibull = storyReport

        LineName.Content = AllProdLines(selectedindexofLine_temp).ToString
       ' CreateSurvivalTable() 'My.Settings.maxuptimecutoff)

        _ActiveDataCollection.Add(New WeibullDisplayEvent("System"))
        For Each i As String In modeNames
            _ActiveDataCollection.Add(New WeibullDisplayEvent(i))
        Next
    End Sub


    Private Sub populateLineChart()

        Dim blankDataTemplate = New DataTemplate("")
        Loss_LineChart.Series.Clear()
        Loss_LineChart.Palette = getLineColors()

        Loss_LineChart.VerticalAxis = New LinearAxis()
        Loss_LineChart.VerticalAxis.Title = "Uptime Probability"


        For x As Integer = 0 To selectedNames.Count - 1

            Dim newSeries As CategoricalSeries = New LineSeries()


            If (selectedNames(x) = "System") Then
                For ii As Integer = 0 To systemYVals.Count - 1
                    Dim cdp As New CategoricalDataPoint()
                    cdp.Value = systemYVals(ii)
                    cdp.Category = systemXVals(ii)
                    cdp.Label = "System - Time: " & systemXVals(ii) & ", UT%: " & systemYVals(ii)
                    newSeries.DataPoints.Add(cdp) 'New CategoricalDataPoint { Value = value, Category = chartCategories[j].ToString("MM/dd"), Label = rndName + ": " + Math.Round(value, 1) + "%" })
                Next
            Else

                Dim targetIndex As Integer = -1
                For ix = 0 To modeNames.Count - 1
                    If selectedNames(x) = modeNames(ix) Then
                        targetIndex = ix
                    End If
                Next


                For ii As Integer = 0 To modeYVals(targetIndex).Count - 1
                    Dim cdp As New CategoricalDataPoint()
                    cdp.Value = modeYVals(targetIndex)(ii)
                    cdp.Category = modeXVals(targetIndex)(ii)
                    cdp.Label = selectedNames(x) & " - Time: " & modeXVals(targetIndex)(ii) & ", UT%: " & modeYVals(targetIndex)(ii)
                    newSeries.DataPoints.Add(cdp) 'New CategoricalDataPoint { Value = value, Category = chartCategories[j].ToString("MM/dd"), Label = rndName + ": " + Math.Round(value, 1) + "%" })
                Next
            End If

            'For (Int() j = 0; j < numberOfTimePeriods; j++)
            '       {
            'Double value = RandomNumberGen.GetRandomDouble(4, 70);
            'newSeries.DataPoints.Add(New CategoricalDataPoint { Value = value, Category = chartCategories[j].ToString("MM/dd"), Label = rndName + ": " + Math.Round(value, 1) + "%" });
            '       }



            newSeries.TrackBallInfoTemplate = blankDataTemplate
            Loss_LineChart.Series.Add(newSeries)

        Next


        ' For (Int() i = 0; i < numberOfLines; i++)
        '    {

        'String rndName = fakeLossNames.OrderBy(s >= Guid.NewGuid()).First(); //please note that this code, while clever, Is silly And needs to be removed

        'CategoricalSeries newSeries = New LineSeries();

        'For (Int() j = 0; j < numberOfTimePeriods; j++)
        '       {
        'Double value = RandomNumberGen.GetRandomDouble(4, 70);
        'newSeries.DataPoints.Add(New CategoricalDataPoint { Value = value, Category = chartCategories[j].ToString("MM/dd"), Label = rndName + ": " + Math.Round(value, 1) + "%" });
        '       }
        'newSeries.TrackBallInfoTemplate = blankDataTemplate;
        'Loss_LineChart.Series.Add(newSeries);
        '   }

        '    Loss_LineChart.HorizontalAxis.LabelInterval = 6;
        '   Loss_LineChart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.None;

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


    
    Public Sub Gridview_SelectionChanged(sender As Object, e As Telerik.Windows.Controls.SelectionChangeEventArgs)
        selectedNames.Clear()

        '  If (e.AddedItems.Count > 0) Then
        '  For i As Integer = 0 To e.AddedItems.Count - 1
        For i As Integer = 0 To Loss_GridView.SelectedItems.Count - 1
            Dim tmpEvent = Loss_GridView.SelectedItems(i)
            Dim lossName = tmpEvent.Name
            selectedNames.Add(lossName)
        Next
        ' End If

        populateLineChart()
    End Sub


    Private Sub ChartTrackBallBehavior_InfoUpdated(sender As Object, e As Telerik.Windows.Controls.ChartView.TrackBallInfoEventArgs)

        Dim tmpString = ""
        For Each info As Telerik.Windows.Controls.ChartView.DataPointInfo In e.Context.DataPointInfos

            ' info.DisplayHeader = "Custom data point header";
            tmpString += info.DataPoint.Label + Environment.NewLine
        Next

        e.Header = tmpString
    End Sub



End Class

Public Class WeibullDisplayEvent
    Public _Name As String
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    Public Sub New(i As String)
        _Name = i
    End Sub
End Class


