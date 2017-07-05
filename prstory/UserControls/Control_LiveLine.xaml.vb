Imports System.Globalization
Imports System.Threading
Imports System.Windows.Media.Animation
Namespace UserControls

    Public Class Control_LiveLine
        Public Sub New()
            InitializeComponent()
        End Sub

  

        Public Sub initialize( rawData as summaryreport, analysisData as summaryreport)
             intermediate = New LiveLineHelper(rawdata, analysisdata)
            LiveLine_GenerateDTViewer()
            LiveLine_TopLoss_GenerateItems()
            LiveLine_TopPlanned_GenerateItems()
            LiveLine_TopDelta_GenerateItems()
            LiveLine_Trend_GenerateChart(DowntimeMetrics.PR, 24)
            LiveLine_UpdateDTViewerLabels()
        End Sub

        Private Property intermediate As LiveLineHelper
        Public Property LiveLineTrends_TimeFrame As Integer = 1
        Public mybrushSelectedCriteria As New SolidColorBrush(Color.FromRgb(50, 205, 240))
        Public mybrushLIGHTBLUEGREEN As New SolidColorBrush(Color.FromRgb(6, 197, 180))

#Region "LiveLine"

        Public Sub LiveLine_onload(sender As Object, e As MouseButtonEventArgs)


        End Sub

        Private Sub LiveLine_UpdateDTViewerLabels()
            LiveLine_label_uptime.Content = "Uptime - " & Math.Round(intermediate.LiveLine_AnalysisPeriodData.PR * 100, 0) & "%"
            LiveLine_label_planned.Content = "Planned - " & Math.Round(intermediate.LiveLine_AnalysisPeriodData.PDTpct * 100, 0) & "%"
            LiveLine_label_unplanned.Content = "Unplanned - " & Math.Round(intermediate.LiveLine_AnalysisPeriodData.UPDTpct * 100, 0) & "%"
        End Sub

        Public Sub LiveLine_GenerateDTViewer()
            LiveLine_ClearDTViewer()
            Dim temprect As Rectangle
            Dim dep As Canvas = LiveLineDTViewerGraphicsCanvas
            Dim templabel As Label
            Dim noofevents As Integer = intermediate.LiveLine_NumberOfEvents


            Dim rectheight As Double = dep.Height
            Dim rectwidth As Double = 0
            Dim sumofalldur As Double = 0
            Dim currentLeftPos As Double = 0

            Dim actualeventdur As New List(Of Double)()
            Dim Colorvalues As New List(Of SolidColorBrush)()
            Colorvalues.Add(mybrushLIGHTBLUEGREEN)
            Colorvalues.Add(bubblecolorRed)
            Colorvalues.Add(BrushColors.mybrushdarkgray)
            Colorvalues.Add(BrushColors.mybrushbrightblue)
            Dim Actualcolor As New SolidColorBrush()

            For j as integer = 0 To noofevents - 1
                actualeventdur.Add(intermediate.LiveLine_ActualDurationOfEachEvent(j))
                sumofalldur = sumofalldur + actualeventdur(j)
            Next

            For i as integer = 0 To noofevents - 1

                rectwidth = (actualeventdur(i) / sumofalldur) * dep.Width
                If intermediate.LiveLine_EventTypes(i) = EventType.Unplanned Then
                    Actualcolor = Colorvalues(1)
                End If
                If intermediate.LiveLine_EventTypes(i) = EventType.Planned Then
                    Actualcolor = Colorvalues(3)
                End If
                If intermediate.LiveLine_EventTypes(i) = EventType.Running Then
                    Actualcolor = Colorvalues(0)
                End If
                If intermediate.LiveLine_EventTypes(i) = EventType.Excluded Then
                    Actualcolor = Colorvalues(2)
                End If

                GenerateRectangleUI(dep, "DTviewrect" & i.ToString(), rectheight, rectwidth, currentLeftPos, 0,
                    Actualcolor, Nothing, 0, AddressOf LiveLine_DTviewer_EventSeleced, AddressOf LiveLine_DTviewer_Eventmousemove, AddressOf LiveLine_DTviewer_Eventmouseleave,
                    0, -1, 1, intermediate.LiveLine_DTviewer_EventNames(i))

                currentLeftPos = currentLeftPos + rectwidth
            Next

            'TimeLabel
            GenerateLabelUI(dep, "DTviewer_TimeLabel", 18, 120, 0, -19,
                Brushes.DarkSlateGray, Brushes.White, 8, Nothing, Nothing, Nothing,
                -1, "")

            templabel = getMenuItem_Label_fromitemindex(LiveLineDTViewerGraphicsCanvas, -1, "", "DTviewer_TimeLabel")
            templabel.Visibility = Visibility.Hidden


            ' Time Highlight rectangle
            GenerateRectangleUI(dep, "DTviewer_TimeHighlight", dep.Height + 10, 10, 0, -5,
                Nothing, BrushColors.mybrushdarkgray, 1, Nothing, Nothing, Nothing)
            temprect = getMenuItem_Rectangle_fromitemindex(dep, -1, "", "DTviewer_TimeHighlight")
            temprect.Visibility = Visibility.Hidden

            'Time Frame Header
            Dim starttime As DateTime = intermediate.LiveLine_SelectedStartTime
            Dim endtime As DateTime = intermediate.LiveLine_SelectedEndTime
            LiveLineDtViewer_TimeFrameHeader.Content = starttime.ToString("MMM", CultureInfo.InvariantCulture) & " " & starttime.ToString("dd", CultureInfo.InvariantCulture) & ", " & starttime.ToString("hh", CultureInfo.InvariantCulture) & ":" & starttime.ToString("mm", CultureInfo.InvariantCulture) & " to " & endtime.ToString("MMM", CultureInfo.InvariantCulture) & " " & endtime.ToString("dd", CultureInfo.InvariantCulture) & ", " & endtime.ToString("hh", CultureInfo.InvariantCulture) & ":" & endtime.ToString("mm", CultureInfo.InvariantCulture)
            LiveLineTrends_TimeFrameHeader.Content = LiveLineDtViewer_TimeFrameHeader.Content

        End Sub

        Public Sub LiveLine_LocateTimeHighlightRectangle(Optional LeftPos As Double = 0, Optional width As Double = 0)
            Dim temprect As Rectangle
            Dim currentLeftPos As Double = 0
            Dim dep As Canvas = LiveLineDTViewerGraphicsCanvas
            temprect = getMenuItem_Rectangle_fromitemindex(dep, -1, "", "DTviewer_TimeHighlight")
            temprect.Width = width + 4
            temprect.Visibility = Visibility.Visible
            currentLeftPos = CDbl(temprect.GetValue(Canvas.LeftProperty))
            System.Windows.Forms.Application.DoEvents()
            AnimateZoomUIElement(currentLeftPos, LeftPos - 2 - width, 0.2, Canvas.LeftProperty, temprect)
        End Sub

        Public Sub LiveLine_ClearDTViewer()
            Dim dep As Canvas = LiveLineDTViewerGraphicsCanvas
            Dim rect As Rectangle
            Dim lbl As Label

            While VisualTreeHelper.GetChildrenCount(dep) <> 0
                If VisualTreeHelper.GetChild(dep, 0).[GetType]().ToString().IndexOf("Rectangle") > -1 Then
                    rect = DirectCast(VisualTreeHelper.GetChild(dep, 0), Rectangle)

                    dep.Children.Remove(rect)
                ElseIf VisualTreeHelper.GetChild(dep, 0).[GetType]().ToString().IndexOf("Label") > -1 Then
                    lbl = DirectCast(VisualTreeHelper.GetChild(dep, 0), Label)

                    dep.Children.Remove(lbl)
                End If
            End While
        End Sub

        Public Sub LiveLine_DTviewer_Eventmousemove(sender As Object, e As MouseEventArgs)
            Dim tempsender As Rectangle = DirectCast(sender, Rectangle)
            tempsender.Opacity = 0.7
            Cursor = Cursors.Hand
            Dim templabel As Label
            Dim tempdate As DateTime
            Dim DTduration As Double = 0
            Dim eventno As Integer = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name.ToString()))
            templabel = getMenuItem_Label_fromitemindex(LiveLineDTViewerGraphicsCanvas, -1, "", "DTviewer_TimeLabel")
            tempdate = intermediate.LiveLine_EventStartTimes(eventno)
            DTduration = intermediate.LiveLine_ActualDurationOfEachEvent(eventno)
            templabel.Content = tempdate.ToString("MMM", CultureInfo.InvariantCulture) & " " & tempdate.ToString("dd", CultureInfo.InvariantCulture) & " " & tempdate.ToString("hh: mm tt", CultureInfo.InvariantCulture) & "  [" & Math.Round(DTduration, 1) & " min]"
            templabel.Visibility = Visibility.Visible
            AnimateZoomUIElement(CDbl(templabel.GetValue(Canvas.LeftProperty)), CDbl(tempsender.GetValue(Canvas.LeftProperty)) + tempsender.Width / 2 - templabel.Width / 2, 0.1, Canvas.LeftProperty, templabel)

        End Sub

        Public Sub LiveLine_DTviewer_Eventmouseleave(sender As Object, e As MouseEventArgs)
            Dim tempsender As Rectangle = DirectCast(sender, Rectangle)
            Dim templabel As Label
            Cursor = Cursors.Arrow
            tempsender.Opacity = 1.0
            templabel = getMenuItem_Label_fromitemindex(LiveLineDTViewerGraphicsCanvas, -1, "", "DTviewer_TimeLabel")
            templabel.Visibility = Visibility.Hidden

        End Sub

        Public Sub LiveLine_DTviewer_EventSeleced(sender As Object, e As MouseButtonEventArgs)
            LiveLine_DTviewer_Eventselectionclear()
            LiveLine_TopLoss_CanvasClearSelection()

            Dim tempsender As Rectangle = DirectCast(sender, Rectangle)

            Dim eventno As Integer = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name.ToString()))
            tempsender.StrokeThickness = 1.0
            tempsender.Stroke = Brushes.Black
            LiveLine_DTview_selectedlossname.Content = intermediate.LiveLine_DTviewer_EventNames(eventno).ToString() & ":  " & Math.Round(intermediate.LiveLine_ActualDurationOfEachEvent(eventno), 1) & " min"

            Dim i As Integer
            Dim canvasno As Integer = 0
            For i = 0 To intermediate.LiveLine_TopLosses.Count - 1
                If intermediate.LiveLine_DTviewer_EventNames(eventno).ToString() = intermediate.LiveLine_TopLosses(i).Item1.ToString() Then
                    canvasno = i
                End If
            Next

            getMenuItem_Canvas_fromitemindex(LiveLine_TopLossGraphicsCanvas, -1, "", "LiveLine_TopLossItem" & canvasno).Background = BrushColors.mybrushgray
        End Sub

        Public Sub LiveLine_DTviewer_Eventselectionclear()
            Dim i As Integer
            For i = 0 To intermediate.LiveLine_NumberOfEvents - 1
                getMenuItem_Rectangle_fromitemindex(LiveLineDTViewerGraphicsCanvas, -1, "", "DTviewrect" & i).StrokeThickness = 0
            Next

        End Sub

        Public Sub LiveLine_TimeFrameChanged(sender As Object, e As MouseButtonEventArgs)
            LiveLine_RefreshTimeframeSelection()
            Dim tempsender As Label = DirectCast(sender, Label)
            tempsender.Background = mybrushSelectedCriteria

            If tempsender.Content.ToString().Contains("7 days") Then
                intermediate.LiveLine_setDTviewerTimeFrame(7)
                LiveLineTrends_TimeFrame = 7
                LiveLine_Trend_GenerateChart(DowntimeMetrics.PR, 7)
            End If
            If tempsender.Content.ToString().Contains("30 days") Then
                intermediate.LiveLine_setDTviewerTimeFrame(30)
                LiveLineTrends_TimeFrame = 30
                LiveLine_Trend_GenerateChart(DowntimeMetrics.PR, 30)
            End If
            If tempsender.Content.ToString().Contains("24 hours") Then
                intermediate.LiveLine_setDTviewerTimeFrame(1)
                LiveLineTrends_TimeFrame = 1
                LiveLine_Trend_GenerateChart(DowntimeMetrics.PR, 24)
            End If

            LiveLine_UpdateDTViewerLabels()

            LiveLine_GenerateDTViewer()
            LiveLine_TopLoss_GenerateItems()
            LiveLine_TopPlanned_GenerateItems()
            LiveLine_TopDelta_GenerateItems()


        End Sub
        Public Sub LiveLine_RefreshTimeframeSelection()
            LiveLine_Last24hours.Background = BrushColors.mybrushdarkgray
            LiveLine_Last7days.Background = BrushColors.mybrushdarkgray
            LiveLine_Last30days.Background = BrushColors.mybrushdarkgray

        End Sub

        Public Sub LiveLine_Trend_SetCharttoOEE(sender As Object, e As MouseButtonEventArgs)
            LiveLine_Trend_GenerateChart(DowntimeMetrics.PR, LiveLineTrends_TimeFrame)
            LiveLine_Trends_OEE.Background = mybrushSelectedCriteria
            LiveLine_Trends_Stops.Background = BrushColors.mybrushdarkgray
        End Sub
        Public Sub LiveLine_Trend_SetCharttoStops(sender As Object, e As MouseButtonEventArgs)
            LiveLine_Trend_GenerateChart(DowntimeMetrics.Stops, LiveLineTrends_TimeFrame)
            LiveLine_Trends_Stops.Background = mybrushSelectedCriteria
            LiveLine_Trends_OEE.Background = BrushColors.mybrushdarkgray
        End Sub


        Public Sub LiveLine_Trend_GenerateChart(losstype As DowntimeMetrics, noofbars As Integer)
            If noofbars = 1 Then
                noofbars = 24
            End If

            LiveLine_Trends_ClearChart()
            Dim dep As Canvas = LiveLineTrendGraphicCanvas
            Dim temprect As Rectangle
            Dim maxbarheight As Double = dep.Height
            Dim actualbarheight As Double = 0
            Dim actuallossvalue As Double = 0
            Dim maxlossvalue As Double = 0
            Dim gapbetweenbars As Double = 0
            Dim actualfieldlabel As String = ""
            Dim tempdate As DateTime
            Dim widthofbar As Double = (dep.Width - (gapbetweenbars * noofbars)) / noofbars
            Dim i As Integer

            If losstype = DowntimeMetrics.Stops Then
                maxlossvalue = intermediate.LiveLine_Trends_MaxStops
            ElseIf losstype = DowntimeMetrics.PR Then
                maxlossvalue = Math.Round(100 * intermediate.LiveLine_Trends_MaxOEE, 1)
            End If

            For i = 0 To noofbars - 1
                If losstype = DowntimeMetrics.Stops Then

                    actuallossvalue = intermediate.LiveLine_TrendsData(i).Item3
                ElseIf losstype = DowntimeMetrics.PR Then
                    actuallossvalue = Math.Round(100 * intermediate.LiveLine_TrendsData(i).Item2, 1)
                End If

                tempdate = intermediate.LiveLine_TrendsData(i).Item1
                Select Case LiveLineTrends_TimeFrame
                    Case 7
                        actualfieldlabel = tempdate.ToString("MMM", CultureInfo.InvariantCulture) & " " & tempdate.ToString("dd", CultureInfo.InvariantCulture)
                        Exit Select
                    Case 1
                        actualfieldlabel = tempdate.ToString("HH", CultureInfo.InvariantCulture) & ":00"
                        Exit Select
                    Case 30
                        actualfieldlabel = tempdate.ToString("MMM", CultureInfo.InvariantCulture) & " " & tempdate.ToString("dd", CultureInfo.InvariantCulture)
                        Exit Select
                    Case Else

                        actualfieldlabel = tempdate.ToString("MMM", CultureInfo.InvariantCulture) & " " & tempdate.ToString("dd", CultureInfo.InvariantCulture)
                        Exit Select

                End Select


                actualbarheight = (actuallossvalue / maxlossvalue) * maxbarheight

                GenerateRectangleUI(dep, "LiveLine_TrendBar" & i, 0.8 * actualbarheight, widthofbar, widthofbar + (i * (widthofbar + gapbetweenbars)), 0.9 * dep.Height,
                    mybrushSelectedCriteria, Brushes.White, 0.5, AddressOf LiveLine_Trends_Clicked, AddressOf Generalmousemove, AddressOf Generalmouseleave,
                    180, -1, 1, "")
                GenerateLabelUI(dep, "LiveLine_TrendDateLabel" & i, 10, widthofbar, widthofbar + ((i - 1) * (widthofbar + gapbetweenbars)), 0.9 * dep.Height,
                    Nothing, BrushColors.mybrushfontgray, 7, Nothing, Nothing, Nothing,
                    -1, actualfieldlabel)
                temprect = getMenuItem_Rectangle_fromitemindex(dep, -1, "", "LiveLine_TrendBar" & i)

                GenerateLabelUI(dep, "LiveLine_DataLabel" & i, 10, widthofbar, widthofbar + ((i - 1) * (widthofbar + gapbetweenbars)), dep.Height - temprect.Height - 20,
                    Nothing, BrushColors.mybrushfontgray, 7, Nothing, Nothing, Nothing,
                    -1, Math.Round(actuallossvalue, 1).ToString())
            Next
        End Sub

        Public Sub LiveLine_Trends_ClearChart()
            Dim dep As Canvas = LiveLineTrendGraphicCanvas
            Dim rect As Rectangle
            Dim lbl As Label

            While VisualTreeHelper.GetChildrenCount(dep) <> 0
                If VisualTreeHelper.GetChild(dep, 0).[GetType]().ToString().IndexOf("Rectangle") > -1 Then
                    rect = DirectCast(VisualTreeHelper.GetChild(dep, 0), Rectangle)

                    dep.Children.Remove(rect)
                ElseIf VisualTreeHelper.GetChild(dep, 0).[GetType]().ToString().IndexOf("Label") > -1 Then
                    lbl = DirectCast(VisualTreeHelper.GetChild(dep, 0), Label)

                    dep.Children.Remove(lbl)
                End If
            End While
        End Sub

        Public Sub LiveLine_Trends_Clicked(sender As Object, e As MouseButtonEventArgs)
            Dim temprect As Rectangle = DirectCast(sender, Rectangle)
            Dim dep As Canvas = LiveLineTrendGraphicCanvas
            LiveLine_LocateTimeHighlightRectangle(CDbl(temprect.GetValue(Canvas.LeftProperty)), temprect.Width)
        End Sub

        Public Sub LiveLine_TopLoss_GenerateItems()
            LiveLine_TopCardsGraphicCanvas_Clear(LiveLine_TopLossGraphicsCanvas)
            Dim rnd As New Random()
            Dim dep As Canvas = LiveLine_TopLossGraphicsCanvas
            Dim tempcanvas As Canvas
            Dim templabel As Label
            Dim itemheight As Double = 30
            Dim itemverticalgap As Double = 5
            Dim gapbetweenlabelandbar As Double = 5
            Dim datalabelwidth As Double = 30
            Dim stopslabelwidth As Double = 50
            Dim lossnamelabelwidth As Double = 150
            Dim itemwidth As Double = 360
            Dim actuallossvalue As Double = 0
            Dim maxlossvalue As Double = Math.Round(100 * intermediate.LiveLine_TopLoss_MaxLossValue, 1)
            Dim maxlossbarwidth As Double = 100
            Dim actualbarwidth As Double = 0
            Dim i As Integer

            If maxlossvalue = 0 Then
                maxlossvalue = 1
            End If

            For i = 0 To intermediate.LiveLine_TopLosses.Count - 1
                GenerateCanvasUI(dep, "LiveLine_TopLossItem" & i, itemheight, itemwidth, 0, itemverticalgap + (i * itemheight))
                tempcanvas = getMenuItem_Canvas_fromitemindex(dep, -1, "", "LiveLine_TopLossItem" & i)
                AddHandler tempcanvas.MouseDown, AddressOf LiveLine_TopLoss_CanvasMouseDown
                AddHandler tempcanvas.MouseMove, AddressOf Generalmousemove
                AddHandler tempcanvas.MouseLeave, AddressOf Generalmouseleave
                GenerateLabelUI(tempcanvas, "LiveLine_TopLoss_lossnameLabel" & i, itemheight, lossnamelabelwidth, 0, 0,
                    Nothing, BrushColors.mybrushfontgray, 12, Nothing, Nothing, Nothing,
                    -1, intermediate.LiveLine_TopLosses(i).Item1, True)
                templabel = getMenuItem_Label_fromitemindex(tempcanvas, -1, "", "LiveLine_TopLoss_lossnameLabel" & i)
                templabel.ToolTip = templabel.Content.ToString()

                actuallossvalue = Math.Round(100 * intermediate.LiveLine_TopLosses(i).Item2, 1)
                actualbarwidth = (actuallossvalue / maxlossvalue) * maxlossbarwidth
                GenerateRectangleUI(tempcanvas, "LiveLine_TopLoss_bar" & i, 0.5 * itemheight, actualbarwidth, lossnamelabelwidth + gapbetweenlabelandbar, 0.25 * itemheight,
                    mybrushSelectedCriteria, Nothing, 0, Nothing, Nothing, Nothing)
                GenerateLabelUI(tempcanvas, "LiveLine_TopLoss_datalabel" & i, itemheight, datalabelwidth, lossnamelabelwidth + gapbetweenlabelandbar + actualbarwidth + 2, 0,
                    Nothing, BrushColors.mybrushfontgray, 9, Nothing, Nothing, Nothing,
                    -1, actuallossvalue & "%", True)
                GenerateLabelUI(tempcanvas, "LiveLine_TopLoss_stopslabel" & i, itemheight, stopslabelwidth, lossnamelabelwidth + gapbetweenlabelandbar + maxlossbarwidth + datalabelwidth + 15, 0,
                    Nothing, BrushColors.mybrushfontgray, 11, Nothing, Nothing, Nothing,
                    -1, intermediate.LiveLine_TopLosses(i).Item3 & " stops", True)

                dep.Height = (itemheight - itemverticalgap) + (i * (itemheight + itemverticalgap))
                AnimateZoomUIElement(0.5 * actualbarwidth, actualbarwidth, 0.1, WidthProperty, getMenuItem_Rectangle_fromitemindex(tempcanvas, -1, "", "LiveLine_TopLoss_bar" & i))
                AnimateZoomUIElement(0.2, 1.0, 0.1, OpacityProperty, tempcanvas)
                System.Windows.Forms.Application.DoEvents()

                Thread.Sleep(10)
            Next


        End Sub

        Public Sub LiveLine_TopLoss_CanvasMouseDown(sender As Object, e As MouseButtonEventArgs)
            LiveLine_TopLoss_CanvasClearSelection()
            Dim tempsender As Canvas = DirectCast(sender, Canvas)
            Dim toplossno As Integer = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name.ToString()))
            Dim lossnamesearch As String = intermediate.LiveLine_TopLosses(toplossno).Item1
            tempsender.Background = BrushColors.mybrushgray
            Dim temprect As Rectangle
            Dim i As Integer
            LiveLine_DTviewer_Eventselectionclear()
            For i = 0 To intermediate.LiveLine_NumberOfEvents - 1
                If intermediate.LiveLine_DTviewer_EventNames(i) = lossnamesearch Then
                    temprect = getMenuItem_Rectangle_fromitemindex(LiveLineDTViewerGraphicsCanvas, -1, "", "DTviewrect" & i)
                    temprect.StrokeThickness = 1.0
                    temprect.Stroke = Brushes.Black
                    LiveLine_DTview_selectedlossname.Content = lossnamesearch
                End If
            Next

        End Sub

        Public Sub LiveLine_TopLoss_CanvasClearSelection()
            Dim tempcanvas As Canvas
            Dim i As Integer

            For i = 0 To intermediate.LiveLine_TopLosses.Count - 1
                tempcanvas = getMenuItem_Canvas_fromitemindex(LiveLine_TopLossGraphicsCanvas, -1, "", "LiveLine_TopLossItem" & i)
                tempcanvas.Background = Nothing
            Next
        End Sub

        Public Sub LiveLine_TopCardsGraphicCanvas_Clear(dep As Canvas)

            Dim cvs As Canvas


            While VisualTreeHelper.GetChildrenCount(dep) <> 0
                If VisualTreeHelper.GetChild(dep, 0).[GetType]().ToString().IndexOf("Canvas") > -1 Then
                    cvs = DirectCast(VisualTreeHelper.GetChild(dep, 0), Canvas)

                    dep.Children.Remove(cvs)
                End If
            End While

        End Sub

        'Top Planned

        Public Sub LiveLine_TopPlanned_GenerateItems()
            LiveLine_TopCardsGraphicCanvas_Clear(LiveLine_TopLosChangeoverraphicsCanvas)
            Dim rnd As New Random()
            Dim dep As Canvas = LiveLine_TopLosChangeoverraphicsCanvas
            Dim tempcanvas As Canvas
            Dim templabel As Label
            Dim itemheight As Double = 30
            Dim itemverticalgap As Double = 5
            Dim gapbetweenlabelandbar As Double = 5
            Dim datalabelwidth As Double = 40
            Dim stopslabelwidth As Double = 50
            Dim lossnamelabelwidth As Double = 150
            Dim itemwidth As Double = 360
            Dim actuallossvalue As Double = 0
            Dim maxlossvalue As Double = intermediate.LiveLine_TopLoss_MaxValue_Planned
            Dim maxlossbarwidth As Double = 100
            Dim actualbarwidth As Double = 0
            Dim i As Integer

            If maxlossvalue = 0 Then
                maxlossvalue = 1
            End If

            For i = 0 To intermediate.LiveLine_Planned.Count - 1
                GenerateCanvasUI(dep, "LiveLine_TopPlannedItem" & i, itemheight, itemwidth, 0, itemverticalgap + (i * itemheight))
                tempcanvas = getMenuItem_Canvas_fromitemindex(dep, -1, "", "LiveLine_TopPlannedItem" & i)
                AddHandler tempcanvas.MouseMove, AddressOf Generalmousemove
                AddHandler tempcanvas.MouseLeave, AddressOf Generalmouseleave
                GenerateLabelUI(tempcanvas, "LiveLine_TopPlanned_lossnameLabel" & i, itemheight, lossnamelabelwidth, 0, 0,
                    Nothing, BrushColors.mybrushfontgray, 12, Nothing, Nothing, Nothing,
                    -1, intermediate.LiveLine_Planned(i).Item1, True)
                templabel = getMenuItem_Label_fromitemindex(tempcanvas, -1, "", "LiveLine_TopPlanned_lossnameLabel" & i)
                templabel.ToolTip = templabel.Content.ToString()

                actuallossvalue = Math.Round(intermediate.LiveLine_Planned(i).Item2)
                actualbarwidth = (actuallossvalue / maxlossvalue) * maxlossbarwidth
                GenerateRectangleUI(tempcanvas, "LiveLine_TopPlanned_bar" & i, 0.5 * itemheight, actualbarwidth, lossnamelabelwidth + gapbetweenlabelandbar, 0.25 * itemheight,
                    mybrushSelectedCriteria, Nothing, 0, Nothing, Nothing, Nothing)
                GenerateLabelUI(tempcanvas, "LiveLine_TopPlanned_datalabel" & i, itemheight, datalabelwidth, lossnamelabelwidth + gapbetweenlabelandbar + actualbarwidth + 2, 0,
                    Nothing, BrushColors.mybrushfontgray, 9, Nothing, Nothing, Nothing,
                    -1, actuallossvalue & "min", True)
                GenerateLabelUI(tempcanvas, "LiveLine_TopPlanned_stopslabel" & i, itemheight, stopslabelwidth, lossnamelabelwidth + gapbetweenlabelandbar + maxlossbarwidth + datalabelwidth + 10, 0,
                    Nothing, BrushColors.mybrushfontgray, 11, Nothing, Nothing, Nothing,
                    -1, Math.Round(intermediate.LiveLine_Planned(i).Item3 * 100, 1) & "% DT", True)

                dep.Height = (itemheight - itemverticalgap) + (i * (itemheight + itemverticalgap))
                AnimateZoomUIElement(0.5 * actualbarwidth, actualbarwidth, 0.1, WidthProperty, getMenuItem_Rectangle_fromitemindex(tempcanvas, -1, "", "LiveLine_TopPlanned_bar" & i))
                AnimateZoomUIElement(0.2, 1.0, 0.1, OpacityProperty, tempcanvas)
                System.Windows.Forms.Application.DoEvents()

                Thread.Sleep(20)
            Next
        End Sub

        Public Sub LiveLine_TopDelta_GenerateItems()
            LiveLine_TopCardsGraphicCanvas_Clear(LiveLine_TopDeltaGraphicsCanvas)
            Dim rnd As New Random()
            Dim dep As Canvas = LiveLine_TopDeltaGraphicsCanvas
            Dim tempcanvas As Canvas
            Dim templabel As Label
            Dim itemheight As Double = 30
            Dim itemverticalgap As Double = 5
            Dim gapbetweenlabelandDT As Double = 10
            Dim gapbetweenDTandstops As Double = 30
            Dim datalabelwidth As Double = 55
            Dim stopslabelwidth As Double = 50
            Dim lossnamelabelwidth As Double = 150
            Dim itemwidth As Double = 360
            Dim actuallossvalue As Double = 0
            Dim deltaiconwidth As Double = 10
            Dim deltaiconheight As Double = 0.5 * itemheight


            Dim actualbarwidth As Double = 0
            Dim i As Integer
            Dim deltaimagefilename As String = ""


            For i = 0 To intermediate.LiveLine_BiggestChanges.Count - 1
                GenerateCanvasUI(dep, "LiveLine_TopDeltaItem" & i, itemheight, itemwidth, 0, itemverticalgap + (i * itemheight))
                tempcanvas = getMenuItem_Canvas_fromitemindex(dep, -1, "", "LiveLine_TopDeltaItem" & i)

                AddHandler tempcanvas.MouseMove, AddressOf Generalmousemove
                AddHandler tempcanvas.MouseLeave, AddressOf Generalmouseleave
                GenerateLabelUI(tempcanvas, "LiveLine_TopDelta_lossnameLabel" & i, itemheight, lossnamelabelwidth, 0, 0,
                    Nothing, BrushColors.mybrushfontgray, 12, Nothing, Nothing, Nothing,
                    -1, intermediate.LiveLine_BiggestChanges(i).Item1, True)
                templabel = getMenuItem_Label_fromitemindex(tempcanvas, -1, "", "LiveLine_TopDelta_lossnameLabel" & i)
                templabel.ToolTip = templabel.Content.ToString()


                'DT
                actuallossvalue = Math.Round(100 * intermediate.LiveLine_BiggestChanges(i).Item2, 1)

                If actuallossvalue >= 0 Then
                    deltaimagefilename = "UpDelta"
                Else
                    deltaimagefilename = "DownDelta"
                End If
                GenerateImageUI(tempcanvas, "LiveLine_TopDelta_DTdeltaicon" & i, deltaiconheight, deltaiconwidth, lossnamelabelwidth + gapbetweenlabelandDT, itemheight / 2 - deltaiconheight / 2,
                    (Convert.ToString(AppDomain.CurrentDomain.BaseDirectory & "\") & deltaimagefilename) & ".png", Nothing, Nothing, Nothing)
                GenerateLabelUI(tempcanvas, "LiveLine_TopDelta_DTlabel" & i, itemheight, datalabelwidth, lossnamelabelwidth + gapbetweenlabelandDT + deltaiconwidth + 2, 0,
                    Nothing, BrushColors.mybrushfontgray, 9, Nothing, Nothing, Nothing,
                    -1, actuallossvalue & "% DT", True)


                'Stops
                actuallossvalue = intermediate.LiveLine_BiggestChanges(i).Item3

                If actuallossvalue >= 0 Then
                    deltaimagefilename = "UpDelta"
                Else
                    deltaimagefilename = "DownDelta"
                End If
                GenerateImageUI(tempcanvas, "LiveLine_TopDelta_Stopsdeltaicon" & i, deltaiconheight, deltaiconwidth, lossnamelabelwidth + gapbetweenlabelandDT + deltaiconwidth + 2 + datalabelwidth + gapbetweenDTandstops, itemheight / 2 - deltaiconheight / 2,
                    (Convert.ToString(AppDomain.CurrentDomain.BaseDirectory & "\") & deltaimagefilename) & ".png", Nothing, Nothing, Nothing)
                GenerateLabelUI(tempcanvas, "LiveLine_TopDelta_stopslabel" & i, itemheight, stopslabelwidth, lossnamelabelwidth + gapbetweenlabelandDT + deltaiconwidth + 2 + datalabelwidth + gapbetweenDTandstops + deltaiconwidth + 2, 0,
                    Nothing, BrushColors.mybrushfontgray, 11, Nothing, Nothing, Nothing,
                    -1, intermediate.LiveLine_BiggestChanges(i).Item3 & " stops", True)

                dep.Height = (itemheight - itemverticalgap) + (i * (itemheight + itemverticalgap))
                AnimateZoomUIElement(0.2, 1.0, 0.1, OpacityProperty, tempcanvas)
                System.Windows.Forms.Application.DoEvents()

                Thread.Sleep(10)
            Next
        End Sub
#End Region

#Region "Crazy UI Stuff"

        Public Sub GenerateCanvasUI(dep As Canvas, canvasname As String, height As Double, width As Double, PosLeft As Double, PosTop As Double,
    Optional Zindex As Integer = -1, Optional canvascolor As SolidColorBrush = Nothing)
            Dim c As Canvas
            c = New Canvas()
            dep.Children.Add(c)
            'dep.Children.Add(c);
            c.Height = height
            c.Width = width
            c.Name = canvasname
            Canvas.SetLeft(c, PosLeft)
            Canvas.SetTop(c, PosTop)
            If Zindex <> -1 Then
                Canvas.SetZIndex(c, Zindex)
            End If
            If canvascolor IsNot Nothing Then

                c.Background = canvascolor
            End If
        End Sub


        Public Sub GenerateLabelUI(dep As Canvas, labelname As String, height As Double, width As Double, PosLeft As Double, PosTop As Double,
        labelfillcolor As SolidColorBrush, labelfontcolor As SolidColorBrush, fontsize As Double, mousedownact As MouseButtonEventHandler, mousemoveact As MouseEventHandler, mouseleaveact As MouseEventHandler,
        Zindex As Integer, Optional content As String = "", Optional isleftaligned As Boolean = False)

            Dim l As Label
            l = New Label()
            dep.Children.Add(l)
            l.Height = height
            l.Width = width
            l.Name = labelname
            Canvas.SetLeft(l, PosLeft)
            Canvas.SetTop(l, PosTop)
            l.Background = labelfillcolor
            l.Foreground = labelfontcolor
            l.FontSize = fontsize
            l.Cursor = Cursors.Hand
            l.Padding = New Thickness(0.5, 0.5, 0.5, 0.5)
            l.VerticalContentAlignment = VerticalAlignment.Center
            If isleftaligned = False Then
                l.HorizontalContentAlignment = HorizontalAlignment.Center
            Else

                l.HorizontalContentAlignment = HorizontalAlignment.Left
            End If
            If Zindex <> -1 Then
                Canvas.SetZIndex(l, Zindex)
            End If

            l.Content = content

            If mousedownact IsNot Nothing Then

                AddHandler l.MouseDown, Sub(sender, e)
                                            mousedownact(sender, e)
                                        End Sub
            End If
            If mousemoveact IsNot Nothing Then

                AddHandler l.MouseMove, Sub(sender, e)
                                            mousemoveact(sender, e)
                                        End Sub
            End If
            If mouseleaveact IsNot Nothing Then

                AddHandler l.MouseLeave, Sub(sender, e)
                                             mouseleaveact(sender, e)
                                         End Sub

                '  AddHandler button.Click, Sub(sender, e)
                '                    MessageBox.Show("Clicked!")
                '                   Dim retval = SomeFunction(value)
                '                  '' etc...
                '             End Sub

            End If


        End Sub


        Public Sub GenerateImageUI(dep As Canvas, Imagename As String, height As Double, width As Double, PosLeft As Double, PosTop As Double,
            source As String, mousedownact As MouseButtonEventHandler, mousemoveact As MouseEventHandler, mouseleaveact As MouseEventHandler, Optional tooltip As String = "", Optional Zindex As Integer = -1)
            Dim I As Image
            I = New Image()
            dep.Children.Add(I)
            I.Height = height
            I.Width = width
            I.Name = Imagename
            Canvas.SetLeft(I, PosLeft)
            Canvas.SetTop(I, PosTop)
            Try
                I.Source = New BitmapImage(New Uri(source))
            Catch
                Dim ixyz As Integer = 0
            End Try

            I.Cursor = Cursors.Hand
            If mousedownact IsNot Nothing Then
                AddHandler I.MouseDown, Sub(sender, e)
                                            mousedownact(sender, e)
                                        End Sub
            End If
            If mousemoveact IsNot Nothing Then

                AddHandler I.MouseMove, Sub(sender, e)
                                            mousemoveact(sender, e)
                                        End Sub
            End If
            If mouseleaveact IsNot Nothing Then

                AddHandler I.MouseLeave, Sub(sender, e)
                                             mouseleaveact(sender, e)
                                         End Sub
            End If

            If tooltip <> "" Then
                I.ToolTip = tooltip
            End If

            If Zindex <> -1 Then
                Canvas.SetZIndex(I, Zindex)
            End If


        End Sub

        Public Sub GenerateRectangleUI(dep As Canvas, rectanglename As String, height As Double, width As Double, PosLeft As Double, PosTop As Double,
            rectcolor As SolidColorBrush, rectborder As SolidColorBrush, strokethickness As Double, mousedownact As MouseButtonEventHandler, mousemoveact As MouseEventHandler, mouseleaveact As MouseEventHandler,
            Optional transformoriginangle As Double = 0, Optional Zindex As Integer = -1, Optional opacity As Double = 1.0, Optional tooltip As String = "", Optional transformmyscale As ScaleTransform = Nothing)

            Dim r As Rectangle
            r = New Rectangle()
            dep.Children.Add(r)
            r.Height = height
            r.Width = width
            r.Name = rectanglename
            Canvas.SetLeft(r, PosLeft)
            Canvas.SetTop(r, PosTop)
            If rectcolor IsNot Nothing Then
                r.Fill = rectcolor
            End If

            r.Stroke = rectborder
            r.StrokeThickness = strokethickness
            r.Opacity = opacity
            Dim myRotateTransform = New RotateTransform()
            myRotateTransform.Angle = transformoriginangle


            If transformmyscale IsNot Nothing Then


                Dim trGrp As TransformGroup
                Dim trRot As RotateTransform
                Dim trScl As ScaleTransform

                myRotateTransform.CenterX = 0.5
                myRotateTransform.CenterY = 0.5

                trScl = transformmyscale
                trRot = myRotateTransform

                trGrp = New TransformGroup()
                trGrp.Children.Add(trRot)
                trGrp.Children.Add(trScl)

                r.RenderTransform = trGrp
            Else
                r.RenderTransform = myRotateTransform
            End If


            If Zindex <> -1 Then
                Canvas.SetZIndex(r, Zindex)
            End If

            If mousedownact IsNot Nothing Then
                AddHandler r.MouseDown, Sub(sender, e)
                                            mousedownact(sender, e)
                                        End Sub
            End If
            If mousemoveact IsNot Nothing Then

                AddHandler r.MouseMove, Sub(sender, e)
                                            mousemoveact(sender, e)
                                        End Sub
            End If
            If mouseleaveact IsNot Nothing Then

                AddHandler r.MouseLeave, Sub(sender, e)
                                             mouseleaveact(sender, e)
                                         End Sub
            End If
            If tooltip <> "" Then
                r.ToolTip = tooltip
            End If


        End Sub

        Public Sub Generalmousemove(sender As Object, e As MouseEventArgs)
            Cursor = Cursors.Hand
            If sender.[GetType]().ToString().IndexOf("Label") > -1 Then
                Dim tempsender As Label = DirectCast(sender, Label)
                tempsender.Opacity = 0.8
            ElseIf sender.[GetType]().ToString().IndexOf("Image") > -1 Then
                Dim tempsender As Image = DirectCast(sender, Image)
                tempsender.Opacity = 0.8
            ElseIf sender.[GetType]().ToString().IndexOf("Rectangle") > -1 Then
                Dim tempsender As Rectangle = DirectCast(sender, Rectangle)
                tempsender.Opacity = 0.8
            ElseIf sender.[GetType]().ToString().IndexOf("Canvas") > -1 Then
                Dim tempsender As Canvas = DirectCast(sender, Canvas)
                tempsender.Opacity = 0.8
            End If
        End Sub
        Public Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
            Cursor = Cursors.Arrow
            If sender.[GetType]().ToString().IndexOf("Label") > -1 Then
                Dim tempsender As Label = DirectCast(sender, Label)
                tempsender.Opacity = 1.0
            ElseIf sender.[GetType]().ToString().IndexOf("Image") > -1 Then
                Dim tempsender As Image = DirectCast(sender, Image)
                tempsender.Opacity = 1.0
            ElseIf sender.[GetType]().ToString().IndexOf("Rectangle") > -1 Then
                Dim tempsender As Rectangle = DirectCast(sender, Rectangle)
                tempsender.Opacity = 1.0
            ElseIf sender.[GetType]().ToString().IndexOf("Canvas") > -1 Then
                Dim tempsender As Canvas = DirectCast(sender, Canvas)
                tempsender.Opacity = 1.0
            End If
        End Sub


        Public Sub AnimateZoomUIElement(from As Double, [to] As Double, durn As Double, depprop As DependencyProperty, AnimatedObject As UIElement)
            ' Standard animation function
            Dim da = New DoubleAnimation()
            ' da will contain the characteristics of the animation
            da.From = from
            ' position, where it starts 
            da.[To] = [to]
            ' position, where it ends
            da.Duration = New Duration(TimeSpan.FromSeconds(durn))
            ' how long animation lasts
            AnimatedObject.BeginAnimation(depprop, da)
            ' Animate object is the subject we are playing with. And Depprop determines what type of UI element it is (rectangle, label, control.. etc)
        End Sub
        Public Function getMenuItem_Label_fromitemindex(dep As DependencyObject, Optional menuitemindex As Integer = -1, Optional stringitemindex As String = "", Optional exactstring As String = "") As Label
            Dim sender As Label = Nothing
            Dim lbl As Label
            Dim j As Integer
            For j = 0 To VisualTreeHelper.GetChildrenCount(dep) - 1
                If VisualTreeHelper.GetChild(dep, j).[GetType]().ToString().IndexOf("Label") > -1 Then
                    lbl = DirectCast(VisualTreeHelper.GetChild(dep, j), Label)

                    If stringitemindex = "" AndAlso menuitemindex <> -1 Then
                        If lbl.Name.IndexOf(menuitemindex.ToString()) > -1 Then
                            sender = lbl
                        End If
                    Else
                        If exactstring = "" Then
                            If lbl.Name.IndexOf(stringitemindex) > -1 Then
                                sender = lbl
                            End If
                        Else
                            If lbl.Name.ToString() = exactstring Then
                                sender = lbl

                            End If

                        End If
                    End If

                End If
            Next

            Return sender
        End Function
        Public Function getMenuItem_Canvas_fromitemindex(dep As DependencyObject, Optional menuitemindex As Integer = -1, Optional stringitemindex As String = "", Optional exactstring As String = "") As Canvas
            Dim sender As Canvas = Nothing
            Dim cnv As Canvas
            Dim j As Integer
            For j = 0 To VisualTreeHelper.GetChildrenCount(dep) - 1
                If VisualTreeHelper.GetChild(dep, j).[GetType]().ToString().IndexOf("Canvas") > -1 Then
                    cnv = DirectCast(VisualTreeHelper.GetChild(dep, j), Canvas)

                    If stringitemindex = "" AndAlso menuitemindex <> -1 Then
                        If cnv.Name.IndexOf(menuitemindex.ToString()) > -1 Then
                            sender = cnv
                        End If
                    Else
                        If exactstring = "" Then
                            If cnv.Name.IndexOf(stringitemindex) > -1 Then
                                sender = cnv
                            End If
                        Else
                            If cnv.Name.ToString() = exactstring Then
                                sender = cnv

                            End If

                        End If
                    End If

                End If
            Next

            Return sender
        End Function
        Public Function getMenuItem_Rectangle_fromitemindex(dep As DependencyObject, Optional menuitemindex As Integer = -1, Optional stringitemindex As String = "", Optional exactstring As String = "") As Rectangle
            Dim sender As Rectangle = Nothing
            Dim rect As Rectangle
            Dim j As Integer
            For j = 0 To VisualTreeHelper.GetChildrenCount(dep) - 1
                If VisualTreeHelper.GetChild(dep, j).[GetType]().ToString().IndexOf("Rectangle") > -1 Then
                    rect = DirectCast(VisualTreeHelper.GetChild(dep, j), Rectangle)

                    If stringitemindex = "" AndAlso menuitemindex <> -1 Then
                        If rect.Name.IndexOf(menuitemindex.ToString()) > -1 Then
                            sender = rect
                        End If
                    Else
                        If exactstring = "" Then
                            If rect.Name.IndexOf(stringitemindex) > -1 Then
                                sender = rect
                            End If
                        Else
                            If rect.Name.ToString() = exactstring Then
                                sender = rect

                            End If

                        End If
                    End If

                End If
            Next

            Return sender
        End Function


#End Region

    End Class


End Namespace