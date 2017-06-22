Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class Window_StopsWatch_Daily
#Region "Variables"
    Private selectedshift As String
    Private tempsender As Object
    Private Const hourbarMAXsize = 60
    Private Const availabilitybarMAXsize = 200
    Dim stopswatchRAWDATA As prStoryStopsWatch_24

    Dim myBrushAquaMarineX As New SolidColorBrush(Colors.Aquamarine)
    Dim myBrushWhiteX As New SolidColorBrush(Colors.White)
    Dim mybrushred As New SolidColorBrush(Colors.OrangeRed)
    Dim mybrushgreen As New SolidColorBrush(Colors.Green)
    Dim mybrushlightgray As New SolidColorBrush(Color.FromRgb(200, 200, 200))

    Dim selecteddate_stopswatch As Date
    Dim newSelectedDate As Date
    Public linearviewcanvaslocation As New Thickness
    Private prstoryReport As prStoryMainPageReport
    Private SelectedFailureMode As String
    Public linearhourbarMaxSize = 51
#End Region

    Public Sub New(ByVal storyReport As prStoryMainPageReport, ByVal initFmode As String)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        prstoryReport = storyReport
        SelectedFailureMode = initFmode
        linearviewcanvaslocation = StopsWatchLinearUI.Margin
    End Sub

#Region "Raw Data Display"
    Dim _ActiveDataCollection As New ObservableCollection(Of DowntimeEvent)()

    'sorting our listview
    Private _lastHeaderClicked_ActiveData As GridViewColumnHeader = Nothing
    Private _lastDirection_ActiveData As ListSortDirection = ListSortDirection.Ascending

    'properties
    Public ReadOnly Property ActiveDataCollection() As ObservableCollection(Of DowntimeEvent)
        Get
            Return _ActiveDataCollection
        End Get
    End Property

    Private Sub displaySelectedDatesRawData(startRawDataDate As Date) ', endRawDataDate As Date)
        Dim startIndex As Integer, endIndex As Integer
        _ActiveDataCollection.Clear()
        With prstoryReport.MainLEDSReport.DT_Report.rawDTdata
            startIndex = .rawConstraintData.IndexOf(New DowntimeEvent(startRawDataDate))
            If startIndex > -1 Then
                endIndex = .rawConstraintData.IndexOf(New DowntimeEvent(DateAdd(DateInterval.Hour, 1, startRawDataDate)))
                If endIndex > startIndex Then
                    For indexIncrementer As Integer = startIndex To endIndex
                        If startRawDataDate <= .rawConstraintData(indexIncrementer).startTime Then
                            If DateAdd(DateInterval.Hour, 1, startRawDataDate) >= .rawConstraintData(indexIncrementer).startTime Then
                                _ActiveDataCollection.Add(.rawConstraintData(indexIncrementer))
                            End If
                        End If
                    Next
                Else
                    If startRawDataDate <= .rawConstraintData(startIndex).startTime And DateAdd(DateInterval.Hour, 1, startRawDataDate) >= .rawConstraintData(startIndex).startTime Then _ActiveDataCollection.Add(.rawConstraintData(startIndex))
                End If
            End If
        End With
    End Sub

#End Region

    Sub stopswatch_loaded()
        UseTrack_StopsWatchMain = True

        stops_calendar.Visibility = Windows.Visibility.Hidden
        selecteddate_stopswatch = prstoryReport.EndDate '.AddDays(-1)
        SelectedDate.Content = Format(selecteddate_stopswatch, "MMMM dd, yyyy")
        stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch, My.Settings.defaultMappingLevel)

        selectedshift = "D"
        AMbutton.Opacity = 0.6
        AMbutton.Background = New SolidColorBrush(Colors.Aquamarine) 'myBrushAquaMarine
        PMbutton.Opacity = 0.2
        PMbutton.Background = New SolidColorBrush(Colors.AliceBlue) 'myBrushWhite

        

        SetCILiconlabel()

        CreateFailureModeList_Stopswatch() 'prstoryReport)
        ShowClassicView()
        GenerateStopsWatch()
    End Sub
    Private Sub ShowClassicView()
        ClassicViewLabel.Background = New SolidColorBrush(Color.FromRgb(101, 222, 200))
        LinearViewLabel.Background = New SolidColorBrush(Color.FromRgb(255, 255, 255))


        StopsWatchClassicUI.Visibility = Visibility.Visible
        StopsWatchLinearUI.Visibility = Visibility.Hidden
        Show_DayNightButtons()
    End Sub
    Private Sub ShowLinearView()
        LinearViewLabel.Background = New SolidColorBrush(Color.FromRgb(101, 222, 200))
        ClassicViewLabel.Background = New SolidColorBrush(Color.FromRgb(255, 255, 255))

        StopsWatchClassicUI.Visibility = Visibility.Hidden
        StopsWatchLinearUI.Visibility = Visibility.Visible
        Hide_DayNightButtons()
        LinearViewShiftsSetup()
        HideAllIcons_LinearView()
    End Sub
    Private Sub Hide_DayNightButtons()
        DayNight.Visibility = Visibility.Hidden
        DayNight_GraySeperator.Visibility = Visibility.Hidden
        DayNight_WhiteSeperator.Visibility = Visibility.Hidden
        AMbutton.Visibility = Visibility.Hidden
        PMbutton.Visibility = Visibility.Hidden
    End Sub
    Private Sub Show_DayNightButtons()
        DayNight.Visibility = Visibility.Visible
        DayNight_GraySeperator.Visibility = Visibility.Visible
        DayNight_WhiteSeperator.Visibility = Visibility.Visible
        AMbutton.Visibility = Visibility.Visible
        PMbutton.Visibility = Visibility.Visible
    End Sub
    Private Sub ManageLinearView_Axes(maxstops As Double)

        rowAlphaAxisLabel1_linear.Content = Math.Round(maxstops)
        rowBetaAxisLabel1_linear.Content = Math.Round(maxstops)
        rowCharlieAxisLabel1_linear.Content = Math.Round(maxstops)

        rowAlphaAxisLabel3_linear.Content = 0
        rowBetaAxisLabel3_linear.Content = 0
        rowCharlieAxisLabel3_linear.Content = 0

    End Sub
    Private Sub LinearViewShiftsSetup()
        Dim i As Integer
        Dim j As Integer = 1
        Dim MaxHourStops_oftwodays As Double

        HideAllIcons_LinearView()

        Select Case AllProdLines(selectedindexofLine_temp).NumberOfShifts
            Case 1
                HideUIelements("Beta")
                HideUIelements("Charlie")
                stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch, My.Settings.defaultMappingLevel)
                stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)

                For i = AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr To AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr + 12
                    AssignHoursvaluestoLabelNames("Alphahour" & j, i)
                    SetLinearView_BarHeights("Alphahour" & j, stopswatchRAWDATA.getModeStops(i + 1), stopswatchRAWDATA.MaxHourStops, stopswatchRAWDATA.isPRout(i + 1), stopswatchRAWDATA.isCIL(i + 1), stopswatchRAWDATA.isChangeover(i + 1))

                    j = j + 1

                Next



            Case 2
                HideUIelements("Charlie")
                stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
                stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)

                MaxHourStops_oftwodays = stopswatchRAWDATA.MaxHourStops

                stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch, My.Settings.defaultMappingLevel)
                stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)

                MaxHourStops_oftwodays = Math.Max(MaxHourStops_oftwodays, stopswatchRAWDATA.MaxHourStops)

                For i = AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr To AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr + 12
                    AssignHoursvaluestoLabelNames("Alphahour" & j, i)

                    If i > 23 Then

                        AssignHoursvaluestoLabelNames("Alphahour" & j, i - 24)

                        stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
                        stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)
                        SetLinearView_BarHeights("Alphahour" & j, stopswatchRAWDATA.getModeStops(i - 23), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i - 23), stopswatchRAWDATA.isCIL(i - 23), stopswatchRAWDATA.isChangeover(i - 23))
                    Else
                        SetLinearView_BarHeights("Alphahour" & j, stopswatchRAWDATA.getModeStops(i + 1), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i + 1), stopswatchRAWDATA.isCIL(i + 1), stopswatchRAWDATA.isChangeover(i + 1))
                    End If

                    j = j + 1

                Next
                j = 1
                For i = AllProdLines(selectedindexofLine_temp).ShiftStartSecond_Hr To AllProdLines(selectedindexofLine_temp).ShiftStartSecond_Hr + 12

                    AssignHoursvaluestoLabelNames("Betahour" & j, i)

                    If i > 23 Then
                        AssignHoursvaluestoLabelNames("Betahour" & j, i - 24)
                        stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
                        stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)
                        SetLinearView_BarHeights("Betahour" & j, stopswatchRAWDATA.getModeStops(i - 23), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i - 23), stopswatchRAWDATA.isCIL(i - 23), stopswatchRAWDATA.isChangeover(i - 23))
                    Else
                        SetLinearView_BarHeights("Betahour" & j, stopswatchRAWDATA.getModeStops(i + 1), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i + 1), stopswatchRAWDATA.isCIL(i + 1), stopswatchRAWDATA.isChangeover(i + 1))
                    End If

                    j = j + 1



                Next

            Case 3


                StopsWatchLinearUI.Margin = New Thickness(linearviewcanvaslocation.Left + 40, linearviewcanvaslocation.Top, linearviewcanvaslocation.Right, linearviewcanvaslocation.Bottom)

                stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
                stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)

                MaxHourStops_oftwodays = stopswatchRAWDATA.MaxHourStops

                stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch, My.Settings.defaultMappingLevel)
                stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)

                MaxHourStops_oftwodays = Math.Max(MaxHourStops_oftwodays, stopswatchRAWDATA.MaxHourStops)

                For i = AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr To AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr + 8
                    AssignHoursvaluestoLabelNames("Alphahour" & j, i)

                    If i > 23 Then
                        AssignHoursvaluestoLabelNames("Alphahour" & j, i - 24)
                        stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
                        stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)



                        SetLinearView_BarHeights("Alphahour" & j, stopswatchRAWDATA.getModeStops(i - 23), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i - 23), stopswatchRAWDATA.isCIL(i - 23), stopswatchRAWDATA.isChangeover(i - 23))
                    Else
                        SetLinearView_BarHeights("Alphahour" & j, stopswatchRAWDATA.getModeStops(i + 1), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i + 1), stopswatchRAWDATA.isCIL(i + 1), stopswatchRAWDATA.isChangeover(i + 1))
                    End If

                    j = j + 1

                Next
                j = 1
                For i = AllProdLines(selectedindexofLine_temp).ShiftStartSecond_Hr To AllProdLines(selectedindexofLine_temp).ShiftStartSecond_Hr + 8
                    AssignHoursvaluestoLabelNames("Betahour" & j, i)

                    If i > 23 Then
                        AssignHoursvaluestoLabelNames("Betahour" & j, i - 24)
                        stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
                        stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)
                        SetLinearView_BarHeights("Betahour" & j, stopswatchRAWDATA.getModeStops(i - 23), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i - 23), stopswatchRAWDATA.isCIL(i - 23), stopswatchRAWDATA.isChangeover(i - 23))
                    Else
                        SetLinearView_BarHeights("Betahour" & j, stopswatchRAWDATA.getModeStops(i + 1), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i + 1), stopswatchRAWDATA.isCIL(i + 1), stopswatchRAWDATA.isChangeover(i + 1))
                    End If
                    j = j + 1

                Next
                j = 1
                For i = AllProdLines(selectedindexofLine_temp).ShiftStartThird_Hr To AllProdLines(selectedindexofLine_temp).ShiftStartThird_Hr + 8
                    AssignHoursvaluestoLabelNames("Charliehour" & j, i)

                    If i > 23 Then
                        AssignHoursvaluestoLabelNames("Charliehour" & j, i - 24)
                        stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
                        stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)
                        SetLinearView_BarHeights("Charliehour" & j, stopswatchRAWDATA.getModeStops(i - 23), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i - 23), stopswatchRAWDATA.isCIL(i - 23), stopswatchRAWDATA.isChangeover(i - 23))
                    Else
                        SetLinearView_BarHeights("Charliehour" & j, stopswatchRAWDATA.getModeStops(i + 1), MaxHourStops_oftwodays, stopswatchRAWDATA.isPRout(i + 1), stopswatchRAWDATA.isCIL(i + 1), stopswatchRAWDATA.isChangeover(i + 1))
                    End If
                    j = j + 1

                Next

                HideUIelements("Alphahour9")
                HideUIelements("Alphahour10")
                HideUIelements("Alphahour11")
                HideUIelements("Alphahour12")
                HideUIelements("Betahour9")
                HideUIelements("Betahour10")
                HideUIelements("Betahour11")
                HideUIelements("Betahour12")
                HideUIelements("Charliehour9")
                HideUIelements("Charliehour10")
                HideUIelements("Charliehour11")
                HideUIelements("Charliehour12")

        End Select

    End Sub
    Private Sub SetLinearView_BarHeights(barname As String, stopsvalue As Double, maxstopsvalue As Double, isPROut As Boolean, isCIL As Boolean, isCO As Boolean)
        Dim rect As Rectangle

        ManageLinearView_Axes(maxstopsvalue)

        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(Me.StopsWatchLinearUI) - 1
            If Not IsNothing(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Rectangle") > 0 Then
                    rect = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(rect.Name, barname) > 0 Then

                        If maxstopsvalue = 0 Then
                            stopsvalue = 0
                            rect.Height = stopsvalue * linearhourbarMaxSize / maxstopsvalue
                            Exit For
                        End If
                        rect.Height = stopsvalue * linearhourbarMaxSize / maxstopsvalue

                    End If

                End If
            End If

        Next

        'Setting color of label for PRout/in and show/hide CIL and CO icon
        Dim lbl As Label
        Dim img As Image
        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(Me.StopsWatchLinearUI) - 1
            If Not IsNothing(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Label") > 0 Then
                    lbl = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(lbl.Name, barname) > 0 Then

                        If isPROut Then

                            lbl.Foreground = mybrushred
                        Else
                            lbl.Foreground = mybrushgreen

                        End If
                    End If

                End If

                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Image") > 0 Then
                    img = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(img.Name, barname) > 0 And InStr(img.Name, "CIL") > 0 Then

                        If isCIL Then

                            img.Visibility = Visibility.Visible
                        Else
                            img.Visibility = Visibility.Hidden

                        End If
                    End If

                    If InStr(img.Name, barname) > 0 And InStr(img.Name, "CO") > 0 Then

                        If isCO Then

                            img.Visibility = Visibility.Visible
                        Else
                            img.Visibility = Visibility.Hidden


                        End If
                    End If

                End If
            End If
        Next

    End Sub

    Private Sub HideAllIcons_LinearView()

        Dim img As Image

        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(Me.StopsWatchLinearUI) - 1
            If Not IsNothing(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Image") > 0 Then
                    img = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(img.Name, "CIL", vbTextCompare) > 0 Or InStr(img.Name, "CO", vbTextCompare) > 0 Then
                        img.Visibility = Visibility.Hidden
                    End If
                End If
            End If
        Next

    End Sub
    Private Sub AssignHoursvaluestoLabelNames(labelname As String, hourvalue As Integer)
        Dim label As Label

        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(Me.StopsWatchLinearUI) - 1
            If Not IsNothing(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Label") > 0 Then
                    label = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(label.Name, labelname) > 0 Then

                        label.Content = hourvalue
                        If hourvalue < 10 Then label.Content = "0" & hourvalue
                    End If

                End If
                End If

        Next



    End Sub

    Private Sub HideUIelements(uiname As String)

        Dim label As Label
        Dim rectangle As Rectangle
        Dim img As Image
        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(Me.StopsWatchLinearUI) - 1


            If Not IsNothing(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Label") > 0 Then
                    label = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(label.Name, uiname) > 0 Then label.Visibility = Visibility.Hidden


                End If
                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Rectangle") > 0 Then
                    rectangle = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(rectangle.Name, uiname) > 0 Then rectangle.Visibility = Visibility.Hidden


                End If


                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Image") > 0 Then
                    img = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(img.Name, uiname) > 0 Then img.Visibility = Visibility.Hidden


                End If
            End If

        Next

    End Sub


    Private Sub CreateFailureModeList_Stopswatch() 'prstoryreport3 As prStoryMainPageReport)
        Dim tmpdtevent3 As DTevent
        Dim i As Integer

        stopswatch_stopslist_combo.Items.Clear()
        stopswatch_stopslist_combo.SelectedValue = SelectedFailureMode
        For i = 0 To 14
            tmpdtevent3 = prstoryReport.getCardEventInfo(31, i)
            If tmpdtevent3.Name <> " " Then
                stopswatch_stopslist_combo.Items.Add(tmpdtevent3.Name)
            End If

        Next i
        stopswatch_stopslist_combo.Items.Add("All failure modes")
    End Sub
    Private Sub OnStopComboSelectionChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs)
        Dim combo = TryCast(sender, ComboBox)

        If combo.SelectedItem IsNot Nothing Then
            selectedfailuremode = combo.SelectedValue
        End If
        SelectedShiftChangeAM()
    End Sub
    Private Sub LaunchCalendar()
        stops_calendar.Visibility = Windows.Visibility.Visible
        SetBlackOutDates()
    End Sub
    Private Sub SetCILiconlabel()
        If AllProdLines(prstoryReport.ParentLineInt).parentSite.Name = SITE_MANDIDEEP Then
            CIL12.Content = "RLS"
        End If

    End Sub
    Private Sub DateScrollClick(sender As Object, e As MouseButtonEventArgs)

        If sender Is NavigationLeft Then
            If selecteddate_stopswatch.AddDays(-1) >= prstoryReport.StartDate Then
                Cursor = Cursors.Wait
                selecteddate_stopswatch = selecteddate_stopswatch.AddDays(-1)
                LoadUIperDateSelected(selecteddate_stopswatch)
                LinearViewShiftsSetup()
            Else
                MsgBox("You have reached the start of the selected date range.", MsgBoxStyle.Information, "Start of date range")
            End If
        ElseIf sender Is NavigationRight Then
            If selecteddate_stopswatch.AddDays(1) <= prstoryReport.EndDate Then
                Cursor = Cursors.Wait
                selecteddate_stopswatch = selecteddate_stopswatch.AddDays(1)
                LoadUIperDateSelected(selecteddate_stopswatch)
                LinearViewShiftsSetup()
            Else
                MsgBox("You have reached the end of the selected date range.", MsgBoxStyle.Information, "End of date range")
            End If

        End If
    End Sub
    Private Sub dateselected()
        LoadUIperDateSelected(stops_calendar.SelectedDate)
        LinearViewShiftsSetup()
    End Sub
    Private Sub LoadUIperDateSelected(dateselected As Date)
        stops_calendar.Visibility = Windows.Visibility.Hidden
        SelectedDate.Content = Format(dateselected, "MMMM dd, yyyy")
        stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, dateselected, My.Settings.defaultMappingLevel)
        selectedshift = "D"
        hourslider.Value = 0.0
        AMbutton.Opacity = 0.6
        AMbutton.Background = myBrushAquaMarineX
        PMbutton.Opacity = 0.2
        PMbutton.Background = myBrushWhiteX
        selectedshift = "N"
        Cursor = Cursors.Arrow
        selecteddate_stopswatch = dateselected
        SelectedShiftChangeAM()
    End Sub
    Private Sub SetBlackOutDates()
        stops_calendar.BlackoutDates.Add(New CalendarDateRange(prstoryReport.EndDate.AddDays(1), New DateTime(2100, 1, 1)))
        stops_calendar.BlackoutDates.Add(New CalendarDateRange(New DateTime(1900, 1, 1), prstoryReport.StartDate))
    End Sub

    Sub GenerateStopsWatch()

        Dim hourNUMBER As Integer
        Dim max_stops As Integer


        stopswatchRAWDATA.setCurrentFailureMode(selectedfailuremode)
        max_stops = stopswatchRAWDATA.MaxHourStops

        If selectedshift = "D" Then
            For hourNUMBER = 1 To 12

                Select Case hourNUMBER

                    Case 2

                        hour1bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour1base.Fill = mybrushred
                        Else
                            hour1base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILIcon1.Visibility = Windows.Visibility.Visible
                        Else
                            CILIcon1.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon1.Visibility = Windows.Visibility.Visible
                        Else
                            COicon1.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour1base.Fill = mybrushlightgray
                        End If


                    Case 3
                        hour2bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour2base.Fill = mybrushred
                        Else
                            hour2base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon2.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon2.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon2.Visibility = Windows.Visibility.Visible
                        Else
                            COicon2.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour2base.Fill = mybrushlightgray
                        End If

                    Case 4
                        hour3bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour3base.Fill = mybrushred
                        Else
                            hour3base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon3.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon3.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon3.Visibility = Windows.Visibility.Visible
                        Else
                            COicon3.Visibility = Windows.Visibility.Hidden
                        End If


                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour3base.Fill = mybrushlightgray
                        End If
                    Case 5
                        hour4bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour4base.Fill = mybrushred
                        Else
                            hour4base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon4.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon4.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon4.Visibility = Windows.Visibility.Visible
                        Else
                            COicon4.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour4base.Fill = mybrushlightgray
                        End If
                    Case 6
                        hour5bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour5base.Fill = mybrushred
                        Else
                            hour5base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon5.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon5.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon5.Visibility = Windows.Visibility.Visible
                        Else
                            COicon5.Visibility = Windows.Visibility.Hidden
                        End If


                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour5base.Fill = mybrushlightgray
                        End If
                    Case 7
                        hour6bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour6base.Fill = mybrushred
                        Else
                            hour6base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon6.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon6.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon6.Visibility = Windows.Visibility.Visible
                        Else
                            COicon6.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour6base.Fill = mybrushlightgray
                        End If
                    Case 8
                        hour7bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour7base.Fill = mybrushred
                        Else
                            hour7base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon7.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon7.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            Coicon7.Visibility = Windows.Visibility.Visible
                        Else
                            Coicon7.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour7base.Fill = mybrushlightgray
                        End If
                    Case 9
                        hour8bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour8base.Fill = mybrushred
                        Else
                            hour8base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            Cilicon8.Visibility = Windows.Visibility.Visible
                        Else
                            Cilicon8.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon8.Visibility = Windows.Visibility.Visible
                        Else
                            COicon8.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour8base.Fill = mybrushlightgray
                        End If
                    Case 10
                        Hour9bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour9base.Fill = mybrushred
                        Else
                            Hour9base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon9.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon9.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon9.Visibility = Windows.Visibility.Visible
                        Else
                            COicon9.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour9base.Fill = mybrushlightgray
                        End If
                    Case 11
                        Hour10bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour10base.Fill = mybrushred
                        Else
                            Hour10base.Fill = mybrushgreen
                        End If
                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon10.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon10.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon10.Visibility = Windows.Visibility.Visible
                        Else
                            COicon10.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour10base.Fill = mybrushlightgray
                        End If
                    Case 12
                        Hour11bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour11base.Fill = mybrushred
                        Else
                            Hour11base.Fill = mybrushgreen
                        End If
                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon11.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon11.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon11.Visibility = Windows.Visibility.Visible
                        Else
                            COicon11.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour11base.Fill = mybrushlightgray
                        End If
                    Case 1
                        Hour12bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour12base.Fill = mybrushred
                        Else
                            Hour12base.Fill = mybrushgreen
                        End If
                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon12.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon12.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon12.Visibility = Windows.Visibility.Visible
                        Else
                            COicon12.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour12base.Fill = mybrushlightgray
                        End If
                End Select


            Next
        ElseIf selectedshift = "N" Then
            For hourNUMBER = 13 To 24

                Select Case hourNUMBER

                    Case 14
                        hour1bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour1base.Fill = mybrushred
                        Else
                            hour1base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILIcon1.Visibility = Windows.Visibility.Visible
                        Else
                            CILIcon1.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon1.Visibility = Windows.Visibility.Visible
                        Else
                            COicon1.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour1base.Fill = mybrushlightgray
                        End If
                    Case 15
                        hour2bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour2base.Fill = mybrushred
                        Else
                            hour2base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon2.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon2.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon2.Visibility = Windows.Visibility.Visible
                        Else
                            COicon2.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour2base.Fill = mybrushlightgray
                        End If
                    Case 16
                        hour3bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour3base.Fill = mybrushred
                        Else
                            hour3base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon3.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon3.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon3.Visibility = Windows.Visibility.Visible
                        Else
                            COicon3.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour3base.Fill = mybrushlightgray
                        End If
                    Case 17
                        hour4bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour4base.Fill = mybrushred
                        Else
                            hour4base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon4.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon4.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon4.Visibility = Windows.Visibility.Visible
                        Else
                            COicon4.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour4base.Fill = mybrushlightgray
                        End If
                    Case 18
                        hour5bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour5base.Fill = mybrushred
                        Else
                            hour5base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon5.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon5.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon5.Visibility = Windows.Visibility.Visible
                        Else
                            COicon5.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour5base.Fill = mybrushlightgray
                        End If
                    Case 19
                        hour6bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour6base.Fill = mybrushred
                        Else
                            hour6base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon6.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon6.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon6.Visibility = Windows.Visibility.Visible
                        Else
                            COicon6.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour6base.Fill = mybrushlightgray
                        End If
                    Case 20
                        hour7bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour7base.Fill = mybrushred
                        Else
                            hour7base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon7.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon7.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            Coicon7.Visibility = Windows.Visibility.Visible
                        Else
                            Coicon7.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour7base.Fill = mybrushlightgray
                        End If
                    Case 21
                        hour8bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            hour8base.Fill = mybrushred
                        Else
                            hour8base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            Cilicon8.Visibility = Windows.Visibility.Visible
                        Else
                            Cilicon8.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon8.Visibility = Windows.Visibility.Visible
                        Else
                            COicon8.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            hour8base.Fill = mybrushlightgray
                        End If
                    Case 22
                        Hour9bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour9base.Fill = mybrushred
                        Else
                            Hour9base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon9.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon9.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon9.Visibility = Windows.Visibility.Visible
                        Else
                            COicon9.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour9base.Fill = mybrushlightgray
                        End If
                    Case 23
                        Hour10bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour10base.Fill = mybrushred
                        Else
                            Hour10base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon10.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon10.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon10.Visibility = Windows.Visibility.Visible
                        Else
                            COicon10.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour10base.Fill = mybrushlightgray
                        End If
                    Case 24
                        Hour11bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour11base.Fill = mybrushred
                        Else
                            Hour11base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon11.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon11.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon11.Visibility = Windows.Visibility.Visible
                        Else
                            COicon11.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour11base.Fill = mybrushlightgray
                        End If
                    Case 13

                        Hour12bar.Height = stopswatchRAWDATA.getModeStops(hourNUMBER) * hourbarMAXsize / max_stops
                        If stopswatchRAWDATA.isPRout(hourNUMBER) Then
                            Hour12base.Fill = mybrushred
                        Else
                            Hour12base.Fill = mybrushgreen
                        End If

                        If stopswatchRAWDATA.isCIL(hourNUMBER) Then
                            CILicon12.Visibility = Windows.Visibility.Visible
                        Else
                            CILicon12.Visibility = Windows.Visibility.Hidden
                        End If
                        If stopswatchRAWDATA.isChangeover(hourNUMBER) Then
                            COicon12.Visibility = Windows.Visibility.Visible
                        Else
                            COicon12.Visibility = Windows.Visibility.Hidden
                        End If

                        If DateValue(selecteddate_stopswatch) = DateValue(prstoryReport.EndDate) And hourNUMBER > Hour(prstoryReport.EndTime) Then
                            Hour12base.Fill = mybrushlightgray
                        End If
                End Select


            Next
        End If

        hourslider.Value = 0.0
        SliderAct()
    End Sub
    Sub StopsWindowExpand()
        WindowExpandIcon.Visibility = Windows.Visibility.Hidden
        WindowCollapseIcon.Visibility = Windows.Visibility.Visible
        Dim hx As Integer
        For hx = 0 To 225 Step 25
            Me.Height = Me.Height + 25
        Next
    End Sub

    Sub StopsWindowCollapse()
        WindowCollapseIcon.Visibility = Windows.Visibility.Hidden
        WindowExpandIcon.Visibility = Windows.Visibility.Visible
        Me.Height = 435
    End Sub

    Sub SelectedShiftChangeAM()
        selectedshift = "D"
        starttimelabel.Content = "12 AM"
        midtimelabel.Content = "6 AM"
        endtimelabel.Content = "11 AM"
        selectedshiftAMPMlabel.Content = "AM"
        AMbutton.Opacity = 0.6
        AMbutton.Background = myBrushAquaMarineX
        PMbutton.Opacity = 0.1
        PMbutton.Background = myBrushWhiteX
        GenerateStopsWatch()
        LinearViewShiftsSetup()
    End Sub
    Sub SelectedShiftChangePM()

        selectedshift = "N"
        starttimelabel.Content = "12 PM"
        midtimelabel.Content = "6 PM"
        endtimelabel.Content = "11 PM"
        selectedshiftAMPMlabel.Content = "PM"
        PMbutton.Opacity = 0.6
        PMbutton.Background = myBrushAquaMarineX
        AMbutton.Opacity = 0.1
        AMbutton.Background = myBrushWhiteX
        GenerateStopsWatch()


    End Sub
    Private Sub Show_Hour_Details_Linear(sender As Object, e As MouseButtonEventArgs)
        Dim selectedhour As Integer
        Dim sendername As String
        Dim myBrush As New SolidColorBrush(Colors.DimGray)


        tempsender = sender
        sendername = sender.name
        selectedhour = -1


        Dim label As Label
        If InStr(sender.GetType().ToString(), "Rectangle") > 0 Then
            For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(Me.StopsWatchLinearUI) - 1
                If Not IsNothing(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)) Then
                    If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Label") > 0 Then
                        label = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                        If InStr(label.Name, Mid(sendername, 1, InStr(sendername, "bar", vbTextCompare) - 1), vbTextCompare) > 0 Then
                            resetallstrokes_linear()

                            If Not IsNothing(tempsender) Then
                                If InStr(tempsender.GetType().ToString, "Rectangle", vbTextCompare) > 0 Then tempsender.StrokeThickness = 0
                            End If

                            selectedhour = onlyDigits(label.Content)
                            sender.Stroke = myBrush
                            sender.StrokeThickness = 2
                            Exit For

                        End If
                    End If
                End If

            Next
        ElseIf InStr(sender.GetType().ToString(), "Label") > 0 Then
            selectedhour = onlyDigits(sender.content)
        End If

        Dim isnextday As Boolean = False

        If selectedhour < AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr Then
            stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch.AddDays(1), My.Settings.defaultMappingLevel)
            stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)
            isnextday = True
        Else
            stopswatchRAWDATA = New prStoryStopsWatch_24(selectedindexofLine_temp, selecteddate_stopswatch, My.Settings.defaultMappingLevel)
            stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)
            isnextday = False

        End If

        Generate_Hourly_PerformanceMetrics("Linear", selectedhour + 1, isnextday)


    End Sub
    Private Sub resetallstrokes_linear()
        Dim rect As Rectangle

        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(Me.StopsWatchLinearUI) - 1
            If Not IsNothing(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i).GetType.ToString, "Rectangle") > 0 Then
                    rect = VisualTreeHelper.GetChild(Me.StopsWatchLinearUI, i)
                    If InStr(rect.Name, "hour") > 0 Then rect.StrokeThickness = 0

                End If
            End If
        Next
    End Sub
    Private Sub Show_Hour_Details(sender As Object, e As MouseButtonEventArgs)

        Dim selectedhour As Integer
        Dim sendername As String
        Dim myBrush As New SolidColorBrush(Colors.DimGray)
        If Not IsNothing(tempsender) Then
            If InStr(tempsender.GetType().ToString, "Rectangle", vbTextCompare) > 0 Then tempsender.StrokeThickness = 0
        End If

        tempsender = sender
        sendername = sender.name
        selectedhour = onlyDigits(sendername)

        If selectedhour <> 12 Then
            hourslider.Value = selectedhour

        Else
            hourslider.Value = 0.0
        End If

        SliderAct()
        sender.Stroke = myBrush
        sender.StrokeThickness = 2
    End Sub
    Sub Generate_Hourly_PerformanceMetrics(shift As String, hour_ofstop As Integer, Optional IsNextDay As Boolean = False)
        Dim tooltipSKUname As String, tmpHeight As Double
        Dim selectedStopsWatchDate As Date

        Select Case shift
            Case "D"
                selectedHourLabel.Content = hour_ofstop - 1 & " " & "AM"
                If hour_ofstop = 1 Then selectedHourLabel.Content = "12 AM"
            Case "N"
                selectedHourLabel.Content = hour_ofstop - 1 & " " & "PM"
                If hour_ofstop = 1 Then selectedHourLabel.Content = "12 PM"
                hour_ofstop = hour_ofstop + 12
            Case "Linear"
                If hour_ofstop - 1 < 12 And hour_ofstop - 1 >= AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr Then selectedHourLabel.Content = hour_ofstop - 1 & " AM"
                If hour_ofstop - 1 > 12 And hour_ofstop - 1 < 24 Then selectedHourLabel.Content = hour_ofstop - 13 & " PM"
                If hour_ofstop - 1 < AllProdLines(selectedindexofLine_temp).ShiftStartFirst_Hr And hour_ofstop - 1 < 12 Then selectedHourLabel.Content = hour_ofstop - 1 & " AM+"
        End Select
        stopswatchRAWDATA.setCurrentFailureMode(SelectedFailureMode)

        tmpHeight = stopswatchRAWDATA.getAvailability(hour_ofstop) * (availabilitybarMAXsize) 'SRO NEW CODE 6/16/15
        Availability_bar.Height = tmpHeight

        availabilitylabeltext.Content = FormatPercent(stopswatchRAWDATA.getAvailability(hour_ofstop))


        stopscounthourlabel.Content = stopswatchRAWDATA.getModeStops(hour_ofstop)
        skunamehourlabel.Content = stopswatchRAWDATA.getSkus(hour_ofstop)
        tooltipSKUname = skunamehourlabel.Content
        skunamehourlabel.ToolTip = tooltipSKUname

        If IsNextDay = False Then
            selectedStopsWatchDate = DateAdd(DateInterval.Hour, -Hour(selecteddate_stopswatch), selecteddate_stopswatch)
            selectedStopsWatchDate = DateAdd(DateInterval.Hour, hour_ofstop - 1, selectedStopsWatchDate)
            displaySelectedDatesRawData(selectedStopsWatchDate)
        Else
            selectedStopsWatchDate = DateAdd(DateInterval.Hour, -Hour(selecteddate_stopswatch), selecteddate_stopswatch)
            selectedStopsWatchDate = DateAdd(DateInterval.Hour, hour_ofstop - 1, selectedStopsWatchDate)
            displaySelectedDatesRawData(selectedStopsWatchDate.AddDays(1))
        End If

    End Sub
    Private Sub SliderAct()
        If StopsWatchLinearUI.Visibility = Visibility.Visible Then Exit Sub
        If Not IsNothing(tempsender) Then
            If InStr(tempsender.GetType().ToString, "Rectangle", vbTextCompare) > 0 Then tempsender.StrokeThickness = 0
        End If

        Dim slidervalue As Integer
        Dim tempbar As New Rectangle

        slidervalue = Int(hourslider.Value) + 1
        If hourslider.Value = 0 Then slidervalue = 1

        Generate_Hourly_PerformanceMetrics(selectedshift, slidervalue)

        Exit Sub
        Select Case slidervalue
            Case 1
                hour1bar.StrokeThickness = 2
            Case 2
                hour2bar.StrokeThickness = 2
            Case 3
                hour3bar.StrokeThickness = 2
            Case 4
                hour4bar.StrokeThickness = 2
            Case 5
                hour5bar.StrokeThickness = 2
            Case 6
                hour6bar.StrokeThickness = 2
            Case 7
                hour7bar.StrokeThickness = 2
            Case 8
                hour8bar.StrokeThickness = 2
            Case 9
                Hour9bar.StrokeThickness = 2
            Case 10
                Hour10bar.StrokeThickness = 2
            Case 11
                Hour11bar.StrokeThickness = 2
            Case 12
                Hour12bar.StrokeThickness = 2
        End Select

    End Sub
    Private Sub BarMove(sender As Object, e As MouseEventArgs)

        sender.Opacity = 0.7

    End Sub

    Private Sub BarLeave(sender As Object, e As MouseEventArgs)
        sender.Opacity = 1

    End Sub
    Sub initializeBars()

    End Sub

End Class
