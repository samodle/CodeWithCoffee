Imports System.ComponentModel
Imports System.Threading
Imports System.Drawing.Font
Imports System.Windows.Media.Fonts
Imports System.Drawing
Imports System.Collections.ObjectModel
Imports MongoDB.Driver
Imports MongoDB.Bson
Imports Newtonsoft.Json
Imports System.IO

Public Class ProductionLineGroup
    Public ReadOnly Property Name As String
        Get
            Return _Name
        End Get
    End Property

    Private _Name As String = "x"

    Public Lines As List(Of Integer)

    Public Sub New(name As String)
        _Name = name
        Lines = New List(Of Integer)
    End Sub


    Public Sub AddLine(lineToString As String)
        For i = 0 To AllProdLines.Count - 1
            If AllProdLines(i).ToString = lineToString Then
                AddLine(i)
                Exit For
            End If
        Next
    End Sub
    Public Sub AddLine(lineName As String, siteName As String)
        For i = 0 To AllProdLines.Count - 1
            If AllProdLines(i).Name = lineName And AllProdLines(i).SiteName = siteName Then
                AddLine(i)
                Exit For
            End If
        Next
    End Sub
    Private Sub AddLine(lineIndex As Integer)
        If Lines.IndexOf(lineIndex) = -1 Then
            Lines.Add(lineIndex)
        End If
    End Sub
End Class

Public Class Window_Multiline

#Region "Constructor"
    Public Sub New(Optional IsForTeamAnalysis As Boolean = False, Optional lineindex As Integer = -1, Optional prstoryreport_forteams As List(Of prStoryMainPageReport) = Nothing, Optional daterange As String = "", Optional paramObj As Object = Nothing)

        ' This call is required by the designer.
        InitializeComponent()
        IsTeamAnalysisinMultiline = IsForTeamAnalysis
        If IsTeamAnalysisinMultiline = True Then
            exportbutton.visibility = visibility.hidden
            currentselectedindex = lineindex
            For i As Integer = 0 To prstoryreport_forteams.Count - 1
                MultiLineRawReports.Add(prstoryreport_forteams(i))
                multiline_LISTofselectedindeces.Add(lineindex)
                DateTimeSelectedLabel.Content = daterange
                MultiTEAMallUIContent = paramObj
                multiline_ListofLineNames_fullname = paramObj(8)
            Next

        End If

    End Sub
#End Region
    Public Sub InitializeMultilineGroups(groupList As List(Of ProductionLineGroup))
        If Not IsNothing(groupList) Then
            For i = 0 To groupList.Count - 1
                _ActiveDataCollection.Add(groupList(i))
            Next
        End If
    End Sub

#Region "Variables"
    Dim DataExport As String = ""
    Dim FileName As String = "PRSTORY_MULTILINE_EXPORT_" & DateTime.Now().Day

    Dim activeLineIndeces As New List(Of Integer)
    Dim multiline_LISTofselectedindeces As New List(Of Integer)
    Dim multiline_currentselectedindex As Integer
    Dim multiline_ListofLineNames_fullname As New List(Of String)
    Dim multiline_ListofPR As New List(Of Double)
    Dim multiline_ListofUPDT As New List(Of Double)
    Dim multiline_ListofPDT As New List(Of Double)
    Dim multiline_ListofSPD As New List(Of Double)
    Dim multiline_ListofMTBF As New List(Of Double)
    Dim multiline_ListofCases As New List(Of Double)
    Dim multiline_ListofActualStops As New List(Of Double)
    Dim multiline_Listofdates As New List(Of String)
    Dim multiline_Listofdates_starttime As New List(Of String)
    Dim multiline_Listofdates_endtime As New List(Of String)
    Dim multiline_Listofdates_endtime_datetimeformat As New List(Of DateTime)
    Public MultiLineRawReports As New List(Of prStoryMainPageReport)
    Public CommonSectorname As String = ""
    Public Commonstartdatetime As String = ""
    Public Commonenddatetime As String = ""
    Public Is1stLineSelected As Boolean = False
    Public IsAllDatesSame As Boolean = True
    Public IsByLossAreaplanned As Boolean = False
    Public IsRollupplanned As Boolean = False
    Public currentselectedindex As Integer = -1
    Public Indexoflastpulledline As Integer = -1
    Public CountofCurrentDataPull As Integer = 0
    Public IsSnakeCurrentlyVisible As Boolean = False
    Public numberofdatapull As Integer = 0
    Public selectedTierRadiobuttoncontent As String = ""
    Public IsTeamAnalysisinMultiline As Boolean = False
    Public MultiTEAMallUIContent() As Object
    Dim PROF_Connectionerror As Boolean = False
    Private prstoryReport As prStoryMainPageReport
    Dim MultilineHTMLthread As Thread
    Dim MultilineHTMLthreas_bylossareachart As Thread
    Dim MultilineHTMLthread_rolluppiechart1 As Thread
    Dim MultilineHTMLthread_rolluppiechart2 As Thread
    Dim _LossTreeList As New ObservableCollection(Of DTevent)()
    Dim _LossTreeList_allselectedlines As New List(Of ObservableCollection(Of DTevent))
    Dim _LossTreeListplanned As New ObservableCollection(Of DTevent)()
    Dim _LossTreeList_allselectedlinesplanned As New List(Of ObservableCollection(Of DTevent))


    Public ReadOnly Property ActiveDataCollection() As ObservableCollection(Of ProdLine)
        Get

            Dim c = New ObservableCollection(Of ProdLine)
            For Each e As ProdLine In AllProdLines
                c.Add(e)
            Next
            Return c

            ' Return allproductionlines'_ActiveDataCollection
        End Get
    End Property

    Public ReadOnly Property ActiveDataCollection2() As ObservableCollection(Of ProductionLineGroup)
        Get
            Return _ActiveDataCollection
        End Get
    End Property


    Dim _ActiveDataCollection As New ObservableCollection(Of ProductionLineGroup)()

    'Dim _LossTreeList_selectedlinetemp As New ArrayList

    Public ReadOnly Property LossTreeList() As ObservableCollection(Of DTevent)
        Get
            Return _LossTreeList
        End Get
    End Property


#End Region
#Region "Threads / Raw Data"
    Dim importTargetsThread As Thread

    Dim snakeThread As Thread
    Dim analysisThread As Thread
    Dim getDTdataThread As Thread
    Dim getDTdataThread2 As Thread
    Dim getDTdataThread3 As Thread
    Dim getProdDataThread As Thread
    Dim progressBarThread As Thread
    Dim inControlThread As Thread
    Dim motionStopsThread As Thread
    Dim motionPRThread As Thread
    Dim motionprstoryThread As Thread
    Dim uptimeViewerThread As Thread
    Dim congratulationmessageThread As Thread
    'Partial Data Arrays
    Dim first10daysDT(,) As Object 'Array
    Dim second10daysDT(,) As Object 'Array
    Dim third10daysDT(,) As Object 'Array
    Dim tmpProdArray As Array
    'data for rate loss
    Dim finalRateLossData(,) As Object

    Dim useThreadingForDataPulling As Boolean = True
#End Region

    Public Sub multiline_onload()
        If IsTeamAnalysisinMultiline = False Then

            UseTrack_Multiline_RawData = False
            UseTrack_Multiline_ByLossAreachartsmain = False
            UseTrack_Multiline_ByLossAreadrilldown1 = False
            UseTrack_Multiline_ByLossAreadrilldown2 = False
            UseTrack_Multiline_ByLossAreadrilldown3 = False
            UseTrack_Multiline_RollupCharts = False
            UseTrack_Multiline_Rollupdrilldown = False

            Splash.Visibility = Visibility.Hidden
            CentralCanvas.Visibility = Visibility.Visible
            ContentCanvas_ListView.Visibility = Visibility.Hidden
            LineAdditionMenuCanvas.Visibility = Visibility.Hidden
            ContentCanvas_ListView.Visibility = Visibility.Hidden
            ' AddLineMainButton.Visibility = Visibility.Hidden
            populatestartandendtimehourandmin()
            figureOutWhichLinesToShow()
            AddItemsToAlllLinesMapping()
            RememberBox.IsChecked = My.Settings.MultilineRememberSelection

        Else

            Splash.Visibility = Visibility.Hidden
            CentralCanvas.Visibility = Visibility.Hidden
            WelcomeCanvas.Visibility = Visibility.Hidden
            Downloadinglinedatalabel.Visibility = Visibility.Hidden
            ContentCanvas_ListView.Visibility = Visibility.Hidden
            LineAdditionMenuCanvas.Visibility = Visibility.Hidden
            ContentCanvas_ListView.Visibility = Visibility.Hidden
            AddLineMainLabel.Visibility = Visibility.Hidden
            AddLineMainButton.Visibility = Visibility.Hidden
            MultilineMainIcon.Visibility = Visibility.Hidden
            TeamWorkMainIcon.Visibility = Visibility.Visible
            Rollupresultslabelheader.Content = "Overall Line Results"
            Benchmarkingresultslabelheader.Content = "Team Results"
            Me.Title = "prstory TEAMWORK"
            BackButton.Visibility = Visibility.Visible
            DateTimeSelectedLabel.Visibility = Visibility.Visible
            SetupMulti_TEAM_ui_All_Lines()
            If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                Chartpostscript2.Visibility = Visibility.Visible
            Else
                Chartpostscript2.Visibility = Visibility.Hidden
            End If

            MultilineHTMLthread = New Thread(AddressOf GenerateSummaryCharts)
            MultilineHTMLthread.Start()
            Thread.Sleep(300)

            ByLossAreaunplannedbtnclicked(LossTreeunplannedbtn, f)
            SummaryChart.Reload(ignoreCache:=True)
            ManageTierComboLabelNames()
            ShowSummaryCharts(SummaryChart, f)
        End If

    End Sub

#Region "MouseMoveLeaveClicketc"
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



    Private Sub exportbtnclicked(sender As Object, e As MouseButtonEventArgs)
        ChartsCanvas.Visibility = Visibility.Hidden
        ExportCanvas.Visibility = Visibility.Visible

        DataViewBtn.Foreground = mybrushdefaultfontgray
        DataViewBtn.Background = mybrushdefaultbackgroundgray

        ChartsViewBtn.Foreground = mybrushdefaultfontgray
        ChartsViewBtn.Background = mybrushdefaultbackgroundgray

        ExportButton.Foreground = mybrushlanguagewhite
        ExportButton.Background = mybrushdefaultfontgray
    End Sub

    Private Sub exportDataSummary(sender As Object, e As MouseButtonEventArgs)
        Try
            CSV_exportString(DataExport, FileName & "_SUMMARY")
        Catch ex As Exception
            MsgBox("Export Error: " & ex.Message, "Export Error", vbCritical)
        End Try

    End Sub

    Private Sub exportRawSummary(sender As Object, e As MouseButtonEventArgs)
        Try
            Dim dataExport2 As String = "Line Name, Start Time, End Time, DT, UT, Reason 1, Reason 2, Reason 3, Reason 4, Tier 1, Tier 2, Tier 3, Fault, PR In/Out"
            dataExport2 += vbCrLf
            For i = 0 To MultiLineRawReports.Count - 1

                Try

                    With MultiLineRawReports(i).MainLEDSReport.DT_Report.rawDTdata
                        dim lineName as string = AllProdLines(multilinerawreports(i)._parentline).tostring()

                        For j = 0 To .rawConstraintData.Count - 1

                            dataExport2 += linename & "," & .rawConstraintData(j).toString_CSV()

                            dataExport2 += vbCrLf
                        Next


                    End With

                Catch ex As Exception
                    dataExport2 += vbCrLf
                    MsgBox("Error exporting data from line " & MultiLineRawReports(i).MainLEDSReport.ParentLine.ToString() & ". We will continue to try to export the rest of the data. Details: " & ex.Message)
               
                    dataExport2 += vbCrLf
                End Try
            Next
            CSV_exportString(dataExport2, FileName & "_RAWDATA")
        Catch ex As Exception
            MsgBox("Export Error: " & ex.Message & ". If this error continues, please reach out to your prstory SPOC or go to the yammer page to directly to ask for help.")
        End Try

    End Sub

    Private Sub dataviewbtnclicked(sender As Object, e As MouseButtonEventArgs)
        DataViewBtn.Foreground = mybrushlanguagewhite
        DataViewBtn.Background = mybrushdefaultfontgray

        ChartsViewBtn.Foreground = mybrushdefaultfontgray
        ChartsViewBtn.Background = mybrushdefaultbackgroundgray

        ExportButton.Foreground = mybrushdefaultfontgray
        ExportButton.Background = mybrushdefaultbackgroundgray

        ChartsCanvas.Visibility = Visibility.Hidden
        ExportButton.Visibility = Visibility.Visible
        ExportCanvas.Visibility = Visibility.Hidden

        UseTrack_Multiline_RawData = True
    End Sub
    Private Sub chartsviewbtnclicked(sender As Object, e As MouseButtonEventArgs)
        ChartsViewBtn.Foreground = mybrushlanguagewhite
        ChartsViewBtn.Background = mybrushdefaultfontgray

        DataViewBtn.Foreground = mybrushdefaultfontgray
        DataViewBtn.Background = mybrushdefaultbackgroundgray

        ExportButton.Foreground = mybrushdefaultfontgray
        ExportButton.Background = mybrushdefaultbackgroundgray

        '  ExportButton.Visibility = Visibility.Hidden
        ChartsCanvas.Visibility = Visibility.Visible
        ExportCanvas.Visibility = Visibility.Hidden
        CloseRollupSplashCanvas(Rollupsplashcanvasclosebtn, f)

    End Sub

    Private Sub Unplannedbtnclicked(sender As Object, e As MouseButtonEventArgs)
        IsRollupplanned = False
        AllLines_plannedButton.Background = mybrushdefaultbackgroundgray
        AllLines_plannedButton.Foreground = mybrushdefaultfontgray

        AllLines_UnplannedButton.Background = mybrushdefaultfontgray
        AllLines_UnplannedButton.Foreground = mybrushlanguagewhite
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        PopulateLossTreeLists(Output_multilinereport, AllLinesMappingLevelComboBox.SelectedValue)

        Dim i As Integer

        For i = 0 To multiline_LISTofselectedindeces.Count - 1
            PopulateLossTreeLists_eachline(Output_multilinereport, i, AllLinesMappingLevelComboBox.SelectedValue)

        Next



    End Sub

    Private Sub plannedbtnclicked(sender As Object, e As MouseButtonEventArgs)
        IsRollupplanned = True

        AllLines_UnplannedButton.Background = mybrushdefaultbackgroundgray
        AllLines_UnplannedButton.Foreground = mybrushdefaultfontgray

        AllLines_plannedButton.Background = mybrushdefaultfontgray
        AllLines_plannedButton.Foreground = mybrushlanguagewhite

        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        PopulateLossTreeListsPlanned(Output_multilinereport, AllLinesMappingLevelComboBox.SelectedValue)


        Dim i As Integer

        For i = 0 To multiline_LISTofselectedindeces.Count - 1
            PopulateLossTreeLists_eachlineplanned(Output_multilinereport, i, AllLinesMappingLevelComboBox.SelectedValue)

        Next
    End Sub
#End Region
#Region "Closing, Launching etc"
    Public Sub CloseAddLine()
        LineAdditionMenuCanvas.Visibility = Visibility.Hidden
    End Sub

    Public Sub LaunchAddLine()
        LineAdditionMenuCanvas.Visibility = Visibility.Visible
        WelcomeCanvas.Visibility = Visibility.Hidden
        CentralCanvas.Visibility = Visibility.Hidden
        AddLinesForAnalysis.Visibility = Visibility.Hidden
        AddLineMainButton.Visibility = Visibility.Visible
        AddLineMainLabel.Visibility = Visibility.Visible
        '  prstory_linedropdown.SelectedIndex = -1

        If My.Settings.MultilineRememberSelection Then
            If My.Settings.MultilineSelection <> "x" And My.Settings.MultilineSelection.Length > 5 Then
                Dim MyArray() As String = My.Settings.MultilineSelection.Split(",")
                Dim MyList As List(Of String) = MyArray.ToList()



                For i = 0 To MyList.Count - 1
                    For j = 0 To AllProdLines.Count - 1
                        If AllProdLines(j).ToString = MyList(i) Then
                            prstory_linedropdown.SelectedItems.Add(AllProdLines(j))
                            j = AllProdLines.Count
                        End If
                    Next
                Next
            End If
        End If
    End Sub
    Private Sub AddLineInitiate()
        Dim i As Integer
        Dim progthread As Thread
        Dim progBar As DelegateUpdateProgressBar = AddressOf updateProgressBar_TimeDriven
        Dim progBarFast As DelegateUpdateProgressBar = AddressOf updateProgressBar_TimeDriven_Fast

        If CheckPreDoAnalyze() <> 0 Then

            'update list of selected lines
            Dim MultilineSelection As String = ""
            If RememberBox.IsChecked Then
                My.Settings.MultilineRememberSelection = True
                For i = 0 To prstory_linedropdown.SelectedItems.Count - 1
                    MultilineSelection = MultilineSelection & prstory_linedropdown.SelectedItems(i).ToString
                    If i <> prstory_linedropdown.SelectedItems.Count - 1 Then
                        MultilineSelection = MultilineSelection & ","
                    End If
                Next
            Else
                My.Settings.MultilineRememberSelection = False
            End If
            My.Settings.MultilineSelection = MultilineSelection
            My.Settings.Save()

            shouldSnakeClose = False
            snakeThread = New Thread(AddressOf PlaySnake)
            progthread = New Thread(AddressOf updateProgressBar_TimeDriven)

            CloseAddLine()
            WelcomeCanvas.Visibility = Visibility.Visible
            Downloadinglinedatalabel.Visibility = Visibility.Visible

            Splash.Visibility = Visibility.Visible
            CentralCanvas.Visibility = Visibility.Visible
            WelcomeCanvas.Visibility = Visibility.Visible
            AddLinesForAnalysis.Visibility = Visibility.Hidden
            System.Windows.Forms.Application.DoEvents()
            If My.Settings.AdvancedSettings_PlaySnake Then snakeThread.Start()
            System.Windows.Forms.Application.DoEvents()

            For i = 0 To numberofdatapull - 1
                'Download data for all selectedlines except the ones which are already downloade, ofcourse
                Do_Analyze(multiline_LISTofselectedindeces(Indexoflastpulledline + i + 1))
                System.Windows.Forms.Application.DoEvents()
                CountofCurrentDataPull = i + 1
            Next
            If Not My.Computer.Keyboard.CtrlKeyDown Then
                shouldSnakeClose = True
            End If
            'Now that all data is download, do the aggregate calculations and generate UI
            ManageDatesandTimeforreport()
            SetupMultiLineUI_AllLines()
            MultilineHTMLthread = New Thread(AddressOf GenerateSummaryCharts)
            MultilineHTMLthread.Start()
            Thread.Sleep(300)
            Downloadinglinedatalabel.Visibility = Visibility.Hidden
            WelcomeCanvas.Visibility = Visibility.Hidden
            CentralCanvas.Visibility = Visibility.Hidden
            Splash.Visibility = Visibility.Hidden

            ByLossAreaunplannedbtnclicked(LossTreeunplannedbtn, f)
            SummaryChart.Reload(ignoreCache:=True)
            ManageTierComboLabelNames()
            ShowSummaryCharts(SummaryChart, f)
        End If
    End Sub


    Private Sub multilinewindowclose(ByVal sender As Object, ByVal e As CancelEventArgs)
        If InStr(sender.ToString, "multiline", vbTextCompare) > 0 And IsTeamAnalysisinMultiline = False Then
            Me.Owner.Visibility = Windows.Visibility.Visible
            Try
                'SendUserAnalyticsDatatoServer_multiline()
            Catch ex As Exception
            End Try
        End If
        SummaryChart.Dispose()
        ByLossAreaChart.Dispose()
        RollupChart1.Dispose()
        RollupChart2.Dispose()

    End Sub

#End Region
#Region "AddLineMenuStuff"
    Private Sub populatestartandendtimehourandmin()

        Dim k As Integer

        For k = 0 To 23
            If k > 9 Then
                starthour.Items.Add(CStr(k))
                endhour.Items.Add(CStr(k))
            Else
                starthour.Items.Add("0" & CStr(k))
                endhour.Items.Add("0" & CStr(k))
            End If
        Next

        For k = 0 To 59

            If k > 9 Then
                startmin.Items.Add(CStr(k))
                endmin.Items.Add(CStr(k))
            Else
                startmin.Items.Add("0" & CStr(k))
                endmin.Items.Add("0" & CStr(k))
            End If
        Next

    End Sub
    Private Sub prstory_linedropdown_SelectionChanged()
        If Is1stLineSelected = False Then
            Is1stLineSelected = True
            SetStartandEndTime()
        End If

        Estimateddownloadtimelabel.Content = "Estimated data download time is " & (prstory_linedropdown.SelectedItems.Count * 0.5) & " min"
    End Sub

    Public Sub AllLinesLossTreeSelectionChanged()

        If System.IO.File.Exists(SERVER_FOLDER_PATH & "pie.js") Then
            'The file exists
        Else
            'the file doesn't exist
            Try
                CreatePie_JS()
            Catch ex As Exception
                DownloadPieJSFiles()
            End Try
        End If
        If AllLinesLossTreeListBox.SelectedItems.Count <> 0 Then
            LaunchRollupSplashCanvas(AllLinesLossTreeListBox.SelectedItems(0).Name.ToString(), AllLinesMappingLevelComboBox.SelectedValue)
        End If
    End Sub




    Public Sub AddItemsToAlllLinesMapping()

        AllLinesMappingLevelComboBox.Items.Add("Tier 1")
        AllLinesMappingLevelComboBox.Items.Add("Tier 2")
        AllLinesMappingLevelComboBox.Items.Add("Tier 3")
        AllLinesMappingLevelComboBox.Items.Add("DTGroup")
        AllLinesMappingLevelComboBox.SelectedValue = "Tier 1"

    End Sub
    Public Sub AlllinesMappingLevelComboBoxSelectionChanged()
        UseTrack_Multiline_Rollupdrilldown = True
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        Dim i As Integer
        If IsRollupplanned = False Then
            PopulateLossTreeLists(Output_multilinereport, AllLinesMappingLevelComboBox.SelectedValue)
            For i = 0 To multiline_LISTofselectedindeces.Count - 1
                PopulateLossTreeLists_eachline(Output_multilinereport, i, AllLinesMappingLevelComboBox.SelectedValue)
            Next


        Else
            PopulateLossTreeListsPlanned(Output_multilinereport, AllLinesMappingLevelComboBox.SelectedValue)
            For i = 0 To multiline_LISTofselectedindeces.Count - 1
                PopulateLossTreeLists_eachlineplanned(Output_multilinereport, i, AllLinesMappingLevelComboBox.SelectedValue)
            Next

        End If
    End Sub

    Public Function FindIndexofStringinList(Llist As List(Of ProdLine), searchstring As ProdLine) As Integer
        Dim i As Integer

        For i = 0 To Llist.Count - 1

            If searchstring.ToString = Llist(i).ToString Then
                Return i
                Exit For
            End If

        Next

        Return -1
    End Function

    Private Sub SetStartandEndTime()
        starthour.SelectedValue = "06"
        endhour.SelectedValue = "06"
        startmin.SelectedValue = "30"
        endmin.SelectedValue = "30"

        Exit Sub
        Try
           startmin.SelectedValue = "00"
            endmin.SelectedValue = "00"

        Catch ex As Exception
            starthour.SelectedValue = "06"
            endhour.SelectedValue = "06"
            startmin.SelectedValue = "00"
            endmin.SelectedValue = "00"
            Exit Sub
        End Try

    End Sub
    Private Sub figureOutWhichLinesToShow()
        Dim lineIncrementer As Integer

        activeLineIndeces.Clear()

        For lineIncrementer = 0 To AllProdLines.Count - 1
            With AllProdLines(lineIncrementer)
                '  If .ToString <> "Line Selection" Then
                '  prstory_linedropdown.Items.Add(.ToString)
                activeLineIndeces.Add(lineIncrementer)
                '   End If
            End With
        Next

    End Sub

    Private Sub figureOutWhichLineGroupsToShow()
        '
        ''        Dim groupList As List(Of String) = New List(Of String)

        '        For lineIncrementer = 0 To AllProductionLines.Count - 1
        '           With AllProductionLines(lineIncrementer)
        '               If .ToString <> "Line Selection" Then
        '                   If (.MultilineGroup <> "") Then
        '                       If Not groupList.Contains(.MultilineGroup) Then
        '                           groupList.Add(.MultilineGroup)
        '                           prstory_linedropdown2.items.add(.MultilineGroup)
        '                       End If
        '                   End If
        '               End If
        '           End With
        '       Next
        '
    End Sub


    Private Sub SwitchSelectMode()
        If SwitchSelectModeLabel.Content = "Select By Group" Then
            SwitchSelectModeLabel.Content = "Select By Line"
            prstory_linedropdown.Visibility = Visibility.Hidden
            ' prstory_linedropdown2.Visibility = Visibility.Visible
            RememberBox.Visibility = Visibility.Hidden
            SelectlineLabel.Content = "Select one or more groups"
        Else
            SwitchSelectModeLabel.Content = "Select By Group"
            prstory_linedropdown.Visibility = Visibility.Visible
            ' prstory_linedropdown2.Visibility = Visibility.Hidden
            RememberBox.Visibility = Visibility.Visible
            SelectlineLabel.Content = "Select one or more lines"
        End If
    End Sub

    Private Sub mtdoptionclicked()
        prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-Day(DateTime.Today) + 1)
        prstory_datepicker_enddate.SelectedDate = DateTime.Today.ToString

    End Sub
    Private Sub last7daysclicked()
        prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-7)
        prstory_datepicker_enddate.SelectedDate = DateTime.Today

    End Sub

    Private Sub yesterdayclicked()
        If Hour(Now) > 6 Then

            prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-1)
            prstory_datepicker_enddate.SelectedDate = DateTime.Today
        Else
            prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-2)
            prstory_datepicker_enddate.SelectedDate = DateTime.Today.AddDays(-1)
        End If
    End Sub
#End Region
#Region "Progress Bar"
    Private Delegate Sub DelegateUpdateProgressBar()
    Private Sub updateProgressBar_TimeDriven()
        MainProgressBar.Visibility = Windows.Visibility.Visible

        With MainProgressBar
            .Value = 0
            While .Value < .Maximum - 2 And Not PROF_Connectionerror

                .Value += 2 / (numberofdatapull)
                System.Windows.Forms.Application.DoEvents()
                Thread.Sleep(200)

            End While

            .Visibility = Windows.Visibility.Hidden

            'System.Windows.Forms.Application.DoEvents()
        End With
    End Sub

    Private Delegate Sub DelegateUpdateProgressBar_Fast()
    Private Sub updateProgressBar_TimeDriven_Fast()
        MainProgressBar.Visibility = Windows.Visibility.Visible

        With MainProgressBar
            .Value = 25
            Thread.Sleep(70)
            .Value = 60
            While .Value < .Maximum - 2 And Not PROF_Connectionerror
                .Value += 10
                System.Windows.Forms.Application.DoEvents()
                Thread.Sleep(150)
            End While
            .Visibility = Windows.Visibility.Hidden
            'System.Windows.Forms.Application.DoEvents()
        End With
    End Sub
    Public Sub PlaySnake()
        Dim snakeS As New Snake
        IsSnakeCurrentlyVisible = True
        snakeS.ShowDialog()
    End Sub
#End Region
#Region "Do_Analyze"
    Private Function CheckPreDoAnalyze() As Integer
        Try
            If IsNothing(prstory_datepicker_enddate.SelectedDate) Or IsNothing(prstory_datepicker_startdate.SelectedDate) Then    'handles user generated  no-date selection error
                MsgBox("Start and end date and time are not properly entered. Please try again.")
                Return 0
            End If


            prstory_datepicker_startdate.SelectedDate = Format(prstory_datepicker_startdate.SelectedDate, "Short Date") & " " & starthour.SelectedValue.ToString & ":" & startmin.SelectedValue.ToString
            prstory_datepicker_enddate.SelectedDate = Format(prstory_datepicker_enddate.SelectedDate, "Short Date") & " " & endhour.SelectedValue.ToString & ":" & endmin.SelectedValue.ToString




            If prstory_datepicker_enddate.SelectedDate < prstory_datepicker_startdate.SelectedDate Then    'handles user generated  no-date selection error
                MsgBox("End date cannot be earlier than start date. Please try again.")
                Return 0
            End If

            If prstory_datepicker_enddate.SelectedDate > Now Then
                MsgBox("End date cannot be in future")
                Return 0
            End If

            If DateDiff("d", prstory_datepicker_startdate.SelectedDate, prstory_datepicker_enddate.SelectedDate) > 89 Then
                MsgBox("Sorry, we are not there yet." & vbNewLine & vbNewLine & "We are still working on getting prstory work for date ranges greater than 99 days.")
                Return 0
            End If

            Dim i As Integer





            If prstory_linedropdown.SelectedItems.Count = 0 Then
                MsgBox("Please select one or more lines")
                Return 0
            Else
                For i = 0 To prstory_linedropdown.SelectedItems.Count - 1
                    If FindIndexofStringinList(AllProdLines, prstory_linedropdown.SelectedItems(i)) <> -1 Then
                        multiline_LISTofselectedindeces.Add(FindIndexofStringinList(AllProdLines, prstory_linedropdown.SelectedItems(i)))

                    End If
                Next
                numberofdatapull = multiline_LISTofselectedindeces.Count - (Indexoflastpulledline + 1)
            End If
        Catch ex As Exception
            MessageBox.Show("Error. Make sure you have selected lines and dates and try again." & ex.Message)
            Return 0
        End Try

        Return numberofdatapull
    End Function
    Public Sub Do_Analyze(indexoflinetopulldata As Integer)
        currentselectedindex = indexoflinetopulldata + 1
        Dim i As Integer, rawDTdataColumns As Integer
        Dim _startTime As Date
        Dim _endTime As Date
        Dim lineToAnalyze As ProdLine
        Dim netEvents As Long
        Dim linenametempforMongo As String

        PROF_Connectionerror = False
        getProdDataThread = New Thread(AddressOf Do_Analyze_Prod)
        getDTdataThread = New Thread(AddressOf Do_Analyze_DT)
        getDTdataThread2 = New Thread(AddressOf Do_Analyze_DT1)
        getDTdataThread3 = New Thread(AddressOf Do_Analyze_DT2)

        Dim paramObj_One(3) As Object 'DOWNTIME
        Dim paramObj_Two(3) As Object 'PRODUCTION
        Dim paramObj_Three(3) As Object ' TEMPORARY Only for lines which are less than 3 months old

        currentselectedindex = activeLineIndeces(currentselectedindex - 1)
        selectedindexofLine_temp = currentselectedindex
        linenametempforMongo = AllProdLines(selectedindexofLine_temp).ToString

        'Send line name to Mongo
        Try
            SendUserAnalyticsDatatoServer_multiline(linenametempforMongo)
        Catch ex As Exception
        End Try


        multiline_ListofLineNames_fullname.Add(AllProdLines(selectedindexofLine_temp).ToString)
        multiline_Listofdates.Add(prstory_datepicker_startdate.SelectedDate.ToString & vbCrLf & prstory_datepicker_enddate.ToString)
        multiline_Listofdates_starttime.Add(prstory_datepicker_startdate.ToString)
        multiline_Listofdates_endtime.Add(prstory_datepicker_enddate.ToString)
        multiline_Listofdates_endtime_datetimeformat.Add(prstory_datepicker_enddate.SelectedDate)

        ReInitializeAllPublicVariables()
        lineToAnalyze = AllProdLines(currentselectedindex)

        If lineToAnalyze.parentModule.SQLprocedurePROD = DefaultProficyProductionProcedure.NA Then
            My.Settings.AdvancedSettings_isAvailabilityMode = True
        End If
        My.Settings.defaultDownTimeField = linetoanalyze.MappingLevelA
        My.Settings.defaultDownTimeField_Secondary = linetoanalyze.MappingLevelB

        CommonSectorname = lineToAnalyze.Sector.ToString
 
        If lineToAnalyze.IsStartupMode = True Or lineToAnalyze.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.GLEDS Then
            useThreadingForDataPulling = False
        End If

        If lineToAnalyze.IsStartupMode = True Then useThreadingForDataPulling = False

        _endTime = prstory_datepicker_enddate.SelectedDate

        If lineToAnalyze.IsStartupMode = True Then
            _startTime = prstory_datepicker_startdate.SelectedDate
        Else
            _startTime = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, _endTime) '-30
        End If

        paramObj_One(0) = currentselectedindex
        paramObj_One(1) = _startTime
        paramObj_One(2) = _endTime

        paramObj_Two(0) = currentselectedindex
        paramObj_Two(1) = DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime)
        paramObj_Two(2) = _endTime

        If lineToAnalyze._isDualConstraint Then useThreadingForDataPulling = False 'Added by SRO to resolve Greensboro panic

        Dim isQQ As Boolean
        If lineToAnalyze.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.QuickQuery Then
            isQQ = True
        Else
            isQQ = False
        End If

        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then getProdDataThread.Start(paramObj_Two)

        While IsNothing(tmpProdArray) And Not My.Settings.AdvancedSettings_isAvailabilityMode And lineToAnalyze.SQLproductionProcedure = DefaultProficyProductionProcedure.QuickQuery_MOT
            Thread.Sleep(500)
        End While

        If lineToAnalyze.SQLdowntimeProcedure = DefaultProficyProductionProcedure.QuickQuery_MOT And Not My.Settings.AdvancedSettings_isAvailabilityMode Then
            If tmpProdArray Is {-1, -1} Or IsNothing(tmpProdArray) Then
                My.Settings.AdvancedSettings_isAvailabilityMode = True
            End If
        End If

        If useThreadingForDataPulling Then
            getDTdataThread.Start(paramObj_One)
            getDTdataThread2.Start(paramObj_One)
            getDTdataThread3.Start(paramObj_One)
        Else

            ''''''''this is temporary'''''''''''the if statement ..
            If lineToAnalyze.IsStartupMode Then
                paramObj_Three(0) = selectedindexofLine_temp
                paramObj_Three(1) = DateAdd(DateInterval.Day, -40, _endTime)   'Format("10/24/2015 07:30:00 AM", "Short Date") & " " & Format("10/24/2015 07:30:00 AM", "Long Time") '"10/24/2015 07:30:00 AM"       'line really started giving data on this day
                paramObj_Three(2) = _endTime
                getDTdataThread.Start(paramObj_Three)
                _startTime = DateAdd(DateInterval.Day, -40, _endTime)
            Else
                '''''''this is permanent''''''''''''''
                getDTdataThread.Start(paramObj_Two)      'do not remove this
            End If
        End If

        If useThreadingForDataPulling Then
            While IsNothing(first10daysDT) And Not PROF_Connectionerror
                Thread.Sleep(500)
            End While
        End If

        If Not PROF_Connectionerror Then
            While IsNothing(third10daysDT)
                Thread.Sleep(500)
            End While
            If useThreadingForDataPulling Then
                While IsNothing(second10daysDT)
                    Thread.Sleep(500)
                End While
                Select Case lineToAnalyze.SQLdowntimeProcedure
                    Case DefaultProficyDowntimeProcedure.OneClick
                        rawDTdataColumns = 20
                    Case DefaultProficyDowntimeProcedure.QuickQuery
                        rawDTdataColumns = 29
                    Case DefaultProficyDowntimeProcedure.RE_CentralServer
                        rawDTdataColumns = 20
                    Case DefaultProficyDowntimeProcedure.GLEDS
                        rawDTdataColumns = 29
                    Case DefaultProficyDowntimeProcedure.Maple
                        rawDTdataColumns = 20
                End Select
            End If
            If useThreadingForDataPulling Then
                netEvents = first10daysDT.GetLength(1) + second10daysDT.GetLength(1) + third10daysDT.GetLength(1)
            Else
                netEvents = third10daysDT.GetLength(1)
            End If

            Dim completeDTarray(rawDTdataColumns, netEvents - 1) As Object
            If useThreadingForDataPulling Then
                For i = 0 To rawDTdataColumns '29
                    Array.Copy(first10daysDT, i * (first10daysDT.GetLength(1)), completeDTarray, i * netEvents, first10daysDT.GetLength(1)) 'first10daysDT.GetLength(1) - 1)
                    System.Array.Copy(second10daysDT, i * (second10daysDT.GetLength(1)), completeDTarray, i * netEvents + first10daysDT.GetLength(1), second10daysDT.GetLength(1))
                    System.Array.Copy(third10daysDT, i * (third10daysDT.GetLength(1)), completeDTarray, i * netEvents + first10daysDT.GetLength(1) + second10daysDT.GetLength(1), third10daysDT.GetLength(1))
                Next
            Else
                completeDTarray = third10daysDT
            End If
            'CHECK FOR DUAL CONSTRAINT
            Dim rateLossData(,) As Object
            If lineToAnalyze._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode <> MultiConstraintAnalysis.SingleConstraint Then

                'If AllProductionSectors(currentselectedindex). = SECTOR_FAMILY Then
                If My.Settings.AdvancedSettings_isAvailabilityMode Then  ' the assumption here is Family Care is only sector that would run availability mode and dual constraint (Blocked/Starved + ...)
                    Try
                        If Not isQQ Then
                            rateLossData = getRawProficyData_OneClick(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                            If IsNothing(rateLossData) Then
                                lineToAnalyze._isDualConstraint = False
                            Else
                                finalRateLossData = PROF_mergeRateLossWithMain_OneClick(rateLossData, completeDTarray)
                            End If
                        Else
                            rateLossData = getRawProficyData(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                            If IsNothing(rateLossData) Then
                                lineToAnalyze._isDualConstraint = False
                            Else
                                finalRateLossData = PROF_mergeRateLossWithMain(rateLossData, completeDTarray)
                            End If
                        End If

                    Catch ex As Exception
                        lineToAnalyze._isDualConstraint = False
                        MessageBox.Show("Error Collecting Blocked/Starved Data for " + lineToAnalyze.Name + ". Please click 'OK' to continue." + ex.Message)
                    End Try
                Else
                    If isQQ Then
                        rateLossData = getRawProficyData(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                        finalRateLossData = PROF_mergeRateLossWithMain(rateLossData, completeDTarray)
                    Else
                        rateLossData = getRawProficyData_OneClick(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                        finalRateLossData = PROF_mergeRateLossWithMain_OneClick(rateLossData, completeDTarray)
                    End If
                End If
            End If
            ''''''''''''''''''''''''

            While IsNothing(tmpProdArray) And Not My.Settings.AdvancedSettings_isAvailabilityMode
                Thread.Sleep(500)
            End While

            With AllProdLines(currentselectedindex) 'lineToAnalyze
                If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
                    .rawProficyProductionData = tmpProdArray
                    .rawProductionData = New ProductionDataset(AllProdLines(currentselectedindex), True)
                End If

                .rawProfStartTime = DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime)
                .rawProfEndTime = _endTime

                If ._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode <> MultiConstraintAnalysis.SingleConstraint Then
                    .rawProficyData = finalRateLossData
                    .rawRateLossData = rateLossData
                    'If My.Settings.populateDTteamsFromPROD Then PROF_setDTteamsFromProd(currentselectedindex)
                    If Not My.Settings.AdvancedSettings_isAvailabilityMode Then PROF_setDTteamsFromProd(currentselectedindex)
                    .rawDowntimeData = New DowntimeDataset(AllProdLines(currentselectedindex), finalRateLossData)
                Else
                    If useThreadingForDataPulling Then
                        .rawProficyData = completeDTarray
                    Else
                        .rawProficyData = third10daysDT
                    End If
                    'If My.Settings.populateDTteamsFromPROD Then PROF_setDTteamsFromProd(currentselectedindex)
                    If Not My.Settings.AdvancedSettings_isAvailabilityMode Then PROF_setDTteamsFromProd(currentselectedindex)
                    If useThreadingForDataPulling Then
                        .rawDowntimeData = New DowntimeDataset(AllProdLines(currentselectedindex), completeDTarray)
                    Else
                        .rawDowntimeData = New DowntimeDataset(AllProdLines(currentselectedindex), third10daysDT)
                    End If
                End If
            End With

            tmpProdArray = Nothing
            first10daysDT = Nothing
            second10daysDT = Nothing
            third10daysDT = Nothing

            prstoryReport = New prStoryMainPageReport(currentselectedindex, prstory_datepicker_startdate.SelectedDate, prstory_datepicker_enddate.SelectedDate)
            MultiLineRawReports.Add(prstoryReport)

        Else 'this means there was a proficy connection error!
            shouldSnakeClose = True
            Thread.Sleep(200)
            ReInitializeAllPublicVariables()
        End If
    End Sub
    Private Sub Do_Analyze_DT(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As String, serverDatabase As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnit
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            serverDatabase = .ServerDatabase
        End With
        STARTx = paramObj(1)
        ENDx = paramObj(2)

        Try

            Select Case preferredDtQuery
                Case DefaultProficyDowntimeProcedure.OneClick
                    third10daysDT = getRawProficyData_OneClick(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    third10daysDT = getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.GLEDS
                    third10daysDT = getGLEDSData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.Maple
                    third10daysDT = getMapleData(ENDx, STARTx, tmpString, prodUnit, serverName, serverUsername, serverPassword, serverDatabase)
            End Select

        Catch ex As System.Runtime.InteropServices.COMException
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION COMException ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Catch ex As Exception
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION DATA ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Finally
        End Try
    End Sub
    Private Sub Do_Analyze_DT1(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As String, serverDatabase As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnit
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            serverDatabase = .ServerDatabase
        End With
        STARTx = DateAdd(DateInterval.Day, -60, paramObj(1)) '-60
        ENDx = DateAdd(DateInterval.Day, -60, paramObj(2)) '-60

        Try

            Select Case preferredDtQuery
                Case DefaultProficyDowntimeProcedure.OneClick
                    first10daysDT = getRawProficyData_OneClick(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    first10daysDT = getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.GLEDS
                    first10daysDT = getGLEDSData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.Maple
                    first10daysDT = getMapleData(ENDx, STARTx, tmpString, prodUnit, serverName, serverUsername, serverPassword, serverDatabase)
            End Select

        Catch ex As System.Runtime.InteropServices.COMException
            ' Debugger.Break()
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION DATA ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Catch ex As Exception
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION DATA ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Finally
        End Try
    End Sub
    Private Sub Do_Analyze_DT2(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date

        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As String, serverDatabase As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnit
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            serverDatabase = .ServerDatabase
        End With

        STARTx = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, paramObj(1)) '-30
        ENDx = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, paramObj(2))    '-30

        Try
            Select Case preferredDtQuery
                Case DefaultProficyDowntimeProcedure.OneClick
                    second10daysDT = getRawProficyData_OneClick(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    second10daysDT = getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.GLEDS
                    second10daysDT = getGLEDSData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.Maple
                    second10daysDT = getMapleData(ENDx, STARTx, tmpString, prodUnit, serverName, serverUsername, serverPassword, serverDatabase)
            End Select
        Catch ex As System.Runtime.InteropServices.COMException
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION DATA ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Catch ex As Exception
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION DATA ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Finally

        End Try
    End Sub
    Private Sub Do_Analyze_Prod(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date)
        Dim STARTx As Date
        Dim ENDx As Date

        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredProdQuery As Integer, prodUnit As String, databaseName As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredProdQuery = .SQLproductionProcedure
            prodUnit = .mainProfProd
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With

        STARTx = paramObj(1)
        ENDx = paramObj(2)

        Try
            Select Case preferredProdQuery
                Case DefaultProficyProductionProcedure.QuickQuery
                    tmpProdArray = getRawProficyProductionData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0)) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0))
                Case DefaultProficyProductionProcedure.SwingRoad
                    tmpProdArray = getRawProficyProductionData_SwingRoad(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyProductionProcedure.Maple
                    tmpProdArray = getMaplePRODData(ENDx, STARTx, prodUnit, prodUnit, serverName, serverUsername, serverPassword, databaseName)
                Case DefaultProficyProductionProcedure.Maple_New
                    tmpProdArray = getMaplePRODData(ENDx, STARTx, prodUnit, prodUnit, serverName, serverUsername, serverPassword, databaseName)


            End Select
        Catch ex As System.Runtime.InteropServices.COMException
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION DATA ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Catch ex As Exception
            If Not PROF_Connectionerror Then
                PROF_Connectionerror = True
                MsgBox("[PRODUCTION DATA ERROR. Try Running In Availability Mode To Avoid This Error] Unable to connect to server. Please check your internet connection and try again. " + ex.Message, vbCritical)
            End If
        Finally
        End Try
    End Sub

#End Region

#Region "MultilineContentUI"
    Public Sub ByLossAreaunplannedbtnclicked(sender As Object, e As MouseButtonEventArgs)

        LossTreeunplannedbtn.Background = mybrushdefaultfontgray
        LossTreeunplannedbtn.Foreground = mybrushlanguagewhite

        LossTreeplannedbtn.Background = mybrushdefaultbackgroundgray
        LossTreeplannedbtn.Foreground = mybrushdefaultfontgray

        Tier1Combo.Items.Clear()
        IsByLossAreaplanned = False
        AddItemstoTier1Combobox()
        ManageTierComboLabelNames()
    End Sub
    Public Sub ByLossAreaplannedbtnclicked(sender As Object, e As MouseButtonEventArgs)
        LossTreeplannedbtn.Background = mybrushdefaultfontgray
        LossTreeplannedbtn.Foreground = mybrushlanguagewhite

        LossTreeunplannedbtn.Background = mybrushdefaultbackgroundgray
        LossTreeunplannedbtn.Foreground = mybrushdefaultfontgray

        Tier1Combo.Items.Clear()
        IsByLossAreaplanned = True
        AddItemstoTier1Combobox()
        ManageTierComboLabelNames()
    End Sub

    Private Sub ManageDatesandTimeforreport()
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)

        Dim i As Integer
        If multiline_Listofdates.Count > 1 Then

            For i = 0 To multiline_Listofdates.Count - 1
                If Output_multilinereport.EndDateList(i) <> multiline_Listofdates_endtime_datetimeformat(i) Then
                    multiline_Listofdates(i) = multiline_Listofdates_starttime(i).ToString & vbCrLf & Output_multilinereport.EndDateList(i).ToString
                    multiline_Listofdates_endtime(i) = Output_multilinereport.EndDateList(i).ToString
                End If
            Next


            For i = 1 To multiline_Listofdates.Count - 1

                If multiline_Listofdates(i) <> multiline_Listofdates(i - 1) Then
                    IsAllDatesSame = False
                Else
                    IsAllDatesSame = True
                End If

            Next
        Else
            DateTimeSelectedLabel.Content = multiline_Listofdates(0)
            DateTimeSelectedLabel.Visibility = Visibility.Visible
            IsAllDatesSame = True
            Exit Sub
        End If


        If IsAllDatesSame = True Then
            DateTimeSelectedLabel.Content = multiline_Listofdates(0)
            DateTimeSelectedLabel.Visibility = Visibility.Visible
        Else
            DateTimeSelectedLabel.Visibility = Visibility.Hidden
        End If


    End Sub
    Private Sub ManageTierComboLabelNames()
        If IsByLossAreaplanned = False Then

            Tier1ComboLabel.Content = "Area"
            Tier2ComboLabel.Content = "Machine"
            Tier3ComboLabel.Content = "Failure Mode"

            If CommonSectorname = SECTOR_FAMILY Then
                Tier1ComboLabel.Content = "Area/Machine"
                Tier2ComboLabel.Content = "Machine Section"
                Tier3ComboLabel.Content = "Fault"

            End If
            If CommonSectorname = SECTOR_BABY Then
                Tier1ComboLabel.Content = "Area"
                Tier2ComboLabel.Content = "Transformation"
                Tier3ComboLabel.Content = "Failure Mode"
            End If
            If CommonSectorname = SECTOR_FEM Then
                Tier1ComboLabel.Content = "Area"
                Tier2ComboLabel.Content = "Transformation"
                Tier3ComboLabel.Content = "Failure Mode"
            End If

        Else
            Tier1ComboLabel.Content = "Area"
            Tier2ComboLabel.Content = "Sub-Area"
            Tier3ComboLabel.Content = ""

        End If
    End Sub

    Private Sub AddItemstoTier1Combobox()
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        Dim i As Integer
        Tier1Combo.Items.Clear()
        Tier2Combo.Items.Clear()
        Tier3Combo.Items.Clear()
        If IsByLossAreaplanned = False Then
            For i = 0 To Output_multilinereport.RollupTier1.Count - 1
                Tier1Combo.Items.Add(Output_multilinereport.RollupTier1(i).Name)
            Next

        Else
            For i = 0 To Output_multilinereport.RollupTier1planned.Count - 1
                Tier1Combo.Items.Add(Output_multilinereport.RollupTier1planned(i).Name)
            Next

        End If




        If Tier1Combo.Items.Count > 0 Then
            Tier1Combo.SelectedIndex = 0
            GenerateSelectedFailureModesLists(Output_multilinereport, Tier1Combo.SelectedValue)
        End If



    End Sub
    Public Sub Tier1Comboselectionchanged()
        UseTrack_Multiline_ByLossAreadrilldown1 = True
        If Tier1Combo.SelectedIndex <> -1 Then
            Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
            Dim i As Integer, t1MTTR As Double = 0
            Tier2Combo.Items.Clear()
            Tier3Combo.Items.Clear()

            Tier2Combo.Items.Add("All")
            Tier3Combo.Items.Add("All")

            If IsByLossAreaplanned = False Then 'UNPLANNED
                Tier1ComboLabel.Content = "Area"
                For i = 0 To Output_multilinereport.RollupTier2.Count - 1
                    If Output_multilinereport.RollupTier2(i).ParentName = Tier1Combo.SelectedValue.ToString Then
                        Tier2Combo.Items.Add(Output_multilinereport.RollupTier2(i).Name)

                    End If
                Next

            Else 'PLANNED
                Dim t1MTTRdt As Double = 0.0
                Dim t1MTTRstops As Double = 0.0
                For i = 0 To Output_multilinereport.RollupTier2planned.Count - 1

                    If Output_multilinereport.RollupTier2planned(i).ParentName = Tier1Combo.SelectedValue.ToString Then
                        Tier2Combo.Items.Add(Output_multilinereport.RollupTier2planned(i).Name)
                        t1MTTRstops += Output_multilinereport.RollupTier2planned(i).Stops
                        t1MTTRdt += Output_multilinereport.RollupTier2planned(i).DT
                    End If
                Next

                If t1MTTRstops > 0 Then
                    t1MTTR = t1MTTRdt / t1MTTRstops
                    Tier1ComboLabel.Content = "Area" ' - " & Math.Round(t1MTTR, 1) & " min MTTR"
                Else
                    Tier1ComboLabel.Content = "Area" ' - " & Math.Round(t1MTTR, 1) & " min MTTR"
                End If

            End If

            Tier2Combo.SelectedValue = "All"
            Tier3Combo.SelectedValue = "All"
            Tier3Combo.IsEnabled = False
        End If
    End Sub
    Public Sub Tier2Comboselectionchanged()
        UseTrack_Multiline_ByLossAreadrilldown2 = True
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        If Tier2Combo.SelectedIndex <> -1 Then
            If Tier2Combo.SelectedItem.ToString <> "All" Then

                Dim i As Integer
                Tier3Combo.IsEnabled = True
                Tier3Combo.Items.Clear()


                Tier3Combo.Items.Add("All")

                If IsByLossAreaplanned = False Then
                    For i = 0 To Output_multilinereport.RollupTier3.Count - 1
                        If Output_multilinereport.RollupTier3(i).ParentName = Tier2Combo.SelectedValue.ToString Then
                            Tier3Combo.Items.Add(Output_multilinereport.RollupTier3(i).Name)
                        End If

                    Next

                Else
                    For i = 0 To Output_multilinereport.RollupTier3planned.Count - 1
                        If Output_multilinereport.RollupTier3planned(i).ParentName = Tier2Combo.SelectedValue.ToString Then
                            Tier3Combo.Items.Add(Output_multilinereport.RollupTier3planned(i).Name)
                        End If

                    Next


                End If


                Tier3Combo.SelectedValue = "All"
                GenerateSelectedFailureModesLists(Output_multilinereport, Tier2Combo.SelectedItem.ToString, "Tier 2")
            Else
                Tier3Combo.IsEnabled = False
                GenerateSelectedFailureModesLists(Output_multilinereport, Tier1Combo.SelectedItem.ToString, "Tier 1")

            End If

        End If
    End Sub
    Public Sub Tier3Comboselectionchanged()
        UseTrack_Multiline_ByLossAreadrilldown3 = True
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        If Tier3Combo.SelectedIndex <> -1 Then

            If Tier3Combo.SelectedItem.ToString <> "All" Then

                GenerateSelectedFailureModesLists(Output_multilinereport, Tier3Combo.SelectedItem.ToString, "Tier 3")
            Else
                If Tier2Combo.SelectedItem.ToString <> "All" Then GenerateSelectedFailureModesLists(Output_multilinereport, Tier2Combo.SelectedItem.ToString, "Tier 2")
            End If
        End If
    End Sub
    Public Sub ByLossAreaComboBoxSelectionChanged()
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        Dim selectedvalue As String
        selectedvalue = ByLossAreaComboBox.SelectedValue
        GenerateSelectedFailureModesLists(Output_multilinereport, selectedvalue, selectedTierRadiobuttoncontent)
    End Sub
    Public Sub GenerateSelectedFailureModesLists(Output_multilinereport As CLS_MultiLineReports, selectedvalue As String, Optional defaulttier As String = "Tier 1")
        Dim selectedfm_dtpctlist As New List(Of Double)
        Dim selectedfm_spdlist As New List(Of Double)
        Dim selectedfm_mtbflist As New List(Of Double)
        Dim selectedfm_stopslist As New List(Of Double)
        Dim selectedfm_MTTRlist As New List(Of Double)
        Dim AnyTierListofLossTreesofEachLine As New List(Of List(Of DTevent))
        Dim AnyTierListofLossTreesofEachLineplanned As New List(Of List(Of DTevent))
        Dim linenum As Integer
        Dim j As Integer
        Dim itemfound As Boolean = False

        If defaulttier = "" Then defaulttier = "Tier 1"  'layer of protection


        Select Case defaulttier
            Case "Tier 1"
                AnyTierListofLossTreesofEachLine = Output_multilinereport.Tier1ListofLossTreeosfEachLine
                AnyTierListofLossTreesofEachLineplanned = Output_multilinereport.Tier1ListofLossTreeosfEachLineplanned
            Case "Tier 2"
                AnyTierListofLossTreesofEachLine = Output_multilinereport.Tier2ListofLossTreeosfEachLine
                AnyTierListofLossTreesofEachLineplanned = Output_multilinereport.Tier2ListofLossTreeosfEachLineplanned
            Case "Tier 3"
                AnyTierListofLossTreesofEachLine = Output_multilinereport.Tier3ListofLossTreeosfEachLine
                AnyTierListofLossTreesofEachLineplanned = Output_multilinereport.Tier3ListofLossTreeosfEachLineplanned
            Case "DTGroup"
                AnyTierListofLossTreesofEachLine = Output_multilinereport.DTGroupListofLossTreeosfEachLine
                AnyTierListofLossTreesofEachLineplanned = Output_multilinereport.DTGroupListofLossTreeosfEachLineplanned


        End Select



        For linenum = 0 To multiline_LISTofselectedindeces.Count - 1
            itemfound = False
            For j = 0 To AnyTierListofLossTreesofEachLine(linenum).Count - 1
                If selectedvalue = AnyTierListofLossTreesofEachLine(linenum)(j).Name Then
                    itemfound = True
                    selectedfm_dtpctlist.Add(AnyTierListofLossTreesofEachLine(linenum)(j).DTpctrounded)
                    selectedfm_spdlist.Add(AnyTierListofLossTreesofEachLine(linenum)(j).SPDrounded)
                    selectedfm_mtbflist.Add(AnyTierListofLossTreesofEachLine(linenum)(j).MTBFrounded)
                    selectedfm_stopslist.Add(AnyTierListofLossTreesofEachLine(linenum)(j).Stops)
                    selectedfm_MTTRlist.Add(AnyTierListofLossTreesofEachLine(linenum)(j).MTTR)
                    GoTo found
                End If

            Next j
            For j = 0 To AnyTierListofLossTreesofEachLineplanned(linenum).Count - 1
                If selectedvalue = AnyTierListofLossTreesofEachLineplanned(linenum)(j).Name Then
                    itemfound = True
                    selectedfm_dtpctlist.Add(AnyTierListofLossTreesofEachLineplanned(linenum)(j).DTpctrounded)
                    selectedfm_spdlist.Add(AnyTierListofLossTreesofEachLineplanned(linenum)(j).SPDrounded)
                    selectedfm_mtbflist.Add(AnyTierListofLossTreesofEachLineplanned(linenum)(j).MTBFrounded)
                    selectedfm_stopslist.Add(AnyTierListofLossTreesofEachLineplanned(linenum)(j).Stops)
                    selectedfm_MTTRlist.Add(AnyTierListofLossTreesofEachLineplanned(linenum)(j).MTTR)
                    Exit For
                End If

            Next j

found:
            If itemfound = False Then
                selectedfm_dtpctlist.Add(0.0)
                selectedfm_spdlist.Add(0.0)
                selectedfm_mtbflist.Add(0.0)
                selectedfm_stopslist.Add(0)
                selectedfm_MTTRlist.Add(0.0)

            End If

        Next linenum

        Dim paramObj(10) As Object
        paramObj(0) = selectedfm_dtpctlist
        paramObj(1) = selectedfm_spdlist
        paramObj(2) = selectedfm_mtbflist
        paramObj(3) = selectedfm_stopslist
        paramObj(4) = multiline_ListofLineNames_fullname
        paramObj(5) = selectedfm_MTTRlist
        paramObj(6) = Tier1Combo.SelectedItem.ToString & "-> "
        paramObj(7) = multiline_Listofdates_starttime
        paramObj(8) = multiline_Listofdates_endtime
        paramObj(9) = IsAllDatesSame
        If Tier2Combo.SelectedIndex <> -1 Then
            paramObj(6) = paramObj(6) & Tier2Combo.SelectedItem.ToString & "-> "
        Else
            paramObj(6) = paramObj(6) & "All" & "-> "
        End If

        If Tier3Combo.SelectedIndex <> -1 Then
            paramObj(6) = paramObj(6) & Tier3Combo.SelectedItem.ToString
        Else
            paramObj(6) = paramObj(6) & "All"
        End If





        MultilineHTMLthreas_bylossareachart = New Thread(AddressOf GenerateByLossAreaCharts)
        MultilineHTMLthreas_bylossareachart.Start(paramObj)

        Thread.Sleep(200)
        ByLossAreaChart.Reload(ignoreCache:=True)
    End Sub

    Private Sub GenerateByLossAreaCharts(ByVal paramObj As Object)

        CreateHTMLMultiLine_Bylossarea(paramObj)


    End Sub
    Private Sub ByLossAreaRadiochecked(sender As Object, e As RoutedEventArgs)
        selectedTierRadiobuttoncontent = sender.content.ToString
        AddItemstoByLossAreaCombobox(sender.content.ToString)
    End Sub
    Private Sub AddItemstoByLossAreaCombobox(Optional Defaulttier As String = "Tier 1", Optional isplanned As Boolean = False)
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        Dim i As Integer
        ByLossAreaComboBox.Items.Clear()

        Select Case Defaulttier

            Case "DTGroup"
                For i = 0 To Output_multilinereport.RollupDTGroup.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupDTGroup(i).Name)
                Next

                For i = 0 To Output_multilinereport.RollupDTGroupplanned.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupDTGroupplanned(i).Name)
                Next



            Case "Tier 1"
                For i = 0 To Output_multilinereport.RollupTier1.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupTier1(i).Name)
                Next

                For i = 0 To Output_multilinereport.RollupTier1planned.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupTier1planned(i).Name)
                Next



            Case "Tier 2"
                For i = 0 To Output_multilinereport.RollupTier2.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupTier2(i).Name)
                Next

                For i = 0 To Output_multilinereport.RollupTier2planned.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupTier2planned(i).Name)
                Next
            Case "Tier 3"
                For i = 0 To Output_multilinereport.RollupTier3.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupTier3(i).Name)
                Next

                For i = 0 To Output_multilinereport.RollupTier3planned.Count - 1
                    ByLossAreaComboBox.Items.Add(Output_multilinereport.RollupTier3planned(i).Name)
                Next


        End Select

        If ByLossAreaComboBox.Items.Count > 0 Then
            ByLossAreaComboBox.SelectedIndex = 0
            GenerateSelectedFailureModesLists(Output_multilinereport, ByLossAreaComboBox.SelectedValue, Defaulttier)
        End If

    End Sub
    Private Sub ShowSummaryCharts(sender As Object, e As MouseButtonEventArgs)
        SummaryChart.Visibility = Visibility.Visible
        SummaryChartbtn.BorderThickness = New Thickness(3, 3, 3, 3)
        SummaryChartbtn.BorderBrush = mybrushdarkgray

        ByLossAreaChart.Visibility = Visibility.Hidden
        LossTreeComboGroupCanvas.Visibility = Visibility.Hidden
        ByLossAreaMenuCanvas.Visibility = Visibility.Hidden
        ByLossAreaChartbtn.BorderThickness = New Thickness(0, 0, 0, 0)
    End Sub

    Private Sub ShowByLossAreaCharts(sender As Object, e As MouseButtonEventArgs)
        ByLossAreaChart.Visibility = Visibility.Visible
        LossTreeComboGroupCanvas.Visibility = Visibility.Visible
        'ByLossAreaMenuCanvas.Visibility = Visibility.Visible
        ByLossAreaChartbtn.BorderThickness = New Thickness(3, 3, 3, 3)
        ByLossAreaChartbtn.BorderBrush = mybrushdarkgray


        SummaryChart.Visibility = Visibility.Hidden
        SummaryChartbtn.BorderThickness = New Thickness(0, 0, 0, 0)
        UseTrack_Multiline_ByLossAreachartsmain = True
    End Sub


    Private Sub GenerateSummaryCharts()
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        Dim paramObj(14) As Object
        paramObj(0) = Output_multilinereport.PRList
        paramObj(1) = Output_multilinereport.SPDList
        paramObj(2) = Output_multilinereport.MTBFList
        paramObj(3) = Output_multilinereport.UPDTList
        paramObj(4) = Output_multilinereport.PDTList
        paramObj(5) = multiline_ListofLineNames_fullname
        paramObj(6) = multiline_Listofdates_starttime
        paramObj(7) = multiline_Listofdates_endtime
        paramObj(8) = IsAllDatesSame
        paramObj(9) = Output_multilinereport.MSUList
        paramObj(10) = Output_multilinereport.CasesList
        paramObj(11) = Output_multilinereport.RateLossList
        paramObj(12) = Output_multilinereport.ListOfAdjustedUnits
        paramObj(13) = Output_multilinereport.MTTRList
        CreateHTMLMultiLine_Summary(paramObj, IsTeamAnalysisinMultiline)

    End Sub
    Private Sub SetupMulti_TEAM_ui_All_Lines()
        ContentCanvas_ListView.Visibility = Visibility.Visible
        MultiLinesReport_InfiniteCanvas.Visibility = Visibility.Visible
        MainScrollViewer.Visibility = Visibility.Visible
        DataViewBtn.Visibility = Visibility.Hidden
        ChartsViewBtn.Visibility = Visibility.Hidden

        AllLines_LineNameandPRLabel.Content = MultiTEAMallUIContent(0).ToString & " - " & MultiTEAMallUIContent(1).ToString
        AllLines_unplannedloss.Content = MultiTEAMallUIContent(2).ToString & " unplanned loss"
        AllLines_plannedloss.Content = MultiTEAMallUIContent(3).ToString & " planned loss"
        AllLines_stopsperday.Content = MultiTEAMallUIContent(4).ToString
        AllLines_mtbf.Content = MultiTEAMallUIContent(5).ToString
        AllLines_schedtime.Content = MultiTEAMallUIContent(6).ToString & " min sched. time"
        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            AllLines_msu.Content = MultiTEAMallUIContent(9).ToString
        Else
            AllLines_msu.Visibility = Visibility.Hidden
        End If
        WeighteddescLabel.Content = ""
        AllLinesLossTreeListBox.Visibility = Visibility.Hidden
        AllLinesMappingLevelComboBox.Visibility = Visibility.Hidden
        AllLinesMappingLevelLabel.Visibility = Visibility.Hidden
        AllLines_UnplannedButton.Visibility = Visibility.Hidden
        AllLines_plannedButton.Visibility = Visibility.Hidden
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)

        Indexoflastpulledline = 0
    End Sub
    Private Sub SetupMultiLineUI_AllLines()
        ContentCanvas_ListView.Visibility = Visibility.Visible
        MultiLinesReport_InfiniteCanvas.Visibility = Visibility.Visible
        MainScrollViewer.Visibility = Visibility.Visible
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        Dim i As Integer

        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            AllLines_LineNameandPRLabel.Content = "All Selected Lines - PR: " & FormatPercent(Output_multilinereport.AvgPR, 1) & "*"
            WeighteddescLabel.Content = "* Volume weighted"
        Else
            AllLines_LineNameandPRLabel.Content = "All Selected Lines - Av.: " & FormatPercent(Output_multilinereport.AvgPR, 1) & "*"
            WeighteddescLabel.Content = "* Schedule time weighted"
        End If

        AllLines_unplannedloss.Content = "Unplanned loss: " & FormatPercent(Output_multilinereport.AvgUPDT, 1) & "*"

        Dim isGRO As Boolean = False 'this variable is to accomodate requests made by Gross Gerrau that may not be optimal for everyone

        For ix As Integer = 0 To multiline_LISTofselectedindeces.Count
            If AllProdLines(multiline_LISTofselectedindeces(i)).SiteName = "GRO" Then
                isGRO = True
            End If

        Next ix

        If isGRO Then
            AllLines_plannedloss.Content = "Planned loss: " & FormatPercent(Output_multilinereport.AvgPDT, 1) & "* Rate loss: " & FormatPercent(Math.Max(1 - Output_multilinereport.AvgPDT - Output_multilinereport.AvgUPDT - Output_multilinereport.AvgPR, 0), 1) & "*"
        Else
            AllLines_plannedloss.Content = "Planned loss: " & FormatPercent(Output_multilinereport.AvgPDT, 1) & "*"
        End If
        AllLines_stopsperday.Content = "Stops per day: " & Math.Round(Output_multilinereport.AvgSPD, 1) '& "*"
        AllLines_mtbf.Content = "MTBF (min) : " & Math.Round(Output_multilinereport.AvgMTBF, 1) '& "*"
        AllLines_mttr.Content = "MTTR (min) : " & Math.Round(Output_multilinereport.AvgMTTR, 1) '& "*"
        AllLines_schedtime.Content = "Total sched. time (min) : " & Math.Round(Output_multilinereport.AvgSchedTime, 1)

        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            AllLines_msu.Content = "Total production(MSU): " & Math.Round(Output_multilinereport.AvgMSU / 1000, 2)
            AllLines_msu.Visibility = Visibility.Visible
        Else
            AllLines_msu.Visibility = Visibility.Hidden
        End If

        DataExport += ",,,,,,,," 'DT%,MTBF,Stops,"
        'now add the tier ones
        Dim masterTier1List As New List(Of String)
        Dim tier1Modes As String = ""
        For i = 0 To Output_multilinereport.RollupTier1.Count - 1
            masterTier1List.Add(Output_multilinereport.RollupTier1(i).Name)
            tier1Modes += Output_multilinereport.RollupTier1(i).Name & ",,,"
        Next

        DataExport = DataExport & tier1Modes & ","


        DataExport += vbCrLf
        DataExport = DataExport & "Line Name, "
        DataExport = DataExport & "PR/Av, "
        DataExport = DataExport & "UPDT, "
        DataExport = DataExport & "PDT, "
        DataExport = DataExport & "SPD, "
        DataExport = DataExport & "MTBF, "
        DataExport = DataExport & "Sched Time, "
        DataExport = DataExport & "MSU, "

        For i = 0 To masterTier1List.Count - 1
            DataExport += "DT (%),MTBF,Stops,"
        Next


        DataExport = DataExport & vbCrLf 'new line

        For i = Indexoflastpulledline + 1 To Output_multilinereport.PRList.Count - 1
            If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, multiline_ListofLineNames_fullname(i) & " - PR: " & FormatPercent(Output_multilinereport.PRList(i), 1), "LineNameandPr", 287, 34, i * 287, 0, True)
            Else
                AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, multiline_ListofLineNames_fullname(i) & " - Av.: " & FormatPercent(Output_multilinereport.PRList(i), 1), "LineNameandPr", 287, 34, i * 287, 0, True)
            End If

            AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, "Unplanned loss: " & FormatPercent(Output_multilinereport.UPDTList(i), 1), "UnplannedLoss", 287, 34, i * 287, 44, False)
            AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, "Planned loss: " & FormatPercent(Output_multilinereport.PDTList(i), 1), "plannedLoss", 287, 34, i * 287, 78, False)
            AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, "Stops per day: " & Math.Round(Output_multilinereport.SPDList(i), 1), "SPD", 287, 34, i * 287, 112, False)
            AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, "MTBF (min): " & Math.Round(Output_multilinereport.MTBFList(i), 1), "MTBF", 287, 34, i * 287, 146, False)
            AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, "Schduled time (min): " & Math.Round(Output_multilinereport.SchedTimeList(i), 1), "SchedTime", 287, 34, i * 287, 181, False)

            DataExport = DataExport & multiline_ListofLineNames_fullname(i) & ", "
            DataExport = DataExport & FormatPercent(Output_multilinereport.PRList(i), 1) & ", "
            DataExport = DataExport & FormatPercent(Output_multilinereport.UPDTList(i), 1) & ", "
            DataExport = DataExport & FormatPercent(Output_multilinereport.PDTList(i), 1) & ", "
            DataExport = DataExport & Math.Round(Output_multilinereport.SPDList(i), 1) & ", "
            DataExport = DataExport & Math.Round(Output_multilinereport.MTBFList(i), 1) & ", "
            DataExport = DataExport & Math.Round(Output_multilinereport.SchedTimeList(i), 1) & ", "

            If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                AddLabelsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, "Production (MSU): " & Math.Round(Output_multilinereport.MSUList(i) / 1000, 2), "MSU", 287, 34, i * 287, 216, False)
                DataExport = DataExport & Math.Round(Output_multilinereport.MSUList(i) / 1000, 2) & ", "
            Else
                DataExport = DataExport & "0, " ' no production
            End If


            'now we need to add the mode data
            Dim tmpList As List(Of DTevent) = MultiLineRawReports(i).MainLEDSReport.DT_Report.getTier1Directory()
            For j = 0 To masterTier1List.Count - 1
                Dim targetIndex As Integer = -1

                For k = 0 To tmpList.Count - 1
                    If tmpList(k).Name = masterTier1List(j) Then
                        targetIndex = k
                        Exit For
                    End If
                Next

                If targetIndex > -1 Then
                    With tmpList(targetIndex)
                        DataExport += .DT * 100 / Math.Max(Output_multilinereport.SchedTimeList(i), 1) & "%," & Math.Round(Output_multilinereport.UptimeList(i) / Math.Max(.Stops, 1), 1) & "," & .Stops & ","
                    End With
                Else
                    DataExport += "0,0,0" & ","
                End If




            Next



            'update raw data export string
            DataExport = DataExport & vbCrLf 'new line

            AddListViewsUI_EachLine(MultiLinesReport_InfiniteCanvas, Output_multilinereport, i, "EachLineLossTree", 277, 205, i * 287, 302, False)

            MultiLinesReport_InfiniteCanvas.Width = 300 + (i * 287)
        Next

        DataExport += vbCrLf

        'back to our regularly scheduled programming
        PopulateLossTreeLists(Output_multilinereport)
        Indexoflastpulledline = multiline_LISTofselectedindeces.Count - 1
    End Sub

    Private Sub AddLabelsUI_EachLine(depcanvas As Canvas, output_multilinereport As CLS_MultiLineReports, reportnum As Integer, labelcontent As String, labelname As String, Lwidth As Double, Lheight As Double, Lleft As Double, LTop As Double, Optional weight As Boolean = False)
        Dim labelinsert As New Label
        depcanvas.Children.Add(labelinsert)
        labelinsert.Name = labelname & reportnum + 1
        labelinsert.Content = labelcontent
        labelinsert.Width = Lwidth
        labelinsert.Height = Lheight
        labelinsert.FontFamily = New Windows.Media.FontFamily("Segoe UI Light")
        labelinsert.FontSize = 16
        If weight = True Then labelinsert.FontWeight = FontWeights.Bold
        labelinsert.Padding = New Thickness(5, 5, 5, 5)
        labelinsert.VerticalContentAlignment = VerticalAlignment.Center
        labelinsert.HorizontalContentAlignment = HorizontalAlignment.Left
        labelinsert.Foreground = mybrushdefaultfontgray
        Canvas.SetLeft(labelinsert, Lleft)
        Canvas.SetTop(labelinsert, LTop)

    End Sub
    Private Sub PopulateLossTreeLists(output_multilinereport As CLS_MultiLineReports, Optional defaulttier As String = "Tier 1")
        Dim i As Integer
        Dim numberofitems As Integer
        _LossTreeList.Clear()

        Select Case defaulttier
            Case "Tier 1"
                numberofitems = output_multilinereport.RollupTier1.Count - 1
                For i = 0 To numberofitems
                    _LossTreeList.Add(output_multilinereport.RollupTier1(i))

                Next
            Case "Tier 2"
                numberofitems = output_multilinereport.RollupTier2.Count - 1
                For i = 0 To numberofitems
                    _LossTreeList.Add(output_multilinereport.RollupTier2(i))
                Next
            Case "Tier 3"
                numberofitems = output_multilinereport.RollupTier3.Count - 1
                For i = 0 To numberofitems
                    _LossTreeList.Add(output_multilinereport.RollupTier3(i))
                Next

            Case "DTGroup"
                numberofitems = output_multilinereport.RollupDTGroup.Count - 1
                For i = 0 To numberofitems
                    _LossTreeList.Add(output_multilinereport.RollupDTGroup(i))
                Next



        End Select
        AllLinesLossTreeListBox.ItemsSource = LossTreeList

    End Sub
    Private Sub PopulateLossTreeListsPlanned(output_multilinereport As CLS_MultiLineReports, Optional defaulttier As String = "Tier 1")
        Dim i As Integer
        Dim numberofitems As Integer
        _LossTreeListplanned.Clear()

        Select Case defaulttier
            Case "Tier 1"
                numberofitems = output_multilinereport.Tier1Directoryplanned.Count - 1
                For i = 0 To numberofitems
                    _LossTreeListplanned.Add(output_multilinereport.RollupTier1planned(i))

                Next
            Case "Tier 2"
                numberofitems = output_multilinereport.Tier2Directoryplanned.Count - 1
                For i = 0 To numberofitems
                    _LossTreeListplanned.Add(output_multilinereport.RollupTier2planned(i))
                Next
            Case "Tier 3"
                numberofitems = output_multilinereport.Tier3Directoryplanned.Count - 1
                For i = 0 To numberofitems
                    _LossTreeListplanned.Add(output_multilinereport.RollupTier3planned(i))
                Next
            Case "DTGroup"
                numberofitems = output_multilinereport.DTgroupDirectoryplanned.Count - 1
                For i = 0 To numberofitems
                    _LossTreeListplanned.Add(output_multilinereport.RollupDTGroupplanned(i))

                Next
        End Select
        AllLinesLossTreeListBox.ItemsSource = _LossTreeListplanned

    End Sub
    Private Sub PopulateLossTreeLists_eachline(output_multilinereport As CLS_MultiLineReports, linenum As Integer, Optional defaulttier As String = "Tier 1")
        Dim i As Integer
        Dim _Losstreelisttemp As New ObservableCollection(Of DTevent)()
        Select Case defaulttier
            Case "Tier 1"
                For i = 0 To output_multilinereport.Tier1ListofLossTreeosfEachLine(linenum).Count - 1
                    _Losstreelisttemp.Add(output_multilinereport.Tier1ListofLossTreeosfEachLine(linenum)(i))
                    '_LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier1ListofLossTreeosfEachLine(linenum)(i)))
                Next
            Case "Tier 2"
                For i = 0 To output_multilinereport.Tier2ListofLossTreeosfEachLine(linenum).Count - 1
                    _Losstreelisttemp.Add(output_multilinereport.Tier2ListofLossTreeosfEachLine(linenum)(i))
                    ' _LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier2ListofLossTreeosfEachLine(linenum)(i)))
                Next
            Case "Tier 3"
                For i = 0 To output_multilinereport.Tier3ListofLossTreeosfEachLine(linenum).Count - 1
                    _Losstreelisttemp.Add(output_multilinereport.Tier3ListofLossTreeosfEachLine(linenum)(i))
                    ' _LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier3ListofLossTreeosfEachLine(linenum)(i)))
                Next
            Case "DTGroup"
                For i = 0 To output_multilinereport.DTGroupListofLossTreeosfEachLine(linenum).Count - 1
                    _Losstreelisttemp.Add(output_multilinereport.DTGroupListofLossTreeosfEachLine(linenum)(i))
                    '_LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier1ListofLossTreeosfEachLine(linenum)(i)))
                Next
        End Select

        If linenum = 0 And Not IsNothing(_LossTreeList_allselectedlines) Then _LossTreeList_allselectedlines.Clear()  ' safety measure to ensure we can over write on a clean slate

        _LossTreeList_allselectedlines.Add(_Losstreelisttemp)

        'Llistview.ItemsSource = _LossTreeList_allselectedlines(linenum)

        Dim clistview As ListView
        For i = 0 To VisualTreeHelper.GetChildrenCount(Me.MultiLinesReport_InfiniteCanvas) - 1
            If Not IsNothing(VisualTreeHelper.GetChild(Me.MultiLinesReport_InfiniteCanvas, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.MultiLinesReport_InfiniteCanvas, i).GetType.ToString, "ListView") > 0 Then
                    clistview = VisualTreeHelper.GetChild(Me.MultiLinesReport_InfiniteCanvas, i)
                    If clistview.Name = "EachLineLossTree" & linenum + 1 Then
                        clistview.ItemsSource = _LossTreeList_allselectedlines(linenum)
                    End If
                End If
            End If
        Next







    End Sub

    Private Sub PopulateLossTreeLists_eachlineplanned(output_multilinereport As CLS_MultiLineReports, linenum As Integer, Optional defaulttier As String = "Tier 1")
        Dim i As Integer
        Dim _Losstreelisttempplanned As New ObservableCollection(Of DTevent)()
        Select Case defaulttier
            Case "Tier 1"
                For i = 0 To output_multilinereport.Tier1ListofLossTreeosfEachLineplanned(linenum).Count - 1
                    _Losstreelisttempplanned.Add(output_multilinereport.Tier1ListofLossTreeosfEachLineplanned(linenum)(i))
                    '_LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier1ListofLossTreeosfeachlineplanned(linenum)(i)))
                Next
            Case "Tier 2"
                For i = 0 To output_multilinereport.Tier2ListofLossTreeosfEachLineplanned(linenum).Count - 1
                    _Losstreelisttempplanned.Add(output_multilinereport.Tier2ListofLossTreeosfEachLineplanned(linenum)(i))
                    ' _LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier2ListofLossTreeosfeachlineplanned(linenum)(i)))
                Next
            Case "Tier 3"
                For i = 0 To output_multilinereport.Tier3ListofLossTreeosfEachLineplanned(linenum).Count - 1
                    _Losstreelisttempplanned.Add(output_multilinereport.Tier3ListofLossTreeosfEachLineplanned(linenum)(i))
                    ' _LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier3ListofLossTreeosfeachlineplanned(linenum)(i)))
                Next
            Case "DTGroup"
                For i = 0 To output_multilinereport.DTGroupListofLossTreeosfEachLineplanned(linenum).Count - 1
                    _Losstreelisttempplanned.Add(output_multilinereport.DTGroupListofLossTreeosfEachLineplanned(linenum)(i))
                    '_LossTreeList_selectedlinetemp.Add(New ObservableCollection(Of DTevent)(output_multilinereport.Tier1ListofLossTreeosfeachlineplanned(linenum)(i)))
                Next
        End Select

        If linenum = 0 And Not IsNothing(_LossTreeList_allselectedlinesplanned) Then _LossTreeList_allselectedlinesplanned.Clear()  ' safety measure to ensure we can over write on a clean slate
        _LossTreeList_allselectedlinesplanned.Add(_Losstreelisttempplanned)

        'Llistview.ItemsSource = _LossTreeList_allselectedlinesplanned(linenum)

        Dim clistview As ListView
        For i = 0 To VisualTreeHelper.GetChildrenCount(Me.MultiLinesReport_InfiniteCanvas) - 1
            If Not IsNothing(VisualTreeHelper.GetChild(Me.MultiLinesReport_InfiniteCanvas, i)) Then
                If InStr(VisualTreeHelper.GetChild(Me.MultiLinesReport_InfiniteCanvas, i).GetType.ToString, "ListView") > 0 Then
                    clistview = VisualTreeHelper.GetChild(Me.MultiLinesReport_InfiniteCanvas, i)
                    If clistview.Name = "EachLineLossTree" & linenum + 1 Then
                        clistview.ItemsSource = _LossTreeList_allselectedlinesplanned(linenum)
                    End If
                End If
            End If
        Next







    End Sub

    Private Sub AddListViewsUI_EachLine(depcanvas As Canvas, output_multilinereport As CLS_MultiLineReports, reportnum As Integer, listviewname As String, Lwidth As Double, Lheight As Double, Lleft As Double, LTop As Double, Optional weight As Boolean = False)
        Dim Listviewinsert As New ListView
        depcanvas.Children.Add(Listviewinsert)
        Listviewinsert.Name = listviewname & reportnum + 1
        Listviewinsert.Width = Lwidth
        Listviewinsert.Height = Lheight
        Listviewinsert.FontFamily = New Windows.Media.FontFamily("Segoe UI Light")
        Listviewinsert.FontSize = 12
        If weight = True Then Listviewinsert.FontWeight = FontWeights.Bold
        Listviewinsert.Padding = New Thickness(0, 0, 0, 0)
        Listviewinsert.HorizontalContentAlignment = HorizontalAlignment.Left
        Listviewinsert.Foreground = mybrushdefaultfontgray
        Canvas.SetLeft(Listviewinsert, Lleft)
        Canvas.SetTop(Listviewinsert, LTop)
        System.Windows.Forms.Application.DoEvents()
        PopulateLossTreeLists_eachline(output_multilinereport, reportnum)
        SetGridViewDynamically(Listviewinsert)
        ' Listviewinsert.

    End Sub
    Private Sub SetGridViewDynamically(listviewinsert As ListView)
        Dim myGridView As New GridView
        myGridView.AllowsColumnReorder = True
        ' myGridView.ColumnHeaderToolTip = "Employee Information"

        Dim gvc1 As New GridViewColumn
        gvc1.DisplayMemberBinding = New Binding("Name")
        gvc1.Header = "Loss Area"
        gvc1.Width = 130
        myGridView.Columns.Add(gvc1)

        Dim gvc2 As New GridViewColumn
        gvc2.DisplayMemberBinding = New Binding("DTpctrounded")

        gvc2.Header = "DT%"
        gvc2.Width = 60
        myGridView.Columns.Add(gvc2)

        Dim gvc3 As New GridViewColumn()
        gvc3.DisplayMemberBinding = New Binding("SPDrounded")
        gvc3.Header = "Stops/Day"
        gvc3.Width = 80
        myGridView.Columns.Add(gvc3)

        Dim gvc4 As New GridViewColumn()
        gvc4.DisplayMemberBinding = New Binding("MTBFrounded")
        gvc4.Header = "MTBF"
        gvc4.Width = 60
        myGridView.Columns.Add(gvc4)

        Dim gvc5 As New GridViewColumn()
        gvc5.DisplayMemberBinding = New Binding("Stops")
        gvc5.Header = "ActualStops"
        gvc5.Width = 80
        myGridView.Columns.Add(gvc5)

        listviewinsert.View = myGridView

        System.Windows.Forms.Application.DoEvents()
        ' listviewinsert.(3) = Format(RS.Fields("Basic").Value, "#0.00")
    End Sub

    Private Sub CloseRollupSplashCanvas(sender As Object, e As MouseButtonEventArgs)
        Rollupsplashcanvas.Visibility = Visibility.Hidden

    End Sub

    Private Sub LaunchRollupSplashCanvas(selectedfailuremode As String, Optional defaulttier As String = "Tier 1")

        Rollupsplashcanvas.Visibility = Visibility.Visible
        Dim Output_multilinereport As New CLS_MultiLineReports(MultiLineRawReports, multiline_LISTofselectedindeces)
        Dim VolumeWeightedSPDList As New List(Of Double)
        Dim VolumeWeightedDTpctList As New List(Of Double)
        Dim linenum As Integer
        Dim j As Integer
        Dim denominator As Double
        Dim paramObj(3) As Object
        Dim tempeventslist As New List(Of List(Of DTevent))

        Select Case defaulttier
            Case "Tier 1"
                If IsRollupplanned = False Then

                    For j = 0 To Output_multilinereport.Tier1ListofLossTreeosfEachLine.Count - 1
                        tempeventslist.Add(Output_multilinereport.Tier1ListofLossTreeosfEachLine(j))
                    Next
                Else
                    For j = 0 To Output_multilinereport.Tier1ListofLossTreeosfEachLineplanned.Count - 1
                        tempeventslist.Add(Output_multilinereport.Tier1ListofLossTreeosfEachLineplanned(j))
                    Next

                End If

            Case "Tier 2"
                If IsRollupplanned = False Then

                    For j = 0 To Output_multilinereport.Tier2ListofLossTreeosfEachLine.Count - 1
                        tempeventslist.Add(Output_multilinereport.Tier2ListofLossTreeosfEachLine(j))
                    Next
                Else
                    For j = 0 To Output_multilinereport.Tier2ListofLossTreeosfEachLineplanned.Count - 1
                        tempeventslist.Add(Output_multilinereport.Tier2ListofLossTreeosfEachLineplanned(j))
                    Next

                End If
            Case "Tier 3"
                If IsRollupplanned = False Then

                    For j = 0 To Output_multilinereport.Tier3ListofLossTreeosfEachLine.Count - 1
                        tempeventslist.Add(Output_multilinereport.Tier3ListofLossTreeosfEachLine(j))
                    Next
                Else
                    For j = 0 To Output_multilinereport.Tier3ListofLossTreeosfEachLineplanned.Count - 1
                        tempeventslist.Add(Output_multilinereport.Tier3ListofLossTreeosfEachLineplanned(j))
                    Next

                End If
            Case "DTGroup"
                If IsRollupplanned = False Then

                    For j = 0 To Output_multilinereport.DTGroupListofLossTreeosfEachLine.Count - 1
                        tempeventslist.Add(Output_multilinereport.DTGroupListofLossTreeosfEachLine(j))
                    Next
                Else
                    For j = 0 To Output_multilinereport.DTGroupListofLossTreeosfEachLineplanned.Count - 1
                        tempeventslist.Add(Output_multilinereport.DTGroupListofLossTreeosfEachLineplanned(j))
                    Next

                End If

        End Select



        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            denominator = Output_multilinereport.AvgMSU
        Else
            denominator = Output_multilinereport.AvgSchedTime
        End If

        For linenum = 0 To multiline_LISTofselectedindeces.Count - 1
            If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                For i = 0 To tempeventslist(linenum).Count - 1
                    If selectedfailuremode = tempeventslist(linenum)(i).Name Then
                        VolumeWeightedDTpctList.Add((tempeventslist(linenum)(i).DTpctspecialrounded * Output_multilinereport.MSUList(linenum)) / denominator)
                        VolumeWeightedSPDList.Add((tempeventslist(linenum)(i).SPDspecialrounded * Output_multilinereport.MSUList(linenum)) / denominator)
                        Exit For
                    End If

                Next i
            Else
                For i = 0 To tempeventslist(linenum).Count - 1
                    If selectedfailuremode = tempeventslist(linenum)(i).Name Then
                        VolumeWeightedDTpctList.Add((tempeventslist(linenum)(i).DTpctspecialrounded * Output_multilinereport.SchedTimeList(linenum)) / denominator)
                        VolumeWeightedSPDList.Add((tempeventslist(linenum)(i).SPDspecialrounded * Output_multilinereport.SchedTimeList(linenum)) / denominator)
                        Exit For
                    End If

                Next i



            End If


        Next linenum


        If VolumeWeightedDTpctList.Count > 0 Then
            RollupChart1.Reload(ignoreCache:=True)
            paramObj(0) = Output_multilinereport.LineNamesList
            paramObj(1) = VolumeWeightedDTpctList
            paramObj(2) = selectedfailuremode
            MultilineHTMLthread_rolluppiechart1 = New Thread(AddressOf CreateHTMLMultiline_RollupChart1_Pie)
            MultilineHTMLthread_rolluppiechart1.Start(paramObj)
            Thread.Sleep(200)
            RollupChart1.Reload(ignoreCache:=True)
            UseTrack_Multiline_RollupCharts = True
        End If

        If VolumeWeightedDTpctList.Count > 0 Then

            RollupChart2.Reload(ignoreCache:=True)
            paramObj(0) = Output_multilinereport.LineNamesList
            paramObj(1) = VolumeWeightedSPDList
            paramObj(2) = selectedfailuremode
            MultilineHTMLthread_rolluppiechart2 = New Thread(AddressOf CreateHTMLMultiline_RollupChart2_Pie)
            MultilineHTMLthread_rolluppiechart2.Start(paramObj)
            Thread.Sleep(200)
            RollupChart2.Reload(ignoreCache:=True)
            UseTrack_Multiline_RollupCharts = True
        End If



    End Sub

    Private Sub Rollupradiochecked(sender As Object, e As RoutedEventArgs)
        If sender Is DTpctradio Then
            RollupChart1.Visibility = Visibility.Visible
            RollupChart2.Visibility = Visibility.Hidden
        Else
            RollupChart2.Visibility = Visibility.Visible
            RollupChart1.Visibility = Visibility.Hidden

        End If

    End Sub



#End Region

#Region "User Analytics"

    Private Sub SendUserAnalyticsDatatoServer_multiline(linename As String)
        Try
            If My.Computer.Network.Ping("prstory.pg.com") Then
                Dim client As MongoClient
                Dim server As MongoServer
                Dim db As MongoDatabase
                Dim col1 As MongoCollection
                client = New MongoClient("mongodb://prstory.pg.com/MongoServer")

                server = client.GetServer()
                db = server.GetDatabase("prstory")
                '  col1 = db.GetCollection(Of BsonDocument)("UseTrack")
                col1 = db("UseTrack")

                Dim currentloginname As String

                currentloginname = Environment.UserName
                Dim NewInfoBson As BsonDocument = New BsonDocument() _
                                                   .Add("who", String.Format(currentloginname)) _
                                                    .Add("when", String.Format(Now(), "MM dd yyyy hh:mm")) _
                                                    .Add("Line", linename) _
                                                    .Add("A", UseTrack_UPDTview) _
                                                    .Add("B", UseTrack_PDTview) _
                                                    .Add("C", UseTrack_PROverallTrends) _
                                                    .Add("D", UseTrack_RawDatawindow_Main) _
                                                    .Add("E", UseTrack_RawDatawindow_Paretos) _
                                                    .Add("F", UseTrack_RawDatawindow_Variance) _
                                                    .Add("G", UseTrack_WeibullMain) _
                                                    .Add("H", UseTrack_WeibullMain_failuremodes) _
                                                    .Add("I", UseTrack_IncontrolMain) _
                                                    .Add("J", UseTrack_IncontrolControlChart) _
                                                    .Add("K", UseTrack_IncontrolControlShift) _
                                                    .Add("L", UseTrack_TopStopsMain) _
                                                    .Add("M", UseTrack_StopsWatchMain) _
                                                    .Add("N", UseTrack_TopStopsTrends) _
                                                    .Add("O", UseTrack_ChangeMapping) _
                                                    .Add("P", UseTrack_Filter) _
                                                    .Add("Q", UseTrack_ExportLossTree) _
                                                    .Add("R", UseTrack_ExportDowntime) _
                                                    .Add("S", UseTrack_ExportProduction) _
                                                    .Add("T", UseTrack_ExportDependency) _
                                                    .Add("U", UseTrack_Notes) _
                                                    .Add("V", UseTrack_Simulation) _
                                                    .Add("W", UseTrack_Notes_PickaLoss) _
                                                    .Add("X", UseTrack_Notes_ExporttoExcel) _
                                                    .Add("Y", UseTrack_TargetsMain) _
                                                    .Add("Z", "oldserver") _
                                                    .Add("Z1", PRSTORY_VERSION_NUMBER) _
                                                    .Add("A0", UseTrack_Multiline_RawData) _
                                                    .Add("A1", UseTrack_Multiline_ByLossAreadrilldown1) _
                                                    .Add("A2", UseTrack_Multiline_ByLossAreadrilldown2) _
                                                    .Add("A3", UseTrack_Multiline_ByLossAreadrilldown3) _
                                                    .Add("A4", UseTrack_Multiline_RollupCharts) _
                                                    .Add("A5", UseTrack_Multiline_Rollupdrilldown) _
                                                    .Add("A6", UseTrack_Multiline_ByLossAreachartsmain)


                col1.Insert(NewInfoBson)
                server.Disconnect()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub


#End Region

#Region "Reinitializationonthefly"

    Sub ReInitializeAllPublicVariables()
        IsRemappingDoneOnce = False

        'bargraphreportwindow_Open = False

        AllProdLines(currentselectedindex).isFilterByBrandcode = False

        datalabelcontent = ""

        '        motionchartsource = 1
        '       bubblenumberpublic = 1

        'MasterDataSet = Nothing
        shouldSnakeClose = False
        PROF_Connectionerror = False
    End Sub

    Private Sub GoToGroupLabel_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If prstory_linedropdown.Visibility = Visibility.Visible Then
            prstory_linedropdown.Visibility = Visibility.Hidden
            RememberBox.Visibility = Visibility.Hidden
            prstory_groupdropdown.Visibility = Visibility.Visible
            SaveGroupLabel.Visibility = Visibility.Visible
            GoToGroupLabel.Content = "Select By Line"
            SelectlineLabel.Content = "Select a group"
            DeleteGroupIcon.Visibility = Visibility.Visible
        Else
            prstory_linedropdown.Visibility = Visibility.Visible
            RememberBox.Visibility = Visibility.Visible
            prstory_groupdropdown.Visibility = Visibility.Hidden
            SaveGroupLabel.Visibility = Visibility.Hidden
            GoToGroupLabel.Content = "Select By Group"
            SelectlineLabel.Content = "Select one or more lines"
            DeleteGroupIcon.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub SaveGroupLabel_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Dim dialog = New Window_TextInput()
        If prstory_linedropdown.SelectedItems.Count > 1 Then
            If (dialog.ShowDialog() = True) Then
                Dim groupName As String = dialog.ResponseText
                If groupName.Length < 2 Then
                    MessageBox.Show("Group Name Must Be At Least 2 Characters", "Name Too Short")
                Else
                    Dim newGroup As ProductionLineGroup = New ProductionLineGroup(groupName)
                    For i = 0 To prstory_linedropdown.SelectedItems.Count - 1
                        newGroup.AddLine((prstory_linedropdown.SelectedItems(i)).ToString())
                    Next

                    'add the group to the active data collection
                    _ActiveDataCollection.Add(newGroup)
                    prstory_groupdropdown.Rebind()

                    'update the local group list
                    JSON_Export(_ActiveDataCollection.ToList())
                    ' dim x as list(of productionlinegroup) = json_import()
                End If
            End If
        Else
            MessageBox.Show("You must select more than 1 line to create a custom group.", "Not Enough Lines Selected")
        End If
    End Sub

    Private Sub prstory_groupdropdown_SelectionChanged(sender As Object, e As Telerik.Windows.Controls.SelectionChangeEventArgs)
        If prstory_groupdropdown.SelectedItems.Count = 1 Then
            'clear all selected lines from the linedropdown
            prstory_linedropdown.SelectedItems.Clear()
            'select the appropriate lines
            For i = 0 To prstory_groupdropdown.SelectedItems(0).Lines.Count - 1
                prstory_linedropdown.SelectedItems.Add(AllProdLines(prstory_groupdropdown.SelectedItems(0).lines(i)))
            Next
        End If
    End Sub

    Public Sub JSON_Export(exportObject As List(Of ProductionLineGroup), Optional FileNameX As String = "C:\\Users\\Public\\prstory\\prstoryLineGroupsJSON", Optional FileType As String = ".txt")

        Dim jsonData As String = JsonConvert.SerializeObject(exportObject)
        Dim FileName As String = FileNameX & FileType
        Dim fcreate As FileStream = File.Open(FileName, FileMode.Create)
        Dim writer As StreamWriter = New StreamWriter(fcreate)

        writer.Write(jsonData)
        writer.Close()
    End Sub

    Public Function JSON_Import_LineGroup(Optional FileNameX As String = "C:\\Users\\Public\\prstory\\prstoryLineGroupsJSON", Optional FileType As String = ".txt") As List(Of ProductionLineGroup)
        Dim json As String = File.ReadAllText(FileNameX & FileType)
        Dim o As List(Of ProductionLineGroup) = JsonConvert.DeserializeObject(Of List(Of ProductionLineGroup))(json)
        Return o
    End Function

    Private Sub DeleteGroupIcon_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If prstory_groupdropdown.SelectedItems.Count = 0 Then
            MessageBox.Show("Please Select A Group.", "No Group Selected")
        Else
            Dim groupName As String = prstory_groupdropdown.SelectedItems(0).name
            For i = 0 To _ActiveDataCollection.Count - 1
                If _ActiveDataCollection(i).Name = groupName Then
                    _ActiveDataCollection.Remove(_ActiveDataCollection(i))
                    i = _ActiveDataCollection.Count
                End If
            Next
            'update the local group list
            JSON_Export(_ActiveDataCollection.ToList())
        End If
    End Sub

#End Region

End Class
