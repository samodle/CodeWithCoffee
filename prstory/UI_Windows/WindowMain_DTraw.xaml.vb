Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Xml 'not sure if i need thsi
Imports System.Threading
Imports System.IO
Imports Awesomium.Core
Imports Awesomium.Windows.Controls
Imports System.Globalization
Imports Telerik.Windows.Controls.ChartView
Imports Telerik.Charting

Public Class RawDataWindow
#Region "Variables"
    Private bw As BackgroundWorker = New BackgroundWorker

    Private parentLine As ProdLine
    Private prstoryReport As prStoryMainPageReport
    Private RowsList As List(Of Integer)
    Private bucketName As String
    Private cardNum
    Private titleName As String

    Private mybrushTabSelected As New SolidColorBrush(Color.FromRgb(44, 152, 195))
    '  Private mybrushTabSelected As New SolidColorBrush(Color.FromRgb(44, 152, 195))
    Private mybrushTabNOTSelected As New SolidColorBrush(Color.FromRgb(202, 221, 228))

    'sorting our listview
    Private _lastHeaderClicked_ActiveData As GridViewColumnHeader = Nothing
    Private _lastDirection_ActiveData As ListSortDirection = ListSortDirection.Ascending
    Private _sortUp As Boolean = True
    Private _lastSortedField As Integer = -1



    ''Trends vars
    Public failuremodeno As Integer
    Public IsSPDActive As Boolean = False
    Public ISDTActive As Boolean = False
    Public ISMTBFActive As Boolean = False
    Public IsLaunchedfromstops_InMOtionChart As Boolean = True
    Public selectedfailuremode_inMotionChart As Integer = 0
    Public IsSShapeActive As Boolean = True
    ''' <summary>
    ''' ''''''''''
    ''' </summary>

    Dim _ActiveDataCollection As New ObservableCollection(Of DowntimeEvent)()
    Dim _ActiveDataSortList As New List(Of DowntimeEvent)
    Public ReadOnly Property ActiveDataCollection() As ObservableCollection(Of DowntimeEvent)
        Get
            Return _ActiveDataCollection
        End Get
    End Property

    Dim _ActiveProductionCollection As New ObservableCollection(Of ProductionEvent)()
    Public ReadOnly Property ActiveProductionCollection() As ObservableCollection(Of ProductionEvent)
        Get
            Return _ActiveProductionCollection
        End Get
    End Property
#End Region

#Region "Show / Hide Displays"
    Private Sub RawDTWindow_onload()
        Cursor = Cursors.Arrow 'LG Code
        ShowRawData()
        HidePareto()
        HideVarianceTab()
        HideTrendsTab()
        HideProdTab()
        Managescreenresolution()
        ManageViewsforPlanned()
        ' ManageColumnDeleteforFamily()
    End Sub
    Private Sub ManageViewsforPlanned()
        If cardNum = 2 Or cardNum = 41 Then
            MotionTrendsTab.Visibility = Visibility.Hidden
        End If

    End Sub

    Private Sub ManageColumnDeleteforFamily()

        If AllProdLines(selectedindexofLine_temp).Sector.ToString.Contains("Fam") Then

            ActiveDataList.Items(6).Width = 0
            ActiveDataList.Items(7).Width = 0
        End If
    End Sub
    Private Sub Managescreenresolution()
        Dim screenWidth As Integer = My.Computer.Screen.Bounds.Width
        Dim screenHeight As Integer = My.Computer.Screen.Bounds.Height

        If screenWidth < 1290 Or screenHeight < 635 Then Me.WindowState = Windows.WindowState.Maximized
    End Sub

#Region "Show/Hide Tabs"

    Private Sub RawData_Prod_TabClicked()
        HideRawData()
        HidePareto()
        ShowProdTab()
        HideVarianceTab()
        HideTrendsTab()

    End Sub
    Private Sub RawDataTabClicked()
        ShowRawData()
        HidePareto()
        HideProdTab()
        HideVarianceTab()
        HideTrendsTab()

    End Sub
    Private Sub ParetoTabClicked()
        UseTrack_RawDatawindow_Paretos = True
        HideRawData()
        HideProdTab()
        ShowPareto()
        HideVarianceTab()
        HideTrendsTab()

    End Sub
    Private Sub VarianceTabClicked()
        UseTrack_RawDatawindow_Variance = True
        HideRawData()
        HideProdTab()
        HidePareto()
        '   ShowVarianceTab()
        HideTrendsTab()

    End Sub
    Private Sub TrendsTabClicked()
        UseTrack_RawDataWindow_Trends = True
        HideRawData()
        HideProdTab()
        HidePareto()
        HideVarianceTab()
        ShowTrendsTab()

    End Sub
    Private Sub HideTrendsTab()
        TrendsCanvas.Visibility = Visibility.Hidden
        ' Dispose_AllTrendCharts_inRawDataWindow()
        MotionTrendsTab.Background = mybrushTabNOTSelected
    End Sub
    Private Sub ShowTrendsTab()
        MotionTrendsTab.Background = mybrushTabSelected
        TrendsCanvas.Visibility = Visibility.Visible
        LaunchtrendsinRawDataWindow()
        Trends_OnStart()
    End Sub
    Private Sub HideVarianceTab()
        '  VarianceHTML.Visibility = Windows.Visibility.Hidden
        CandlestickTab.Background = mybrushTabNOTSelected
    End Sub
    Private Sub ShowVarianceTab()
        CandlestickTab.Background = mybrushTabSelected
        '  VarianceHTML.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub HideProdTab()
        RawDataTab_Prod.Background = mybrushTabNOTSelected
        'ParetoChart.Visibility = Windows.Visibility.Hidden
        ActiveProdDataList.Visibility = Visibility.Hidden

    End Sub
    Private Sub ShowProdTab()
        RawDataTab_Prod.Background = mybrushTabSelected
        ParetoHTML.Visibility = Visibility.Hidden
        ActiveProdDataList.Visibility = Visibility.Visible
        ExportRawDatabutton.Visibility = Visibility.Visible
    End Sub
    Private Sub HidePareto()
        ParetoTab.Background = mybrushTabNOTSelected
        ParetoHTML.Visibility = Visibility.Hidden
        ReasonSelection.Visibility = Visibility.Hidden
        ParetoMainLabel.Visibility = Visibility.Hidden
    End Sub
    Private Sub ShowPareto()
        ParetoTab.Background = mybrushTabSelected
        ParetoHTML.Visibility = Visibility.Visible

        ReasonSelection.Visibility = Visibility.Visible
        ParetoMainLabel.Visibility = Visibility.Visible
    End Sub
    Private Sub HideRawData()
        ActiveDataList.Visibility = Visibility.Hidden
        RawDataTab_DT.Background = mybrushTabNOTSelected
        ExportRawDatabutton.Visibility = Visibility.Hidden
        ActiveDataList.Visibility = Visibility.Hidden
    End Sub
    Private Sub ShowRawData()
        ActiveDataList.Visibility = Visibility.Visible
        RawDataTab_DT.Background = mybrushTabSelected
        ExportRawDatabutton.Visibility = Visibility.Visible
        ActiveDataList.Visibility = Visibility.Visible

    End Sub
#End Region

    Private Sub Dispose_AllTrendCharts_inRawDataWindow()
        MotionChartD.Dispose()
        MotionChartS.Dispose()
        MotionChartD_Monthly.Dispose()
        MotionChartS_Monthly.Dispose()
        MotionChartD_Weekly.Dispose()
        MotionChartS_Weekly.Dispose()
        MotionChart_MTBF.Dispose()
        MotionChart_MTBF_Weekly.Dispose()
        MotionChart_MTBF_Monthly.Dispose()
    End Sub
    Sub toggleDTprod()
        With prodDTtoggleButton
            If .Content.contains("Prod") Then
                .Content = "View Downtime Data"
                ActiveProdDataList.Visibility = Visibility.Visible
                ActiveDataList.Visibility = Visibility.Hidden
            Else
                .Content = "View Production Data"
                ActiveProdDataList.Visibility = Visibility.Hidden
                ActiveDataList.Visibility = Visibility.Visible
            End If
        End With
    End Sub
#End Region

    Public Sub CSV_exportVisibleDataList()
        If RawDataTab_DT.Background Is mybrushTabSelected Then CSV_exportRawLEDsDataFromList(parentLine, ActiveDataCollection, prstoryReport.StartTime, prstoryReport.EndTime, bucketName)
        If RawDataTab_Prod.Background Is mybrushTabSelected Then CSV_exportRawProdDataFromList(parentLine, ActiveProductionCollection, prstoryReport.StartTime, prstoryReport.EndTime, bucketName)

        CSV_exportRawLEDsDataFromList(parentLine, ActiveDataCollection, prstoryReport.StartTime, prstoryReport.EndTime, bucketName)
    End Sub

#Region "Setting Data List For Charts"
    Friend ExportListAM As New List(Of DTevent)


    Sub dropDown_SelectionChange()
        Select Case ReasonSelection.SelectedValue.ToString
            Case "Tier 1"
                setExportListAM(DowntimeField.Tier1)
            Case "Tier 2"
                setExportListAM(DowntimeField.Tier2)
            Case "Tier 3"
                setExportListAM(DowntimeField.Tier3)
            Case "Reason 1"
                setExportListAM(DowntimeField.Reason1)
            Case "Reason 2"
                setExportListAM(DowntimeField.Reason2)
            Case "Reason 3"
                setExportListAM(DowntimeField.Reason3)
            Case "Reason 4"
                setExportListAM(DowntimeField.Reason4)
            Case "Fault"
                setExportListAM(DowntimeField.Fault)
            Case "Location"
                setExportListAM(DowntimeField.Location)
            Case "DT Group"
                setExportListAM(DowntimeField.DTGroup)
            Case "SKU (Product Code)"
                setExportListAM(DowntimeField.ProductCode)
            Case "SKU (Product Name)"
                setExportListAM(DowntimeField.Product)
            Case "Team"
                setExportListAM(DowntimeField.Team)


        End Select
        ' System.Threading.Thread.Sleep(200)
        ' ParetoHTML.Reload(ignoreCache:=True)

    End Sub


    'sets the export list
    Sub setExportListAM(dtField As Integer)
        ' Dim HTMLthread As Thread
        Dim i As Integer
        ExportListAM.Clear()
        Select Case dtField
            Case DowntimeField.Reason1
                For i = 0 To Reason1Directory.Count - 1
                    ExportListAM.Add(Reason1Directory(i))
                Next
            Case DowntimeField.Reason2
                For i = 0 To Reason2Directory.Count - 1
                    ExportListAM.Add(Reason2Directory(i))
                Next
            Case DowntimeField.Reason3
                For i = 0 To Reason3Directory.Count - 1
                    ExportListAM.Add(Reason3Directory(i))
                Next
            Case DowntimeField.Reason4
                For i = 0 To Reason4Directory.Count - 1
                    ExportListAM.Add(Reason4Directory(i))
                Next
            Case DowntimeField.Fault
                For i = 0 To FaultDirectory.Count - 1
                    ExportListAM.Add(FaultDirectory(i))
                Next
            Case DowntimeField.DTGroup
                For i = 0 To DTgroupDirectory.Count - 1
                    ExportListAM.Add(DTgroupDirectory(i))
                Next
            Case DowntimeField.Location
                For i = 0 To LocationDirectory.Count - 1
                    ExportListAM.Add(LocationDirectory(i))
                Next
            Case DowntimeField.ProductCode
                For i = 0 To SKUDirectory.Count - 1
                    ExportListAM.Add(SKUDirectory(i))
                Next
            Case DowntimeField.Product
                For i = 0 To SKU_withdescDirectory.Count - 1
                    ExportListAM.Add(SKU_withdescDirectory(i))
                Next
            Case DowntimeField.Tier1
                For i = 0 To Tier1Directory.Count - 1
                    ExportListAM.Add(Tier1Directory(i))
                Next
            Case DowntimeField.Tier2
                For i = 0 To Tier2Directory.Count - 1
                    ExportListAM.Add(Tier2Directory(i))
                Next
            Case DowntimeField.Tier3
                For i = 0 To Tier3Directory.Count - 1
                    ExportListAM.Add(Tier3Directory(i))
                Next
            Case DowntimeField.Team
                For i = 0 To TeamDirectory.Count - 1
                    ExportListAM.Add(TeamDirectory(i))
                Next
        End Select

        ExportListAM.Sort()
        For i = 0 To ExportListAM.Count - 1
            ExportListAM(i).DTpct = prstoryReport.schedTime
            ExportListAM(i).UT = (prstoryReport.schedTime - (prstoryReport.UPDT * prstoryReport.schedTime) - (prstoryReport.PDT * prstoryReport.schedTime) - (prstoryReport.rateLoss * prstoryReport.schedTime))
        Next

        '  Dim paramObj(1) As Object
        '  paramObj(0) = ExportListAM
        If ExportListAM.Count > 0 Then
            ' HTMLthread = New Thread(AddressOf exportAMchart)
            ' HTMLthread.Start(paramObj)

            exportAMchart_Telerik()
        Else
            ParetoTab.Visibility = Visibility.Hidden
        End If
    End Sub

    'all populated on raw data window initialization
    Friend FaultDirectory As New List(Of DTevent)
    Friend Reason1Directory As New List(Of DTevent)

    Friend DTgroupDirectory As New List(Of DTevent)
    Friend LocationDirectory As New List(Of DTevent)

    Friend Reason2Directory As New List(Of DTevent)
    Friend Reason3Directory As New List(Of DTevent)
    Friend Reason4Directory As New List(Of DTevent)

    Friend TeamDirectory As New List(Of DTevent)
    Friend SKUDirectory As New List(Of DTevent)
    Friend Tier1Directory As New List(Of DTevent)
    Friend Tier2Directory As New List(Of DTevent)
    Friend Tier3Directory As New List(Of DTevent)

    Friend SKU_withdescDirectory As New List(Of DTevent)  ' LG code experiment



    'Friend SKUdirectory As New List(Of dtevent)

    Private Sub createParetoDirectories_Unplanned()
        Dim tmpIndex As Integer
        'look at all the unplanned data
        For i As Integer = 0 To prstoryReport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData.Count - 1
            If RowsList.IndexOf(i) > -1 Then
                With prstoryReport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData(i)
                    'R1
                    tmpIndex = Reason1Directory.IndexOf(New DTevent(.Reason1, 0))
                    If tmpIndex = -1 Then
                        Reason1Directory.Add(New DTevent(.Reason1, .DT, i))
                    Else
                        Reason1Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'R2
                    tmpIndex = Reason2Directory.IndexOf(New DTevent(.Reason2, 0))
                    If tmpIndex = -1 Then
                        Reason2Directory.Add(New DTevent(.Reason2, .DT, i))
                    Else
                        Reason2Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'R3
                    tmpIndex = Reason3Directory.IndexOf(New DTevent(.Reason3, 0))
                    If tmpIndex = -1 Then
                        Reason3Directory.Add(New DTevent(.Reason3, .DT, i))
                    Else
                        Reason3Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'R4
                    tmpIndex = Reason4Directory.IndexOf(New DTevent(.Reason4, 0))
                    If tmpIndex = -1 Then
                        Reason4Directory.Add(New DTevent(.Reason4, .DT, i))
                    Else
                        Reason4Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'fault
                    tmpIndex = FaultDirectory.IndexOf(New DTevent(.Fault, 0))
                    If tmpIndex = -1 Then
                        FaultDirectory.Add(New DTevent(.Fault, .DT, i))
                    Else
                        FaultDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'dtgroup
                    tmpIndex = DTgroupDirectory.IndexOf(New DTevent(.DTGroup, 0))
                    If tmpIndex = -1 Then
                        DTgroupDirectory.Add(New DTevent(.DTGroup, .DT, i))
                    Else
                        DTgroupDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'location
                    tmpIndex = LocationDirectory.IndexOf(New DTevent(.Location, 0))
                    If tmpIndex = -1 Then
                        LocationDirectory.Add(New DTevent(.Location, .DT, i))
                    Else
                        LocationDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'sku
                    tmpIndex = SKUDirectory.IndexOf(New DTevent(.ProductCode, 0))
                    If tmpIndex = -1 Then
                        SKUDirectory.Add(New DTevent(.ProductCode, .DT, i))
                    Else
                        SKUDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'tier1
                    tmpIndex = Tier1Directory.IndexOf(New DTevent(.Tier1, 0))
                    If tmpIndex = -1 Then
                        Tier1Directory.Add(New DTevent(.Tier1, .DT, i))
                    Else
                        Tier1Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'tier2
                    tmpIndex = Tier2Directory.IndexOf(New DTevent(.Tier2, 0))
                    If tmpIndex = -1 Then
                        Tier2Directory.Add(New DTevent(.Tier2, .DT, i))
                    Else
                        Tier2Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'tier3
                    tmpIndex = Tier3Directory.IndexOf(New DTevent(.Tier3, 0))
                    If tmpIndex = -1 Then
                        Tier3Directory.Add(New DTevent(.Tier3, .DT, i))
                    Else
                        Tier3Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'team
                    tmpIndex = TeamDirectory.IndexOf(New DTevent(.Team, 0))
                    If tmpIndex = -1 Then
                        TeamDirectory.Add(New DTevent(.Team, .DT, i))
                    Else
                        TeamDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'SKU_withdescription
                    tmpIndex = SKU_withdescDirectory.IndexOf(New DTevent(.ProductCode & "-" & .Product, 0))
                    If tmpIndex = -1 Then
                        SKU_withdescDirectory.Add(New DTevent(.ProductCode & "-" & .Product, .DT, i))
                    Else
                        SKU_withdescDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If

                End With
            End If
        Next

    End Sub

    Private Sub CreateParetoDirectories_Planned()
        Dim tmpIndex As Integer
        'look at all the planned data
        For i As Integer = 0 To prstoryReport.MainLEDSReport.DT_Report.rawDTdata.PlannedData.Count - 1
            If RowsList.IndexOf(i) > -1 Then
                With prstoryReport.MainLEDSReport.DT_Report.rawDTdata.PlannedData(i)
                    'R1
                    tmpIndex = Reason1Directory.IndexOf(New DTevent(.Reason1, 0))
                    If tmpIndex = -1 Then
                        Reason1Directory.Add(New DTevent(.Reason1, .DT, i))
                    Else
                        Reason1Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'R2
                    tmpIndex = Reason2Directory.IndexOf(New DTevent(.Reason2, 0))
                    If tmpIndex = -1 Then
                        Reason2Directory.Add(New DTevent(.Reason2, .DT, i))
                    Else
                        Reason2Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'R3
                    tmpIndex = Reason3Directory.IndexOf(New DTevent(.Reason3, 0))
                    If tmpIndex = -1 Then
                        Reason3Directory.Add(New DTevent(.Reason3, .DT, i))
                    Else
                        Reason3Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'R4
                    tmpIndex = Reason4Directory.IndexOf(New DTevent(.Reason4, 0))
                    If tmpIndex = -1 Then
                        Reason4Directory.Add(New DTevent(.Reason4, .DT, i))
                    Else
                        Reason4Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'fault
                    tmpIndex = FaultDirectory.IndexOf(New DTevent(.Fault, 0))
                    If tmpIndex = -1 Then
                        FaultDirectory.Add(New DTevent(.Fault, .DT, i))
                    Else
                        FaultDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'dtgroup
                    tmpIndex = DTgroupDirectory.IndexOf(New DTevent(.DTGroup, 0))
                    If tmpIndex = -1 Then
                        DTgroupDirectory.Add(New DTevent(.DTGroup, .DT, i))
                    Else
                        DTgroupDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'location
                    tmpIndex = LocationDirectory.IndexOf(New DTevent(.Location, 0))
                    If tmpIndex = -1 Then
                        LocationDirectory.Add(New DTevent(.Location, .DT, i))
                    Else
                        LocationDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'tier1
                    tmpIndex = Tier1Directory.IndexOf(New DTevent(.Tier1, 0))
                    If tmpIndex = -1 Then
                        Tier1Directory.Add(New DTevent(.Tier1, .DT, i))
                    Else
                        Tier1Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'tier2
                    tmpIndex = Tier2Directory.IndexOf(New DTevent(.Tier2, 0))
                    If tmpIndex = -1 Then
                        Tier2Directory.Add(New DTevent(.Tier2, .DT, i))
                    Else
                        Tier2Directory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'team
                    tmpIndex = TeamDirectory.IndexOf(New DTevent(.Team, 0))
                    If tmpIndex = -1 Then
                        TeamDirectory.Add(New DTevent(.Team, .DT, i))
                    Else
                        TeamDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'SKU
                    tmpIndex = SKUDirectory.IndexOf(New DTevent(.ProductCode, 0))
                    If tmpIndex = -1 Then
                        SKUDirectory.Add(New DTevent(.ProductCode, .DT, i))
                    Else
                        SKUDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                    'SKU_withdescription
                    tmpIndex = SKU_withdescDirectory.IndexOf(New DTevent(.ProductCode & "-" & .Product, 0))
                    If tmpIndex = -1 Then
                        SKU_withdescDirectory.Add(New DTevent(.ProductCode & "-" & .Product, .DT, i))
                    Else
                        SKU_withdescDirectory(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End With
            End If
        Next
    End Sub
#End Region

    Public Sub updateValues(ByVal cardNumber As Integer, ByVal rawRows As List(Of Integer), ByVal storyReport As prStoryMainPageReport, titleNam As String)
        ' This call is required by the designer.
        InitializeComponent()

        RowsList = rawRows
        cardNum = cardNumber
        prstoryReport = storyReport
        TitleLabel.Content = titleName
        titleName = titleNam
        Title = titleNam + " - Downtime Explorer"
    End Sub

#Region "Construction"
    Public Sub New(ByVal cardNumber As Integer, ByVal rawRows As List(Of Integer), ByVal storyReport As prStoryMainPageReport, titleNam As String)
        ' This call is required by the designer.
        InitializeComponent()

        RowsList = rawRows
        cardNum = cardNumber
        prstoryReport = storyReport
        TitleLabel.Content = titleName
        titleName = titleNam

        FinishConstruction()
    End Sub

    Public Sub FinishConstruction()

        'retain old mapping 
        mappinglevel1 = My.Settings.defaultDownTimeField
        mappinglevel2 = My.Settings.defaultDownTimeField_Secondary

        ''

        IsRemappingDoneOnce = True
        Select Case cardNum
            case -1
                'dont do anything
            Case 1
                My.Settings.defaultDownTimeField = DowntimeField.Tier1
                My.Settings.defaultDownTimeField_Secondary = -1
            Case 3
                My.Settings.defaultDownTimeField = DowntimeField.Tier2
                My.Settings.defaultDownTimeField_Secondary = -1
            Case 4
                My.Settings.defaultDownTimeField = DowntimeField.Tier3
                My.Settings.defaultDownTimeField_Secondary = -1
            Case 5
                My.Settings.defaultDownTimeField = DowntimeField.Tier3
                My.Settings.defaultDownTimeField_Secondary = -1
            Case 6
                My.Settings.defaultDownTimeField = DowntimeField.Tier3
                My.Settings.defaultDownTimeField_Secondary = -1
        End Select

        

        '''''''''''''''''''''''''''''''''''''''''''''

        CandlestickTab.Visibility = Visibility.Hidden
        'If cardNum <> prStoryCard.Changeover Then CandlestickTab.Visibility = Windows.Visibility.Hidden

        parentLine = AllProdLines(prstoryReport.ParentLineInt)

        prstoryReport.reMapReport()
        AllProdLines(selectedindexofLine_temp).reMapRawData()

        'production data
        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            If (Not prstoryReport.MainLEDSReport.PROD_Report.isNothing_rawProdData) Then
                With prstoryReport.MainLEDSReport.PROD_Report
                    If .RawData.Count > 0 Then
                        For i As Integer = 0 To .RawData.Count - 1
                            _ActiveProductionCollection.Add(.RawData(i))
                        Next
                    End If
                End With
            Else
                RawDataTab_Prod.Visibility = Visibility.Hidden
            End If
        Else
            RawDataTab_Prod.Visibility = Visibility.Hidden
        End If

        'DT
        If RowsList.Count > 0 Then
            bucketName = ""
            If RowsList.Count > 0 Then
                'RowsList = FirstRowsList
                If cardNum = prStoryCard.Changeover Or cardNum = prStoryCard.Planned Then
                    generateCollectionFromRows_Planned()
                    CreateParetoDirectories_Planned()
                Else
                    generateCollectionFromRows_Unplanned()
                    createParetoDirectories_Unplanned()
                End If


                setExportListAM(DowntimeField.Reason1)

                'set dropdown
                ReasonSelection.Items.Add("Tier 1")
                ReasonSelection.Items.Add("Tier 2")
                ReasonSelection.Items.Add("Tier 3")
                ReasonSelection.Items.Add("SKU (Product Code)")
                ReasonSelection.Items.Add("SKU (Product Name)")
                ReasonSelection.Items.Add("Team")
                ReasonSelection.Items.Add("Reason 1")
                ReasonSelection.Items.Add("Reason 2")
                If Not AllProdLines(selectedindexofLine_temp).Sector.Contains("Fam") Or (AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyMaking) Then
                    ReasonSelection.Items.Add("Reason 3")
                    ReasonSelection.Items.Add("Reason 4")
                End If
                ReasonSelection.Items.Add("Fault")
                ReasonSelection.Items.Add("Location")
                ReasonSelection.Items.Add("DT Group")


                ReasonSelection.SelectedIndex = 0

                If cardNum = 1 And titleName <> "Total" Then
                    ReasonSelection.SelectedIndex = 1
                ElseIf cardNum = 1 And titleName = "Total" Then
                    ReasonSelection.SelectedIndex = 0
                    MotionTrendsTab.Visibility = Visibility.Hidden
                ElseIf cardNum = 2 And titleName <> "Total" Then
                    ReasonSelection.SelectedIndex = 1
                ElseIf cardNum = 2 And titleName = "Total" Then
                    ReasonSelection.SelectedIndex = 0
                ElseIf cardNum = 3 Then
                    ReasonSelection.SelectedIndex = 2
                ElseIf cardNum = 4 Or cardNum = 5 Or cardNum = 6 Then
                    If Not AllProdLines(selectedindexofLine_temp).Sector.Contains("Fam") Then
                        ReasonSelection.SelectedIndex = 9
                    Else
                        ReasonSelection.SelectedIndex = 8
                    End If

                End If

                ' TitleLabel.Content = titleName
                RawDTWindow_onload()
            Else
                MsgBox("Error! you wanted to see raw data...but there was none :(")
                Debugger.Break()
            End If
        Else
            MsgBox("No Stops For Selected Field!", vbExclamation, "Mode Selected With No Stops/DT")
        End If

        'BusyIndicator.IsBusy = False
    End Sub
    'populate / repopulate our collection
    Public Sub regenerateRawData(NewRowsList As List(Of Integer))
        ActiveDataCollection.Clear()
        RowsList = NewRowsList
        If cardNum = prStoryCard.Changeover Or cardNum = prStoryCard.Planned Then
            generateCollectionFromRows_Planned()
        Else
            generateCollectionFromRows_Unplanned()
        End If
    End Sub

    Private Sub generateCollectionFromRows_Planned()
        Dim listIncrementer As Integer, actualRow As Integer
        For listIncrementer = 0 To RowsList.Count - 1
            actualRow = RowsList(listIncrementer)
            ActiveDataCollection.Add(prstoryReport.MainLEDSReport.DT_Report.rawDTdata.PlannedData(actualRow))  'New DTeventComplete(parentLine, actualRow))
        Next
    End Sub
    Private Sub generateCollectionFromRows_Unplanned()
        Dim listIncrementer As Integer, actualRow As Integer
        For listIncrementer = 0 To RowsList.Count - 1
            actualRow = RowsList(listIncrementer)
            ActiveDataCollection.Add(prstoryReport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData(actualRow))
        Next
    End Sub
#End Region

#Region "Column Sorting"
    Private Sub GridViewColumnHeaderClickedHandler_activedata(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim headerClicked As GridViewColumnHeader = TryCast(e.OriginalSource, GridViewColumnHeader)
        ' Dim direction As ListSortDirection
        Dim tmpString As String, listIncrementer As Integer
        Dim currentSortField As Integer = -1
        '  Dim eventsToTransfer As Integer

        If headerClicked IsNot Nothing Then
            If headerClicked.Role <> GridViewColumnHeaderRole.Padding Then

                _ActiveDataSortList.Clear()

                For listIncrementer = 0 To ActiveDataCollection.Count - 1
                    _ActiveDataSortList.Add(ActiveDataCollection(listIncrementer))
                Next


                tmpString = headerClicked.Content
                Select Case tmpString
                    Case "DT"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.DT)
                        Next
                        currentSortField = DowntimeField.DT
                    Case "UT"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.UT)
                        Next
                        currentSortField = DowntimeField.UT
                    Case "Start Time"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.startTime)
                        Next
                        currentSortField = DowntimeField.startTime
                    Case "End Time"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.endTime)
                        Next
                        currentSortField = DowntimeField.endTime
                    Case "Reason 1"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.Reason1)
                        Next
                        currentSortField = DowntimeField.Reason1
                    Case "Reason 2"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.Reason2)
                        Next
                        currentSortField = DowntimeField.Reason2
                    Case "Reason 3"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.Reason3)
                        Next
                        currentSortField = DowntimeField.Reason3
                    Case "Reason 4"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.Reason4)
                        Next
                        currentSortField = DowntimeField.Reason4
                    Case "Fault"
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataSortList(listIncrementer).setSortParam(DowntimeField.Fault)
                        Next
                        currentSortField = DowntimeField.Fault
                End Select

                _ActiveDataSortList.Sort()
                _ActiveDataCollection.Clear()

                If currentSortField <> _lastSortedField Then
                    For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                        _ActiveDataCollection.Add(_ActiveDataSortList(listIncrementer))
                    Next
                Else
                    If _sortUp Then
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataCollection.Add(_ActiveDataSortList(_ActiveDataSortList.Count - 1 - listIncrementer))
                        Next
                    Else
                        For listIncrementer = 0 To _ActiveDataSortList.Count - 1
                            _ActiveDataCollection.Add(_ActiveDataSortList(listIncrementer))
                        Next
                    End If
                    _sortUp = Not _sortUp
                End If
                _lastSortedField = currentSortField
            End If
        End If
    End Sub

    Private Sub GridViewColumnHeaderClickedHandler_activedata_OLD(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim headerClicked As GridViewColumnHeader = TryCast(e.OriginalSource, GridViewColumnHeader)
        Dim direction As ListSortDirection

        If headerClicked IsNot Nothing Then
            If headerClicked.Role <> GridViewColumnHeaderRole.Padding Then
                If headerClicked IsNot _lastHeaderClicked_ActiveData Then
                    direction = ListSortDirection.Ascending
                Else
                    If _lastDirection_ActiveData = ListSortDirection.Ascending Then
                        direction = ListSortDirection.Descending
                    Else
                        direction = ListSortDirection.Ascending
                    End If
                End If

                Dim header As String = TryCast(headerClicked.Column.Header, String)
                ''  Sort_ActiveData(header, direction)

                If direction = ListSortDirection.Ascending Then
                    headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowUp"), DataTemplate)
                Else
                    headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowDown"), DataTemplate)
                End If

                ' Remove arrow from previously sorted header
                If _lastHeaderClicked_ActiveData IsNot Nothing AndAlso _lastHeaderClicked_ActiveData IsNot headerClicked Then
                    _lastHeaderClicked_ActiveData.Column.HeaderTemplate = Nothing
                End If


                _lastHeaderClicked_ActiveData = headerClicked
                _lastDirection_ActiveData = direction
            End If
        End If
    End Sub







    '  Private Sub Sort_ActiveData(ByVal sortBy As String, ByVal direction As ListSortDirection)
    ' Dim dataView As ICollectionView = CollectionViewSource.GetDefaultView(ActiveDataList.ItemsSource)
    '
    '        dataView.SortDescriptions.Clear()
    '    Dim sd As New SortDescription(sortBy, direction)
    '        dataView.SortDescriptions.Add(sd)
    '        dataView.Refresh()
    '    End Sub
#End Region

    Public Sub SelectionChangedEventHandler()

    End Sub

#Region "HTML"
    Private Sub exportAMchart(ByVal paramObj As Object)
        If True Then
            exportAMchart(paramObj)
        Else
            Dim fsT As Object
            Dim fileName As String
            Dim dataList As List(Of DTevent) = paramObj(0)
            Dim us As New CultureInfo("en-US")
            fsT = CreateObject("ADODB.Stream")
            fsT.Type = 2 'Specify stream type - we want To save text/string data.
            fsT.Charset = "utf-8" 'Specify charset For the source text data.
            fsT.Open() 'Open the stream And write binary data To the object

            fileName = SERVER_FOLDER_PATH & "RawPareto.html"

            fsT.WriteText("<!DOCTYPE HTML PUBLIC " & Chr(34) & "-'//W3C'//DTD HTML 4.01'//EN" & Chr(34) & " " & Chr(34) & "http:'//www.w3.org/TR/html4/strict.dtd" & Chr(34) & ">" & vbCrLf)
            fsT.WriteText("<html>" & vbCrLf)
            fsT.WriteText("<head>" & vbCrLf)
            fsT.WriteText("<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & ">" & vbCrLf)
            fsT.WriteText("<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " href=" & Chr(34) & "style.css" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf)
            fsT.WriteText("<script src=" & Chr(34) & "amcharts.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
            fsT.WriteText("<script src=" & Chr(34) & "serial.js" & Chr(34) & " type=" & Chr(34) & "text/javascript" & Chr(34) & "></script>" & vbCrLf)
            fsT.WriteText("<script>" & vbCrLf)
            fsT.WriteText("var chart;" & vbCrLf)
            fsT.WriteText("var chartData = [" & vbCrLf)

            For eventIncrementer As Integer = 0 To ExportListAM.Count - 1
                fsT.WriteText("{" & Chr(34) & "fault" & Chr(34) & ": " & Chr(34) & dataList(eventIncrementer).Name & Chr(34) & "," & vbCrLf)
                fsT.WriteText(Chr(34) & "dt" & Chr(34) & ": " & (Math.Round((CDbl(dataList(eventIncrementer).DTpct) * 100), 2)).ToString("######0.00", us) & "," & vbCrLf)
                fsT.WriteText(Chr(34) & "stops per day" & Chr(34) & ": " & (Math.Round(CDbl(dataList(eventIncrementer).SPD), 1)).ToString("######0.0", us) & "," & vbCrLf)
                fsT.WriteText(Chr(34) & "pct" & Chr(34) & ": " & (dataList(eventIncrementer).DT_Display).ToString("######0.0", us) & "," & vbCrLf)
                fsT.WriteText(Chr(34) & "stops" & Chr(34) & ": " & (Math.Round(dataList(eventIncrementer).Stops, 0)) & "," & vbCrLf)
                fsT.WriteText(Chr(34) & "mtbf" & Chr(34) & ": " & (Math.Round(CDbl(dataList(eventIncrementer).MTBF), 1)).ToString("######0.0", us) & vbCrLf)
                If eventIncrementer = dataList.Count - 1 Then
                    fsT.WriteText("}" & vbCrLf)
                Else
                    fsT.WriteText("}," & vbCrLf)
                End If
            Next

            fsT.WriteText("];" & vbCrLf)


            fsT.WriteText("AmCharts.ready(function () {" & vbCrLf)
            '// SERIAL CHART
            fsT.WriteText("chart = new AmCharts.AmSerialChart()" & vbCrLf)

            fsT.WriteText("chart.dataProvider = chartData; " & vbCrLf)
            fsT.WriteText("chart.categoryField = " & Chr(34) & "fault" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.startDuration = 1; " & vbCrLf)
            fsT.WriteText("chart.export; " & vbCrLf)

            fsT.WriteText("var categoryAxis = chart.categoryAxis; " & vbCrLf)
            fsT.WriteText("categoryAxis.gridPosition = " & Chr(34) & "start" & Chr(34) & "; " & vbCrLf)
            If dataList.Count > 15 Then
                fsT.WriteText("categoryAxis.labelRotation = 45;" & vbCrLf)
            ElseIf dataList.Count > 6 And dataList.Count < 16 Then
                fsT.WriteText("categoryAxis.labelRotation = 45;" & vbCrLf)
            Else
                fsT.WriteText("categoryAxis.labelRotation = 0;" & vbCrLf)
            End If
            fsT.WriteText("chart.addListener('init', function() {" & vbCrLf)
            fsT.WriteText("chart.categoryAxis.addListener('rollOverItem', function(event) {" & vbCrLf)
            fsT.WriteText("event.target.setAttr('cursor', 'default' );" & vbCrLf)
            fsT.WriteText("event.chart.balloon.followCursor( true );" & vbCrLf)
            fsT.WriteText("event.chart.balloon.showBalloon(event.serialDataItem.dataContext.fault);" & vbCrLf)
            fsT.WriteText(" } );" & vbCrLf)
            fsT.WriteText("chart.categoryAxis.addListener( 'rollOutItem', function( event ) {" & vbCrLf)
            fsT.WriteText(" event.chart.balloon.hide();" & vbCrLf)
            fsT.WriteText("} );" & vbCrLf)
            fsT.WriteText("} )" & vbCrLf)

            fsT.WriteText("var valueAxis = new AmCharts.ValueAxis(); " & vbCrLf)
            fsT.WriteText("valueAxis.axisColor = " & Chr(34) & "#2C99C3" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("valueAxis.gridAlpha = 0; " & vbCrLf)
            fsT.WriteText("valueAxis.axisThickness = 2; " & vbCrLf)
            fsT.WriteText("chart.addValueAxis(valueAxis); " & vbCrLf)

            '// second value axis (on the right)
            fsT.WriteText("var valueAxis2 = new AmCharts.ValueAxis(); " & vbCrLf)
            fsT.WriteText("valueAxis2.position = " & Chr(34) & "right" & Chr(34) & "; " & vbCrLf) '// this line makes the axis to appear on the right
            fsT.WriteText("valueAxis2.axisColor = " & Chr(34) & "#fcd202" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("valueAxis2.gridAlpha = 0; " & vbCrLf)
            fsT.WriteText("valueAxis.unit = " & Chr(34) & "%" & Chr(34) & ";" & vbCrLf)
            fsT.WriteText("valueAxis2.axisThickness = 2; " & vbCrLf)
            fsT.WriteText("chart.addValueAxis(valueAxis2); " & vbCrLf)

            '// third value axis (on the right)
            fsT.WriteText("var valueAxis3 = new AmCharts.ValueAxis(); " & vbCrLf)
            fsT.WriteText("valueAxis3.position = " & Chr(34) & "right" & Chr(34) & "; " & vbCrLf) '// this line makes the axis to appear on the right
            fsT.writetext("valueAxis3.offset = 30; " & vbCrLf)
            fsT.WriteText("valueAxis3.axisColor = " & Chr(34) & "#FF8C00" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("valueAxis3.gridAlpha = 0; " & vbCrLf)
            fsT.WriteText("valueAxis3.axisThickness = 2; " & vbCrLf)
            fsT.WriteText("chart.addValueAxis(valueAxis3); " & vbCrLf)

            '// GRAPHS
            '// column graph DT%
            fsT.WriteText("var graph1 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph1.type = " & Chr(34) & "column" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText(" graph1.title = " & Chr(34) & "Downtime" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph1.lineColor = " & Chr(34) & "#2C99C3" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph1.valueField = " & Chr(34) & "dt" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph1.lineAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph1.fillAlphas = 1; " & vbCrLf)
            fsT.WriteText("graph1.labelText = " & Chr(34) & "[[value]]%" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph1.alphaField = " & Chr(34) & "alpha" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph1.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b>[[pct]] min / [[dt]] % </b> </span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph1); " & vbCrLf)

            '// line 2nd SPD
            fsT.WriteText("var graph2 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph2.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph2.title = " & Chr(34) & "Stops per day" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph2.lineColor = " & Chr(34) & "#fcd202" & Chr(34) & "; " & vbCrLf)
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
            fsT.WriteText("graph2.balloonText = " & Chr(34) & "<span style='font-size:13px;'><b>[[value]] stops per day / [[stops]] stops </b></span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph2);" & vbCrLf)

            '// line 3rd MTBF
            fsT.WriteText("var graph3 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph3.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph3.title = " & Chr(34) & "MTBF (min)" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph3.lineColor = " & Chr(34) & "#fcd202" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph3.valueField = " & Chr(34) & "mtbf" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph3.lineThickness = 0; " & vbCrLf)
            fsT.WriteText("graph3.lineAlpha = 0; " & vbCrLf)
            fsT.WriteText("graph3.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph3.bulletBorderThickness = 3; " & vbCrLf)
            fsT.WriteText("graph3.bulletBorderColor = " & Chr(34) & "#FF8C00" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph3.bulletBorderAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph3.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph3.valueAxis = valueAxis3; " & vbCrLf)
            fsT.WriteText("graph3.balloonText = " & Chr(34) & "<span style='font-size:13px;'>[[title]]:<b>[[value]]</b></span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph3);" & vbCrLf)

            '// line 4th Actual Stops
            fsT.WriteText("var graph4 = new AmCharts.AmGraph(); " & vbCrLf)
            fsT.WriteText("graph4.type = " & Chr(34) & "line" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.title = " & Chr(34) & "Total Stops" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.lineColor = " & Chr(34) & "#cc66ff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.valueField = " & Chr(34) & "stops" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.lineThickness = 0; " & vbCrLf)
            fsT.WriteText("graph4.lineAlpha = 0; " & vbCrLf)
            fsT.WriteText("graph4.bullet = " & Chr(34) & "round" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.bulletBorderThickness = 3; " & vbCrLf)
            fsT.WriteText("graph4.bulletBorderColor = " & Chr(34) & "#cc66ff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.bulletBorderAlpha = 1; " & vbCrLf)
            fsT.WriteText("graph4.bulletColor = " & Chr(34) & "#ffffff" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("graph4.valueAxis = valueAxis2; " & vbCrLf)
            fsT.WriteText("graph4.balloonText = " & Chr(34) & "<span style='font-size:13px;'>[[title]]:<b>[[value]]</b></span>" & Chr(34) & "; " & vbCrLf)
            fsT.WriteText("chart.addGraph(graph4);" & vbCrLf)


            '// LEGEND
            fsT.WriteText("var legend = new AmCharts.AmLegend();" & vbCrLf)
            fsT.WriteText("legend.useGraphSettings = true;" & vbCrLf)
            fsT.WriteText("chart.addLegend(legend);" & vbCrLf)
            fsT.WriteText("chart.hideGraph(graph3);" & vbCrLf)
            fsT.WriteText("chart.hideGraph(graph4);" & vbCrLf)

            '// WRITE
            fsT.WriteText("chart.write(" & Chr(34) & "chartdiv" & Chr(34) & ");" & vbCrLf)
            fsT.WriteText("});" & vbCrLf)
            fsT.WriteText("</script>" & vbCrLf)
            fsT.WriteText("</head>" & vbCrLf)

            fsT.WriteText("<body>" & vbCrLf)

            If dataList.Count > 10 Then
                fsT.WriteText("<div id=" & Chr(34) & "chartdiv" & Chr(34) & " style=" & Chr(34) & "width:2000px; height:480px;" & Chr(34) & "></div>" & vbCrLf)
            Else
                fsT.WriteText("<div id=" & Chr(34) & "chartdiv" & Chr(34) & " style=" & Chr(34) & "width:100%; height:480px;" & Chr(34) & "></div>" & vbCrLf)
            End If

            'wrap it up
            fsT.WriteText("</body>" & vbCrLf)
            fsT.WriteText("</html>" & vbCrLf)

            Try
                fsT.SaveToFile(fileName, 2) 'Save binary data To disk
            Catch ex As Exception

            End Try
            fsT = Nothing

        End If
    End Sub

    Dim isDowntime As Boolean = True
    Dim isMTBF As Boolean = False
    Dim isSPD As Boolean = True
    Dim isStops As Boolean = False
    Dim isDTMin As Boolean = False

    Private Sub exportAMchart_Telerik()
        Dim blankDataTemplate = New DataTemplate("")
        ParetoHTMLChart.Series.Clear()
        ParetoHTMLChart.Palette = getLineColors()
        Dim xyz = New LinearAxis()
        xyz.Minimum = 0
        ParetoHTMLChart.VerticalAxis = xyz

        Dim secondaryVAxis = New LinearAxis()
        secondaryVAxis.HorizontalLocation = AxisHorizontalLocation.Right

        Dim thirdVAxis = New LinearAxis()
        thirdVAxis.HorizontalLocation = AxisHorizontalLocation.Right

        Dim dtSeries As CategoricalSeries = New BarSeries()
        Dim stopSeries As CategoricalSeries = New PointSeries()
        Dim spdSeries As CategoricalSeries = New PointSeries()
        Dim mtbfSeries As CategoricalSeries = New PointSeries()
        Dim dtminSeries As CategoricalSeries = New BarSeries()

        For eventIncrementer As Integer = 0 To ExportListAM.Count - 1

            Dim dtPoint As CategoricalDataPoint = New CategoricalDataPoint()
            dtPoint.Value = Math.Round(CDbl(ExportListAM(eventIncrementer).DTpct) * 100, 2)
            dtPoint.Label = Math.Round(CDbl(ExportListAM(eventIncrementer).DTpct) * 100, 1) & "%"
            dtPoint.Category = ExportListAM(eventIncrementer).Name

            dtSeries.DataPoints.Add(dtPoint)

            Dim spdPoint As CategoricalDataPoint = New CategoricalDataPoint()
            spdPoint.Value = Math.Round(CDbl(ExportListAM(eventIncrementer).SPD), 2)
            spdPoint.Label = Math.Round(CDbl(ExportListAM(eventIncrementer).SPD), 1)
            spdPoint.Category = ExportListAM(eventIncrementer).Name

            spdSeries.DataPoints.Add(spdPoint)

            Dim mtbfPoint As CategoricalDataPoint = New CategoricalDataPoint()
            mtbfPoint.Value = Math.Round(CDbl(ExportListAM(eventIncrementer).MTBF), 2)
            mtbfPoint.Label = Math.Round(CDbl(ExportListAM(eventIncrementer).MTBF), 1)
            mtbfPoint.Category = ExportListAM(eventIncrementer).Name

            mtbfSeries.DataPoints.Add(mtbfPoint)

            Dim stopPoint As CategoricalDataPoint = New CategoricalDataPoint()
            stopPoint.Value = CDbl(ExportListAM(eventIncrementer).Stops)
            stopPoint.Label = CDbl(ExportListAM(eventIncrementer).Stops)
            stopPoint.Category = ExportListAM(eventIncrementer).Name

            stopSeries.DataPoints.Add(stopPoint)


            Dim dtminPoint As CategoricalDataPoint = New CategoricalDataPoint()
            dtminPoint.Value = Math.Round(CDbl(ExportListAM(eventIncrementer).DT), 2)
            dtminPoint.Label = Math.Round(CDbl(ExportListAM(eventIncrementer).DT), 1) & " min"
            dtminPoint.Category = ExportListAM(eventIncrementer).Name

            dtminSeries.DataPoints.Add(dtminPoint)
        Next

        If isDowntime Then
            dtSeries.ShowLabels = True
            ParetoHTMLChart.Series.Add(dtSeries)
        End If

        If isSPD Then
            spdSeries.ShowLabels = True
            spdSeries.VerticalAxis = secondaryVAxis
            ParetoHTMLChart.Series.Add(spdSeries)
        End If

        If isMTBF Then
            If Not isDowntime And Not isSPD And Not isStops Then
                mtbfSeries.ShowLabels = True
            Else
                mtbfSeries.ShowLabels = False
            End If
            mtbfSeries.VerticalAxis = thirdVAxis
            ParetoHTMLChart.Series.Add(mtbfSeries)
        End If

        If isStops Then
            If Not isDowntime And Not isSPD And Not isMTBF Then
                stopSeries.ShowLabels = True
            Else
                stopSeries.ShowLabels = False
            End If
            stopSeries.VerticalAxis = secondaryVAxis
            ParetoHTMLChart.Series.Add(stopSeries)
        End If

        if isdtmin then
            dtminseries.showlabels = true
            paretohtmlchart.series.add(dtminseries)
        End If

        ParetoHTMLChart.HorizontalAxis.LabelInterval = 1
        ParetoHTMLChart.HorizontalAxis.LabelFitMode = AxisLabelFitMode.MultiLine

    End Sub



    Public Function getLineColors2() As ChartPalette
        Dim tmp As ChartPalette = New ChartPalette()
        addPaletteEntry(tmp, 255, 102, 0)
        Return tmp
    End Function
    Public Function getLineColors() As ChartPalette
        Dim tmp As ChartPalette = New ChartPalette()
        If isDowntime Then
            addPaletteEntry(tmp, 44, 153, 195) 'dt pct
        End If
        If isSPD Then
            addPaletteEntry(tmp, 252, 210, 2) 'spd
        End If
        If isMTBF Then
            addPaletteEntry(tmp, 255, 140, 0) 'mtbf
        End If
        If isStops Then
            addPaletteEntry(tmp, 207, 111, 255) 'stops
        End If
        If isDTMin Then
            addPaletteEntry(tmp, 44, 153, 195) 'dt pct
        End If
        Return tmp
    End Function
    Private Sub addPaletteEntry(ByRef palette As ChartPalette, R As Byte, G As Byte, B As Byte)

        Dim tmp = New PaletteEntry()
        tmp.Fill = New SolidColorBrush(Color.FromRgb(R, G, B))
        tmp.Stroke = New SolidColorBrush(Color.FromRgb(R, G, B))
        palette.GlobalEntries.Add(tmp)
    End Sub


#End Region

#Region "Opening and Closing"
    Public Sub RawdatawindowClose(ByVal sender As Object, ByVal e As CancelEventArgs)



        If InStr(sender.ToString, "rawdatawindow", vbTextCompare) > 0 Then
            My.Settings.defaultDownTimeField = mappinglevel1
            My.Settings.defaultDownTimeField_Secondary = mappinglevel2
            prstoryReport.reMapReport()
            AllProdLines(selectedindexofLine_temp).reMapRawData()
            IsRemappingDoneOnce = True
            '  ParetoHTML.Dispose()
            'VarianceHTML.Dispose()
            Dispose_AllTrendCharts_inRawDataWindow()
        End If





    End Sub

#End Region
#Region "Trends"
    Private bargraphReportWindow_forraw As bargraphreportwindow

    Private Sub UpdateTrendChart(YVals As List(Of Double), XVals As List(Of DateTime))
        TrendChart.Series.Clear()
        TrendChart.Palette = getLineColors2()

        Dim xyz = New LinearAxis()
        xyz.Minimum = 0
        TrendChart.VerticalAxis = xyz

        Dim series As CategoricalSeries = New LineSeries()
        For i As Integer = 0 To YVals.Count - 1
            Dim x As CategoricalDataPoint = New CategoricalDataPoint()
            x.Value = YVals(i)
            x.Category = XVals(i)
            series.DataPoints.Add(x)
        Next
        TrendChart.Series.Add(series)
    End Sub

    Public Enum TrendKPIEnum
        Stops
        DT
        MTBF
    End Enum
    Public Enum TrendTimeEnum
        Day
        Week
        Month
    End Enum
    Public TrendKPI As TrendKPIEnum = TrendKPIEnum.Stops
    Public TrendTime As TrendTimeEnum = TrendTimeEnum.Day

    Private GoTime As Boolean = False

    Private stopsMotionReport As MotionReport

    Private Sub LaunchtrendsinRawDataWindow()
        stopsMotionReport = New MotionReport(AllProdLines(selectedindexofLine_temp), starttimeselected, endtimeselected, prstoryReport.MainLEDSReport.DT_Report.UnplannedEventDirectory, 1)

        If GoTime Then
            Dim yval As New List(Of Double)
            Dim xval As New List(Of DateTime)

            Select Case TrendKPI
                Case TrendKPIEnum.Stops
                    Select Case TrendTime
                        Case TrendTimeEnum.Day
                            For timeIncrementer = 0 To stopsMotionReport.DailyReports.Count - 1
                                '   fsT.Writetext("{date: new Date('" & Format(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_Date(timeIncrementer, False, failuremodeno), "MM dd yyyy") & "')" & "," & vbCrLf)
                                '   If isDT Then
                                '      fsT.Writetext("UnplannedDowntime: " & (Math.Round(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd(timeIncrementer, True, failuremodeno), 1)).ToString("######0.0", us) & "," & vbCrLf)
                                '  Else
                                '       fsT.Writetext("UnplannedDowntime: " & (Math.Round(rawData.getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd(timeIncrementer, False, failuremodeno))).ToString("######0.0", us) & "," & vbCrLf)
                                '   End If
                                '
                                '                                If timeIncrementer <> rawData.DailyReports.Count - 1 Then
                                '                                    fsT.Writetext("}," & vbCrLf)
                                '                               End If
                            Next
                        Case TrendTimeEnum.Week

                        Case Else 'month

                    End Select
                Case TrendKPIEnum.DT
                    Select Case TrendTime
                        Case TrendTimeEnum.Day

                        Case TrendTimeEnum.Week

                        Case Else 'month

                    End Select
                Case Else 'mtbf
                    Select Case TrendTime
                        Case TrendTimeEnum.Day

                        Case TrendTimeEnum.Week

                        Case Else 'month

                    End Select
            End Select
        Else
            Dim i As Integer
            Dim selectedfailuremodeindex As Integer = -1

            'updating the top stops list after remapping
            If TitleLabel.Content <> "Total" Then
                For i = 0 To prstoryReport.TopStopsList.Count - 1
                    If TitleLabel.Content = prstoryReport.TopStopsList(i).Name Then
                        selectedfailuremodeindex = i
                        Exit For
                    End If
                Next
            Else
                Exit Sub
            End If

            Try
                exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, True, selectedfailuremodeindex)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, False, selectedfailuremodeindex)

                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, True, selectedfailuremodeindex)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, False, selectedfailuremodeindex)

                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, True, selectedfailuremodeindex)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, False, selectedfailuremodeindex)

                exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF(stopsMotionReport, True, selectedfailuremodeindex)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Weekly(stopsMotionReport, True, selectedfailuremodeindex)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Monthly(stopsMotionReport, True, selectedfailuremodeindex)

                Dim sourcestringS As String
                Dim sourcestringD As String

                failuremodeno = selectedfailuremodeindex
                motionchartsource = 31
                IsLaunchedfromstops_InMOtionChart = True
                selectedfailuremode_inMotionChart = failuremodeno

                sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
                MotionChartS.Source = New Uri(sourcestringS)

                sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D.html"
                MotionChartD.Source = New Uri(sourcestringD)
                losscardnamelabel.Content = TitleLabel.Content & " losses over last 3 months"


                sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S_Weekly.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
                MotionChartS_Weekly.Source = New Uri(sourcestringS)

                sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D_Weekly.html"
                MotionChartD_Weekly.Source = New Uri(sourcestringD)


                sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S_Monthly.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
                MotionChartS_Monthly.Source = New Uri(sourcestringS)

                sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D_Monthly.html"
                MotionChartD_Monthly.Source = New Uri(sourcestringD)


                sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
                MotionChart_MTBF.Source = New Uri(sourcestringS)

                sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF_Monthly.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
                MotionChart_MTBF_Monthly.Source = New Uri(sourcestringS)

                sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF_Weekly.html"
                MotionChart_MTBF_Weekly.Source = New Uri(sourcestringS)

            Catch ex As Exception
                MsgBox("The trends chart could not be loaded. Let the developer know about this problem and provide detail info such as which line was being analyzed and what time frame was selected.")
                Exit Sub
            End Try
        End If
    End Sub
    Public Sub setBargraphReportWindow_forraw(parentWindow As bargraphreportwindow)
        bargraphReportWindow_forraw = parentWindow
    End Sub

    Private Sub Trends_OnStart()
        Dailybtn.Background = mybrushbrightorange
        Weeklybtn.Background = mybrushdarkgray
        Monthlybtn.Background = mybrushdarkgray
        SetSourceString()
    End Sub
    Private Sub SetGeneralTrendsDefaultSource()
        Dim sourcestringS As String
        Dim sourcestringD As String


        motionchartsource = 31
        IsLaunchedfromstops_InMOtionChart = True
        selectedfailuremode_inMotionChart = failuremodeno

        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S.html"
        MotionChartS.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D.html"
        MotionChartD.Source = New Uri(sourcestringD)
        losscardnamelabel.Content = TitleLabel.Content & " losses over last 3 months"


        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S_Weekly.html"
        MotionChartS_Weekly.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D_Weekly.html"
        MotionChartD_Weekly.Source = New Uri(sourcestringD)


        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S_Monthly.html"
        MotionChartS_Monthly.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D_Monthly.html"
        MotionChartD_Monthly.Source = New Uri(sourcestringD)


        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF.html"
        MotionChart_MTBF.Source = New Uri(sourcestringS)

        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF_Monthly.html"
        MotionChart_MTBF_Monthly.Source = New Uri(sourcestringS)

        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF_Weekly.html"
        MotionChart_MTBF_Weekly.Source = New Uri(sourcestringS)


    End Sub
    Private Sub SetSShapeTrendsDefaultSource()
        Dim sourcestringS As String
        Dim sourcestringD As String


        motionchartsource = 31
        IsLaunchedfromstops_InMOtionChart = True
        selectedfailuremode_inMotionChart = failuremodeno

        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "Sshape" & motionchartsource & "_" & failuremodeno & "S.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
        MotionChartS.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "Sshape" & motionchartsource & "_" & failuremodeno & "D.html"
        MotionChartD.Source = New Uri(sourcestringD)

        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "Sshape" & motionchartsource & "_" & failuremodeno & "MTBF.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
        MotionChart_MTBF.Source = New Uri(sourcestringS)



    End Sub
    Private Sub SetSourceString()

        Dim sourcestringS As String
        Dim sourcestringD As String

        prclicked()
        Select Case motionchartsource
            Case 31
                'losscardnamelabel.Content = "Top losses in analysis period"
                'losscardnamelabel.Content =

                stopclicked()


                Exit Sub
            Case 0
                losscardnamelabel.Content = "Line Performance"
                ' stopsbutton.Visibility = Windows.Visibility.Hidden
                ' prbutton.Visibility = Windows.Visibility.Hidden
                If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                    prbutton.Content = "PR"

                Else
                    prbutton.Content = "Av."
                End If
                mtbfbutton.Visibility = Visibility.Hidden
                losscardnamelabel.Content = "Line performance in last 3 months"
                UseTrack_PROverallTrends = True
        End Select


        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "S.html"
        MotionChartS.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "D.html"
        MotionChartD.Source = New Uri(sourcestringD)


        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "S_Weekly.html"
        MotionChartS_Weekly.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "D_Weekly.html"
        MotionChartD_Weekly.Source = New Uri(sourcestringD)

        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "S_Monthly.html"
        MotionChartS_Monthly.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "D_MOnthly.html"
        MotionChartD_Monthly.Source = New Uri(sourcestringD)




    End Sub

    Private Sub stopclicked()
        stopsbutton.Opacity = 1.0
        prbutton.Opacity = 0.2
        mtbfbutton.Opacity = 0.2
        MotionChartS.Visibility = Visibility.Visible
        MotionChartD.Visibility = Visibility.Hidden
        MotionChartD_Weekly.Visibility = Visibility.Hidden
        MotionChartD_Monthly.Visibility = Visibility.Hidden
        MotionChart_MTBF.Visibility = Visibility.Hidden
        MotionChart_MTBF_Monthly.Visibility = Visibility.Hidden
        MotionChart_MTBF_Weekly.Visibility = Visibility.Hidden

        IsSPDActive = True
        ISDTActive = False
        ISMTBFActive = False

        If GoTime Then
            TrendKPI = TrendKPIEnum.Stops
            LaunchtrendsinRawDataWindow()
        End If

        If IsSShapeActive = False Then
            DailyClicked()
        End If
    End Sub

    Private Sub prclicked()
        stopsbutton.Opacity = 0.2
        prbutton.Opacity = 1.0
        mtbfbutton.Opacity = 0.2
        MotionChartD.Visibility = Visibility.Visible
        MotionChartS.Visibility = Visibility.Hidden
        MotionChartS_Weekly.Visibility = Visibility.Hidden
        MotionChartS_Monthly.Visibility = Visibility.Hidden
        MotionChart_MTBF.Visibility = Visibility.Hidden
        MotionChart_MTBF_Monthly.Visibility = Visibility.Hidden
        MotionChart_MTBF_Weekly.Visibility = Visibility.Hidden
        IsSPDActive = False
        ISDTActive = True
        ISMTBFActive = False

        If GoTime Then
            TrendKPI = TrendKPIEnum.DT
            LaunchtrendsinRawDataWindow()
        End If

        If IsSShapeActive = False Then
            DailyClicked()
        End If
    End Sub

    Private Sub mtbfclicked()
        stopsbutton.Opacity = 0.2
        prbutton.Opacity = 0.2
        mtbfbutton.Opacity = 1.0
        MotionChartD.Visibility = Visibility.Hidden
        MotionChartS.Visibility = Visibility.Hidden
        MotionChartS_Weekly.Visibility = Visibility.Hidden
        MotionChartS_Monthly.Visibility = Visibility.Hidden
        MotionChart_MTBF.Visibility = Visibility.Visible
        MotionChart_MTBF_Monthly.Visibility = Visibility.Hidden
        MotionChart_MTBF_Weekly.Visibility = Visibility.Hidden
        IsSPDActive = False
        ISDTActive = False
        ISMTBFActive = True

        If GoTime Then
            TrendKPI = TrendKPIEnum.MTBF
            LaunchtrendsinRawDataWindow()
        End If

        If IsSShapeActive = False Then
            DailyClicked()
        End If
    End Sub

    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        'sender.Opacity = 0.7
    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        ' sender.Opacity = 1.0
    End Sub

    Private Sub DailyClicked()
        Dailybtn.Background = mybrushbrightorange
        Weeklybtn.Background = mybrushdarkgray
        Monthlybtn.Background = mybrushdarkgray


        If IsSPDActive Then

            MotionChartS.Visibility = Visibility.Visible
            MotionChartS_Weekly.Visibility = Visibility.Hidden
            MotionChartS_Monthly.Visibility = Visibility.Hidden

        ElseIf ISDTActive Then
            MotionChartD.Visibility = Visibility.Visible
            MotionChartD_Weekly.Visibility = Visibility.Hidden
            MotionChartD_Monthly.Visibility = Visibility.Hidden
        ElseIf ISMTBFActive Then
            MotionChart_MTBF.Visibility = Visibility.Visible
            MotionChart_MTBF_Weekly.Visibility = Visibility.Hidden
            MotionChart_MTBF_Monthly.Visibility = Visibility.Hidden

        End If

        If GoTime Then
            TrendTime = TrendTimeEnum.Day
            LaunchtrendsinRawDataWindow()
        End If


    End Sub

    Private Sub WeeklyClicked(sender As Object, e As RoutedEventArgs)
        Dailybtn.Background = mybrushdarkgray
        Weeklybtn.Background = mybrushbrightorange
        Monthlybtn.Background = mybrushdarkgray

        If IsSPDActive Then

            MotionChartS.Visibility = Visibility.Hidden
            MotionChartS_Weekly.Visibility = Visibility.Visible
            MotionChartS_Monthly.Visibility = Visibility.Hidden


        ElseIf ISDTActive Then
            MotionChartD.Visibility = Visibility.Hidden
            MotionChartD_Weekly.Visibility = Visibility.Visible
            MotionChartD_Monthly.Visibility = Visibility.Hidden
        ElseIf ISMTBFActive Then
            MotionChart_MTBF.Visibility = Visibility.Hidden
            MotionChart_MTBF_Weekly.Visibility = Visibility.Visible
            MotionChart_MTBF_Monthly.Visibility = Visibility.Hidden
        End If

        If GoTime Then
            TrendTime = TrendTimeEnum.Week
            LaunchtrendsinRawDataWindow()
        End If

    End Sub
    Private Sub MonthlyClicked(sender As Object, e As RoutedEventArgs)
        Dailybtn.Background = mybrushdarkgray
        Weeklybtn.Background = mybrushdarkgray
        Monthlybtn.Background = mybrushbrightorange

        If IsSPDActive Then

            MotionChartS.Visibility = Visibility.Hidden
            MotionChartS_Weekly.Visibility = Visibility.Hidden
            MotionChartS_Monthly.Visibility = Visibility.Visible


        ElseIf ISDTActive Then
            MotionChartD.Visibility = Visibility.Hidden
            MotionChartD_Weekly.Visibility = Visibility.Hidden
            MotionChartD_Monthly.Visibility = Visibility.Visible
        ElseIf ISMTBFActive Then
            MotionChart_MTBF.Visibility = Visibility.Hidden
            MotionChart_MTBF_Weekly.Visibility = Visibility.Hidden
            MotionChart_MTBF_Monthly.Visibility = Visibility.Visible
        End If

        If GoTime Then
            TrendTime = TrendTimeEnum.Month
            LaunchtrendsinRawDataWindow()
        End If

    End Sub

    Private Sub TrendsRadioChecked(sender As Object, e As RoutedEventArgs)
        System.Windows.Forms.Application.DoEvents()

        If TrendSelection_General.IsChecked = True Then
            IsSShapeActive = False
            SetGeneralTrendsButtonNames()
            SetGeneralTrendsDefaultSource()
            Dailybtn.Visibility = Visibility.Visible
            Monthlybtn.Visibility = Visibility.Visible
            Weeklybtn.Visibility = Visibility.Visible
            prclicked()
        ElseIf TrendSelection_SShape.IsChecked = True Then
            IsSShapeActive = True
            SetSShapeButtonNames()
            prclicked()
            losscardnamelabel.Content = TitleLabel.Content & ": S-Shape Growth Trends"
            SetSShapeTrendsDefaultSource()
            Dailybtn.Visibility = Visibility.Hidden
            Monthlybtn.Visibility = Visibility.Hidden
            Weeklybtn.Visibility = Visibility.Hidden

        End If

    End Sub
    Private Sub SetGeneralTrendsButtonNames()
        prbutton.Content = "DT%"
        stopsbutton.Content = "STOPS/D"
        mtbfbutton.Content = "MTBF"
    End Sub
    Private Sub SetSShapeButtonNames()
        prbutton.Content = "Av."
        stopsbutton.Content = "STOPS"
        mtbfbutton.Content = "MTBF"
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#Region "Paretos KPI Selection"
    Private Sub Label_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If isDowntime Then
            isDowntime = False
            DT.Opacity = 0.4
        Else
            isDowntime = True
            DT.Opacity = 1.0
        End If
        exportAMchart_Telerik()
    End Sub

    Private Sub Label_MouseDown_1(sender As Object, e As MouseButtonEventArgs)
        If isSPD Then
            isSPD = False
            SPD.Opacity = 0.4
        Else
            isSPD = True
            SPD.Opacity = 1.0
        End If
        exportAMchart_Telerik()
    End Sub

    Private Sub Label_MouseDown_2(sender As Object, e As MouseButtonEventArgs)
        If isMTBF Then
            isMTBF = False
            MTBF.Opacity = 0.4
        Else
            isMTBF = True
            MTBF.Opacity = 1.0
        End If
        exportAMchart_Telerik()
    End Sub

    Private Sub Label_MouseDown_3(sender As Object, e As MouseButtonEventArgs)
        If isStops Then
            isStops = False
            STOPS.Opacity = 0.4
        Else
            isStops = True
            STOPS.Opacity = 1.0
        End If
        exportAMchart_Telerik()
    End Sub
    Private Sub Label_MouseDown_4(sender As Object, e As MouseButtonEventArgs)
        If isDTMin Then
            isDTMin = False
            DTPCT.Opacity = 0.4
        Else
            isDTMin = True
            DTPCT.Opacity = 1.0
        End If
        exportAMchart_Telerik()
    End Sub
#End Region
#End Region
End Class
