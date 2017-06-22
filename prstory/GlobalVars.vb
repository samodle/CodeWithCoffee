Module prStoryGlobalVars
    Public shouldSnakeClose As Boolean
    Public MasterDataSet As inControlReport

    Public selectedindexofLine_temp As Integer
    Public starttimeselected As Date
    Public endtimeselected As Date

    Public isAPRILFOOLS As Boolean = False 'SET TO FALSE BEFORE UPDATE

    Public PRSTORY_VERSION_NUMBER As String = ""
End Module


Module Runtime_Variables
    Public incontrolAnalysisSelectedStartdate As Date
    Public incontrolAnalysisSelectedEnddate As Date
    Public stopbubblenames(0 To 14) As String
    Public stopbubblestops(0 To 14) As Double
    Public stopbubblePR(0 To 14) As Double
    Public stopbubbleMTBF(0 To 14) As Double
    Public stopbubble90daystopsperday(0 To 14) As Double
    Public stopbubbleAnalysisstopsperday(0 To 14) As Double

    Public topstopname(0 To 14) As String

    Public cardnameLabeltext(0 To 100) As String
    Public tempreasonlevel As String
    Public datalabelcontent As String
    Public bubblenumberpublic As Integer
    Public motionchartsource As Integer

    Public bargraphreportwindow_Open As Boolean
    Public IsRemappingDoneOnce As Boolean
    Public IsAnalyzeButtonClickSource_Analyze As Boolean = False
    Public IsNotesChanged As Boolean = False
    Public IsPickMode As Boolean = False
    Public IsSimulationMode As Boolean = False
    Public WhichPickRow As Integer = -1
    Public OriginStopSim As Boolean = False
    Public IsIncontrol_Shiftmode As Boolean = False

    Public UseTrack_UPDTview As Boolean = False
    Public UseTrack_PDTview As Boolean = False
    Public UseTrack_PROverallTrends As Boolean = False
    Public UseTrack_RawDatawindow_Main As Boolean = False
    Public UseTrack_RawDatawindow_Paretos As Boolean = False
    Public UseTrack_RawDatawindow_Variance As Boolean = False
    Public UseTrack_WeibullMain As Boolean = False
    Public UseTrack_WeibullMain_failuremodes As Boolean = False
    Public UseTrack_IncontrolMain As Boolean = False
    Public UseTrack_IncontrolControlChart As Boolean = False
    Public UseTrack_IncontrolControlShift As Boolean = False
    Public UseTrack_TopStopsMain As Boolean = False
    Public UseTrack_StopsWatchMain As Boolean = False
    Public UseTrack_TopStopsTrends As Boolean = False
    Public UseTrack_ChangeMapping As Boolean = False
    Public UseTrack_Filter As Boolean = False
    Public UseTrack_ExportLossTree As Boolean = False
    Public UseTrack_ExportDowntime As Boolean = False
    Public UseTrack_ExportProduction As Boolean = False
    Public UseTrack_ExportDependency As Boolean = False
    Public UseTrack_Notes As Boolean = False
    Public UseTrack_Simulation As Boolean = False
    Public UseTrack_Notes_PickaLoss As Boolean = False
    Public UseTrack_Notes_ExporttoExcel As Boolean = False
    Public UseTrack_TargetsMain As Boolean = False
    Public UseTrack_RawDataWindow_Trends As Boolean = False

    Public UseTrack_Multiline_ByLossAreachartsmain As Boolean = False
    Public UseTrack_Multiline_RollupCharts As Boolean = False
    Public UseTrack_Multiline_ByLossAreadrilldown1 As Boolean = False
    Public UseTrack_Multiline_ByLossAreadrilldown2 As Boolean = False
    Public UseTrack_Multiline_ByLossAreadrilldown3 As Boolean = False
    Public UseTrack_Multiline_RawData As Boolean = False
    Public UseTrack_Multiline_Rollupdrilldown As Boolean = False


    Public ErrorFunctionName As String = ""

    Public IsNotesCreateMode As Boolean = False
    Public currentNotesfilename As String = ""


    Public mappinglevel1 As Integer ' Only used to retain original mapping while opening rawdatawindow
    Public mappinglevel2 As Integer ' Only used to retain original mapping while opening rawdatawindow

    Public f As MouseButtonEventArgs

    Public datapull_duration As Integer = 99
    Public DeactivateIncontrol As Boolean = False
    Public IsExcludedEventsIncluded As Boolean = False
End Module

