Imports System.Linq

Public Class LiveLineHelper

    public  LiveLine_AnalysisPeriodData As SummaryReport

    public  LiveLine_AnalysisPeriodData_PreviousPeriod As SummaryReport

    public  rawData As SummaryReport 'all data for a given line
    public  AnalysisPeriodData As SummaryReport 'analysis period data for the line

    public sub New(rawdatax as summaryreport, analysisperioddatax as summaryreport)
        rawdata = rawdatax
        analysisperioddata = analysisperioddatax

        liveline_initialize()
    End sub

    Private Sub LiveLine_initialize()
        LiveLine_TimeFrameInDays = 1
        LiveLine_AnalysisPeriodData = rawData.getSubset(AnalysisPeriodData.endTime.AddDays((LiveLine_TimeFrameInDays * -1)), AnalysisPeriodData.endTime)
        LiveLine_AnalysisPeriodData_PreviousPeriod = rawData.getSubset(AnalysisPeriodData.endTime.AddDays(((2 * LiveLine_TimeFrameInDays) * -1)), AnalysisPeriodData.endTime.AddDays((LiveLine_TimeFrameInDays * -1)))
        LiveLine_AnalysisPeriodData.reMapDowntime(LiveLine_Mapping_A, LiveLine_Mapping_B)
        LiveLine_PopulateIntermediateSheet()
    End Sub

    Public ReadOnly Property LiveLine_SelectedStartTime As DateTime
        Get
            Return Me.LiveLine_AnalysisPeriodData.startTime
        End Get
    End Property

    Public ReadOnly Property LiveLine_SelectedEndTime As DateTime
        Get
            Return Me.LiveLine_AnalysisPeriodData.endTime
        End Get
    End Property

    Private Property LiveLine_TimeFrameInDays As Integer

    Private Sub LiveLine_PopulateIntermediateSheet()
        LiveLine_ActualDurationOfEachEvent.Clear()
        LiveLine_EventStartTimes.Clear()
        LiveLine_DTviewer_EventNames.Clear()
        LiveLine_EventTypes.Clear()
        LiveLine_TopLosses.Clear()
        Dim i As Integer = 0
        Do While (i < Me.LiveLine_AnalysisPeriodData.rawData.Count)
            If (Me.LiveLine_AnalysisPeriodData.rawData(i).UT > 0) Then
                'uptime
                LiveLine_ActualDurationOfEachEvent.Add(Me.LiveLine_AnalysisPeriodData.rawData(i).UT)
                If Me.LiveLine_AnalysisPeriodData.rawData(i).isExcluded Then
                    LiveLine_EventTypes.Add(EventType.Excluded)
                Else
                    LiveLine_EventTypes.Add(EventType.Running)
                End If

                LiveLine_DTviewer_EventNames.Add("uptime")
                LiveLine_EventStartTimes.Add(Me.LiveLine_AnalysisPeriodData.rawData(i).startTime_UT)
                'downtime
                LiveLine_ActualDurationOfEachEvent.Add(Me.LiveLine_AnalysisPeriodData.rawData(i).DT)
                If Me.LiveLine_AnalysisPeriodData.rawData(i).isUnplanned Then
                    LiveLine_EventTypes.Add(EventType.Unplanned)
                ElseIf Me.LiveLine_AnalysisPeriodData.rawData(i).isPlanned Then
                    LiveLine_EventTypes.Add(EventType.Planned)
                ElseIf Me.LiveLine_AnalysisPeriodData.rawData(i).isExcluded Then
                    LiveLine_EventTypes.Add(EventType.Excluded)
                Else
                    LiveLine_EventTypes.Add(EventType.Excluded)
                End If

                LiveLine_DTviewer_EventNames.Add(Me.LiveLine_AnalysisPeriodData.rawData(i).MappedField)
                LiveLine_EventStartTimes.Add(Me.LiveLine_AnalysisPeriodData.rawData(i).startTime)
            End If

            i = (i + 1)
        Loop

        'Top Losses
        Dim prevPeriodOEE As Double
        Dim prevPeriodStops As Double
        Dim prevPeriodIndex As Integer
        Dim topLosses = New List(Of Tuple(Of String, Double, Integer, Double, Integer))
        Dim topLosses_Planned = New List(Of Tuple(of string, double, double))
        i = 0
        Do While (i < Me.LiveLine_AnalysisPeriodData.DT_Report.MappedDirectory.Count)
            Me.LiveLine_AnalysisPeriodData.DT_Report.MappedDirectory(i).SchedTime = Me.LiveLine_AnalysisPeriodData.schedTime
            prevPeriodIndex = Me.LiveLine_AnalysisPeriodData_PreviousPeriod.DT_Report.MappedDirectory.IndexOf(New DTevent(Me.LiveLine_AnalysisPeriodData.DT_Report.MappedDirectory(i).Name, 0))
            If (prevPeriodIndex >= 0) Then
                Me.LiveLine_AnalysisPeriodData_PreviousPeriod.DT_Report.MappedDirectory(prevPeriodIndex).SchedTime = Me.LiveLine_AnalysisPeriodData_PreviousPeriod.schedTime
                prevPeriodOEE = Me.LiveLine_AnalysisPeriodData_PreviousPeriod.DT_Report.MappedDirectory(prevPeriodIndex).DTpct
                prevPeriodStops = Me.LiveLine_AnalysisPeriodData_PreviousPeriod.DT_Report.MappedDirectory(prevPeriodIndex).Stops
            Else
                prevPeriodOEE = 0
                prevPeriodStops = 0
            End If

            topLosses.Add(New Tuple(Of String, Double, Integer, Double, Integer)(Me.LiveLine_AnalysisPeriodData.DT_Report.MappedDirectory(i).Name, Me.LiveLine_AnalysisPeriodData.DT_Report.MappedDirectory(i).DTpct, CType(Me.LiveLine_AnalysisPeriodData.DT_Report.MappedDirectory(i).Stops, Integer), prevPeriodOEE, CType(prevPeriodStops, Integer)))
            i = (i + 1)
        Loop

        LiveLine_TopLosses = topLosses.OrderBy(function(x) x.Item2).ToList
        LiveLine_TopLosses.Reverse()
        'planned
        i = 0

        dim tmpDirectory as list(of dtevent) = Me.LiveLine_AnalysisPeriodData.DT_Report.getPlannedEventDirectory(DowntimeField.Tier1)

        Do While (i < tmpDirectory.Count) 'this is questionable at best !TODO
            tmpDirectory(i).SchedTime = Me.LiveLine_AnalysisPeriodData.schedTime
            topLosses_Planned.Add(New Tuple(Of String, Double, Double)(tmpDirectory(i).Name, tmpDirectory(i).DT, tmpDirectory(i).DTpct))
            i = (i + 1)
        Loop

        LiveLine_Planned = topLosses_Planned.OrderBy(function(x) x.Item2).ToList
        LiveLine_Planned.Reverse()
        'trends
        Me.LiveLine_populateIntermediate_Trends()
        'update the biggest changes
        LiveLine_BiggestChanges.Clear()
        i = 0
        Do While (i < LiveLine_TopLosses.Count)
            LiveLine_BiggestChanges.Add(New Tuple(Of String, Double, Double)(LiveLine_TopLosses(i).Item1, (LiveLine_TopLosses(i).Item4 - LiveLine_TopLosses(i).Item2), (LiveLine_TopLosses(i).Item5 - LiveLine_TopLosses(i).Item3)))
            i = (i + 1)
        Loop

    End Sub

#Region "Mapping"

    Public Property LiveLine_Mapping_A As DowntimeField = DowntimeField.Tier2
    Public Property LiveLine_Mapping_B As DowntimeField = DowntimeField.NA

    Public Sub LiveLine_ReMap(ByVal MappingA As DowntimeField, ByVal MappingB As DowntimeField)
        Me.LiveLine_Mapping_A = MappingA
        Me.LiveLine_Mapping_B = MappingB
        Me.LiveLine_AnalysisPeriodData.reMapDowntime(MappingA, MappingB)
        Me.LiveLine_AnalysisPeriodData_PreviousPeriod.reMapDowntime(MappingA, MappingB)
        Me.LiveLine_PopulateIntermediateSheet()
    End Sub
#End Region

#Region "DTviewer"

    Public Sub LiveLine_setDTviewerTimeFrame(ByVal timeFrameInDays As Integer)
        Me.LiveLine_TimeFrameInDays = timeFrameInDays
        Me.LiveLine_AnalysisPeriodData = rawData.getSubset(AnalysisPeriodData.endTime.AddDays((timeFrameInDays * -1)), AnalysisPeriodData.endTime)
        Me.LiveLine_AnalysisPeriodData_PreviousPeriod = rawData.getSubset(AnalysisPeriodData.endTime.AddDays(((2 * Me.LiveLine_TimeFrameInDays) _
                            * -1)), AnalysisPeriodData.endTime.AddDays((Me.LiveLine_TimeFrameInDays * -1)))
        Me.LiveLine_PopulateIntermediateSheet()
    End Sub

    Public Property LiveLine_ActualDurationOfEachEvent As List(Of Double) = new List(of double)

    Public ReadOnly Property LiveLine_NumberOfEvents As Integer
        Get
            Return Me.LiveLine_ActualDurationOfEachEvent.Count
        End Get
    End Property

    Public Property LiveLine_EventTypes As List(Of EventType) = New List(Of EventType)
    Public Property LiveLine_EventStartTimes As List(Of DateTime) = New List(Of DateTime)
    Public Property LiveLine_DTviewer_EventNames As List(Of String) = New List(Of String)
#End Region

#Region "Top Loss / Biggest Changes"

    Public LiveLine_TopLosses As List(Of Tuple(Of String, Double, Integer, Double, Integer)) = New List(Of Tuple(Of String, Double, Integer, Double, Integer))

    'double -> dtpct/OEE impact, 2nd double -> duration. sorted by longest DT
    Public LiveLine_Planned As List(Of Tuple(Of String, Double, Double)) = New List(Of Tuple(Of String, Double, Double))

    Public LiveLine_BiggestChanges As List(Of Tuple(Of String, Double, Double)) = New List(Of Tuple(Of String, Double, Double))

    Public ReadOnly Property LiveLine_TopLoss_MaxValue_Planned() As Double
        Get
            Return If(LiveLine_Planned.Count = 0, 0, LiveLine_Planned.[Select](Function(t) t.Item2).ToList().Max())
        End Get
    End Property

    Public ReadOnly Property LiveLine_TopLoss_MaxLossValue() As Double
        Get
            Return If(LiveLine_TopLosses.Count = 0, 0, LiveLine_TopLosses.[Select](Function(t) t.Item2).ToList().Max())
        End Get
    End Property
#End Region

#Region "Trend"

    Public LiveLine_TrendsData As List(Of Tuple(Of DateTime, Double, Integer)) = New List(Of Tuple(Of DateTime, Double, Integer))

    Private LiveLine_TrendRawData As List(Of SummaryReport) = New List(Of SummaryReport)

    Public ReadOnly Property LiveLine_Trends_MaxOEE() As Double
        Get
            Return LiveLine_TrendsData.[Select](Function(t) t.Item2).ToList().Max()
        End Get
    End Property
    Public ReadOnly Property LiveLine_Trends_MaxStops() As Double
        Get
            Return LiveLine_TrendsData.[Select](Function(t) t.Item3).ToList().Max()
        End Get
    End Property

    Private Sub LiveLine_populateIntermediate_Trends()
        Dim TrendTimePeriod_Hours As Double = 1
        Dim TrendTimePeriod_Number As Double = 24
        Dim tmpStartTime As DateTime
        Dim tmpEndTime As DateTime
        Me.LiveLine_TrendsData.Clear()
        Me.LiveLine_TrendRawData.Clear()
        'figure out the time periods to display
        If (Me.LiveLine_TimeFrameInDays < 7) Then
            TrendTimePeriod_Hours = 1
            TrendTimePeriod_Number = 24
        ElseIf (Me.LiveLine_TimeFrameInDays < 30) Then
            TrendTimePeriod_Hours = 24
            TrendTimePeriod_Number = 7
        Else
            TrendTimePeriod_Hours = 24
            TrendTimePeriod_Number = 30
        End If

        'populate the raw data list
        Dim i As Integer = 0
        Do While (i < TrendTimePeriod_Number)
            tmpEndTime = rawData.endTime.AddHours(((i * TrendTimePeriod_Hours) _
                            * -1))
            tmpStartTime = tmpEndTime.AddHours((TrendTimePeriod_Hours * -1))
            Me.LiveLine_TrendRawData.Add(rawData.getSubset(tmpStartTime, tmpEndTime))
            i = (i + 1)
        Loop

        Me.LiveLine_TrendRawData.Reverse()
        'convert raw data list to intermediate sheet
        i = 0
        Do While (i < Me.LiveLine_TrendRawData.Count)
            Me.LiveLine_TrendsData.Add(New Tuple(Of DateTime, Double, Integer)(Me.LiveLine_TrendRawData(i).startTime, Me.LiveLine_TrendRawData(i).PR, CType(Me.LiveLine_TrendRawData(i).Stops, Integer)))
            i = (i + 1)
        Loop

    End Sub
#End Region
End Class
