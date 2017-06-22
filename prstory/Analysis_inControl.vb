Module inControlVars
    Public Const inCONTROL_SCHED_TIME_CUTOFF As Integer = 180
    Public Const inCONTROL_EventsToShow As Integer = 15
End Module
Public Class inControlReport

#Region "Variables & Properties"
    Private parentLine As ProdLine
    Private rawDTdata As Array
    Private mappingCol As Integer
    Private startDate As Date 'baseline data start date
    Private _analysisStartDate As Date
    Private _analysisEndDate As Date
    Private timePeriods As Integer
    Private _RankByStops As Integer = False
    Private timePeriodsInAnalysisPeriod As Integer

    Private _maxStops As Integer
    Private _maxDT As Integer

    Public ReadOnly Property maxStops As Integer
        Get
            Return _maxStops  'SRONEW AnalysisPeriodReport.EventMaxStops '_maxStops
        End Get
    End Property
    Public ReadOnly Property maxDTpct As Double
        Get
            If AnalysisPeriodReport.schedTime = 0 Then Return 0
            Return _maxDT / AnalysisPeriodReport.schedTime   'SRONEW AnalysisPeriodReport.EventMaxDT / AnalysisPeriodReport.schedTime
        End Get
    End Property

    Public ReadOnly Property AnalysisStartDate As Date
        Get
            Return _analysisStartDate
        End Get
    End Property
    Public ReadOnly Property AnalysisEndDate As Date
        Get
            Return _analysisEndDate
        End Get
    End Property

    Public ReadOnly Property MostRecentStartDate As Date
        Get
            '   Return AnalysisStartDate 'DailyReports(DailyReports.Count - 1).StartTime
            Return DailyReports(DailyReports.Count - 1).StartTime
        End Get
    End Property
    Public ReadOnly Property MostRecentEndDate As Date
        Get
            '  Return AnalysisEndDate 'DailyReports(DailyReports.Count - 1).EndTime
            Return DailyReports(DailyReports.Count - 1).EndTime
        End Get
    End Property
    Private Property RankByStops As Boolean
        Get
            Return _RankByStops
        End Get
        Set(value As Boolean)
            If Not value.Equals(_RankByStops) Then
                _RankByStops = value
                For i = 0 To DailyReports.Count - 1
                    'SRONEW     DailyReports(i).sortUnplanned(_RankByStops)
                Next
            End If
        End Set
    End Property
    Public ReadOnly Property SPCNumberOfDays As Integer
        Get
            Return DailyReports.Count
        End Get
    End Property

    Private DailyReports As New List(Of DowntimeReport)
    Private AnalysisPeriodReport As DowntimeReport

    Friend inControlEvents As New List(Of inControlDTevent)
#End Region

#Region "Construction & Reinitialization"
    Public Sub analyzeNewTimePeriod(startDate As Date, endDate As Date)
        mappingCol = My.Settings.defaultMappingLevel
        _analysisStartDate = startDate
        _analysisEndDate = endDate
        AnalysisPeriodReport = New DowntimeReport(parentLine, startDate, endDate)
        generateDailyAnalysis()
    End Sub

    'constructor
    Public Sub New(myParentLine As ProdLine, analysisStartDate As Date, analysisEndDate As Date) ', mappingColumn As Integer)
        Dim timePeriodIncrementer As Integer, tmpProfReport As DowntimeReport
        parentLine = myParentLine
        rawDTdata = parentLine.rawProficyData
        mappingCol = My.Settings.defaultMappingLevel 'mappingColumn
        _analysisEndDate = analysisEndDate
        _analysisStartDate = analysisStartDate

        'adjust our start date
        startDate = DateAdd(DateInterval.Second, -Second(parentLine.rawProfStartTime), parentLine.rawProfStartTime)
        startDate = DateAdd(DateInterval.Minute, -Second(startDate), startDate)
        If Hour(startDate) > parentLine.ShiftStartFirst_Hr Then ' need to go to the next day
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
            startDate = DateAdd(DateInterval.Day, 1, startDate)
        Else 'naw this is cool
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
        End If
        timePeriods = DateDiff(DateInterval.Day, startDate, parentLine.rawProfEndTime)

        'lets look at our analysis data
        timePeriodsInAnalysisPeriod = 1 'DateDiff(DateInterval.Day, analysisStartDate, analysisEndDate)

        'create our data list
        For timePeriodIncrementer = 0 To timePeriods - 1
            tmpProfReport = New DowntimeReport(parentLine, DateAdd(DateInterval.Day, timePeriodIncrementer, startDate), DateAdd(DateInterval.Day, timePeriodIncrementer + 1, startDate))

            If tmpProfReport.schedTime > inCONTROL_SCHED_TIME_CUTOFF Then
                DailyReports.Add(tmpProfReport)  ' the error is probably here 
            End If
        Next

        If DailyReports.Count <> 0 Then
            AnalysisPeriodReport = DailyReports(DailyReports.Count - 1)



            _analysisEndDate = AnalysisPeriodReport.EndTime
            _analysisStartDate = AnalysisPeriodReport.StartTime
            generateDailyAnalysis()
            DeactivateIncontrol = False
        End If
    End Sub
#End Region


    Private Sub generateDailyAnalysis()
        Dim eventIncrementer As Integer
        If AnalysisPeriodReport.schedTime = 0 Or AnalysisPeriodReport.UT = 0 Then
            inControlEvents.Clear()
            ' MsgBox("No Production In Selected Time Period! Please Select Another Time Period", vbExclamation, "No PR:IN Production") ' LG Code
        Else
            'redo it (or do it)
            inControlEvents.Clear()

            _maxDT = 0
            _maxStops = 0
            With AnalysisPeriodReport.UnplannedEventDirectory
                For i As Integer = 0 To .Count - 1
                    If AnalysisPeriodReport.UnplannedEventDirectory(i).DT > _maxDT Then _maxDT = AnalysisPeriodReport.UnplannedEventDirectory(i).DT
                    If AnalysisPeriodReport.UnplannedEventDirectory(i).Stops > _maxStops Then _maxStops = AnalysisPeriodReport.UnplannedEventDirectory(i).Stops
                Next
            End With

            With AnalysisPeriodReport
                For eventIncrementer = 0 To Math.Min(inCONTROL_EventsToShow - 1, .UnplannedEventDirectory.Count - 1)
                    inControlEvents.Add(New inControlDTevent(DailyReports, .UnplannedEventDirectory(eventIncrementer), .UT, .schedTime, timePeriodsInAnalysisPeriod))
                Next
            End With
        End If
    End Sub
End Class

Public Class MotionReport
    Private parentLine As ProdLine
    Private rawDTdata As Array
    Private mappingCol As Integer
    Private startDate As Date 'baseline data start date
    Private timePeriods As Integer
    Private eventList As New List(Of DTevent)

    Friend DailyReports As New List(Of SummaryReport)
    Friend motionEvents As New List(Of motionDTevent)
    Friend motionEvents_All15 As New List(Of motionDTevent) ' LG Code

    Public Function getHTMLtitleDataString() As String
        Dim tmpString As String = "", i As Integer
        For i = 0 To motionEvents.Count - 1
            tmpString = tmpString & "data.addColumn('number', '" & motionEvents(i).Name & "');" & vbCrLf
        Next
        Return tmpString
    End Function
    Public Function getHTMLdataString(dayNumber As Integer, isDT As Boolean) As String
        Dim tmpString As String = "", i As Integer, tmpTime As Date, tmpDT As Double
        tmpTime = DailyReports(dayNumber).startTime

        tmpString = tmpString & "[new Date (" & Year(tmpTime) & "," & Month(tmpTime) - 1 & "," & Day(tmpTime) & ")," & vbCrLf

        If motionEvents.Count > 0 Then

            For i = 0 To motionEvents.Count - 2
                If isDT Then
                    tmpDT = motionEvents(i).DailyDTpct(dayNumber)
                    '     If tmpDT > 1.1 Then tmpDT = 1.1
                    tmpString = tmpString & tmpDT & "," & vbCrLf
                Else
                    tmpString = tmpString & motionEvents(i).DailySPD(dayNumber) & "," & vbCrLf
                End If
            Next
            i = motionEvents.Count - 1
            If isDT Then
                tmpDT = motionEvents(i).DailyDTpct(dayNumber)
                '   If tmpDT > 1.1 Then tmpDT = 1.1
                tmpString = tmpString & tmpDT & "]" '& vbCrLf
            Else
                tmpString = tmpString & motionEvents(i).DailySPD(dayNumber) & "]" '& vbCrLf
            End If

        Else
            tmpString = tmpString & "0]"
        End If

        Return tmpString
    End Function

    Public Function getHTMLtitleDataString_selectedfailuremode(failuremodeno As Integer) As String
        Dim tmpString As String = ""

        tmpString = tmpString & "data.addColumn('number', '" & motionEvents_All15(failuremodeno).Name & "');" & vbCrLf

        Return tmpString
    End Function
    Public Function getHTMLdataString_selectedfailuremode(dayNumber As Integer, isDT As Boolean, failuremodeno As Integer) As String
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime

        tmpString = tmpString & "[new Date (" & Year(tmpTime) & "," & Month(tmpTime) - 1 & "," & Day(tmpTime) & ")," & vbCrLf

        If motionEvents_All15.Count > 0 Then


            If isDT Then
                tmpString = tmpString & motionEvents_All15(failuremodeno).DailyDTpct(dayNumber) & "," & vbCrLf
            Else
                tmpString = tmpString & motionEvents_All15(failuremodeno).DailySPD(dayNumber) & "," & vbCrLf
            End If


            tmpString = tmpString & "]"
        Else
            tmpString = tmpString & "0]"
        End If

        Return tmpString
    End Function
    Public Function getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd(dayNumber As Integer, isDT As Boolean, failuremodeno As Integer) As Double


        If motionEvents_All15.Count > 0 Then


            If isDT Then
                Return motionEvents_All15(failuremodeno).DailyDTpct(dayNumber)
            Else
                Return motionEvents_All15(failuremodeno).DailySPD(dayNumber)
            End If
        Else
            Return 0


        End If


    End Function
    Public Function getHTMLdataString_selectedfailuremode_AMCHARTS_Date(dayNumber As Integer, isDT As Boolean, failuremodeno As Integer) As Object
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime

        Return DateValue(tmpTime)


    End Function

    'constructor
    Public Sub New(myParentLine As ProdLine, analysisStartDate As Date, analysisEndDate As Date, analysisEventList As List(Of DTevent), daysInEachAnalysisPeriod As Integer) ', mappingColumn As Integer)
        Dim timePeriodIncrementer As Integer, tmpProfReport As SummaryReport, analysisDays As Integer
        parentLine = myParentLine
        rawDTdata = parentLine.rawProficyData
        mappingCol = My.Settings.defaultMappingLevel 'mappingColumn
        analysisDays = daysInEachAnalysisPeriod
        eventList = analysisEventList
        'adjust our start date
        startDate = DateAdd(DateInterval.Second, -Second(parentLine.rawProfStartTime), parentLine.rawProfStartTime)
        startDate = DateAdd(DateInterval.Minute, -Second(startDate), startDate)
        If Hour(startDate) > parentLine.ShiftStartFirst_Hr Then ' need to go to the next day
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
            startDate = DateAdd(DateInterval.Day, 1, startDate)
        Else 'naw this is cool
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
        End If
        timePeriods = DateDiff(DateInterval.Day, startDate, parentLine.rawProfEndTime) / analysisDays

        'create our data list
        For timePeriodIncrementer = 0 To timePeriods - 1
            tmpProfReport = New SummaryReport(parentLine, DateAdd(DateInterval.Day, timePeriodIncrementer, startDate), DateAdd(DateInterval.Day, timePeriodIncrementer + analysisDays, startDate))
            DailyReports.Add(tmpProfReport)
        Next
        generateDailyAnalysis()
    End Sub

    Private Sub generateDailyAnalysis()
        Dim eventIncrementer As Integer

        'find out maxes and mins for the chart
        '   With AnalyssisPeriodReport
        '   .sortUnplanned(False) 'sort by dt
        '  End With
        'redo it (or do it)
        motionEvents.Clear()
        motionEvents_All15.Clear()
        ' With eventlist
        For eventIncrementer = 0 To eventList.Count - 1  ' Math.Min(14, eventList.Count - 1)
            motionEvents_All15.Add(New motionDTevent(DailyReports, eventList(eventIncrementer)))
        Next
        eventIncrementer = 0
        For eventIncrementer = 0 To Math.Min(4, eventList.Count - 1)
            motionEvents.Add(New motionDTevent(DailyReports, eventList(eventIncrementer)))
        Next


        '  End With
    End Sub


   
    Public Function getHTMLdataString_selectedfailuremode_AMCHARTS_Date_Monthly(monthNumber As Integer, isDT As Boolean, failuremodeno As Integer) As String
 

        Select Case monthNumber
            Case 3
                Return (DailyReports(DailyReports.Count - 30).startTime.ToShortDateString & " to " & DailyReports(DailyReports.Count - 1).startTime.ToShortDateString)
            Case 2
                Return (DailyReports(DailyReports.Count - 60).startTime.ToShortDateString & " to " & DailyReports(DailyReports.Count - 31).startTime.ToShortDateString)
            Case 1
                Return (DailyReports(DailyReports.Count - 90).startTime.ToShortDateString & " to " & DailyReports(DailyReports.Count - 61).startTime.ToShortDateString)
        End Select
       return "error!"
    End Function
    Public Function getHTMLdataString_selectedfailuremode_AMCHARTS_Date_Weekly(startday As Integer, isDT As Boolean, failuremodeno As Integer, endday As Integer) As String
        Return (DailyReports(startday).startTime.ToString("MM/dd") & "-" & DailyReports(endday).startTime.ToString("MM/dd"))
    End Function

    Public Function getHTMLdataString_AMCharts_DTpctorSPD_Weekly(startday As Integer, isDT As Boolean, failuremodeno As Integer, endday As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0

        If isDT Then
            For i = startday To endday Step 1

                instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyDTpct(i) * DailyReports(i).schedTime)
                instasum_denum = instasum_denum + (DailyReports(i).schedTime)

            Next i
        Else
            For i = startday To endday Step 1

                instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailySPD(i) * DailyReports(i).schedTime)
                instasum_denum = instasum_denum + (DailyReports(i).schedTime)

            Next i

        End If

        Return ((instasum_num / instasum_denum))



    End Function
    Public Function getHTMLdataString_selectedfailuremode_AMCHARTS_DTpctORspd_Monthly(monthNumber As Integer, isDT As Boolean, failuremodeno As Integer) As Double


        If motionEvents_All15.Count > 0 Then
            Dim i As Integer
            Dim instasum_num As Double = 0
            Dim instasum_denum As Double = 0

            If isDT Then



                Select Case monthNumber
                    Case 3

                        For i = motionEvents_All15(failuremodeno).DailyDTpct.Count - 1 To motionEvents_All15(failuremodeno).DailyDTpct.Count - 30 Step -1

                            instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyDTpct(i) * DailyReports(i).schedTime)
                            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                        Next i
                        Return ((instasum_num / instasum_denum))

                    Case 2
                        For i = motionEvents_All15(failuremodeno).DailyDTpct.Count - 31 To motionEvents_All15(failuremodeno).DailyDTpct.Count - 60 Step -1

                            instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyDTpct(i) * DailyReports(i).schedTime)
                            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                        Next i
                        Return ((instasum_num / instasum_denum))

                    Case 1
                        For i = motionEvents_All15(failuremodeno).DailyDTpct.Count - 61 To motionEvents_All15(failuremodeno).DailyDTpct.Count - 90 Step -1

                            instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyDTpct(i) * DailyReports(i).schedTime)
                            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                        Next i
                        Return ((instasum_num / instasum_denum))
                End Select

            Else

                Select Case monthNumber
                    Case 3

                        For i = motionEvents_All15(failuremodeno).DailySPD.Count - 1 To motionEvents_All15(failuremodeno).DailySPD.Count - 30 Step -1

                            instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailySPD(i) * DailyReports(i).schedTime)
                            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                        Next i
                        Return ((instasum_num / instasum_denum))

                    Case 2
                        For i = motionEvents_All15(failuremodeno).DailySPD.Count - 31 To motionEvents_All15(failuremodeno).DailySPD.Count - 60 Step -1

                            instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailySPD(i) * DailyReports(i).schedTime)
                            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                        Next i
                        Return ((instasum_num / instasum_denum))

                    Case 1
                        For i = motionEvents_All15(failuremodeno).DailySPD.Count - 61 To motionEvents_All15(failuremodeno).DailySPD.Count - 90 Step -1

                            instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailySPD(i) * DailyReports(i).schedTime)
                            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                        Next i
                        Return ((instasum_num / instasum_denum))
                End Select




            End If
        Else
            Return 0


        End If


    End Function



















    ''''' '''''''''''''''''''''''''''''''''''''''''''''''''
    'AM CHarts MTBF
    'MTBF NEW CODE 
    '''''''''''''
    Public Function getHTMLdataString_selectedfailuremode_AMCHARTS_MTBF(dayNumber As Integer, isDT As Boolean, failuremodeno As Integer) As Double


        If motionEvents_All15.Count > 0 Then



            Return motionEvents_All15(failuremodeno).DailyMTBF(dayNumber)

        Else
            Return 0


        End If


    End Function
    Public Function getHTMLdataString_AMCharts_MTBF_Weekly(startday As Integer, isDT As Boolean, failuremodeno As Integer, endday As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0

        If isDT Then
            For i = startday To endday Step 1
                instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyUT(i))
                instasum_denum = instasum_denum + (motionEvents_All15(failuremodeno).DailyStops(i))

            Next i
        Else


        End If

        Return ((instasum_num / instasum_denum))



    End Function
    Public Function getHTMLdataString_AMCharts_MTBF_Monthly(monthNumber As Integer, isDT As Boolean, failuremodeno As Integer) As Double
        If motionEvents_All15.Count > 0 Then
            Dim i As Integer
            Dim instasum_num As Double = 0
            Dim instasum_denum As Double = 0




            Select Case monthNumber
                Case 3

                    For i = motionEvents_All15(failuremodeno).DailyDTpct.Count - 1 To motionEvents_All15(failuremodeno).DailyDTpct.Count - 30 Step -1

                        instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyUT(i))
                        instasum_denum = instasum_denum + (motionEvents_All15(failuremodeno).DailyStops(i))

                    Next i
                    Return ((instasum_num / instasum_denum))

                Case 2
                    For i = motionEvents_All15(failuremodeno).DailyDTpct.Count - 31 To motionEvents_All15(failuremodeno).DailyDTpct.Count - 60 Step -1

                        instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyUT(i))
                        instasum_denum = instasum_denum + (motionEvents_All15(failuremodeno).DailyStops(i))

                    Next i
                    Return ((instasum_num / instasum_denum))

                Case 1
                    For i = motionEvents_All15(failuremodeno).DailyDTpct.Count - 61 To motionEvents_All15(failuremodeno).DailyDTpct.Count - 90 Step -1

                        instasum_num = instasum_num + (motionEvents_All15(failuremodeno).DailyUT(i))
                        instasum_denum = instasum_denum + (motionEvents_All15(failuremodeno).DailyStops(i))

                    Next i
                    Return ((instasum_num / instasum_denum))
            End Select
        Else
            Return 0
        End If
    End Function
End Class

Public Class Motion_LinePRReport
    Private parentLine As ProdLine
    Private rawDTdata As Array
    Private startDate As Date 'baseline data start date
    Private timePeriods As Integer

    Friend DailyReports As New List(Of SummaryReport)

    Public ReadOnly Property lineName As String
        Get
            Return parentLine.Name
        End Get
    End Property

#Region "HTML Strings"
    Public Function getHTMLdataString(dayNumber As Integer) As String
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime
        With DailyReports(dayNumber)
            'dt pct , spd, mtbf
            tmpString = tmpString & "[new Date (" & Year(tmpTime) & "," & Month(tmpTime) - 1 & "," & Day(tmpTime) & ")," & vbCrLf
            tmpString = tmpString & Math.Min(1.1, Math.Max(.UPDTpct, -0.01)) * 100 & ", " & vbCrLf & Math.Min(1.1, Math.Max(.PDTpct, -0.01)) * 100 & ", " & vbCrLf & Math.Min(1.1, Math.Max(.PR, -0.01)) * 100 & "]" '& vbCrLf & .RateLossPct * 100 & ", 'RateLoss%']"
            Return tmpString
        End With
    End Function
    Public Function getHTMLdataString_AMCHarts_DateObj(dayNumber As Integer) As Object
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime
        With DailyReports(dayNumber)
            Return DateValue(tmpTime)
        End With
    End Function
    Public Function getHTMLdataString_AMCharts_UPDT(dayNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime
        With DailyReports(dayNumber)
            Return (Math.Min(1.1, Math.Max(.UPDTpct, -0.01))) * 100
        End With
    End Function
    Public Function getHTMLdataString_AMCharts_PDT(dayNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime
        With DailyReports(dayNumber)
            Return (Math.Min(1.1, Math.Max(.PDTpct, -0.01))) * 100
        End With
    End Function
    Public Function getHTMLdataString_AMCharts_PR(dayNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime
        With DailyReports(dayNumber)
            Return (Math.Min(1.1, Math.Max(.PR, -0.01))) * 100
        End With
    End Function
    Public Function getHTMLdataString_AMCharts_SPD(dayNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime
        With DailyReports(dayNumber)
            Return .SPD
        End With
    End Function
    Public Function getHTMLdataString_AMCharts_MTBF(dayNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double
        Dim tmpString As String = "", tmpTime As Date
        tmpTime = DailyReports(dayNumber).startTime
        With DailyReports(dayNumber)
            Return .UT / .Stops
        End With
    End Function

    Public Function getHTMLdataString_AMCharts_SPD_Weekly(startday As Integer, endday As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0

        For i = startday To endday Step 1

            instasum_num = instasum_num + (DailyReports(i).SPD * DailyReports(i).schedTime)
            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

        Next i
        Return (instasum_num / instasum_denum)



    End Function
    Public Function getHTMLdataString_AMCharts_MTBF_Weekly(startday As Integer, endday As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0

        For i = startday To endday Step 1

            instasum_num = instasum_num + ((DailyReports(i).UT / DailyReports(i).Stops) * DailyReports(i).schedTime)
            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

        Next i
        Return (instasum_num / instasum_denum)



    End Function
    Public Function getHTMLdataString_AMCharts_SPD_Monthly(monthNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0
        Select Case monthNumber
            Case 3
                For i = DailyReports.Count - 1 To DailyReports.Count - 30 Step -1

                    instasum_num = instasum_num + (DailyReports(i).SPD * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum))

            Case 2
                For i = DailyReports.Count - 31 To DailyReports.Count - 60 Step -1
                    instasum_num = instasum_num + (DailyReports(i).SPD * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum))

            Case 1
                For i = DailyReports.Count - 61 To DailyReports.Count - 90 Step -1
                    instasum_num = instasum_num + (DailyReports(i).SPD * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return (instasum_num / instasum_denum)
        End Select
        'With DailyReports(dayNumber)
        'Return (Math.Min(1.1, .PR)) * 100
        'End With
    End Function
    Public Function getHTMLdataString_AMCharts_MTBF_Monthly(monthNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0
        Select Case monthNumber
            Case 3
                For i = DailyReports.Count - 1 To DailyReports.Count - 30 Step -1

                    instasum_num = instasum_num + ((DailyReports(i).UT / DailyReports(i).Stops) * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum))

            Case 2
                For i = DailyReports.Count - 31 To DailyReports.Count - 60 Step -1
                    instasum_num = instasum_num + ((DailyReports(i).UT / DailyReports(i).Stops) * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum))

            Case 1
                For i = DailyReports.Count - 61 To DailyReports.Count - 90 Step -1
                    instasum_num = instasum_num + ((DailyReports(i).UT / DailyReports(i).Stops) * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return (instasum_num / instasum_denum)
        End Select
        'With DailyReports(dayNumber)
        'Return (Math.Min(1.1, .PR)) * 100
        'End With
    End Function


    Public Function getHTMLdataString_AMCharts_Dateobj_Monthly(monthNumber As Integer) As String

        Select Case monthNumber
            Case 3
                Return (DailyReports(DailyReports.Count - 30).startTime.ToShortDateString & " to " & DailyReports(DailyReports.Count - 1).startTime.ToShortDateString)
            Case 2
                Return (DailyReports(DailyReports.Count - 60).startTime.ToShortDateString & " to " & DailyReports(DailyReports.Count - 31).startTime.ToShortDateString)
            Case 1
                Return (DailyReports(DailyReports.Count - 90).startTime.ToShortDateString & " to " & DailyReports(DailyReports.Count - 61).startTime.ToShortDateString)
        End Select

        return "error!"
    End Function
    Public Function getHTMLdataString_AMCharts_UPDT_Monthly(monthNumber As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0
        Select Case monthNumber
            Case 3
                For i = DailyReports.Count - 1 To DailyReports.Count - 30 Step -1

                    instasum_num = instasum_num + (DailyReports(i).UPDTpct * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)

            Case 2
                For i = DailyReports.Count - 31 To DailyReports.Count - 60 Step -1
                    instasum_num = instasum_num + (DailyReports(i).UPDTpct * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)

            Case 1
                For i = DailyReports.Count - 61 To DailyReports.Count - 90 Step -1
                    instasum_num = instasum_num + (DailyReports(i).UPDTpct * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)
        End Select
        'With DailyReports(dayNumber)
        'Return (Math.Min(1.1, .PR)) * 100
        'End With
    End Function
    Public Function getHTMLdataString_AMCharts_PDT_Monthly(monthNumber As Integer) As Double
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0
        Select Case monthNumber
            Case 3
                For i = DailyReports.Count - 1 To DailyReports.Count - 30 Step -1

                    instasum_num = instasum_num + (DailyReports(i).PDTpct * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)

            Case 2
                For i = DailyReports.Count - 31 To DailyReports.Count - 60 Step -1
                    instasum_num = instasum_num + (DailyReports(i).PDTpct * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)

            Case 1
                For i = DailyReports.Count - 61 To DailyReports.Count - 90 Step -1
                    instasum_num = instasum_num + (DailyReports(i).PDTpct * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)
        End Select
    End Function
    Public Function getHTMLdataString_AMCharts_PR_Monthly(monthNumber As Integer) As Double
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0
        Select Case monthNumber
            Case 3
                For i = DailyReports.Count - 1 To DailyReports.Count - 30 Step -1


                    instasum_num = instasum_num + (DailyReports(i).PR * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)

            Case 2
                For i = DailyReports.Count - 31 To DailyReports.Count - 60 Step -1

                    instasum_num = instasum_num + (DailyReports(i).PR * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)

            Case 1
                For i = DailyReports.Count - 61 To DailyReports.Count - 90 Step -1

                    instasum_num = instasum_num + (DailyReports(i).PR * DailyReports(i).schedTime)
                    instasum_denum = instasum_denum + (DailyReports(i).schedTime)

                Next i
                Return ((instasum_num / instasum_denum) * 100)
        End Select

    End Function
    Public Function getHTMLdataString_AMCharts_Dateobj_Weekly(startday As Integer, endday As Integer) As String


        Return (DailyReports(startday).startTime.ToString("MM/dd") & "-" & DailyReports(endday).startTime.ToString("MM/dd"))

    End Function
    Public Function getHTMLdataString_AMCharts_UPDT_Weekly(startday As Integer, endday As Integer) As Double

        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0

        For i = startday To endday Step 1

            instasum_num = instasum_num + (DailyReports(i).UPDTpct * DailyReports(i).schedTime)
            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

        Next i
        Return ((instasum_num / instasum_denum) * 100)



    End Function
    Public Function getHTMLdataString_AMCharts_PDT_Weekly(startday As Integer, endday As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0

        For i = startday To endday Step 1

            instasum_num = instasum_num + (DailyReports(i).PDTpct * DailyReports(i).schedTime)
            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

        Next i
        Return ((instasum_num / instasum_denum) * 100)


    End Function
    Public Function getHTMLdataString_AMCharts_PR_Weekly(startday As Integer, endday As Integer) As Double
        'Dim tmpSched As Double, tmpSPD As Double, tmpMTBF As Double

        'tmpTime = DailyReports(dayNumber).startTime
        Dim i As Integer
        Dim instasum_num As Double = 0
        Dim instasum_denum As Double = 0

        For i = startday To endday Step 1

            instasum_num = instasum_num + (DailyReports(i).PR * DailyReports(i).schedTime)
            instasum_denum = instasum_denum + (DailyReports(i).schedTime)

        Next i
        Return ((instasum_num / instasum_denum) * 100)



    End Function

#End Region


    Public Function findMtDStartDay() As Integer
        Dim i As Integer
        Dim selectedmonth As Integer = Month(endtimeselected)
        Dim startdayfound As Boolean = False


        For i = 0 To DailyReports.Count - 1
            If Month(DailyReports(i).startTime) = selectedmonth Then
                startdayfound = True
                Exit For
            End If
        Next
        If startdayfound = True Then
            Return i
        Else
            Return -1
        End If
    End Function
    Public Function findAnyDayinDailyReports(dateobj As Object) As Integer
        Dim i As Integer
        Dim selectedmonth As Integer = Month(dateobj)
        Dim selecteddate As Integer = Day(dateobj)
        Dim startdayfound As Boolean = False


        For i = 0 To DailyReports.Count - 1
            If Month(DailyReports(i).startTime) = selectedmonth And Day(DailyReports(i).startTime) + 1 >= selecteddate Then
                startdayfound = True
                Exit For
            End If
        Next
        If startdayfound = True Then
            Return i
        Else
            Return -1
        End If
    End Function

    Public Function getDailyData_Date(daynumber As Integer) As Date
        Return DailyReports(daynumber).startTime
    End Function
    Public Function getDailyData_Stops(daynumber As Integer) As Double
        Return DailyReports(daynumber).Stops
    End Function
    Public Function getDailyData_SchedTime(daynumber As Integer) As Double
        Return DailyReports(daynumber).schedTime
    End Function
    Public Function getDailyData_Uptime(daynumber As Integer) As Double
        Return DailyReports(daynumber).UT
    End Function
    'constructor
    Public Sub New(myParentLine As ProdLine, daysInEachAnalysisPeriod As Integer, optional startHourOffset as double = 0, optional endHourOffset as double = 0) ', mappingColumn As Integer)
        Dim timePeriodIncrementer As Integer, tmpProfReport As SummaryReport, analysisDays As Integer
        parentLine = myParentLine
        rawDTdata = parentLine.rawProficyData

        analysisDays = daysInEachAnalysisPeriod
        'eventList = analysisEventList
        'adjust our start date
        startDate = DateAdd(DateInterval.Second, -Second(parentLine.rawProfStartTime), parentLine.rawProfStartTime)
        startDate = DateAdd(DateInterval.Minute, -Second(startDate), startDate)
        If Hour(startDate) > parentLine.ShiftStartFirst_Hr Then ' need to go to the next day
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
            startDate = DateAdd(DateInterval.Day, 1, startDate)
        Else 'naw this is cool
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
        End If
        timePeriods = DateDiff(DateInterval.Day, startDate, parentLine.rawProfEndTime) / analysisDays

        'create our data list
        For timePeriodIncrementer = 0 To timePeriods - 1
            tmpProfReport = New SummaryReport(parentLine, DateAdd(DateInterval.Day, timePeriodIncrementer, startDate), DateAdd(DateInterval.Day, timePeriodIncrementer + analysisDays, startDate))
            DailyReports.Add(tmpProfReport)
        Next
        '  generateDailyAnalysis()
    End Sub


End Class

Public Class Motion_prstory
    Private parentLine As ProdLine
    Private mappingCol As Integer
    Private startDate As Date
    Private timePeriods As Integer

    Friend DailyReports As New List(Of prStoryMainPageReport)

    ' Public Function getHTMLdataTitleString(cardNumber As Integer)
    ' Dim tmpString As String = "", i As Integer
    ' With DailyReports(0)
    ' tmpString = tmpString & "data.addColumn('date', 'Date');" & vbCrLf
    '     For i = 0 To .getCardEventFields(cardNumber) - 1
    '         tmpString = tmpString & "data.addColumn('number', '" & getprStoryCardField(parentLine.prStoryMapping, cardNumber, i) & "');" & vbCrLf 'getCardEventInfo(cardNumber, i).Name & "');" & vbCrLf
    '     Next
    '  End With
    'Return tmpString
    ' End Function

    Public Function getHTMLdataTitleString_PlannedUnplanned(cardNumber As Integer)
        Dim tmpString As String = "", i As Integer
        With DailyReports(0)
            tmpString = tmpString & "data.addColumn('date', 'Date');" & vbCrLf
            For i = 1 To .getCardEventFields(cardNumber) - 2
                '         tmpString = tmpString & "data.addColumn('number', '" & getprStoryCardField(parentLine.prStoryMapping, cardNumber, i) & "');" & vbCrLf 'getCardEventInfo(cardNumber, i).Name & "');" & vbCrLf
            Next
        End With
        Return tmpString
    End Function

    Public Function getHTMLdataString_PlannedUnplanned(dayNumber As Integer, cardNumber As Integer, isDT As Boolean) As String
        Dim tmpString As String = "", tmpTime As Date, i As Integer, tmpIndex As Integer
        tmpTime = DailyReports(dayNumber).MainLEDSReport.startTime
        tmpString = tmpString & "[new Date (" & Year(tmpTime) & "," & Month(tmpTime) - 1 & "," & Day(tmpTime) & ")," & vbCrLf

        With DailyReports(dayNumber)
            For i = 1 To .getCardEventFields(cardNumber) - 3

                If tmpIndex = -1 Then
                    tmpString = tmpString & 0 & "," & vbCrLf
                End If
            Next
            '   i = .getCardEventFields(cardNumber) - 2
            '  tmpField = getprStoryCardField(parentLine.prStoryMapping, cardNumber, i)
            '   tmpIndex = DailyReports(i).getListIndexFromName(cardNumber, tmpField)
            If tmpIndex = -1 Then
                tmpString = tmpString & 0 & "]"
            Else
                If isDT Then
                    '           tmpString = tmpString & .getCardEventInfo(cardNumber, tmpIndex).DTpct * 100 & "]"
                Else
                    '          tmpString = tmpString & .getCardEventInfo(cardNumber, tmpIndex).SPD & "]"
                End If
            End If

        End With
        Return tmpString
    End Function

    'constructor
    Public Sub New(myParentLine As ProdLine, daysInEachAnalysisPeriod As Integer) ', mappingColumn As Integer)
        Dim timePeriodIncrementer As Integer, tmpstoryReport As prStoryMainPageReport, analysisDays As Integer
        parentLine = myParentLine
        mappingCol = My.Settings.defaultMappingLevel 'mappingColumn
        analysisDays = daysInEachAnalysisPeriod
        'adjust our start date
        startDate = DateAdd(DateInterval.Second, -Second(parentLine.rawProfStartTime), parentLine.rawProfStartTime)
        startDate = DateAdd(DateInterval.Minute, -Second(startDate), startDate)
        If Hour(startDate) > parentLine.ShiftStartFirst_Hr Then
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
            startDate = DateAdd(DateInterval.Day, 1, startDate)
        Else
            startDate = DateAdd(DateInterval.Hour, -Hour(startDate), startDate)
            startDate = DateAdd(DateInterval.Hour, parentLine.ShiftStartFirst_Hr, startDate)
        End If
        timePeriods = DateDiff(DateInterval.Day, startDate, parentLine.rawProfEndTime) / analysisDays

        For timePeriodIncrementer = 0 To timePeriods - 1
            tmpstoryReport = New prStoryMainPageReport(AllProdLines.IndexOf(parentLine), DateAdd(DateInterval.Day, timePeriodIncrementer, startDate), DateAdd(DateInterval.Day, timePeriodIncrementer + analysisDays, startDate))
            DailyReports.Add(tmpstoryReport)
        Next
    End Sub
End Class

Public Class motionDTevent
    Public ReadOnly Property Name As String
        Get
            Return baseEvent.Name
        End Get
    End Property
    Public Function getHTMLdataString(dayNumber As Integer) As String
        'dt pct , spd, mtbf
        Return DailyDTpct(dayNumber) & "," & DailySPD(dayNumber) & "," & DailyMTBF(dayNumber)
    End Function

    Private baseEvent As DTevent

    'primary 'raw' metrics
    Friend DailyStops As New List(Of Integer)

    Friend DailyDT As New List(Of Integer)

    Friend DailySched As New List(Of Double)
    Friend DailyUT As New List(Of Double)
    Friend DailyDTpct As New List(Of Double)
    Friend DailyStopspct As New List(Of Double)
    Friend DailyNetStops As New List(Of Double)
    'secondary metrics
    Friend DailyMTTR As New List(Of Double)
    Friend DailyMTBF As New List(Of Double)

    'tertiary metrics
    Friend DailySPD As New List(Of Double)
    Friend DailyInvMTDF As New List(Of Double)

    'LEDs data
    Private DailyReports As List(Of SummaryReport)

    'constructor
    Public Sub New(inputDailyReports As List(Of SummaryReport), baselineDTevent As DTevent) 'inTargetEvent As String) ', targetDay As Integer)
        DailyReports = inputDailyReports
        baseEvent = baselineDTevent
        AnalyzeInControlData()
    End Sub

    Private Sub AnalyzeInControlData()
        Dim timePeriodIncrementer As Integer, tmpIndex As Integer

        'set stops and dt
        For timePeriodIncrementer = 0 To DailyReports.Count - 1
            tmpIndex = DailyReports(timePeriodIncrementer).DT_Report.UnplannedEventDirectory.IndexOf(baseEvent)
            '  With DailyReports(timePeriodIncrementer).DT_Report.UnplannedEventDirectory(tmpIndex)
            If tmpIndex > -1 Then
                DailyStops.Add(DailyReports(timePeriodIncrementer).DT_Report.UnplannedEventDirectory(tmpIndex).Stops)
                DailyDT.Add(DailyReports(timePeriodIncrementer).DT_Report.UnplannedEventDirectory(tmpIndex).DT)
            Else
                DailyStops.Add(0)
                DailyDT.Add(0)
            End If

            With DailyReports(timePeriodIncrementer)
                DailySched.Add(.schedTime)
                DailyUT.Add(.UT_DT)
                DailyNetStops.Add(.Stops)
            End With
        Next
        'finish it up
        SetSecondaryAndTertiaryMetrics()
    End Sub

    Private Sub SetSecondaryAndTertiaryMetrics()
        For i = 0 To DailyStops.Count - 1
            If DailyStops(i) = 0 Or DailySched(i) = 0 Or DailyUT(i) = 0 Then
                DailyMTTR.Add(0)
                DailyMTBF.Add(0)
                DailyInvMTDF.Add(0)
                DailySPD.Add(0)
                DailyDTpct.Add(0)
                DailyStopspct.Add(0)
            Else
                DailyMTTR.Add(DailyDT(i) / DailyStops(i))
                DailySPD.Add(DailyStops(i) * 1440 / DailySched(i))
                DailyMTBF.Add(DailyUT(i) / DailyStops(i))
                DailyInvMTDF.Add(1 / DailyMTBF(i) * 1440)
                DailyDTpct.Add(DailyDT(i) * 100 / DailySched(i))
                DailyStopspct.Add(DailyStops(i) / DailyNetStops(i))
            End If
        Next
    End Sub

End Class
Public Class inControlDTevent
    Private _targetEvent As String
    Private _daysInTargetPeriod As Integer
    Private _UT As Double
    Private _schedTime As Double
    Private baseEvent As DTevent
    Friend DailyStops As New List(Of Integer)
    Friend DailyDT As New List(Of Integer)
    Friend DailyMTTR As New List(Of Double)
    Friend DailyMTBF As New List(Of Double)
    Friend DailySPD As New List(Of Double)
    Friend DailyInvMTDF As New List(Of Double)
    Private RawMu_InvMTDF As Double
    Private RawSigma_InvMTDF As Double
    Private RawMu_SPD As Double
    Private RawSigma_SPD As Double
    Friend AdjMu_SPD As Double
    Friend AdjSigma_SPD As Double
    Friend AdjMu_InvMTDF As Double
    Friend AdjSigma_InvMTDF As Double
    Private AdjDistFromMean_InvMTDF As Double
    Private DailyReports As List(Of DowntimeReport)

    Public Sub New(inputDailyReports As List(Of DowntimeReport), baselineDTevent As DTevent, baselineUT As Double, baselineSched As Double, daysInTargetPeriod As Integer) 'inTargetEvent As String) ', targetDay As Integer)
        _targetEvent = baselineDTevent.Name
        _schedTime = baselineSched
        _UT = baselineUT
        baseEvent = baselineDTevent
        DailyReports = inputDailyReports
        _daysInTargetPeriod = daysInTargetPeriod
        AnalyzeInControlData()
    End Sub

    Private Sub AnalyzeInControlData()
        Dim timePeriodIncrementer As Integer, tmpIndex As Integer
        Dim tmpEvent As DTevent = New DTevent(_targetEvent, 0)
        For timePeriodIncrementer = 0 To DailyReports.Count - 1
            With DailyReports(timePeriodIncrementer)
                tmpIndex = .UnplannedEventDirectory.IndexOf(tmpEvent)
                If tmpIndex > -1 Then
                    DailyStops.Add(.UnplannedEventDirectory(tmpIndex).Stops)
                    DailyDT.Add(.UnplannedEventDirectory(tmpIndex).DT)
                Else
                    DailyStops.Add(0)
                    DailyDT.Add(0)
                End If
            End With
        Next
        SetSecondaryAndTertiaryMetrics()
        SetMeansAndDevs_MTDF()
        SetMeansAndDevs_SPD()
        SetDistanceFromMean()
    End Sub

    Private Sub SetSecondaryAndTertiaryMetrics()
        For i = 0 To DailyStops.Count - 1
            If DailyStops(i) = 0 Then
                DailyMTTR.Add(0)
                DailyMTBF.Add(0)
                DailyInvMTDF.Add(0)
                DailySPD.Add(0)
            Else
                DailyMTTR.Add(DailyDT(i) / DailyStops(i))
                DailySPD.Add(DailyStops(i) * 1440 / DailyReports(i).schedTime)
                DailyMTBF.Add(DailyReports(i).UT / DailyStops(i))
                DailyInvMTDF.Add(1 / DailyMTBF(i) * 1440)
            End If
        Next
    End Sub
    Private Sub SetMeansAndDevs_MTDF()
        Dim Squares As New List(Of Double)
        Dim filteredList As New List(Of Double)
        Dim SquareAvg As Double, tmpMeanDist As Double
        RawMu_InvMTDF = DailyInvMTDF.Average()
        For Each value As Double In DailyInvMTDF
            Squares.Add(Math.Pow(value - RawMu_InvMTDF, 2))
        Next
        SquareAvg = Squares.Average
        RawSigma_InvMTDF = Math.Sqrt(SquareAvg)
        Squares.Clear()
        For i = 0 To DailyStops.Count - 1
            tmpMeanDist = (DailyInvMTDF(i) - RawMu_InvMTDF) / RawSigma_InvMTDF
            If (DailyStops(i) > 10 And tmpMeanDist > 2.5) Or (DailyStops(i) < 11 And tmpMeanDist > 3) Then
            Else
                filteredList.Add(DailyInvMTDF(i))
            End If
        Next
        AdjMu_InvMTDF = filteredList.Average
        For Each value As Double In filteredList
            Squares.Add(Math.Pow(value - AdjMu_InvMTDF, 2))
        Next
        SquareAvg = Squares.Average
        AdjSigma_InvMTDF = Math.Sqrt(SquareAvg)
    End Sub
    Private Sub SetMeansAndDevs_SPD()
        Dim Squares As New List(Of Double)
        Dim filteredList As New List(Of Double)
        Dim SquareAvg As Double, tmpMeanDist As Double
        RawMu_SPD = DailySPD.Average()
        For Each value As Double In DailySPD
            Squares.Add(Math.Pow(value - RawMu_SPD, 2))
        Next
        SquareAvg = Squares.Average
        RawSigma_SPD = Math.Sqrt(SquareAvg)
        Squares.Clear()
        For i = 0 To DailyStops.Count - 1
            tmpMeanDist = (DailySPD(i) - RawMu_SPD) / RawSigma_SPD
            If (DailyStops(i) > 10 And tmpMeanDist > 2.5) Or (DailyStops(i) < 11 And tmpMeanDist > 3) Then
            Else
                filteredList.Add(DailySPD(i))
            End If
        Next
        AdjMu_SPD = filteredList.Average
        For Each value As Double In filteredList
            Squares.Add(Math.Pow(value - AdjMu_SPD, 2))
        Next
        SquareAvg = Squares.Average
        AdjSigma_SPD = Math.Sqrt(SquareAvg)
    End Sub
    Private Sub SetDistanceFromMean()
        Dim tmpInvMTDF As Double, tmpMTBF As Double
        With baseEvent
            tmpMTBF = _UT / baseEvent.Stops
            tmpInvMTDF = 1440 / tmpMTBF
            AdjDistFromMean_InvMTDF = (tmpInvMTDF - AdjMu_InvMTDF) / AdjSigma_InvMTDF
        End With
    End Sub

    Public ReadOnly Property Stops As Integer
        Get
            Return baseEvent.Stops
        End Get
    End Property
    Public ReadOnly Property SPD As Double
        Get
            Return baseEvent.Stops / _schedTime * 1440
        End Get
    End Property
    Public ReadOnly Property DTpct As Double
        Get
            Return baseEvent.DT / _schedTime
        End Get
    End Property
    Public ReadOnly Property WesternRulesScore As Integer
        Get
            Return getWesternRulesScore()
        End Get
    End Property
    Private Function getWesternRulesScore() As Integer
        Dim tmpScore As Integer = 1
        Dim lastDay As Integer = DailyInvMTDF.Count - 1
        Dim dayIncrementer As Integer
        Dim isRuleOnePass As Boolean = True
        Dim isRuleTwoPass As Boolean = True
        Dim isRuleThreePass As Boolean = True
        Dim isRuleFourPass As Boolean = True
        Dim isRuleFivePass As Boolean = True
        Dim isRuleSixPass As Boolean = True
        Dim testDays As Integer, failDays As Integer

        Dim oneDevAbove As Double = AdjMu_InvMTDF + AdjSigma_InvMTDF
        Dim oneDevBelow As Double = AdjMu_InvMTDF - AdjSigma_InvMTDF
        Dim twoDevsAbove As Double = AdjMu_InvMTDF + 2 * AdjSigma_InvMTDF
        Dim threeDevsAbove As Double = AdjMu_InvMTDF + 3 * AdjSigma_InvMTDF

        If My.Settings.inControl_useRule1 Then
            For dayIncrementer = 0 To _daysInTargetPeriod - 1
                If DailyInvMTDF(lastDay - dayIncrementer) > threeDevsAbove Or DailyInvMTDF(lastDay - dayIncrementer - 1) > threeDevsAbove Then
                    isRuleOnePass = False
                    dayIncrementer = _daysInTargetPeriod
                End If
            Next
        End If

        If My.Settings.inControl_useRule2 Then
            For dayIncrementer = 0 To _daysInTargetPeriod - 1
                If DailyInvMTDF(lastDay - dayIncrementer) > twoDevsAbove And DailyInvMTDF(lastDay - dayIncrementer - 1) > twoDevsAbove _
                   Or DailyInvMTDF(lastDay - dayIncrementer - 2) > twoDevsAbove And DailyInvMTDF(lastDay - dayIncrementer - 1) > twoDevsAbove _
                    Or DailyInvMTDF(lastDay - dayIncrementer - 2) > twoDevsAbove And DailyInvMTDF(lastDay - dayIncrementer - 3) > twoDevsAbove Then
                    isRuleTwoPass = False
                    dayIncrementer = _daysInTargetPeriod
                End If
            Next
        End If

        If My.Settings.inControl_useRule3 Then
            For dayIncrementer = 0 To _daysInTargetPeriod - 1
                failDays = 0
                For testDays = 0 To 4
                    If DailyInvMTDF(lastDay - dayIncrementer - testDays) > oneDevAbove Then failDays = failDays + 1
                Next
                If failDays > 3 Then
                    isRuleThreePass = False
                    dayIncrementer = _daysInTargetPeriod
                End If
            Next
        End If

        If My.Settings.inControl_useRule4 Then
            For dayIncrementer = 0 To _daysInTargetPeriod - 1
                failDays = 0
                For testDays = 0 To 4
                    If DailyInvMTDF(lastDay - dayIncrementer - testDays) < DailyInvMTDF(lastDay - dayIncrementer - testDays - 1) Then failDays = failDays + 1
                Next
                If failDays = 5 Then
                    isRuleFourPass = False
                    dayIncrementer = _daysInTargetPeriod
                End If
            Next
        End If

        If My.Settings.inControl_useRule5 Then
            For dayIncrementer = 0 To _daysInTargetPeriod - 1
                failDays = 0
                For testDays = 0 To 8
                    If DailyInvMTDF(lastDay - dayIncrementer - testDays) > AdjMu_InvMTDF Then failDays = failDays + 1
                Next
                If failDays = 9 Then
                    isRuleFivePass = False
                    dayIncrementer = _daysInTargetPeriod
                End If
            Next
        End If

        If My.Settings.inControl_useRule6 Then
            For dayIncrementer = 0 To _daysInTargetPeriod - 1
                failDays = 0
                For testDays = 0 To 7
                    If DailyInvMTDF(lastDay - dayIncrementer - testDays) > oneDevAbove Or DailyInvMTDF(lastDay - dayIncrementer - testDays) < oneDevBelow Then
                        failDays = failDays + 1
                    End If
                Next
                If failDays = 8 Then
                    isRuleSixPass = False
                    dayIncrementer = _daysInTargetPeriod
                End If
            Next
        End If

        If Not isRuleOnePass Then
            tmpScore = tmpScore + 1
        End If
        If Not isRuleTwoPass Then
            tmpScore = tmpScore + 1
        End If
        If Not isRuleThreePass Then
            tmpScore = tmpScore + 1
        End If
        If Not isRuleFourPass Then
            tmpScore = tmpScore + 1
        End If
        If Not isRuleFivePass Then
            tmpScore = tmpScore + 1
        End If
        If Not isRuleSixPass Then
            tmpScore = tmpScore + 1
        End If
        Return tmpScore
    End Function

    Public ReadOnly Property ChronicSporadicRanking
        Get
            Select Case AdjDistFromMean_InvMTDF
                Case Is < 1
                    If AdjDistFromMean_InvMTDF < 0.3 Then
                        Return 1 - AdjDistFromMean_InvMTDF '1
                    ElseIf AdjDistFromMean_InvMTDF < 0.6 Then
                        Return 1.6 + AdjDistFromMean_InvMTDF < 0.6  '2
                    Else
                        Return 2.2 + AdjDistFromMean_InvMTDF '3
                    End If
                Case Is < 1.8
                    If AdjDistFromMean_InvMTDF < 1.5 Then
                        Return 2.8 + AdjDistFromMean_InvMTDF '4
                    Else
                        Return 3.4 + AdjDistFromMean_InvMTDF '5 
                    End If
                Case Is < 3
                    If AdjDistFromMean_InvMTDF < 2.2 Then
                        Return 3.9 + AdjDistFromMean_InvMTDF
                    ElseIf AdjDistFromMean_InvMTDF < 2.7 Then
                        Return 4.6 + AdjDistFromMean_InvMTDF '7
                    Else
                        Return 5.2 + AdjDistFromMean_InvMTDF '8
                    End If
                Case Is < 4 '3 to 3.9
                    Return 5.5 + AdjDistFromMean_InvMTDF '9
                Case Else
                    Return 10 '10
            End Select

        End Get
    End Property
    Public ReadOnly Property Name As String
        Get
            Return _targetEvent
        End Get
    End Property
End Class