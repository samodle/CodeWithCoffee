Public Class DowntimeDataset
#Region "Variables & Properties"
    Friend rawConstraintData As New List(Of DowntimeEvent)

    Friend UnplannedData As New List(Of DowntimeEvent)
    Friend PlannedData As New List(Of DowntimeEvent)
    Friend CurtailmentData As New List(Of DowntimeEvent)

    Friend CILdata As New List(Of DowntimeEvent)
    Friend COdata As New List(Of DowntimeEvent)

    Friend SplitData As New List(Of DowntimeEvent)

    Private parentLine As ProdLine

    'properties
    Public ReadOnly Property StartDate As Date
        Get
            Return rawConstraintData(0).startTime_UT
        End Get
    End Property
    Public ReadOnly Property EndDate As Date
        Get
            Return rawConstraintData(rawConstraintData.Count - 1).endTime
        End Get
    End Property
    Public ReadOnly Property Stops As Long
        Get
            Return UnplannedData.Count ' - SplitData.Count
        End Get
    End Property
    Public ReadOnly Property numChangeovers As Long
        Get
            Return COdata.Count
        End Get
    End Property
    Public ReadOnly Property numCILs As Long
        Get
            Return CILdata.Count
        End Get
    End Property
#End Region

    Public Sub reMapData(MappingFieldA As Integer, Optional MappingFieldB As Integer = -1)
        Dim i As Integer
        If MappingFieldB = -1 Then
            For i = 0 To rawConstraintData.Count - 1
                rawConstraintData(i).MappedField = rawConstraintData(i).getFieldFromInteger(MappingFieldA)
            Next
            For i = 0 To UnplannedData.Count - 1
                UnplannedData(i).MappedField = UnplannedData(i).getFieldFromInteger(MappingFieldA)
            Next 'lg code
            For i = 0 To PlannedData.Count - 1
                PlannedData(i).MappedField = PlannedData(i).getFieldFromInteger(MappingFieldA)
            Next
        Else 'Dual Mapping!!!
            For i = 0 To rawConstraintData.Count - 1
                rawConstraintData(i).MappedField = rawConstraintData(i).getFieldFromInteger(MappingFieldA) & "-" & rawConstraintData(i).getFieldFromInteger(MappingFieldB)
            Next
            For i = 0 To UnplannedData.Count - 1
                UnplannedData(i).MappedField = UnplannedData(i).getFieldFromInteger(MappingFieldA) & "-" & UnplannedData(i).getFieldFromInteger(MappingFieldB)
            Next
            For i = 0 To PlannedData.Count - 1
                PlannedData(i).MappedField = PlannedData(i).getFieldFromInteger(MappingFieldA) & "-" & PlannedData(i).getFieldFromInteger(MappingFieldB)
            Next
        End If
    End Sub

#Region "Filtering Data Sets"
    Public Sub reFilterData_ClearAllFilters()
        Dim i As Integer
        initializeDataForFiltering()
        For i = 0 To rawConstraintData.Count - 1
            rawConstraintData(i).isFiltered = False
        Next
        For i = 0 To UnplannedData.Count - 1
            UnplannedData(i).isFiltered = False
        Next
        For i = 0 To PlannedData.Count - 1
            PlannedData(i).isFiltered = False
        Next
        finalizeFiltering()
    End Sub

    Public Sub reFilterData_Team(inclusionList As List(Of String))
        initializeDataForFiltering()
        For i As Integer = 0 To rawConstraintData.Count - 1
            If inclusionList.IndexOf(rawConstraintData(i).Team) = -1 Then rawConstraintData(i).isFiltered = True
        Next
        finalizeFiltering()
    End Sub
    Public Sub reFilterData_Format(inclusionList As List(Of String))
        initializeDataForFiltering()
        For i As Integer = 0 To rawConstraintData.Count - 1
            If inclusionList.IndexOf(rawConstraintData(i).Format) = -1 Then rawConstraintData(i).isFiltered = True
        Next
        finalizeFiltering()
    End Sub
    Public Sub reFilterData_ProductGroup(inclusionList As List(Of String))
        initializeDataForFiltering()
        For i As Integer = 0 To rawConstraintData.Count - 1
            If inclusionList.IndexOf(rawConstraintData(i).ProductGroup) = -1 Then rawConstraintData(i).isFiltered = True
        Next
        finalizeFiltering()
    End Sub
    Public Sub reFilterData_SKU(inclusionList As List(Of String))
        initializeDataForFiltering()
        For i As Integer = 0 To rawConstraintData.Count - 1
            If inclusionList.IndexOf(rawConstraintData(i).Product) = -1 Then rawConstraintData(i).isFiltered = True
        Next
        finalizeFiltering()
    End Sub
    Public Sub reFilterData_Shape(inclusionList As List(Of String))
        initializeDataForFiltering()
        For i As Integer = 0 To rawConstraintData.Count - 1
            If inclusionList.IndexOf(rawConstraintData(i).Shape) = -1 Then rawConstraintData(i).isFiltered = True
        Next
        finalizeFiltering()
    End Sub
    Private Sub initializeDataForFiltering()
        UnplannedData.Clear()
        PlannedData.Clear()
        COdata.Clear()
        CILdata.Clear()
        SplitData.Clear()
    End Sub
    Private Sub finalizeFiltering()
        For eventIncrementer As Integer = 0 To rawConstraintData.Count - 1
            If Not rawConstraintData(eventIncrementer).isExcluded Then
                If rawConstraintData(eventIncrementer).isUnplanned Then
                    UnplannedData.Add(rawConstraintData(eventIncrementer))
                    If rawConstraintData(eventIncrementer).isSplitEvent Then SplitData.Add(rawConstraintData(eventIncrementer))
                ElseIf rawConstraintData(eventIncrementer).isPlanned Then
                    PlannedData.Add(rawConstraintData(eventIncrementer))
                    If rawConstraintData(eventIncrementer).isCIL Then
                        CILdata.Add(rawConstraintData(eventIncrementer))
                    ElseIf rawConstraintData(eventIncrementer).isChangeover Then
                        COdata.Add(rawConstraintData(eventIncrementer))
                    End If
                End If
            End If
        Next
    End Sub
#End Region

#Region "Constructors"
    Public Sub New(pLine As ProdLine)
        parentLine = pLine
    End Sub
    Public Sub New(pline As ProdLine, singleEvent As DowntimeEvent)
        Me.New(pline)
        rawConstraintData.Add(singleEvent)
        If singleEvent.IsCurtailment Then
            If singleEvent.isPlanned Then
                PlannedData.Add(singleEvent)
                If singleEvent.isCIL Then
                    CILdata.Add(singleEvent)
                ElseIf singleEvent.isChangeover Then
                    COdata.Add(singleEvent)
                End If
            ElseIf singleEvent.isUnplanned Then
                UnplannedData.Add(singleEvent)
                If singleEvent.isSplitEvent Then SplitData.Add(singleEvent)
            End If
        Else
            CurtailmentData.Add(singleEvent)
        End If
    End Sub

    Public Sub New(pLine As ProdLine, rawData As Array)
        Me.New(pLine)
        'Dim netDowntimeEvents As Long
        Dim rowIncrementer As Integer, tmpEvent As DowntimeEvent

        rowIncrementer = 0
        tmpEvent = New DowntimeEvent(parentLine, rowIncrementer)
        rawConstraintData.Add(tmpEvent)

        If tmpEvent.IsCurtailment Then
            CurtailmentData.Add(tmpEvent)
        Else
            If tmpEvent.isUnplanned Then
                UnplannedData.Add(tmpEvent)
                If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
            ElseIf tmpEvent.isPlanned Then
                PlannedData.Add(tmpEvent)
                If tmpEvent.isCIL Then
                    CILdata.Add(tmpEvent)
                ElseIf tmpEvent.isChangeover Then
                    COdata.Add(tmpEvent)
                End If
            End If
        End If

        'all the rest
        For rowIncrementer = 1 To rawData.GetLength(1) - 1

            '  If rowIncrementer < rawConstraintData.Count Then 'maple

            tmpEvent = New DowntimeEvent(parentLine, rowIncrementer)


            If parentLine.SQLdowntimeProcedure <> DefaultProficyDowntimeProcedure.Maple Then

                'DO THIS IF ITS NOT MAPLE
                If tmpEvent.UT = rawConstraintData(rowIncrementer - 1).UT Then 'ERROR CHECKING!
                    With rawConstraintData(rowIncrementer - 1)
                        If DateDiff(DateInterval.Second, .endTime, DateAdd(DateInterval.Second, -tmpEvent.UT * 60, tmpEvent.startTime)) > 5 Then

                            'fixing for GLEDS 0 Downtime issue
                            If parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.GLEDS Then
                                rawConstraintData.Add(tmpEvent)
                            Else
                                .endTime = tmpEvent.endTime
                                .DT = DateDiff(DateInterval.Second, .startTime, .endTime) / 60

                                rawConstraintData.Add(tmpEvent) 'added 9/22, w/o this line if there is this problem in the second record the application will crash
                            End If
                        Else
                            'naw its ok
                            rawConstraintData.Add(tmpEvent)

                            If tmpEvent.IsCurtailment Then
                                CurtailmentData.Add(tmpEvent)
                            Else
                                If tmpEvent.isUnplanned Then
                                    UnplannedData.Add(tmpEvent)
                                    If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
                                ElseIf tmpEvent.isPlanned Then
                                    PlannedData.Add(tmpEvent)
                                    If tmpEvent.isCIL Then
                                        CILdata.Add(tmpEvent)
                                    ElseIf tmpEvent.isChangeover Then
                                        COdata.Add(tmpEvent)
                                    End If
                                End If
                            End If

                        End If
                    End With
                Else
                    rawConstraintData.Add(tmpEvent)

                    If tmpEvent.IsCurtailment Then
                        CurtailmentData.Add(tmpEvent)
                    Else
                        If tmpEvent.isUnplanned Then
                            UnplannedData.Add(tmpEvent)
                            If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
                        ElseIf tmpEvent.isPlanned Then
                            PlannedData.Add(tmpEvent)
                            If tmpEvent.isCIL Then
                                CILdata.Add(tmpEvent)
                            ElseIf tmpEvent.isChangeover Then
                                COdata.Add(tmpEvent)
                            End If
                        End If
                    End If

                End If
                '   End If

                'DO THIS IF IT IS MAPLE
            Else
                If tmpEvent.UT = rawConstraintData(rowIncrementer - 1).UT Then 'ERROR CHECKING!
                    With rawConstraintData(rowIncrementer - 1)
                        If DateDiff(DateInterval.Second, .endTime, DateAdd(DateInterval.Second, -tmpEvent.UT * 60, tmpEvent.startTime)) > 5 Then

                            'fixing for GLEDS 0 Downtime issue
                            If parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.Maple Then
                                rawConstraintData.Add(tmpEvent)
                            Else
                                .endTime = tmpEvent.endTime
                                .DT = DateDiff(DateInterval.Second, .startTime, .endTime) / 60

                            End If
                        Else
                            'naw its ok
                            rawConstraintData.Add(tmpEvent)

                            If tmpEvent.isUnplanned Then
                                UnplannedData.Add(tmpEvent)
                                If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
                            ElseIf tmpEvent.isPlanned Then
                                PlannedData.Add(tmpEvent)
                                If tmpEvent.isCIL Then
                                    CILdata.Add(tmpEvent)
                                ElseIf tmpEvent.isChangeover Then
                                    COdata.Add(tmpEvent)
                                End If
                            End If

                        End If
                    End With
                Else
                    rawConstraintData.Add(tmpEvent)

                    If tmpEvent.isUnplanned Then
                        UnplannedData.Add(tmpEvent)
                        If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
                    ElseIf tmpEvent.isPlanned Then
                        PlannedData.Add(tmpEvent)
                        If tmpEvent.isCIL Then
                            CILdata.Add(tmpEvent)
                        ElseIf tmpEvent.isChangeover Then
                            COdata.Add(tmpEvent)
                        End If
                    End If

                End If

            End If
            'THIS IS THE END OF THOSE IF MAPLE ETC
        Next
        'End If

        'CLEANING 
        Dim doWeReDoAnalysis As Boolean = False

        If rawConstraintData.Count > 5 Then
            For rowIncrementer = 2 To rawConstraintData.Count - 2
                If rowIncrementer < rawConstraintData.Count - 2 Then
                    If rawConstraintData(rowIncrementer).startTime = rawConstraintData(rowIncrementer - 1).startTime Then
                        rawConstraintData.RemoveAt(rowIncrementer - 1)
                        doWeReDoAnalysis = True
                    End If
                End If
            Next rowIncrementer
        End If

        If doWeReDoAnalysis Then
            initializeDataForFiltering()
            finalizeFiltering()
        End If


        'combine all the planned events
        If My.Settings.PDT_maxMinutesBetweenEvents > 0 Then CombinePlannedEvents()

    End Sub

    Public Sub New(pLine As ProdLine, rawD As List(Of DowntimeEvent))
        parentLine = pLine
        rawConstraintData = rawD
        initializeDataForFiltering()
        finalizeFiltering()
    End Sub

#End Region

    Public Function getSubset(startDate As Date, endDate As Date, Optional MappingFieldA As Integer = -1, Optional MappingFieldB As Integer = -1) As DowntimeDataset
        Dim listIndexA As Integer, listIndexB As Integer, tmpEvent As DowntimeEvent, tmpData As DowntimeDataset, eventIncrementer As Integer
        listIndexA = rawConstraintData.IndexOf(New DowntimeEvent(startDate))
        If listIndexA = -1 Then
            If startDate < rawConstraintData(0).endTime Then
                listIndexA = 0
            End If
        End If
        listIndexB = rawConstraintData.IndexOf(New DowntimeEvent(endDate))
        If listIndexB = -1 Then
            With rawConstraintData(rawConstraintData.Count - 1)
                If endDate.Equals(.endTime) Then
                    listIndexB = rawConstraintData.Count - 1
                ElseIf endDate > .endTime Then

                    If Not rawConstraintData(rawConstraintData.Count - 1).isExcluded Then

                        rawConstraintData.Add(New DowntimeEvent(DateDiff(DateInterval.Second, .endTime, endDate) / 60, endDate, endDate, False))
                        listIndexB = rawConstraintData.Count - 1

                        'maybe this event will help with our listIndexA...
                        If listIndexA = -1 Then
                            If startDate >= rawConstraintData(rawConstraintData.Count - 1).startTime_UT Then listIndexA = rawConstraintData.Count - 1
                        End If

                    Else
                        listIndexB = rawConstraintData.Count - 1
                        If listIndexA = -1 Then
                            listIndexA = rawConstraintData.Count - 1
                        End If
                    End If

                End If
            End With
            'check if both are before
            If listIndexA = 0 And listIndexB = -1 And endDate < rawConstraintData(0).startTime_UT Then
                tmpEvent = New DowntimeEvent(0, endDate, endDate, True)
                tmpEvent.adjustMyStartTime(startDate)
                tmpEvent.adjustMyEndTime(endDate)
                Return New DowntimeDataset(parentLine, tmpEvent)

            End If
        End If


ReAnalyzeData:

        'one line only
        If listIndexA = listIndexB And listIndexA > -1 Then
            tmpEvent = New DowntimeEvent(parentLine, rawConstraintData(listIndexA).rowNum)
            tmpEvent.adjustMyStartTime(startDate)
            tmpEvent.adjustMyEndTime(endDate)
            Return New DowntimeDataset(parentLine, tmpEvent)

            'multiple lines of data
        ElseIf listIndexA < listIndexB And listIndexA > -1 Then
            tmpData = New DowntimeDataset(parentLine)
            tmpEvent = New DowntimeEvent(parentLine, rawConstraintData(listIndexA).rowNum)
            tmpEvent.adjustMyStartTime(startDate)

            If tmpEvent.endTime < startDate Then
                tmpEvent.DT = 0 ' = true              
            End If

            tmpData.rawConstraintData.Add(tmpEvent)
            If Not tmpEvent.isExcluded Then
                If tmpEvent.isUnplanned Then
                    tmpData.UnplannedData.Add(tmpEvent)
                    If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
                ElseIf tmpEvent.isPlanned Then
                    tmpData.PlannedData.Add(tmpEvent)
                    If tmpEvent.isCIL Then
                        tmpData.CILdata.Add(tmpEvent)
                    ElseIf tmpEvent.isChangeover Then
                        tmpData.COdata.Add(tmpEvent)
                    End If
                End If
            End If
            For eventIncrementer = listIndexA + 1 To listIndexB - 1
                tmpEvent = rawConstraintData(eventIncrementer)

                tmpData.rawConstraintData.Add(tmpEvent)
                If Not tmpEvent.isExcluded Then
                    If tmpEvent.isUnplanned Then
                        tmpData.UnplannedData.Add(tmpEvent)
                        If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
                    ElseIf tmpEvent.isPlanned Then
                        tmpData.PlannedData.Add(tmpEvent)
                        If tmpEvent.isCIL Then
                            tmpData.CILdata.Add(tmpEvent)
                        ElseIf tmpEvent.isChangeover Then
                            tmpData.COdata.Add(tmpEvent)
                        End If
                    End If
                End If
            Next

            'check if the last event is uptime only and act accordingly!!!

            If rawConstraintData(listIndexB).DT > 0 Then
                tmpEvent = New DowntimeEvent(parentLine, rawConstraintData(listIndexB).rowNum)
                tmpEvent.adjustMyEndTime(endDate)
                tmpData.rawConstraintData.Add(tmpEvent)
                If Not rawConstraintData(listIndexB).isExcluded Then
                    If tmpEvent.isUnplanned Then
                        tmpData.UnplannedData.Add(tmpEvent)
                        If tmpEvent.isSplitEvent Then SplitData.Add(tmpEvent)
                    ElseIf tmpEvent.isPlanned Then
                        tmpData.PlannedData.Add(tmpEvent)
                    End If
                End If
            Else 'uptime only event
                tmpData.rawConstraintData.Add(rawConstraintData(listIndexB))
            End If

            'set it!
            If My.Settings.PDT_maxMinutesBetweenEvents > 0 Then tmpData.CombinePlannedEvents()
            Return tmpData

            'error!
        ElseIf listIndexA = -1 Or listIndexB = -1 Then
            If parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.GLEDS Or parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.Maple Then
                If listIndexA = -1 And listIndexB = -1 Then
                    listIndexA = 0
                    listIndexB = rawConstraintData.Count - 1
                ElseIf listIndexB = -1 Then
                    listIndexB = listIndexA
                    While rawConstraintData(listIndexB).endTime < endDate And listIndexB < rawConstraintData.Count
                        listIndexB += 1
                    End While
                Else
                    listIndexA = listIndexB
                    While rawConstraintData(listIndexA).startTime > startDate And listIndexA > 0
                        listIndexA -= 1
                    End While
                End If
                GoTo ReAnalyzeData

            Else
                '  Messagebox.Show("Date Range Exception/CLS_DataInterface: Trend data incomplete.")
                Throw New dateRangeException("")
            End If
        Else
            Throw New dateRangeException("")
        End If

    End Function

    Public Sub CombinePlannedEvents()
        Dim eventIncrementer As Integer = 0
        While eventIncrementer < PlannedData.Count - 1
            With PlannedData(eventIncrementer)
                If .isPlanned Then

                End If
            End With
            eventIncrementer += 1
        End While
    End Sub
End Class

Public Class ProductionDataset
#Region "Variables & Properties"
    Friend rawProductionData As New List(Of ProductionEvent)

    Private Property _parentLine As ProdLine

    Public ReadOnly Property startTime As Date
        Get
            Return rawProductionData(0).startTime
        End Get
    End Property
    Public ReadOnly Property endTime As Date
        Get
            Return rawProductionData(rawProductionData.Count - 1).endTime
        End Get
    End Property
#End Region

#Region "Construction"
    Public Sub New(pLine As ProdLine)
        _parentLine = pLine
    End Sub
    Public Sub New(pLine As ProdLine, analyzeData As Boolean)
        Me.New(pLine)
        Dim rowIncrementer As Integer

        For rowIncrementer = 0 To pLine.rawProficyProductionData.GetLength(1) - 1
            rawProductionData.Add(New ProductionEvent(_parentLine, rowIncrementer))
        Next
    End Sub

#End Region

    Public Function getSubset(ByRef startDate As Date, ByRef endDate As Date) As ProductionDataset
        Dim tmpData As New ProductionDataset(_parentLine)
        Dim startIndex As Integer, endIndex As Integer, rowIncrementer As Integer

        'find the indices
        If startDate > rawProductionData(rawProductionData.Count - 1).endTime Or endDate < rawProductionData(0).startTime Then Throw New columnStubbingException

        startIndex = rawProductionData.IndexOf(New ProductionEvent(startDate))
        If startIndex = -1 Then
            If startDate < rawProductionData(0).startTime Then
                startIndex = 0
                startDate = rawProductionData(0).startTime
            Else
                Debugger.Break()
            End If
        End If

        If startDate = rawProductionData(startIndex).endTime Then startIndex += 1

        endIndex = rawProductionData.IndexOf(New ProductionEvent(endDate))

        If endIndex = -1 Then
            If endDate > rawProductionData(rawProductionData.Count - 1).endTime Then
                endDate = rawProductionData(rawProductionData.Count - 1).endTime
                endIndex = rawProductionData.Count - 1
            Else
                Debugger.Break()
            End If
        End If

        If endDate < rawProductionData(endIndex).endTime Then
            If startIndex <> endIndex Then endIndex -= 1
        End If


        If startIndex = endIndex Then
            If rawProductionData(startIndex).startTime = startDate And rawProductionData(startIndex).endTime = endDate Then
                tmpData.rawProductionData.Add(rawProductionData(startIndex))
            Else
                Throw New columnStubbingException
            End If
            Return tmpData
        Else
            For rowIncrementer = startIndex To endIndex
                tmpData.rawProductionData.Add(rawProductionData(rowIncrementer))
            Next rowIncrementer

            If tmpData.rawProductionData.Count > 0 Then
                startDate = tmpData.rawProductionData(0).startTime
                endDate = tmpData.rawProductionData(tmpData.rawProductionData.Count - 1).endTime
            End If
            Return tmpData
        End If

    End Function

End Class


Public Class DowntimeEvent
    Implements IComparable(Of DowntimeEvent)
    Implements IEquatable(Of DowntimeEvent)

    Public Overrides Function toString() As String
        Dim x As String = ""
        If isSplitEvent Then x = "SPLIT   "
        Return x & "S/E: " & _startTime & "/" & _endTime & "  DT/UT:  " & _DT & "/" & _UT & "   " & _DTGroup
    End Function

#Region "Raw Data Mapping Constants"
    'proficy downtime setup
    Private Enum DownTimeColumn 'Downtime Explorer
        StartTime = 0
        Endtime = 1
        DT = 2
        UT = 3
        MasterProdUnit = 4
        Location = 5
        Fault = 6
        Reason1 = 7
        Reason2 = 8
        Reason3 = 9
        Reason4 = 10
        StopClass = 11 ' LG Code
        PR_InOut = 12
        Team = 13
        Shift = 14
        PlannedUnplanned = 15
        DTGroup = 16
        Product = 19
        ProductCode = 20
        ProductGroup = 21
        Comment = 22
        Max = 29
    End Enum

    Private Enum DownTimeColumn_OneClick
        StartTime = 0
        EndTime = 1
        DT = 2
        UT = 3
        'MasterProdUnit = 4
        Location = 4
        Fault = 5
        Reason1 = 6
        Reason2 = 7
        Reason3 = 8
        Reason4 = 9
        '  SpliEvent = 10  'LG Code
        'StopClass = 12
        Team = 11 '13
        Shift = 12 '14
        Product = 13 '19
        ProductCode = 14 '20
        PlannedUnplanned = 16 '15
        DTGroup = 17 '16
        PR_InOut = 19 '12
        Comment = 20 '22
        Max = 20
    End Enum
    Private Enum DownTimeColumn_GLEDS
        StartTime = 0
        Endtime = 1
        DT = 2
        UT = 3

        Location = 4
        Fault = 5
        Reason1 = 6
        Reason2 = 7
        Reason3 = 8
        Reason4 = 9

        Team = 11 '13
        Shift = 12 '14
        ' Product = 13 '19

        DTGroup = 15 '16
        PlannedUnplanned = 16 '15
        PR_InOut = 17 '12
        Product = 18 '20
        Comment = 20 '22
        Max = 20
    End Enum
    Private Enum DownTimeColumn_OneClickALT
        StartTime = 0
        Endtime = 1
        DT = 2
        UT = 3
        Location = 4
        Fault = 5
        Reason1 = 6
        Reason2 = 7
        Reason3 = 8
        Reason4 = 9
        SplitEvent = 10
        Team = 11 '13
        Shift = 12 '14
        Product = 13 '19
        ProductCode = 14 '20
        Category = 15
        Schedule = 16 '15
        Subsystem = 17 '16
        PR_InOut = 19 '12
        Comment = 20 '22
        Max = 20
    End Enum
    Private Enum DownTimeColumn_RE_CentralServer_Raw
        StartTime = 0
        Endtime = 1
        UT = 2
        DT = 3
        Mpu = 4
        Location = 5
        Fault = 7
        Reason1 = 8
        Reason2 = 9
        Reason3 = 10
        Reason4 = 11

        SplitEvent = 6
        Team = 12 '13
        Shift = 13 '14
        Product = 20 '19
        ProductCode = 18 '20
        Category = 15
        Schedule = 16 '15
        Subsystem = 17 '16
        PR_InOut = 21 '12
        Comment = 22 '22
        Max = 22

        '0  StartTime,
        '1      EndTime,
        '2    UpTime,
        '3  DownTime,
        '4 MasterProdUnit,
        '5Location,
        '6Split,
        '7Fault,
        '8Reason1,
        '9Reason2,
        '10Reason3,
        '11Reason4,
        '12Team,
        '13 Shift,

        '14                   Cat1,
        '15                 Cat2,
        '16               Cat3,
        '17             Cat4,
        '18          Product,
        '19       ProductGroup,
        '20     Brand,
        '21 LineStatus,
        '22  Comments


    End Enum


    ' Public Enum DowntimeField
    '     startTime = 0
    '     endTime = 1
    '     DT = 2
    '     UT = 3
    '     MasterProdUnit = 4
    '     Location = 5
    '     Fault = 6
    '     Reason1 = 7
    '     Reason2 = 8
    '     Reason3 = 9
    '     Reason4 = 10
    '     PR_inout = 11
    '     Team = 12
    '     PlannedUnplanned = 13
    '     DTGroup = 14
    '     Product = 15
    '     ProductCode = 16
    '     Comment = 17
    ' 'MAPPED DATA
    '     Tier1 = 18
    '     Tier2 = 19
    '     Tier3 = 20
    '     Format = 21
    '     Shape = 22
    '     Classification = 23
    '     OneClick = 24
    ' End Enum



#End Region

    Public Function toString_CSV() As String
        Dim x As String = ""
        Try
            x = "," & _startTime.ToString().Replace(",", ";") & "," & _endTime.ToString().Replace(",", ";") & "," & _DT.ToString().Replace(",", ";") & "," & _UT.ToString().Replace(",", ";") & "," & _Reason1.ToString().Replace(",", ";") & "," & _Reason2.ToString().Replace(",", ";") & "," & _Reason3.ToString().Replace(",", ";") & "," & _Reason4.ToString().Replace(",", ";") & "," & _Tier1 & "," & _Tier2.ToString().Replace(",", ";") & "," & _Tier3.ToString().Replace(",", ";") & "," & _Fault.ToString().Replace(",", ";") & "," & _PR_inout.ToString().Replace(",", ";") & "," & _PlannedUnplanned.ToString().Replace(",", ";") & "," & _ProductCode.ToString().Replace(",", ";") & "," & _DTGroup.ToString().Replace(",", ";")
        Catch e As Exception
            x = ", data err"
        End Try

        Return x
    End Function

    Public Function getFieldFromInteger(TargetField As Integer) As String
        Select Case TargetField
            Case DowntimeField.Location
                Return _Location
            Case DowntimeField.Fault
                Return _Fault
            Case DowntimeField.Reason1
                Return _Reason1
            Case DowntimeField.Reason2
                Return _Reason2
            Case DowntimeField.Reason3
                Return _Reason3
            Case DowntimeField.Reason4
                Return _Reason4
                ' Case DowntimeField.PR_inout
                ' Case DowntimeField.Team
                ' Case DowntimeField.PlannedUnplanned
            Case DowntimeField.Stopclass  ' 
                Return _StopClass
            Case DowntimeField.DTGroup
                Return _DTGroup
                ' Case DowntimeField.Product
            Case DowntimeField.ProductCode
                Return _ProductCode
                ' Case DowntimeField.Comment 'NOPE
                ' 'MAPPED DATA
            Case DowntimeField.Tier1
                Return _Tier1
            Case DowntimeField.Tier2
                Return _Tier2
            Case DowntimeField.Tier3
                Return _Tier3
            Case DowntimeField.Format
                Return _Format
            Case DowntimeField.Shape
                Return _Shape

            Case DowntimeField.ProductGroup
                Return _ProductGroup
            Case Else

                Return _Tier1 'SRO 6/20/17. added this as default, so tool will still be usable in event of strange mapping
                'Throw New unknownMappingException
        End Select
    End Function

#Region "Custom Sorting Params"
    Private _isStandardSort = True
    Private _sortField As Integer
    Public Sub setSortParam(dtField As Integer)
        If dtField = DowntimeField.startTime Then
            _isStandardSort = True
        Else
            _sortField = dtField
            _isStandardSort = False
        End If
    End Sub
#End Region

#Region "Variables & Properties"
    'Private rawData As Array
    Private Property parentLine As ProdLine
    Friend Property rowNum As Long

    'Booleans
    Private _isExcluded As Boolean
    Private _isUnplanned As Boolean
    Private _isPlanned As Boolean
    Private _isChangeover As Boolean = False
    Private _isCIL As Boolean = False
    Private _isFiltered As Boolean = False
    Public Property IsCurtailment As Boolean = False
    Public ReadOnly Property isUnplanned As Boolean
        Get
            Return _isUnplanned
        End Get
    End Property
    Public ReadOnly Property isExcluded As Boolean
        Get
            Return (_isExcluded Or _isFiltered)
        End Get
    End Property
    Public ReadOnly Property isExcluded_NoFilter As Boolean
        Get
            Return _isExcluded
        End Get
    End Property
    Public ReadOnly Property isPlanned As Boolean
        Get
            Return _isPlanned
        End Get
    End Property
    Public ReadOnly Property isChangeover As Boolean
        Get
            Return _isChangeover
        End Get
    End Property
    Public ReadOnly Property isCIL As Boolean
        Get
            Return _isCIL
        End Get
    End Property
    Public WriteOnly Property isFiltered As Boolean
        Set(value As Boolean)
            _isFiltered = value
        End Set
    End Property

    'RAW DATA
    Private _startTime As Date
    Private _endTime As Date
    Private _DT As Double
    Private _UT As Double
    Private _MasterProdUnit As String
    Private _Location As String
    Private _Fault As String
    Private _Reason1 As String
    Private _Reason2 As String
    Private _Reason3 As String
    Private _Reason4 As String
    Private _StopClass As String
    Private _PR_inout As String
    Private _Team As String
    Private _PlannedUnplanned As String
    Private _DTGroup As String
    Private _Product As String
    Private _ProductCode As String
    Private _ProductGroup As String
    Private _Comment As String
    'MAPPED DATA
    Private _Tier1 As String = BLANK_INDICATOR
    Private _Tier2 As String = BLANK_INDICATOR
    Private _Tier3 As String = BLANK_INDICATOR
    Private _Format As String = ""
    Private _Shape As String = ""
    Private _Classification As String = ""

    Private _Mapping As String ' = BLANK_INDICATOR

    ''''''THESE ARE KNOWN RAW DATA FIELDS''''''
    'date / time fields
    Public ReadOnly Property startTime_UT As Date
        Get
            Return DateAdd(DateInterval.Second, -60 * _UT, _startTime)
        End Get
    End Property

    Public Property startTime As Date
        Get
            Return _startTime
        End Get
        Set(value As Date)
            _startTime = value
        End Set
    End Property
    Public Property endTime As Date
        Get
            Return _endTime
        End Get
        Set(value As Date)
            _endTime = value
        End Set
    End Property

    Public ReadOnly Property startTime_24hr As String
        Get
            Return _startTime.ToString("MM/dd/yyyy HH:mm:ss")
        End Get
    End Property
    Public ReadOnly Property endTime_24hr As String
        Get
            Return _endTime.ToString("MM/dd/yyyy HH:mm:ss")
        End Get
    End Property


    Public Property DT As Double
        Get
            Return _DT
        End Get
        Set(value As Double)
            _DT = value
        End Set
    End Property
    Public ReadOnly Property DT_display As Double
        Get
            Return Math.Round(_DT, 2)
        End Get
    End Property
    Public ReadOnly Property UT_display As Double
        Get
            Return Math.Round(_UT, 2)
        End Get
    End Property
    Public ReadOnly Property UT As Double
        Get
            Return _UT
        End Get
    End Property

    'top level leds fields
    Public ReadOnly Property Location As String
        Get
            Return _Location
        End Get
    End Property
    Public ReadOnly Property Fault As String
        Get
            Return _Fault
        End Get
    End Property
    Public Property DTGroup As String
        Get
            Return _DTGroup
        End Get
        Set(value As String)
            _DTGroup = value
        End Set
    End Property
    'tree level fields
    Public ReadOnly Property Reason1 As String
        Get
            Return _Reason1
        End Get
    End Property
    Public ReadOnly Property Reason2 As String
        Get
            Return _Reason2
        End Get
    End Property
    Public ReadOnly Property Reason3 As String
        Get
            Return _Reason3
        End Get
    End Property
    Public ReadOnly Property Reason4 As String
        Get
            Return _Reason4
        End Get
    End Property
    Public ReadOnly Property StopClass As String
        Get
            Return _StopClass
        End Get
    End Property
    'comment
    Public ReadOnly Property Comment As String
        Get
            Return _Comment
        End Get
    End Property

    Public ReadOnly Property PlannedUnplanned As String
        Get
            Return _PlannedUnplanned
        End Get
    End Property
    Public Property Team As String
        Get
            Return _Team
        End Get
        Set(value As String)
            _Team = value
        End Set
    End Property
    Public ReadOnly Property PR_inout As String
        Get
            Return _PR_inout
        End Get
    End Property
    Public Property ProductCode As String
        Get
            Return _ProductCode
        End Get
        Set(value As String)
            _ProductCode = value
        End Set
    End Property
    Public ReadOnly Property ProductGroup As String
        Get
            Return _ProductGroup
        End Get
    End Property
    Public Property Product As String
        Get
            Return _Product
        End Get
        Set(value As String)
            _Product = value
        End Set
    End Property
    Public ReadOnly Property MasterProductionUnit As String
        Get
            Return _MasterProdUnit
        End Get
    End Property

    '''''''THESE ARE MAPPED DATA FIELDS
    Public Property Tier1 As String
        Get
            Return _Tier1
        End Get
        Set(value As String)
            _Tier1 = value
        End Set
    End Property
    Public Property Tier2 As String
        Get
            Return _Tier2
        End Get
        Set(value As String)
            _Tier2 = value
        End Set
    End Property
    Public Property Tier3 As String
        Get
            Return _Tier3
        End Get
        Set(value As String)
            _Tier3 = value
        End Set
    End Property
    Public Property Format As String
        Get
            Return _Format
        End Get
        Set(value As String)
            _Format = value
        End Set
    End Property
    Public Property Shape As String
        Get
            Return _Shape
        End Get
        Set(value As String)
            _Shape = value
        End Set
    End Property
    Public ReadOnly Property Classification As String
        Get
            Return _Classification
        End Get
    End Property


    Public Property MappedField As String
        Get
            Return _Mapping
        End Get
        Set(value As String)
            _Mapping = value
        End Set
    End Property

    Public isSplitEvent As Boolean = False
#End Region

#Region "Construction"
    Public Sub New(startDate As Date)
        _startTime = startDate
        _endTime = startDate
    End Sub

    Public Sub New(parentLineIn As ProdLine, row As Long)
        parentLine = parentLineIn
        rowNum = row

        If IsNothing(parentLine.rawProficyData) Then 'assumes this means there is sum demo data
            populateFromDowntimeEvent(parentLine.rawDowntimeData.rawConstraintData(rowNum), False)
        Else
            Select Case parentLine.SQLdowntimeProcedure
                Case DefaultProficyDowntimeProcedure.OneClick
                    If parentLineIn.SiteName = SITE_ALBANY Or parentLineIn.SiteName = SITE_OXNARD Then
                        initializeFields_OneClickALTERNATE()
                    Else
                        initializeFields_OneClick()
                    End If
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    initializeFields()
                Case DefaultProficyDowntimeProcedure.RE_CentralServer
                    initializeFields()
                Case DefaultProficyDowntimeProcedure.GLEDS
                    initializeFields_GLEDS2()
                Case DefaultProficyDowntimeProcedure.Maple
                    initializeFields_Other()
            End Select
        End If
        finalizeInitialization()
    End Sub

    'constructor for uptime only events
    Public Sub New(uptimeOnly As Double, startTime As Date, endTime As Date, isExcluded As Boolean)
        _DT = 0
        _startTime = startTime
        _endTime = endTime
        _UT = uptimeOnly

        _isExcluded = isExcluded
        _isPlanned = False
        _isUnplanned = False
        _isChangeover = False
        _isCIL = False
    End Sub

    Public Sub New(pLine As ProdLine, ByVal sDate As Date, ByVal eDate As Date, ByVal dt As Double, ByVal ut As Double)
        Dim strLevelOne() As String

        parentLine = pLine

        _startTime = sDate
        _endTime = eDate
        _DT = dt
        _UT = ut
        strLevelOne = {"Making", "Packaging", "Bat Cave"}
        _Location = strLevelOne(GetRandom(0, strLevelOne.Length))
        _MasterProdUnit = BLANK_INDICATOR
        strLevelOne = {"SERVO FAULT", "OTTERS. LOTS OF OTTERS.", "ERROR", "LINE DOWN", "no error code"}
        _Fault = strLevelOne(GetRandom(0, strLevelOne.Length)) & "  [" & GetRandom(1, 15) & "]"
        strLevelOne = {" Fault A", "Utilities", " Matl C", "Error 7D", "Filler Jam", "Casepacker", "CR77", "D1", "D2", "A3", "A2", "H123", "321P"}
        _Reason1 = strLevelOne(GetRandom(0, strLevelOne.Length))
        '   strLevelOne = {"", "", ""}
        _Reason2 = strLevelOne(GetRandom(0, strLevelOne.Length))
        ' strLevelOne = {"", "", ""}
        _Reason3 = strLevelOne(GetRandom(0, strLevelOne.Length))
        ' strLevelOne = {"", "", ""}
        _Reason4 = strLevelOne(GetRandom(0, strLevelOne.Length))
        ' strLevelOne = {"", "", ""}
        _StopClass = strLevelOne(GetRandom(0, strLevelOne.Length))
        _PR_inout = "PR:In"
        _Team = strLevelOne(GetRandom(0, strLevelOne.Length))
        strLevelOne = {"Unplanned", "Planned", "Unplanned", "Unplanned"}
        _PlannedUnplanned = strLevelOne(GetRandom(0, strLevelOne.Length))
        '    strLevelOne = {"", "", ""}
        _DTGroup = strLevelOne(GetRandom(0, strLevelOne.Length))
        '   strLevelOne = {"", "", ""}
        _Product = strLevelOne(GetRandom(0, strLevelOne.Length))
        '   strLevelOne = {"", "", ""}
        _ProductCode = strLevelOne(GetRandom(0, strLevelOne.Length))
        '  strLevelOne = {"", "", ""}
        _ProductGroup = strLevelOne(GetRandom(0, strLevelOne.Length))
        strLevelOne = {"6", "loose bird", "seriously...another bird", "do or do not, there is no try", "salt and corrosion. the infamous old enemies of the crime fighter", "abc", "words, words, words...", ".", "dance party"}
        _Comment = strLevelOne(GetRandom(0, strLevelOne.Length))

        strLevelOne = {" Fault A", " Matl C", "Error 7D", "Filler Jam", "Casepacker", "CR77", "D1", "D2", "A3", "A2", "H123", "321P"}
        _Tier1 = strLevelOne(GetRandom(0, strLevelOne.Length))
        _Tier2 = strLevelOne(GetRandom(0, strLevelOne.Length))
        _Tier3 = strLevelOne(GetRandom(0, strLevelOne.Length))

        finalizeInitialization()

    End Sub

    Public Sub populateFromDowntimeEvent(ByVal other As DowntimeEvent, Optional finalize As Boolean = True)
        _startTime = other.startTime
        _endTime = other.endTime
        _DT = other.DT
        _UT = other.UT
        _Location = other.Location
        _MasterProdUnit = other.MasterProductionUnit
        _Fault = other.Fault
        _Reason1 = other.Reason1
        _Reason2 = other.Reason2
        _Reason3 = other.Reason3
        _Reason4 = other.Reason4
        _StopClass = other.StopClass
        _PR_inout = other.PR_inout
        _Team = other.Team
        _PlannedUnplanned = other.PlannedUnplanned
        _DTGroup = other.DTGroup
        _Product = other.Product
        _ProductGroup = other.ProductGroup
        _Comment = other.Comment

        _Tier1 = other.Tier1
        _Tier2 = other.Tier2
        _Tier3 = other.Tier3

        If finalize Then
            finalizeInitialization()
        End If
    End Sub

    Private Sub finalizeInitialization()
        If isDTExcluded() Then
            _isExcluded = True
            _isUnplanned = False
            _isPlanned = False
            _isCIL = False
            _isChangeover = False
        Else
            _isExcluded = False
            If isDTUnplanned() Then
                _isUnplanned = True
                _isPlanned = False
            Else
                _isUnplanned = False
                _isPlanned = True
                If isDTChangeover() Then
                    _isChangeover = True
                ElseIf isDTCIL() Then
                    _isCIL = True
                End If
            End If
        End If

        mapMyTiers()

        If My.Settings.defaultDownTimeField_Secondary = -1 Then
            _Mapping = getFieldFromInteger(My.Settings.defaultDownTimeField)
        Else
            _Mapping = getFieldFromInteger(My.Settings.defaultDownTimeField) & "-" & getFieldFromInteger(My.Settings.defaultDownTimeField_Secondary)
        End If
    End Sub

#End Region

#Region "isExcluded / isUnplanned / isPlanned / CO, CIL, etc"

#Region "Summary"
    Private Function isDTUnplanned() As Boolean
        Select Case parentLine.Mapping_DTschedPlannedUnplanned
            Case DTsched_Mapping.Greensboro
                Return isDTUnplanned_GBO()
            Case DTsched_Mapping.SkinCare
                Return isDTUnplanned_GBO()
            Case DTsched_Mapping.Phenoix
                Return isDTUnplanned_PHX()
            Case DTsched_Mapping.SwingRoad
                Return isDTUnplanned_SwingRoad()
            Case DTsched_Mapping.HuangpuHC
                Return isDTUnplanned_HuangpuHC()
            Case DTsched_Mapping.APDO
                Return isDTUnplanned_APDO()
            Case DTsched_Mapping.Ukraine
                Return isDTUnplanned_Ukraine()
            Case DTsched_Mapping.Belleville
                Return isDTUnplanned_Belleville()
            Case DTsched_Mapping.Mariscala
                Return isDTUnplanned_Mariscala()
            Case DTsched_Mapping.BabyCare
                Return isDTUnplanned_BabyCare()
            Case DTsched_Mapping.ModPack
                isDTPlanned_ModPack()
                Return isDTUnplanned_ModPACK()
            Case DTsched_Mapping.Albany
                Return isDTUnplanned_Albany()
            Case DTsched_Mapping.Hyderabad
                Return isDTUnplanned_Hyderabad()
            Case DTsched_Mapping.Budapest
                Return isDTUnplanned_Budapest()
            Case DTsched_Mapping.Rakona
                Return isDTUnplanned_Rakona()
            Case DTsched_Mapping.SingaporePioneer
                Return isDTUnplanned_SingaporePioneer()
            Case DTsched_Mapping.Fem_LuisCustom
                Return isDTUnplanned_Fem_LuisCustom()
            Case Else
                Throw New unknownMappingException
        End Select
    End Function
    Private Function isDTExcluded() As Boolean
        If IsExcludedEventsIncluded = True Then
            If _Reason2.Contains("Curtailment") Then
                IsCurtailment = True
                Return True
            Else
                Return False
            End If
        Else

            If My.Settings.EnableTimeSpanExclusion Then
                If _startTime.Hour < My.Settings.Exclude_StartHour Or _startTime.Hour > My.Settings.Exclude_EndHour Then
                    Return True
                ElseIf _startTime.Hour = My.Settings.Exclude_StartHour Then
                    If _startTime.Minute < My.Settings.Exclude_StartMinutes Then
                        Return True
                    End If
                ElseIf _startTime.Hour = My.Settings.Exclude_EndHour Then
                    If _startTime.Minute > My.Settings.Exclude_EndMinutes Then
                        Return True
                    End If
                End If

            End If


            If _DT >= My.Settings.AdvancedSettings_DTcutoff Then Return True 'make sure DT Cutoff works regardless of mapping
            If _UT >= My.Settings.AdvancedSettings_UTcutoff Then Return True

            Select Case parentLine.Mapping_DTschedPlannedUnplanned
                Case DTsched_Mapping.Greensboro
                    Return isDTExcluded_GBO()
                Case DTsched_Mapping.SkinCare
                    Return isDTExcluded_SkinCare()
                Case DTsched_Mapping.Phenoix
                    Return isDTExcluded_PHX()
                Case DTsched_Mapping.HuangpuHC
                    Return isDTExcluded_HuangpuHC()
                Case DTsched_Mapping.SwingRoad
                    Return isDTExcluded_SwingRoad()
                Case DTsched_Mapping.APDO
                    Return isDTExcluded_APDO()
                Case DTsched_Mapping.Ukraine
                    Return isDTExcluded_Ukraine()
                Case DTsched_Mapping.Belleville
                    Return isDTExcluded_Belleville()
                Case DTsched_Mapping.Mariscala
                    Return isDTExcluded_Mariscala()
                Case DTsched_Mapping.BabyCare
                    Return isDTExcluded_BabyCare()
                Case DTsched_Mapping.ModPack
                    Return isDTExcluded_ModPACK()
                Case DTsched_Mapping.Albany
                    If parentLine.Name = "ModPACK" Then
                        If Reason2.Contains("Curtailment") Then Return True
                    End If
                    Return isDTExcluded_Albany()
                Case DTsched_Mapping.Hyderabad
                    Return isDTExcluded_Hyderabad()
                Case DTsched_Mapping.Budapest
                    Return isDTExcluded_Budapest()
                Case DTsched_Mapping.Rakona
                    Return isDTExcluded_Rakona()
                Case DTsched_Mapping.SingaporePioneer
                    Return isDTExcluded_SingaporePioneer()
                Case DTsched_Mapping.ModPack
                    Return isDTExcluded_ModPACK()
                Case DTsched_Mapping.Fem_LuisCustom
                    Return isDTExcluded_Fem_LuisCustom()
                Case Else
                    Throw New unknownMappingException
            End Select

        End If

    End Function
    Private Function isDTCIL() As Boolean
        Select Case parentLine.Mapping_DTschedPlannedUnplanned
            Case DTsched_Mapping.Greensboro
                Return isDTCil_GBO()
            Case DTsched_Mapping.SkinCare
                Return isDTCil_GBO()
            Case DTsched_Mapping.Phenoix
                Return isDTCil_PHX()
            Case DTsched_Mapping.SwingRoad
                Return isDTCil_SwingRoad()
            Case DTsched_Mapping.HuangpuHC
                Return isDTCil_HuangpuHC()
            Case DTsched_Mapping.APDO
                Return isDTCil_APDO()
            Case DTsched_Mapping.Ukraine
                Return isDTCil_Ukraine()
            Case DTsched_Mapping.Belleville
                Return isDTCil_Belleville()
            Case DTsched_Mapping.Mariscala
                Return isDTCil_Mariscala()
            Case DTsched_Mapping.BabyCare
                Return isDTCil_BabyCare()
            Case DTsched_Mapping.ModPack
                Return isDTCil_Albany()
            Case DTsched_Mapping.Albany
                Return isDTCil_Albany()
            Case DTsched_Mapping.Hyderabad
                Return isDTCil_Hyderabad()
            Case DTsched_Mapping.Budapest
                Return isDTCil_Budapest()
            Case DTsched_Mapping.Rakona
                Return isDTCil_Rakona()
            Case DTsched_Mapping.SingaporePioneer
                Return isDTCil_SingaporePioneer()
            Case DTsched_Mapping.Fem_LuisCustom
                Return isDTCil_Fem_LuisCustom()
            Case Else
                Throw New unknownMappingException
        End Select
    End Function
    Private Function isDTChangeover() As Boolean
        Select Case parentLine.Mapping_DTschedPlannedUnplanned
            Case DTsched_Mapping.Greensboro
                Return isDTChangeover_GBO()
            Case DTsched_Mapping.SkinCare
                Return isDTChangeover_GBO()
            Case DTsched_Mapping.Phenoix
                Return isDTChangeover_PHX()
            Case DTsched_Mapping.SwingRoad
                Return isDTChangeover_SwingRoad()
            Case DTsched_Mapping.HuangpuHC
                Return isDTChangeover_HuangpuHC()
            Case DTsched_Mapping.APDO
                Return isDTChangeover_APDO()
            Case DTsched_Mapping.Ukraine
                Return isDTChangeover_Ukraine()
            Case DTsched_Mapping.Belleville
                Return isDTChangeover_Belleville()
            Case DTsched_Mapping.Mariscala
                Return isDTChangeover_Mariscala()
            Case DTsched_Mapping.BabyCare
                Return isDTChangeover_BabyCare()
            Case DTsched_Mapping.ModPack
                Return isDTChangeover_Albany()
            Case DTsched_Mapping.Albany
                Return isDTChangeover_Albany()
            Case DTsched_Mapping.Hyderabad
                Return isDTChangeover_Hyderabad()
            Case DTsched_Mapping.Budapest
                Return isDTChangeover_Budapest()
            Case DTsched_Mapping.Rakona
                Return isDTChangeover_Rakona()
            Case DTsched_Mapping.SingaporePioneer
                Return isDTChangeover_SingaporePioneer()
            Case DTsched_Mapping.Fem_LuisCustom
                Return isDTChangeover_Fem_LuisCustom()
            Case Else
                Throw New unknownMappingException
        End Select
    End Function
#End Region

#Region "Fem_LuisCustom"
    Private Function isDTUnplanned_Fem_LuisCustom() As Boolean
        If isDTPlanned_Fem_LuisCustom() Then Return False
        Return True
    End Function

    Private Function isDTExcluded_Fem_LuisCustom() As Boolean
        If Left(_PR_inout, 4) = "PR O" Then Return True
        If _Reason1.Contains("STNU") Then Return True
        If _Reason1.Contains("Non-Scheduled") Then Return True
        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_Fem_LuisCustom() As Boolean
        If _Location.Contains("No Area") Then
            If _Reason1.Contains("CHANGEOVER PARTS HANGING") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("EXTERNAL REASONS") Then
                If _Reason2.Contains("Programmed") Then
                    Return True
                End If
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("IWS") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("MPSa") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("PREVENTIVE MAINTENACE (PM)") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("Programmed") Then
                If _Reason2.Contains("Programmed") Then
                    Return True
                End If
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("SAFETY") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("CHANGE OVER") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("DOWNTIME PLANEJADO") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("EO") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("PARADA INTENCIONAL") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("SEGURANÇA") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("SHUTDOWN") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("CIL") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("EO Não Vendável") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("EO Vendável") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("TREINAMENTO") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("Startup Shutdown") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("REUNIÃO") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("Projeto") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("PtD") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("Manutenção Planejada") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("INTENTIONAL STOP") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("AUTONOMOUS MAINTENANCE") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("4HR SURVIVAL") Then
                Return True
            End If
        End If

        If _Location.Contains("OFFLINE") Then
            If _Reason2.Contains("SHUT DOWN") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("CHANGEOVER") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("COMISSIONING") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("E/O") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("PROGRESSIVE MAINTENANCE") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("PROJECTS") Then
                Return True
            End If
        End If

        If _Location.Contains("No Area") Then
            If _Reason1.Contains("SURVIVAL") Then
                Return True
            End If
        End If

        Return False
    End Function
    Private Function isDTChangeover_Fem_LuisCustom() As Boolean
        If _Reason2 = "Changeover" Then Return True
        Return False
    End Function
    Public Function isDTCil_Fem_LuisCustom() As Boolean
        If _Reason3.Contains("CIL") Then Return True
        Return False
    End Function
#End Region

#Region "Singapore Pioneer"


    Private Function isDTUnplanned_SingaporePioneer() As Boolean
        If isDTPlanned_SingaporePioneer() Then Return False
        Return True
    End Function

    Private Function isDTExcluded_SingaporePioneer() As Boolean
        If Left(_PR_inout, 4) = "PR O" Then Return True
        If _Reason1.Contains("STNU") Then Return True
        If _Reason1.Contains("Non-Scheduled") Then Return True

        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_SingaporePioneer() As Boolean
        If _Reason1.Contains("Planned") Then
            Return True
        ElseIf _Reason2.Contains("Induced") Then
            Return True
        End If
        Return False
    End Function
    Private Function isDTChangeover_SingaporePioneer() As Boolean
        If _Reason2 = "Changeover" Then Return True
        Return False
    End Function
    Public Function isDTCil_SingaporePioneer() As Boolean
        If _Reason3.Contains("CIL") Then Return True
        Return False
    End Function
#End Region

#Region "SwingRoad"
    'swing road
    Private Function isDTUnplanned_SwingRoad() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_SwingRoad() Then
                Return False
            Else
                Return True
            End If

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_SwingRoad() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 5) = "OEE O" Then
            Return True
        ElseIf Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If

        If Left(_Reason1, 4) = "Excl" Then Return True
        If _DT > 1500 Then Return True


        If _Reason2 = "No Scheduled Production" Or _Reason2.Equals("Line Not Scheduled") Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then   ' LG Code 'PHC Phoenix Hack
            Return True ' LG Code ' PHC Phoenix Hack
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If

        Return False
    End Function
    Private Function isDTPlanned_SwingRoad() As Boolean

        If Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_SwingRoad() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True
        End If
        Return False
    End Function
    Public Function isDTCil_SwingRoad() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
        End If
        Return False
    End Function
#End Region
#Region "PHX"
    'PHEONIX
    Private Function isDTUnplanned_PHX() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_PHX() Then
                Return False
            Else
                Return True
            End If

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_PHX() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True
        ' LG Code

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then   ' LG Code 'PHC Phoenix Hack
            Return True ' LG Code ' PHC Phoenix Hack
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If

        If _DT > 800 Then
            Return True
            '  If isDTPlanned_PHX() Then Return True
        End If
        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_PHX() As Boolean

        If Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_PHX() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True
        End If
        Return False
    End Function
    Public Function isDTCil_PHX() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
        End If
        Return False
    End Function
#End Region
#Region "GBO"
    'GREENSBORO
    Private Function isDTUnplanned_GBO() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_GBO() Then
                Return False
            Else
                Return True
            End If

        ElseIf _Reason1.Contains("CIL/RLS") Or _Reason1.Contains("CIL / RLS") Then
            Return False

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_GBO() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If
        If Left(_Reason1, 4) = "Excl" Then Return True

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then
            Return True
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If

        If _DT > 800 And _Reason3.Contains("SHUTDOWN") Then Return True

        If parentLine.parentModule.Name = BS_APDO Then
            If _Reason1 = "Supply Losses" Then
                Return True
            ElseIf _Reason1 = "Starved" Then
                Return True
            End If
        End If

        Return False
    End Function
    Private Function isDTPlanned_GBO() As Boolean

        If Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _Reason1.Contains("CIL/RLS") Or _Reason1.Contains("CIL / RLS") Or _Reason1.Equals("CIL") Then
            Return True
        ElseIf _Location.Contains("Planned") Or _Location.Contains("PLANNED") Then
            Return True
        ElseIf _Reason1.Contains("01_Plan Stop") Or _Reason2.Contains("INTENTIONAL STOP") Then
            Return True
        ElseIf _Reason1.Contains("IWS") Or _Reason1.Contains("CHANGEOVER") Then
            Return True


        ElseIf _Fault.Equals("PLC 099 PLANNED STOP") Then
            Return False
        ElseIf _Reason1.Contains("MANTENIMIENTO PLANIFICADO") Or _Reason2.Contains("(PM02)") Or _Reason1.Contains("(PM02)") Then
            Return True
        ElseIf parentLine.parentSite.Name = "Montornes" Then
            Return isDTPlanned_Montornes()
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function

    Private Function isDTPlanned_Montornes() As Boolean

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("CAMBIO TALLA") Then
                    Return True
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("CAMBIO TALLA") Then
                    If _Reason4.Equals("CAMBIO TALLA") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("CAMBIO UNIDADES") Then
                    Return True
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("CAMBIO UNIDADES") Then
                    If _Reason4.Equals("CAMBIO UNIDADES") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("MANTENIMIENTO") Then
                Return True
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("MANTENIMIENTO") Then
                If _Reason3.Equals("MANTENIMIENTO AUXILIAR") Then
                    If _Reason4.Equals("MANTENIMIENTO AUXILIAR") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("MANTENIMIENTO") Then
                If _Reason3.Equals("MANTENIMIENTO PROGRAMADO") Then
                    If _Reason4.Equals("MANTENIMIENTO PROGRAMADO") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("MANTENIMIENTO") Then
                If _Reason3.Equals("NUEVO MODO DE FALLO") Then
                    If _Reason4.Equals("INTRODUCIR NUEVA CAUSA RAIZ DE FALLO") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO ADMINISTRATIVO") Then
                If _Reason3.Equals("FORMACIÓN") Then
                    Return True
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO ADMINISTRATIVO") Then
                If _Reason3.Equals("FORMACIÓN") Then
                    If _Reason4.Equals("FORMACIÓN") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO ADMINISTRATIVO") Then
                If _Reason3.Equals("PARO FIN DE SEMANA") Then
                    Return True
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO ADMINISTRATIVO") Then
                If _Reason3.Equals("PARO FIN DE SEMANA") Then
                    If _Reason4.Equals("PARO FIN DE SEMANA") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO ADMINISTRATIVO") Then
                If _Reason3.Equals("REUNION EQUIPO MENSUAL") Then
                    Return True
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO ADMINISTRATIVO") Then
                If _Reason3.Equals("REUNION EQUIPO MENSUAL") Then
                    If _Reason4.Equals("REUNION EQUIPO MENSUAL") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO IWS/EET") Then
                Return True
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO IWS/EET") Then
                If _Reason3.Equals("CIL") Then
                    If _Reason4.Equals("CIL") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO IWS/EET") Then
                If _Reason3.Equals("NUEVO MODO DE FALLO") Then
                    If _Reason4.Equals("INTRODUCIR NUEVA CAUSA RAIZ DE FALLO") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PARO IWS/EET") Then
                If _Reason3.Equals("PARO PLANIFICADO AM/PM/EET") Then
                    If _Reason4.Equals("PARO PLANIFICADO AM/PM/EET") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PROYECTOS") Then
                Return True
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PROYECTOS") Then
                If _Reason3.Equals("CONTROL DE CALIDAD") Then
                    If _Reason4.Equals("CONTROL DE CALIDAD") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PROYECTOS") Then
                If _Reason3.Equals("EO/PRUEBA DE MATERIA PRIMA") Then
                    If _Reason4.Equals("EO/PRUEBA DE MATERIA PRIMA") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PROYECTOS") Then
                If _Reason3.Equals("NUEVO MODO DE FALLO") Then
                    If _Reason4.Equals("INTRODUCIR NUEVA CAUSA RAIZ DE FALLO") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PROYECTOS") Then
                If _Reason3.Equals("OTRAS EO'S/PRUEBAS") Then
                    If _Reason4.Equals("OTRAS EO'S/PRUEBAS") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("OPERACION LINEA") Then
            If _Reason2.Equals("PROYECTOS") Then
                If _Reason3.Equals("TEST") Then
                    If _Reason4.Equals("TEST") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            Return True
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Fina y Segura") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Fina y Segura") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Maxi Ausonia") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Maxi Ausonia") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Maxi Lambada") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Maxi Lambada") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Mona Lisa") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Mona Lisa") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Normal Ausonia") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Normal Ausonia") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Normal Lambada") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Normal Lambada") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("P0") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("P0") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Tanga") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Tanga") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Venus") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE TALLA") Then
                If _Reason3.Equals("Venus") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("MEGAPACK") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("MEGAPACK") Then
                    If _Reason4.Equals("Añadir MGPack") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("MEGAPACK") Then
                    If _Reason4.Equals("Quitar MGPack") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("Perfume") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("Perfume") Then
                    If _Reason4.Equals("Añadir Perfume") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("Perfume") Then
                    If _Reason4.Equals("Quitar Perfume") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("Unidades") Then
                    If _Reason4.Equals("Ajustes tras arrancada") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CAMBIO TALLA / UDS") Then
            If _Reason2.Equals("CAMBIO DE UNIDADES / FORMATO") Then
                If _Reason3.Equals("Unidades") Then
                    If _Reason4.Equals("Cambio físico") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CIL") Then
            Return True
        End If

        If _Reason1.Equals("CIL") Then
            If _Reason2.Equals("CIL") Then
                If _Reason3.Equals("CIL") Then
                    If _Reason4.Equals("CIL") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CIL") Then
            If _Reason2.Equals("Preparacion de tours/auditorias") Then
                If _Reason3.Equals("Preparacion de tours/auditorias") Then
                    If _Reason4.Equals("Preparacion de tours/auditorias") Then
                        Return True
                    End If
                End If
            End If
        End If

        If _Reason1.Equals("CIL") Then
            If _Reason2.Equals("RLS") Then
                If _Reason3.Equals("RLS") Then
                    If _Reason4.Equals("RLS") Then
                        Return True
                    End If
                End If
            End If
        End If











        Return False
    End Function
    Private Function isDTChangeover_GBO() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True 'Greensboro
            If InStr(_Reason1, "990", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "991", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "992", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_DTGroup, "C/O", vbTextCompare) > 0 Then Return True ' LG Code Greensboro
            If InStr(_Reason2, "Format Change", vbTextCompare) > 0 Then Return True ' LG Code Iowa City IC

        End If
        Return False
    End Function
    Public Function isDTCil_GBO() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
            If _Reason2 = "RLS" Then Return True
        End If

        Return False
    End Function
#End Region

#Region "Budapest"
    'GREENSBORO
    Private Function isDTUnplanned_Budapest() As Boolean
        If isDTPlanned_Budapest() Then
            Return False
        Else
            Return True
        End If

        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_Budapest() Then
                Return False
            Else
                Return True
            End If

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_Budapest() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If

        Return False
    End Function
    Private Function isDTPlanned_Budapest() As Boolean
        If _Reason1.Equals("RLS") Or _Reason1.Equals("TRAINING") Then
            Return True

        ElseIf _Reason1.Equals("ATALLAS SCO") Then

            Return True

        ElseIf _Reason1.Equals("TRAINING") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS TCO") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS TCCO") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS CCO OS2 DUO - SINGLE") Then

            Return True

        ElseIf _Reason1.Equals("EO") Then

            Return True

        ElseIf _Reason1.Equals("GEPEGYSEG FELELOS STOP") Then

            Return True

        ElseIf _Reason1.Equals("PM02 - KARBANTARTAS") Then

            Return True

        ElseIf _Reason1.Equals("SOR LEALLITAS - INDITAS") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS  KAZETTAS - ZOMITOS") Then

            Return True

        ElseIf _Reason1.Equals("SOR INDITAS") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS  ZOMITOS - KAZETTAS") Then

            Return True

        ElseIf _Reason1.Equals("SEGEDBERENDEZESEK") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS STCO") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS PMC") Then

            Return True

        ElseIf _Reason1.Equals("TEK") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS SINGLE - DUO") Then

            Return True

        ElseIf _Reason1.Equals("MOI_STOP") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS DUO - SINGLE") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS CCO OM2 DUO - SINGLE") Then

            Return True

        ElseIf _Reason1.Equals("FLOW TO WORK") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS CCO OS2 SINGLE - DUO") Then

            Return True

        ElseIf _Reason1.Equals("ATALLAS CCO OM2 SINGLE - DUO") Then

            Return True

        ElseIf _Reason1.Equals("MEETING") Then

            Return True

        ElseIf _Reason1.Equals("SOR LEALLITAS") Then

            Return True

        ElseIf _Reason1.Equals("KVALIFIKACIO") Then

            Return True

        ElseIf _Reason1.Equals("RLS") Then

            Return True

        ElseIf _Reason1.Equals("KIURITES") Then

            Return True


        ElseIf _Reason1.Contains("ATALLAS TCO") Then
            Return True

        ElseIf _Reason1.Contains("TRAINING") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T5 IP 2.0 - T5 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1S NW - T1 IP3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS CCO") Then
            Return True
        ElseIf _Reason1.Contains("ATALLAS TCO") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS SCO") Then
            Return True

        ElseIf _Reason1.Contains("SOR LEALLITAS - INDITAS") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 NW FRESH - T3 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("MEETING") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T5 IP 2.0 - T5 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS PMC") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3S IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS SINGLE - DUO") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS DUO - SINGLE") Then
            Return True

        ElseIf _Reason1.Contains("TEK") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 FRESH - T5 IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3S IP 2.0 - T3 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 FRESH - T1 IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 2.0 FRESH - T3 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 FRESH - T3S IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 NW FRESH - T3 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1S IP2.0 - T1 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP2.0 - T1 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1    T1 IP3.0 - T1S IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 FRESH - T3S IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3  T3 IP 3.0 FRESH - T3 IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("PM02 - KARBANTARTAS") Then
            Return True

        ElseIf _Reason1.Contains("KVALIFIKACIO") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T3 IP 2.0 NW -T5 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("GEPEGYSEG FELELOS STOP") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 FRESH - T3 NW") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T3S IP 2.0 - T5 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1S NW FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 NW - T3 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("UTILITY") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3 NW") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 FRESH - T1 IP3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 - T5 IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS SCO LCC1   T2S IP2.0 - T1 IP3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3  T3 IP 2.0 - T3 IP 3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1 IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3S IP 2.0 - T3 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0  FRESH - T3 IP 2.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T3 IP 2.0 NW -T5 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 FRESH - T3 IP 2.0 NW") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 FRESH - T3 NW FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3  T3 IP 3.0 - T3 IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1S NW FRESH - T1 IP3.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1S NW") Then
            Return True

        ElseIf _Reason1.Contains("EO") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 FRESH - T1S  IP 2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 NW - T3 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3  T3 IP 2.0 - T3 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC4 T3S IP 2.0 - T5 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1S NW - T1 IP3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC1   T1S IP2.0 - T1 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS SCO LCC1   T2S IP2.0 - T1 IP3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 2.0 FRESH - T3 IP 3.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3 NW FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3 IP 2.0 FRESH") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS SCO LCC1   T1 IP3.0 FRESH - T2S IP2.0") Then
            Return True

        ElseIf _Reason1.Contains("ATALLAS SCO LCC1   T1 IP3.0 - T2S IP2.0") Then
            Return True
        Else
            Return False
        End If


    End Function
    Private Function isDTChangeover_Budapest() As Boolean
        If _isPlanned Then

        End If
        Return False
    End Function
    Public Function isDTCil_Budapest() As Boolean
        If _isPlanned Then
            If _Reason1 = "RLS" Then Return True
        End If

        Return False
    End Function
#End Region

#Region "APDO"
    'GREENSBORO
    Private Function isDTUnplanned_APDO() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_APDO() Then
                Return False
            Else
                Return True
            End If

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_APDO() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True
        ' LG Code

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then   ' LG Code 'PHC Phoenix Hack
            Return True ' LG Code ' PHC Phoenix Hack
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If

        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_APDO() As Boolean

        If Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_APDO() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True 'Greensboro
            If InStr(_Reason1, "990", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "991", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "992", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_DTGroup, "C/O", vbTextCompare) > 0 Then Return True ' LG Code Greensboro
            If InStr(_Reason2, "Format Change", vbTextCompare) > 0 Then Return True ' LG Code Iowa City IC

        End If
        Return False
    End Function
    Public Function isDTCil_APDO() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
            If _Reason2 = "RLS" Then Return True
        End If

        Return False
    End Function
#End Region

#Region "Huangpu HC"

    Private Function isDTUnplanned_HuangpuHC() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_GBO() Then
                Return False
            Else
                Return True
            End If

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_HuangpuHC() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True
        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then   ' LG Code 'PHC Phoenix Hack
            Return True ' LG Code ' PHC Phoenix Hack
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If

        If _Reason3.Equals("No Plan") Then Return True
        If _DT > 400 Then Return True
        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_HuangpuHC() As Boolean

        If Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_HuangpuHC() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True 'Greensboro
            If InStr(_Reason1, "990", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "991", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "992", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_DTGroup, "C/O", vbTextCompare) > 0 Then Return True ' LG Code Greensboro
            If InStr(_Reason2, "Format Change", vbTextCompare) > 0 Then Return True ' LG Code Iowa City IC

        End If
        Return False
    End Function
    Public Function isDTCil_HuangpuHC() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
            If _Reason2 = "RLS" Then Return True
        End If

        Return False
    End Function
#End Region

#Region "Skin Care"
    Public Function isDTExcluded_SkinCare() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True
        ' LG Code

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then   ' LG Code 'PHC Phoenix Hack
            Return True ' LG Code ' PHC Phoenix Hack
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If
        Return False
    End Function
#End Region

#Region "Ukraine"
    'GREENSBORO
    Private Function isDTUnplanned_Ukraine() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_Ukraine() Then
                Return False
            Else
                Return True
            End If

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_Ukraine() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True
        ' LG Code

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then   ' LG Code 'PHC Phoenix Hack
            Return True ' LG Code ' PHC Phoenix Hack
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If
        Return False

    End Function
    Private Function isDTPlanned_Ukraine() As Boolean
        If isDTCil_Ukraine() Or isDTChangeover_Ukraine() Then
            Return True
        ElseIf _Reason1.Contains("ЗАМЕНА") Then
            Return True
        ElseIf Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_Ukraine() As Boolean
        If _Reason2 = "ПЕРЕХОД_" Or _Reason1 = "ПЕРЕХОД_" Then Return True
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True 'Greensboro
            If InStr(_DTGroup, "C/O", vbTextCompare) > 0 Then Return True ' LG Code Greensboro
            If InStr(_Reason2, "Format Change", vbTextCompare) > 0 Then Return True ' LG Code Iowa City IC

        End If
        Return False
    End Function
    Public Function isDTCil_Ukraine() As Boolean
        If _Reason2 = "CIL_" Or _Reason1 = "CIL_" Then Return True 'ПЕРЕХОД_

        Return False
    End Function
#End Region

#Region "Belleville"
    Private Function isDTUnplanned_Belleville() As Boolean

        If isDTPlanned_Belleville() Then
            Return False
        Else
            Return True
        End If

    End Function
    Public Function isDTExcluded_Belleville() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True


        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True
        ' LG Code

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True

        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If

        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If
        Return False
    End Function
    Private Function isDTPlanned_Belleville() As Boolean
        If isDTCil_Belleville() Or isDTChangeover_Belleville() Then
            Return True
        ElseIf _Reason1.Contains("Planned DownTime") Or _Reason1.Contains("Planned") Then
            Return True
        ElseIf _Reason1.Contains("EO") Then
            Return True
        ElseIf _Reason1.Contains("MEETING") Or _Reason1.Contains("01_Plan Stop") Then
            Return True
        ElseIf _Reason3.Contains("CIL") Or _Reason3.Contains("RLS") Or _Reason2.Contains("CAMBIO DE REFERENCIA") Or _Reason2.Contains("CAMBIO DE TALLA") Or _Reason2.Contains("PARO ADMINISTRATIVO") Or _Reason2.Contains("PROYECTOS") Then
            Return True
        ElseIf _Reason1.Contains("Logistics") Or _Reason2.Contains("CHANGE OVER") Or _Reason1.Contains("CHANGE OVER") Or _Reason2.Contains("IWS SHUTDOWN") Or _Reason2.Contains("PROGRAMADOS") Then
            Return True
        ElseIf _Location.Contains("No Area") And parentLine.Name.Contains("Tepeji") Then   'this is true for Tepeji fem care...          
            If _Reason1.Contains("OPERATIONAL") Or _Reason1.Contains("OPERACION") Or _Reason1.Contains("CHANGE OVER") Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_Belleville() As Boolean
        If _Reason1 = "Changeover" Then Return True
        Return False
    End Function
    Public Function isDTCil_Belleville() As Boolean
        If _Reason1 = "CIL" Then Return True
        Return False
    End Function
#End Region

#Region "Mariscala"
    Private Function isDTUnplanned_Mariscala() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_Mariscala() Then
                Return False
            Else
                Return True
            End If

        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_Mariscala() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Or Left(_PR_inout, 3) = "PRO" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True
        ' LG Code

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If
        If isDTPlanned_Mariscala() And _DT > 400 Then Return True

        Return False
    End Function
    Private Function isDTPlanned_Mariscala() As Boolean

        If Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_Mariscala() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True
            If InStr(_DTGroup, "C/O", vbTextCompare) > 0 Then Return True
            If InStr(_Reason2, "Format Change", vbTextCompare) > 0 Then Return True

        End If
        Return False
    End Function
    Public Function isDTCil_Mariscala() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
            If _Reason2 = "RLS" Then Return True
        End If

        Return False
    End Function
#End Region

#Region "BabyCare"

    Private Function isDTUnplanned_BabyCare() As Boolean
        If _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            If isDTPlanned_BabyCare() Then
                Return False
            Else
                Return True
            End If
        ElseIf _PlannedUnplanned.Contains("Planned") Then
            Return False
        ElseIf Left(_PlannedUnplanned, 1) = "U" Then
            Return True
        ElseIf Len(_PlannedUnplanned) < 2 Then
            Return True 'chck 4 blanks
        Else
            Return False
        End If
    End Function
    Public Function isDTExcluded_BabyCare() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False
        If _DT = 0 Then Return True
        ' If _UT = 0 Then Return True
        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If
        ' LG Code
        If Left(_Reason1, 4) = "Excl" Then Return True

        If _Reason2 = "No Scheduled Production" Then
            Return True
        ElseIf _Reason2 = "No Production Scheduled" Then
            Return True
        ElseIf _Reason3 = "EOW" And _DT > inCONTROL_SCHED_TIME_CUTOFF Then   ' LG Code 'PHC Phoenix Hack
            Return True ' LG Code ' PHC Phoenix Hack
        ElseIf Left(_Reason2, 4) = "Excl" Then
            Return True
        End If
        ' LG Code
        If _DTGroup = "STNU" Or _DTGroup = "Line Not Scheduled" Then
            Return True
        End If

        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_BabyCare() As Boolean

        If Left(_PlannedUnplanned, 1) = "P" Then
            Return True
        ElseIf Left(_Location, 7) = "Planned" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_BabyCare() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True
            If InStr(_Reason1, "990", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "991", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "992", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_DTGroup, "C/O", vbTextCompare) > 0 Then Return True
            If InStr(_Reason2, "Format Change", vbTextCompare) > 0 Then Return True

        End If
        Return False
    End Function
    Public Function isDTCil_BabyCare() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
            If _Reason2 = "RLS" Then Return True
        End If

        Return False
    End Function
#End Region

#Region "Albany"

    Private Function isDTUnplanned_Albany() As Boolean

        If isDTPlanned_Albany() Then
            Return False
        Else
            Return True
        End If

    End Function
    Public Function isDTExcluded_Albany() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If _DTGroup.Contains("Curtail") Then Return True
        If _Reason2.Contains("Curtailment") Then Return True

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        End If


        'Special exclusion rules for Paper Making
        If AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyMaking Then
            If _Reason1.Contains("Not Applicable") And _Reason4.Contains("Holiday") Then Return True
            If _Reason2.Contains("Curtail") Then Return True
            If _Reason2.Contains("Holiday") Then Return True
            If _Reason2.Contains("EO") Then Return True
            If _Reason1.Contains("Natural Causes") Then Return True
            If _Reason1.Contains("Production Planning") Then Return True
        End If


        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_Albany() As Boolean


        ''''''''''''''''''''''''
        If AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyCareUnitOP_Wrapper Or AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyCareUnitOP_ModPACK Then
            If _Reason1.Contains("STARVE") Or _Reason1.Contains("Starv") Then
                Return True
            ElseIf _Reason1.Contains("BLOCK") Or _Reason1.Contains("Block") Or _Reason1.Contains("Starved") Or _Reason1.Contains("Blocked") Then
                Return True
            ElseIf _Reason1.Contains("No Backlog") Then
                Return True
            ElseIf _Location.Contains("Starv") Then
                Return True
            ElseIf _Location.Contains("Block") Then
                Return True
            ElseIf _Fault.Contains("Starv") Then
                Return True
            ElseIf _Fault.Contains("Block") Then
                Return True
            End If
        ElseIf AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyCareUnitOP_mf Then
            If _Reason1.Contains("STARVE") Or _Reason1.Contains("Starv") Then
                Return True
            ElseIf _Reason1.Contains("BLOCK") Or _Reason1.Contains("Block") Then
                Return True
            ElseIf _Fault.Contains("BLOCK") Or _Fault.Contains("Block") Then
                Return True
            ElseIf _Fault.Contains("STARVE") Or _Fault.Contains("STARVE") Then
                Return True
            ElseIf _Reason1.Contains("Backlog") Then
                Return True
            ElseIf _Reason1.Contains("Die Jam") And _Reason2.Contains("Turned") Then
                Return True
            ElseIf _Reason1.Contains("Tipped") And _Reason2.Contains("Turned") Then
                Return True
            ElseIf _Reason1.Contains("Fault") And _Reason2.Contains("Incomplete") Then
                Return True
            ElseIf _Reason1.Contains("Fault") And _Reason2.Contains("Tipped") Then
                Return True
            ElseIf _Reason1.Contains("Fault") And _Reason2.Contains("Turned") Then
                Return True
            ElseIf _Reason1.Contains("Slug") Then
                Return True
            ElseIf _Reason1.Contains("Divert") Then
                Return True
            ElseIf _Reason1.Contains("Disch") Then
                Return True
            ElseIf _Reason1.Contains("CUSTOMER") Then
                Return True
            ElseIf _Location.Contains("Starved") Then
                Return True
            End If

        ElseIf AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyCareUnitOP_ACP Then
            If _Reason2.Contains("DISC") And _Reason2.Contains("FULL") Then
                Return True
            ElseIf _Reason1.Contains("FCC Down") Then
                Return True
            ElseIf _Reason1.Contains("STARVE") Or _Reason1.Contains("Starv") Then
                Return True
            ElseIf _Location.Contains("Package Conveyor") Then
                Return True
            ElseIf _Reason1.Contains("Divert") Then
                Return True
            ElseIf _Location.Contains("Starved") Then
                Return True
            ElseIf _Fault.Contains("Block") Or _Fault.Contains("BLOCK") Then
                Return True
            ElseIf _Fault.Contains("STARVE") Or _Fault.Contains("starve") Then
                Return True
            End If
        ElseIf AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyMaking Then
            If _Location.Contains("DT") And _Reason2.Contains("Brand Change") Then
                Return True

            ElseIf _Location.Contains("DT") And _Reason2.Contains("Downday") Then
                Return True
            ElseIf _Location.Contains("SB") And _Reason3.Contains("Brand Change") Then
                Return True

            ElseIf _Fault.Contains("Sheet") And _Reason2.Contains("Creping Blade") Then
                Return True
            ElseIf _Fault.Contains("Sheet") And _Reason3.Contains("Creping Blade") Then
                Return True
            ElseIf _Location.Contains("Operations DT") And _Fault.Contains("Operators") And _Reason1.Contains("Operators") And _Reason4.Contains("Proces/Operational") Then
                Return True
            ElseIf _Reason1.Contains("Not Applicable") And _Reason4.Contains("Outage") Then
                Return True
            ElseIf _Reason1.Contains("Not Applicable") And _Reason4.Contains("Downday") Then
                Return True
            ElseIf _Reason1.Contains("Not Applicable") And _Reason4.Contains("Brandswing") Then
                Return True
            ElseIf _Reason1.Contains("Production Planning") And _Reason2.Contains("Outage") Then
                Return True
            ElseIf _Reason1.Contains("Production Planning") And _Reason3.Contains("Process/Operational") Then
                Return True
            ElseIf _Reason2.Contains("Outage") Or _Reason2.Contains("Downday") Then
                Return True

            End If
        End If

        'For full line


        If _Reason1.Contains("Planned") Then
            Return True
        ElseIf _Reason1.Contains("UWS05") Then
            Return True
        ElseIf _DTGroup = "Changeover" Then
            Return True

        ElseIf _Reason1.Contains("Product Change") Then

            Return True
        ElseIf _Reason1.Contains("Poly Roll") Then

            Return True
        ElseIf _Reason1.Contains("Planned Intervention") Then

            Return True
        ElseIf _Reason1.Contains("GEN01") And _Reason2.Contains("Maintenance") Then

            Return True
        ElseIf _Reason1.Contains("GEN01") And _Reason2.Contains("Blowdown") Then

            Return True
        ElseIf _Reason1.Contains("GEN01") And _Reason2.Contains("CIL") Then

            Return True
        ElseIf _Reason1.Contains("GEN01") And _Reason2.Contains("Clean") Then

            Return True
        ElseIf _Reason1.Contains("GEN01") And _Reason2.Contains("centerline") Then

            Return True
        ElseIf _Reason1.Contains("GEN01") And _Reason2.Contains("AM") Then

            Return True
        ElseIf _DTGroup.Contains("Holiday/Curtail") Then

            Return True
        ElseIf _DTGroup.Contains("E.O./Projects") Then

            Return True
        ElseIf _DTGroup.Contains("Special Causes") Then

            Return True
        ElseIf _DTGroup.Contains("PR/Poly Change") Then
            Return True

        ElseIf _DTGroup.Contains("Planned Intervention") Then
            Return True
        ElseIf _DTGroup.Contains("Planned Hygiene/Cleaning") Then
            Return True

            ''''''''Blocked starved considered planned for wrapper''''''''''''''


        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_Albany() As Boolean
        If _isPlanned Then
            If _DTGroup = "Changeover" Then Return True
            If _Reason1.Contains("Change") Then Return True
            If _Reason2.Contains("Change") Then Return True
        End If
        Return False
    End Function
    Public Function isDTCil_Albany() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If _Reason2.Contains("autonomous") Then Return True
            If _Reason1.Contains("RLS") Then Return True
            If _Reason2.Contains("RLS") Then Return True
            If _Tier1.Contains("RLS") Then Return True
        End If

        Return False
    End Function


    Private Function isDTUnplanned_ModPACK() As Boolean
        If _Reason2.Contains("QF") Or _Reason2.Contains("Quality - Film") Then
            _Tier1 = "Quality Film"
            _Tier2 = _Reason2
            _Tier3 = _Fault
            Return True
        ElseIf _Reason1.Contains("WRP10 Incoming Quality - Rolls") Or _Reason2.Contains("QR") Or _Reason2.Contains("Quality - Rolls") Then
            _Tier1 = "Quality Rolls"
            _Tier2 = _Reason2
            _Tier3 = _Fault
            Return True
        ElseIf _Reason2.Contains("QP Package Ears") Then
            _Tier1 = "830, 835"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True
        ElseIf _Reason2.Contains("QP Poor / No Lap Seal") Then
            _Tier1 = "820, 825"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True
        ElseIf _Reason2.Contains("QP Packages Stuck Together") Then
            _Tier1 = "820, 825"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True
        ElseIf _Reason1.Contains("WRP36 Package Separation Failure") Then
            _Tier1 = "820, 825"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True
        ElseIf _Fault.Contains("805") Or _Fault.Contains("810") Or _Fault.Contains("815") Then
            _Tier1 = "805, 810, 815"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True
        ElseIf _Fault.Contains("820") Or _Fault.Contains("825") Then
            _Tier1 = "820, 825"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True
        ElseIf _Fault.Contains("835") Or _Fault.Contains("830") Then
            _Tier1 = "830, 835"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True

        ElseIf _Reason2.Contains("QP Log Compressibility") Or _Reason2.Contains("QP Paper Downstream") Or _Reason2.Contains("QP Low Log Compressibility") Or _Reason2.Contains("QP High Log Compressibility") Then
            _Tier1 = "Quality Rolls"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True


        ElseIf _Location.Contains("WRP") And (_Reason2.Contains("Mechanical") Or _Reason2.Contains("Electrical")) Then
            _Tier1 = "OTHER PRIMARY"
            _Tier2 = _Reason2
            _Tier3 = _Reason3
            Return True
        End If

        Return False
    End Function
    Public Function isDTExcluded_ModPACK() As Boolean

        Dim r3Num As Integer = 100000

        If _Fault.Contains("[4942] UN110 Remote suspend command from Pack Master") Then
            Return True
        ElseIf _DT > 30 Then
            Return True
        ElseIf _Fault.Contains("4950") Then
            Return True
        ElseIf _Reason2.Contains("QR Standing Roll") Then
            Return True
        ElseIf _Reason2.Contains("Tipped") Then
            Return True
        ElseIf _Fault.Contains("[2] Blocked") Or _Fault.Contains("[1] Starved") Then
            Return True
        ElseIf _Reason1.Contains("MP98 Starved") Then
            Return True
        ElseIf _Reason2.Contains("Upstream") Then
            Return True
        ElseIf _Reason1.Contains("M99 Blicked") Then
            Return True
        ElseIf _Reason2.Contains("Codedater") Or _Reason2.Contains("Downstream") Or _Reason2.Contains("QK") Or _Reason2.Contains("Quality - KDF's") Then
            Return True
        ElseIf _Reason1.Contains("Codedater") Then
            Return True
        ElseIf _Fault.Contains("840") Then
            Return True
        ElseIf _Fault.Contains("845") Then
            Return True
        ElseIf _Fault.Contains("850") Then
            Return True
        ElseIf _Fault.Contains("855") Then
            Return True
        ElseIf _Fault.Contains("856") Then
            Return True
        ElseIf _Fault.Contains("860") Then
            Return True
        ElseIf _Fault.Contains("865") Then
            Return True
        ElseIf _Fault.Contains("870") Then
            Return True
        ElseIf _Fault.Contains("875") Then
            Return True
        ElseIf _Fault.Contains("880") Then
            Return True
        ElseIf Int32.TryParse(_Reason3, r3Num) Then
            If r3Num < 80000 Then
                Return True
            End If
        ElseIf _Reason1.Contains("Aux") Then
            Return True
        ElseIf _Reason1.Contains("Planned Intervention") And _Reason2.Contains("cap work") Then
            Return True
        ElseIf _Reason1.Contains("Planned Intervention") And _Reason2.Contains("outage") Then
            Return True
        ElseIf _Reason1.Contains("special causes") And _Reason2.Contains("curtailment") Then
            Return True
        ElseIf _Reason2.Contains("Mechanical") Or _Reason2.Contains("Electrical") Then
            If _Location.Contains("PKH") Or _Location.Contains("BND") Or _Location.Contains("ACP") Then
                Return True
            End If
        ElseIf _Reason1.Contains("special causes") Then
            Return True
        ElseIf _Reason1.Contains("blowdown") Then
            Return True
        ElseIf _Reason1.Contains("WND") Then
            If _Reason2.Contains("Troubleshooting") Or _Reason2.Contains("Re-centerlining") Or _Reason2.Contains("Set-up error") Or _Reason2.Contains("Adjust") Then
                Return True
            End If
        End If


        Return False
        Exit Function

    End Function
    Private Function isDTPlanned_ModPack() As Boolean
        If _Reason2.Contains("Changeover") Then
            _Tier1 = "CHANGEOVER"
            _Tier2 = _Team
            Return True
        ElseIf _Reason2.Contains("CIL") Or _Reason2.Contains("RLS") Or _Reason2.Contains("autonomous") Or _Reason2.Contains("maint") Or _Reason2.Contains("blowdown") Or _Reason2.Contains("centerline") Then
            _Tier1 = "CIL"
            _Tier2 = _Team
            Return True
        ElseIf _Reason1.Contains("Planned Intervention") Then
            _Tier1 = "Planned Intervention"
            _Tier2 = _Team
            Return True
        ElseIf _Reason1.Contains("special causes") And (_Reason2.Contains("Meetings") Or _Reason2.Contains("Safety")) Then
            _Tier1 = "Planned Intervention"
            _Tier2 = _Team
        End If

        Return False
    End Function



#End Region

#Region "Hyderabad" 'F&HC
    Private Function isDTUnplanned_Hyderabad() As Boolean

        If isDTPlanned_Hyderabad() Then
            Return False
        Else
            Return True
        End If

    End Function
    Public Function isDTExcluded_Hyderabad() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf _Reason1.Contains("NOT SCHEDULED") Or _Reason2.Contains("NOT SCHEDULED") Or _Reason3.Contains("NOT SCHEDULED") Or _Reason4.Contains("NOT SCHEDULED") Then
            Return True
        End If
        Return False
    End Function
    Private Function isDTPlanned_Hyderabad() As Boolean
        If _PlannedUnplanned.Contains("Planned") Then
            Return True

        ElseIf _DTGroup = "Changeover" Then
            Return True
        ElseIf _PlannedUnplanned.Equals(BLANK_INDICATOR) Then
            Return False
        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_Hyderabad() As Boolean
        If _isPlanned Then
            If _DTGroup = "Changeover" Then Return True
        End If
        Return False
    End Function
    Public Function isDTCil_Hyderabad() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If _Reason2.Contains("autonomous") Then Return True
        End If

        Return False
    End Function

#End Region

#Region "Rakona"
    Private Function isDTUnplanned_Rakona() As Boolean
        If isDTPlanned_Rakona() Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Function isDTExcluded_Rakona() As Boolean
        'If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        'If _DT > 1000 Or _DT < 0 Then Return True

        If Left(_DTGroup, 4) = "Idle" Or Left(_PR_inout, 4) = "PR O" Then
            Return True
            '  ElseIf Left(_ProductCode, 1) = "S" Then
            '     Return True
        End If
        Return False

    End Function
    Private Function isDTPlanned_Rakona() As Boolean

        If Left(_DTGroup, 7) = "Planned" Then
            Return True
        ElseIf _DTGroup = "Planned Downtime" Then
            Return True

        Else
            Return False
        End If
    End Function
    Private Function isDTChangeover_Rakona() As Boolean
        If _isPlanned Then
            If _Reason2 = "Changeover" Then Return True
            If Left(_DTGroup, 3) = "CO" Then Return True
            If InStr(_Reason1, "990", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "991", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_Reason1, "992", vbTextCompare) > 0 Then Return True ' LG Code Baby Care
            If InStr(_DTGroup, "C/O", vbTextCompare) > 0 Then Return True
            If InStr(_Reason2, "Format Change", vbTextCompare) > 0 Then Return True ' LG Code Iowa City IC

        End If
        Return False
    End Function
    Public Function isDTCil_Rakona() As Boolean
        If _isPlanned Then
            If _Reason2 = "CIL" Then Return True
            If Right(_DTGroup, 3) = "CIL" Then Return True
            If _Reason2 = "RLS" Then Return True
        End If

        Return False
    End Function
#End Region

#End Region

#Region "Initialize From Raw Data Array"
    'initialization
    Private Sub initializeFields()
        With parentLine
            If IsDBNull(.rawProficyData(DownTimeColumn.StartTime, rowNum)) Then
                _startTime = ""
            Else
                _startTime = .rawProficyData(DownTimeColumn.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Endtime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyData(DownTimeColumn.Endtime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.DT, rowNum)) Then
                _DT = 0
            Else
                _DT = .rawProficyData(DownTimeColumn.DT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.UT, rowNum)) Then
                _UT = 0
            Else
                _UT = .rawProficyData(DownTimeColumn.UT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Location, rowNum)) Then
                _Location = BLANK_INDICATOR
            Else
                _Location = .rawProficyData(DownTimeColumn.Location, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.MasterProdUnit, rowNum)) Then
                _MasterProdUnit = BLANK_INDICATOR
            Else
                _MasterProdUnit = .rawProficyData(DownTimeColumn.MasterProdUnit, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Fault, rowNum)) Then
                _Fault = BLANK_INDICATOR '""
            Else
                _Fault = .rawProficyData(DownTimeColumn.Fault, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Reason1, rowNum)) Then
                _Reason1 = BLANK_INDICATOR '""
            Else
                _Reason1 = .rawProficyData(DownTimeColumn.Reason1, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Reason2, rowNum)) Then
                _Reason2 = BLANK_INDICATOR '""
            Else
                _Reason2 = .rawProficyData(DownTimeColumn.Reason2, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Reason3, rowNum)) Then
                _Reason3 = BLANK_INDICATOR '""
            Else
                _Reason3 = .rawProficyData(DownTimeColumn.Reason3, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Reason4, rowNum)) Then
                _Reason4 = BLANK_INDICATOR '""
            Else
                _Reason4 = .rawProficyData(DownTimeColumn.Reason4, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.StopClass, rowNum)) Then
                _StopClass = BLANK_INDICATOR '""
            Else
                _StopClass = .rawProficyData(DownTimeColumn.StopClass, rowNum)
            End If

            If IsDBNull(.rawProficyData(DownTimeColumn.PR_InOut, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyData(DownTimeColumn.PR_InOut, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyData(DownTimeColumn.Team, rowNum)
                If _Team = "" Then
                    _Team = BLANK_INDICATOR
                End If
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.PlannedUnplanned, rowNum)) Then
                _PlannedUnplanned = BLANK_INDICATOR '""
            Else
                _PlannedUnplanned = .rawProficyData(DownTimeColumn.PlannedUnplanned, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.DTGroup, rowNum)) Then
                _DTGroup = BLANK_INDICATOR '""
            Else
                _DTGroup = .rawProficyData(DownTimeColumn.DTGroup, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyData(DownTimeColumn.Product, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyData(DownTimeColumn.ProductCode, rowNum)
            End If
            Try
                If IsDBNull(.rawProficyData(DownTimeColumn.ProductGroup, rowNum)) Then
                    _ProductGroup = BLANK_INDICATOR '""
                Else
                    _ProductGroup = .rawProficyData(DownTimeColumn.ProductGroup, rowNum)
                End If
                If IsDBNull(.rawProficyData(DownTimeColumn.Comment, rowNum)) Then
                    _Comment = ""
                Else
                    _Comment = .rawProficyData(DownTimeColumn.Comment, rowNum)
                End If
            Catch ex As Exception
                _ProductGroup = BLANK_INDICATOR
                _Comment = ""
            End Try
        End With
    End Sub


    Private Sub initializeFields_OneClickALTERNATE()
        With parentLine
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)) Then
                _startTime = "" '"0" 'LG Code
            Else
                _startTime = .rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DT, rowNum)) Then
                _DT = 0
            Else
                _DT = .rawProficyData(DownTimeColumn_OneClick.DT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.UT, rowNum)) Then
                _UT = 0 ' '"0" 'LG code
            Else
                _UT = .rawProficyData(DownTimeColumn_OneClick.UT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Location, rowNum)) Then
                _Location = BLANK_INDICATOR '""
            Else
                _Location = .rawProficyData(DownTimeColumn_OneClick.Location, rowNum)
            End If
            'If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)) Then
            _MasterProdUnit = BLANK_INDICATOR '"" '"0" 'LG Code
            _StopClass = BLANK_INDICATOR '""
            _ProductGroup = BLANK_INDICATOR
            'Else
            ' _MasterProdUnit = .rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)
            ' End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)) Then
                _Fault = BLANK_INDICATOR '""
            Else
                _Fault = .rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)) Then
                _Reason1 = BLANK_INDICATOR '""
            Else
                _Reason1 = .rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)) Then
                _Reason2 = BLANK_INDICATOR '""
            Else
                _Reason2 = .rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)) Then
                _Reason3 = BLANK_INDICATOR '""
            Else
                _Reason3 = .rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)) Then
                _Reason4 = BLANK_INDICATOR '""
            Else
                _Reason4 = .rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)
            End If



            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.PR_InOut, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyData(DownTimeColumn_OneClick.PR_InOut, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyData(DownTimeColumn_OneClick.Team, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.PlannedUnplanned, rowNum)) Then
                _PlannedUnplanned = BLANK_INDICATOR '""
            Else
                _PlannedUnplanned = .rawProficyData(DownTimeColumn_OneClick.PlannedUnplanned, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)) Then
                _DTGroup = BLANK_INDICATOR '""
            Else
                _DTGroup = .rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyData(DownTimeColumn_OneClick.Product, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Comment, rowNum)) Then
                _Comment = ""
            Else
                _Comment = .rawProficyData(DownTimeColumn_OneClick.Comment, rowNum)
            End If
        End With
    End Sub


    Private Sub initializeFields_OneClick()
        With parentLine
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)) Then
                _startTime = "" '"0" 'LG Code
            Else
                _startTime = .rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DT, rowNum)) Then
                _DT = 0
            Else
                _DT = .rawProficyData(DownTimeColumn_OneClick.DT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.UT, rowNum)) Then
                _UT = 0 ' '"0" 'LG code
            Else
                _UT = .rawProficyData(DownTimeColumn_OneClick.UT, rowNum)
                If _UT < -10 Then _UT = 0
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Location, rowNum)) Then
                _Location = BLANK_INDICATOR '""
            Else
                _Location = .rawProficyData(DownTimeColumn_OneClick.Location, rowNum)
            End If
            'If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)) Then
            _MasterProdUnit = BLANK_INDICATOR '"" '"0" 'LG Code
            _StopClass = BLANK_INDICATOR '""
            _ProductGroup = BLANK_INDICATOR
            'Else
            ' _MasterProdUnit = .rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)
            ' End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)) Then
                _Fault = BLANK_INDICATOR '""
            Else
                _Fault = .rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)) Then
                _Reason1 = BLANK_INDICATOR '""
            Else
                _Reason1 = .rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)) Then
                _Reason2 = BLANK_INDICATOR '""
            Else
                _Reason2 = .rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)) Then
                _Reason3 = BLANK_INDICATOR '""
            Else
                _Reason3 = .rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)) Then
                _Reason4 = BLANK_INDICATOR '""
            Else
                _Reason4 = .rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)
            End If



            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.PR_InOut, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyData(DownTimeColumn_OneClick.PR_InOut, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyData(DownTimeColumn_OneClick.Team, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.PlannedUnplanned, rowNum)) Then
                _PlannedUnplanned = BLANK_INDICATOR '""
            Else
                _PlannedUnplanned = .rawProficyData(DownTimeColumn_OneClick.PlannedUnplanned, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)) Then
                _DTGroup = BLANK_INDICATOR '""
            Else
                _DTGroup = .rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyData(DownTimeColumn_OneClick.Product, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Comment, rowNum)) Then
                _Comment = ""
            Else
                _Comment = .rawProficyData(DownTimeColumn_OneClick.Comment, rowNum)
            End If
        End With
    End Sub

    Private Sub initializeFields_GLEDS2()
        With parentLine
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.StartTime, rowNum)) Then
                _startTime = "" '"0" 'LG Code
            Else
                _startTime = .rawProficyData(DownTimeColumn_GLEDS.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Endtime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyData(DownTimeColumn_GLEDS.Endtime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.DT, rowNum)) Then
                _DT = 0
            Else
                _DT = .rawProficyData(DownTimeColumn_GLEDS.DT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.UT, rowNum)) Then
                _UT = 0 ' '"0" 'LG code
            Else
                _UT = .rawProficyData(DownTimeColumn_GLEDS.UT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Location, rowNum)) Then
                _Location = BLANK_INDICATOR '""
            Else
                _Location = .rawProficyData(DownTimeColumn_GLEDS.Location, rowNum)
            End If
            'If IsDBNull(.rawProficyData(DowntimeColumn_GLEDS.MasterProdUnit, rowNum)) Then
            _MasterProdUnit = BLANK_INDICATOR '"" '"0" 'LG Code
            _StopClass = BLANK_INDICATOR '""
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Product, rowNum)) Then
                _ProductGroup = BLANK_INDICATOR '""
            Else
                _ProductGroup = .rawProficyData(DownTimeColumn_GLEDS.Product, rowNum)
            End If
            'Else
            ' _MasterProdUnit = .rawProficyData(DowntimeColumn_GLEDS.MasterProdUnit, rowNum)
            ' End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Fault, rowNum)) Then
                _Fault = BLANK_INDICATOR '""
            Else
                _Fault = .rawProficyData(DownTimeColumn_GLEDS.Fault, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Reason1, rowNum)) Then
                _Reason1 = BLANK_INDICATOR '""
            Else
                _Reason1 = .rawProficyData(DownTimeColumn_GLEDS.Reason1, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Reason2, rowNum)) Then
                _Reason2 = BLANK_INDICATOR '""
            Else
                _Reason2 = .rawProficyData(DownTimeColumn_GLEDS.Reason2, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Reason3, rowNum)) Then
                _Reason3 = BLANK_INDICATOR '""
            Else
                _Reason3 = .rawProficyData(DownTimeColumn_GLEDS.Reason3, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Reason4, rowNum)) Then
                _Reason4 = BLANK_INDICATOR '""
            Else
                _Reason4 = .rawProficyData(DownTimeColumn_GLEDS.Reason4, rowNum)
            End If



            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.PR_InOut, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyData(DownTimeColumn_GLEDS.PR_InOut, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyData(DownTimeColumn_GLEDS.Team, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Location, rowNum)) Then
                '_PlannedUnplanned = BLANK_INDICATOR '""
                _PlannedUnplanned = "Unplanned"
            Else
                _PlannedUnplanned = .rawProficyData(DownTimeColumn_GLEDS.Location, rowNum)
                If Not _PlannedUnplanned.Contains("Planned") Then _PlannedUnplanned = "Unplanned"
            End If
            '  If IsDBNull(.rawProficyData(DowntimeColumn_GLEDS.DTGroup, rowNum)) Then
            _DTGroup = BLANK_INDICATOR '""
            '  Else
            ' _DTGroup = .rawProficyData(DowntimeColumn_GLEDS.DTGroup, rowNum)
            ' End If

            If IsDBNull(.rawProficyData(DownTimeColumn_GLEDS.Product, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyData(DownTimeColumn_GLEDS.Product, rowNum)
            End If
            _Comment = ""
        End With
        _Product = _ProductCode
    End Sub

    Private Sub initializeFields_GLEDS()
        With parentLine
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)) Then
                _startTime = "" '"0" 'LG Code
            Else
                _startTime = .rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DT, rowNum)) Then
                _DT = 0
            Else
                _DT = .rawProficyData(DownTimeColumn_OneClick.DT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.UT, rowNum)) Then
                _UT = 0 ' '"0" 'LG code
            Else
                _UT = .rawProficyData(DownTimeColumn_OneClick.UT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Location, rowNum)) Then
                _Location = BLANK_INDICATOR '""
            Else
                _Location = .rawProficyData(DownTimeColumn_OneClick.Location, rowNum)
            End If
            'If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)) Then
            _MasterProdUnit = BLANK_INDICATOR '"" '"0" 'LG Code
            _StopClass = BLANK_INDICATOR '""
            _ProductGroup = BLANK_INDICATOR
            'Else
            ' _MasterProdUnit = .rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)
            ' End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)) Then
                _Fault = BLANK_INDICATOR '""
            Else
                _Fault = .rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)) Then
                _Reason1 = BLANK_INDICATOR '""
            Else
                _Reason1 = .rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)) Then
                _Reason2 = BLANK_INDICATOR '""
            Else
                _Reason2 = .rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)) Then
                _Reason3 = BLANK_INDICATOR '""
            Else
                _Reason3 = .rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)) Then
                _Reason4 = BLANK_INDICATOR '""
            Else
                _Reason4 = .rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)
            End If



            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.PR_InOut - 1, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyData(DownTimeColumn_OneClick.PR_InOut - 1, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyData(DownTimeColumn_OneClick.Team, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Location, rowNum)) Then
                '_PlannedUnplanned = BLANK_INDICATOR '""
                _PlannedUnplanned = "Unplanned"
            Else
                _PlannedUnplanned = .rawProficyData(DownTimeColumn_OneClick.Location, rowNum)
                If Not _PlannedUnplanned.Contains("Planned") Then _PlannedUnplanned = "Unplanned"
            End If
            '  If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)) Then
            _DTGroup = BLANK_INDICATOR '""
            '  Else
            ' _DTGroup = .rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)
            ' End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyData(DownTimeColumn_OneClick.Product, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)
            End If
            _Comment = ""
        End With
    End Sub


    Private Sub initializeFields_Other()
        With parentLine
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)) Then
                _startTime = "" '"0" 'LG Code
            Else
                _startTime = .rawProficyData(DownTimeColumn_OneClick.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyData(DownTimeColumn_OneClick.EndTime, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DT, rowNum)) Then
                _DT = 0
            Else
                _DT = .rawProficyData(DownTimeColumn_OneClick.DT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.UT, rowNum)) Then
                _UT = 0 ' '"0" 'LG code
            Else
                _UT = .rawProficyData(DownTimeColumn_OneClick.UT, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Location, rowNum)) Then
                _Location = BLANK_INDICATOR '""
            Else
                _Location = .rawProficyData(DownTimeColumn_OneClick.Location, rowNum)
            End If
            'If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)) Then
            _MasterProdUnit = BLANK_INDICATOR '"" '"0" 'LG Code
            _StopClass = BLANK_INDICATOR '""
            _ProductGroup = BLANK_INDICATOR
            'Else
            ' _MasterProdUnit = .rawProficyData(DownTimeColumn_OneClick.MasterProdUnit, rowNum)
            ' End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)) Then
                _Fault = BLANK_INDICATOR '""
            Else
                _Fault = .rawProficyData(DownTimeColumn_OneClick.Fault, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)) Then
                _Reason1 = BLANK_INDICATOR '""
            Else
                _Reason1 = .rawProficyData(DownTimeColumn_OneClick.Reason1, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)) Then
                _Reason2 = BLANK_INDICATOR '""
            Else
                _Reason2 = .rawProficyData(DownTimeColumn_OneClick.Reason2, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)) Then
                _Reason3 = BLANK_INDICATOR '""
            Else
                _Reason3 = .rawProficyData(DownTimeColumn_OneClick.Reason3, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)) Then
                _Reason4 = BLANK_INDICATOR '""
            Else
                _Reason4 = .rawProficyData(DownTimeColumn_OneClick.Reason4, rowNum)
            End If

            If Not IsDBNull(.rawProficyData(10, rowNum)) Then
                If .rawProficyData(10, rowNum) = "S" Then
                    isSplitEvent = True
                End If
            End If


            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.PR_InOut, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyData(DownTimeColumn_OneClick.PR_InOut, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyData(DownTimeColumn_OneClick.Team, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.PlannedUnplanned, rowNum)) Then
                _PlannedUnplanned = BLANK_INDICATOR '""
            Else
                _PlannedUnplanned = .rawProficyData(DownTimeColumn_OneClick.PlannedUnplanned, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)) Then
                _DTGroup = BLANK_INDICATOR '""
            Else
                _DTGroup = .rawProficyData(DownTimeColumn_OneClick.DTGroup, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyData(DownTimeColumn_OneClick.Product, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyData(DownTimeColumn_OneClick.ProductCode, rowNum)
            End If
            If IsDBNull(.rawProficyData(DownTimeColumn_OneClick.Comment, rowNum)) Then
                _Comment = ""
            Else
                _Comment = .rawProficyData(DownTimeColumn_OneClick.Comment, rowNum)
            End If
        End With
    End Sub


#End Region

#Region "Other"
    Private Sub mapMyFormats()
        Select Case parentLine.formatMapping
            Case MappingByFormat.SkinCare
                getSkinCareFormatMapping(Me)
                _ProductCode = _Format
            Case MappingByFormat.NoMapping
            Case Else
                Throw New unknownMappingException(parentLine.formatMapping)
        End Select
    End Sub
    Private Sub mapMyShapes()
        Select Case parentLine.shapeMapping
            Case MappingByShape.SkinCare
                getSkinCareShapeMapping(Me)
            Case MappingByShape.NoMapping
            Case Else
                Throw New unknownMappingException(parentLine.shapeMapping)
        End Select
    End Sub
    Private Sub mapMyTiers()
        If _Reason1 = BLANK_INDICATOR Then _Reason1 = _Fault

        Select Case parentLine.prStoryMapping
            Case prStoryMapping.OralCare
                getOralCareprstoryMapping(Me)
            Case prStoryMapping.OralCareNau
                getOralCareNauprstoryMapping(Me)
            Case prStoryMapping.OralCareGross
                getOralCareGrossprstoryMapping(Me)
            Case prStoryMapping.Hyderabad
                getHyderabadprstoryMapping(Me)
            Case prStoryMapping.OralCare_DF
                getOralCareprstoryMappingDF(Me)
            Case prStoryMapping.OralCareNau_DF
                getOralCareNauprstoryMappingDF(Me)
            Case prStoryMapping.IowaCity
                getIowaCityprstoryMapping(Me)
            Case prStoryMapping.Phoenix
                getPhoenixprstoryMapping(Me)
            Case prStoryMapping.Pheonix_D
                getPheonix_DprstoryMapping(Me)
            Case prStoryMapping.SkinCare
                getSkinCareprstoryMapping(Me)
            Case prStoryMapping.SwingRoad
                getSwingRoadprstoryMapping(Me)
            Case prStoryMapping.SwingRoad_6
                getSwingRoadprstoryMapping_67(Me)
            Case prStoryMapping.SwingRoad_7
                getSwingRoadprstoryMapping_67(Me)
            Case prStoryMapping.Mandideep
                getMandideepprstoryMapping(Me)
            Case prStoryMapping.Mandideep_Fem
                getMandideepFemprstoryMapping(Me)
            Case prStoryMapping.GENERIC
                getGENERICprstoryMapping(Me)
            Case prStoryMapping.HuangPu
                getHuangpuprstoryMapping(Me)
            Case prStoryMapping.APDO_I
                getAPDOprstoryMapping_I(Me)
            Case prStoryMapping.APDO_J
                getAPDOprstoryMapping_J(Me)
            Case prStoryMapping.Boryspil
                getBoryspilprstoryMapping(Me)
            Case prStoryMapping.FemCare_Pads
                getBellevilleprstoryMapping(Me)
            Case prStoryMapping.FemCare_Pads_Huangpu
                getFemCare_Pads_HPMapping(Me)
            Case prStoryMapping.TepejiFem
                getTepejiFemprstoryMapping(Me)
            Case prStoryMapping.Mariscala
                getMariscalaprstoryMapping(Me)
            Case prStoryMapping.Puffs
                getPuffsMapping(Me)
            Case prStoryMapping.Albany
                getAlbanyprstoryMapping(Me)
            Case prStoryMapping.FamilyCareUnitOP_Wrapper
                getFamilyCareUnitOP_WrapperprstoryMapping(Me)
            Case prStoryMapping.FamilyCareUnitOP_ModPACK
                If _Tier1 = BLANK_INDICATOR Then
                    _isExcluded = True
                End If
                getFamilyCareUnitOP_ModPACK(Me)
            Case prStoryMapping.JijonaUltra
                getJijonaUltraprstoryMapping(Me)
            Case prStoryMapping.FamilyCareUnitOP_Napkins
                getFamilyCareUnitOP_Napkins(Me)
            Case prStoryMapping.FamilyCareUnitOP_mf
                getFamilyCareUnitOP_MFprstoryMapping(Me)
                If Not _isExcluded And _Tier1.Contains("Unscheduled") Then
                    _DT = 0
                End If

            Case prStoryMapping.FamilyCareUnitOP_Palletizer
                getFamilyCareUnitOP_PalletizerprstoryMapping(Me)
            Case prStoryMapping.Fem_LCC_HPU
                getFemLCCHPUMapping(Me)
            Case prStoryMapping.FamilyMaking
                getFamilyMakingprstoryMapping(Me)
            Case prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER
                ' getstraightprstorymapping(me)
                getFamilyCareUnitOP_Stretchwrapper(Me)
            Case prStoryMapping.FamilyCareUnitOP_Stacker
                '  getstraightprstorymapping(me)
                getFamilyCareUnitOP_Stacker(Me)
            Case prStoryMapping.BudapestFGC
                getBudapestFGCprstoryMapping(Me)
            Case prStoryMapping.BudapestLCC
                getBudapestLCCprstoryMapping(Me)
            Case prStoryMapping.NaucalpanPHC_B
                getNaucalpanPHC_BprstoryMapping(Me)
            Case prStoryMapping.NaucalpanPHC_J
                getNaucalpanPHC_JprstoryMapping(Me)
            Case prStoryMapping.NaucalpanPHC_Mex
                getNaucalpanPHC_MexprstoryMapping(Me)
            Case prStoryMapping.NaucalpanPHC_Vita1
                getNaucalpanPHC_Vita1prstoryMapping(Me)
            Case prStoryMapping.Rakona
                getRAKONAprstoryMapping(Me)
            Case prStoryMapping.SingaporePioneer
                getSingaporePioneerMapping(Me)
            Case prStoryMapping.ICOC
                getICOCprstoryMapping(Me)
            Case prStoryMapping.ICOC_Making
                getICOCMakingprstoryMapping(Me)
            Case prStoryMapping.Rio
                getRIOprstoryMapping(Me)
            Case prStoryMapping.Mariscala2
                getMariscalaprstoryMapping2(Me)
            Case prStoryMapping.STRAIGHT
                getSTRAIGHTprstoryMapping(Me)
            Case prStoryMapping.FamilyCareUnitOP_ACP
                getFamilyCareUnitOP_ACPprstoryMapping(Me)
            Case prStoryMapping.STRAIGHTPLANNEDPlusOne
                getSTRAIGHTPLANNEDPLUSONEprstoryMapping(Me)
            Case prStoryMapping.STRAIGHTPLANNEDPlusTwo
                getSTRAIGHTPLANNEDPLUSTWOprstoryMapping(Me)
            Case prStoryMapping.STRAIGHTPlusOne
                getSTRAIGHTPLUSONEprstoryMapping(Me)
            Case prStoryMapping.Fem_LuisCustom
                getFem_LuisCustomMapping(Me)
            Case prStoryMapping.OralCareCrux
                getOralCareCruxprstoryMapping(Me)
            Case prStoryMapping.NoMappingAvailable

            Case Else
                Throw New unknownMappingException(parentLine.prStoryMapping)
        End Select

        '  _tier1 = _reason1
        Try
            'this is where we handle 'post mapping'
            If My.Settings.PostMap_Enable Then
                ' Dim t As DowntimeField
                If My.Settings.PostMap_Field1 > -1 Then
                    '     t = My.Settings.PostMap_Field1
                    _Tier1 = getFieldFromInteger(My.Settings.PostMap_Field1)
                End If
                If My.Settings.PostMap_Field2 > -1 Then
                    '     t = My.Settings.PostMap_Field2
                    _Tier1 = getFieldFromInteger(My.Settings.PostMap_Field2)
                End If
                If My.Settings.PostMap_Field3 > -1 Then
                    '    t = My.Settings.PostMap_Field3
                    _Tier1 = getFieldFromInteger(My.Settings.PostMap_Field3)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Custom Mapping Error. Please Disable Custom Mapping And Rerun The Report. " & ex.Message)
        End Try
    End Sub

    'modify event based on a start and end time
    Public Sub adjustMyStartTime(startTime As Date)
        Dim timeDifference As Double
        timeDifference = DateDiff(DateInterval.Second, startTime, _startTime)
        Select Case timeDifference
            Case Is < 0 '_startTime is earlier: we need to cut the all the ut & and dt
                _UT = 0
                _DT = DateDiff(DateInterval.Second, startTime, _endTime) / 60
                _startTime = startTime
            Case 0
                ' do nothing
            Case Is > 0 ' _startTime is later: just cut the ut
                _UT = timeDifference / 60
            Case Else
                Throw New dateRangeException
        End Select

    End Sub
    Public Sub adjustMyEndTime(endTimeX As Date)

        Dim timeDifference As Double
        timeDifference = DateDiff(DateInterval.Second, endTimeX, _startTime)
        Select Case timeDifference
            Case Is < 0 '_startTime is earlier
                'ORIGINAL -> BUGGY!!! _DT = DateDiff(DateInterval.Second, endTimeX, _endTime) / 60
                _DT = (-1 * timeDifference) / 60
                _endTime = endTimeX
                'THE FOLLOWING CASES ARE UPTIME ONLY!!!!
            Case 0
                _endTime = endTimeX
                _DT = 0
                _Mapping = ""
                _isUnplanned = False
            Case Is > 0 ' _startTime is later: just cut the ut
                _endTime = endTimeX
                _DT = 0
                _Mapping = ""
                _isUnplanned = False
                _UT = DateDiff(DateInterval.Second, startTime_UT, endTimeX) / 60
            Case Else
                Throw New dateRangeException
        End Select
    End Sub
#End Region

#Region "Imlementation of Sortable and Equitable"
    'IMPLEMENTATION OF SORTABLE AND EQUITABLE
    Public Function CompareTo(ByVal Other As DowntimeEvent) As Integer Implements System.IComparable(Of DowntimeEvent).CompareTo
        If _isStandardSort Then
            Return Me._startTime.CompareTo(Other.startTime)
        Else
            Select Case _sortField
                Case DowntimeField.DT
                    Return Me.DT.CompareTo(Other.DT)
                Case DowntimeField.UT
                    Return Me.UT.CompareTo(Other.UT)
                Case DowntimeField.Reason1
                    Return Me.Reason1.CompareTo(Other.Reason1)
                Case DowntimeField.Reason2
                    Return Me.Reason2.CompareTo(Other.Reason2)
                Case DowntimeField.Reason3
                    Return Me.Reason3.CompareTo(Other.Reason3)
                Case DowntimeField.Reason4
                    Return Me.Reason4.CompareTo(Other.Reason4)
                Case DowntimeField.endTime
                    Return Me.endTime.CompareTo(Other.endTime)
                Case DowntimeField.Fault
                    Return Me.Fault.CompareTo(Other.Fault)
                Case DowntimeField.ProductCode
                    Return Me.ProductCode.CompareTo(Other.ProductCode)
                Case DowntimeField.ProductGroup
                    Return Me.ProductGroup.CompareTo(Other.ProductGroup)
                Case Else
                    Throw New Exception("Unknown Downtime Field. CLS.DataInterface")
            End Select
        End If
    End Function
    'equality - find index of
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As DowntimeEvent = TryCast(obj, DowntimeEvent)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As DowntimeEvent) As Boolean _
        Implements IEquatable(Of DowntimeEvent).Equals
        If other.startTime >= startTime_UT And other.startTime <= _endTime Then
            Return True
        ElseIf DateDiff(DateInterval.Second, other.startTime, startTime_UT) = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region
End Class

Public Class ProductionEvent
    Implements IComparable(Of ProductionEvent)
    Implements IEquatable(Of ProductionEvent)

#Region "Raw Data Mapping Constants"
    'proficy prod setup
    Private Enum ProductionColumn
        StartTime = 0
        EndTime = 1
        ProductionUnit = 2
        ProductCode = 3
        Product = 4
        ProductionStatus = 5
        Shift = 6
        Team = 7
        ActualUnits = 8
        ActualCases = 9
        AdjustedCases = 10
        AdjustedUnits = 11
        StatUnits = 12
        ActualRate = 13
        TargetRate = 14
        SchedTime = 15
        UnitsPerCase = 16
        StatConversion = 17
    End Enum


    Private Enum ProductionColumn_Maple
        StartTime = 0
        EndTime = 1
        ProductionUnit = 3
        ProductCode = 4
        Product = 5
        ProductionStatus = 6
        Shift = 10
        Team = 9
        ActualUnits = 11
        ActualCases = 12
        AdjustedCases = 14
        AdjustedUnits = 13
        StatUnits = 15
        ActualRate = 16
        TargetRate = 17
        SchedTime = 18
        UnitsPerCase = 19
        ' StatConversion = 17
    End Enum

    Public Enum ProductionColumn_Maple_New
        Start_Time = 0
        End_Time = 1
        Duration = 2
        Production_Unit = 3
        Product_Code = 4
        Product = 5
        LineStatus = 6 'PR In Out aka Line Status
        Line_State = 7
        Line_Substate = 8
        TEAM = 9
        SHIFT = 10
        Actual_Units = 11
        Actual_Cases = 12
        Adjusted_Units = 13
        Adjusted_Cases = 14
        Stat_Units = 15
        Actual_Rate = 16
        Target_Rate = 17
        Line_Scheduled_Time = 18
        Constraint_Scheduled_Time = 19
        Constraint_Uptime = 20
        Units_Per_Case = 21
        Number_Of_Constraints = 22
        Constraints_Running = 23
    End Enum




    Private Enum ProductionColumn_SwingRoad
        StartTime = 1
        EndTime = 2
        ProductionUnit = 0
        ProductCode = 16
        Product = 5
        ProductionStatus = 6
        Shift = 3
        Team = 4

        ' ActualUnits = 8
        ActualCases = 9
        AdjustedCases = 9
        '  AdjustedUnits = 9
        StatUnits = 15
        ActualRate = 12
        TargetRate = 11
        SchedTime = 8
        UnitsPerCase = 14
        '  StatConversion = 17
    End Enum
#End Region

    Public Overrides Function toString() As String
        Return "S/E: " & _startTime & "/" & _endTime & "      " & PR_inout & "     " & RateGainMinutes & "     " & ProductionMinutes
    End Function

#Region "Variables & Properties"
    'Private rawData As Array
    Private Property parentLine As ProdLine
    Friend rowNum As Long

    'Booleans
    Private _isExcluded As Boolean

    Public ReadOnly Property isExcluded As Boolean
        Get
            Return (_isExcluded Or _isFiltered)
        End Get
    End Property
    Public WriteOnly Property isFiltered As Boolean
        Set(value As Boolean)
            _isFiltered = value
        End Set
    End Property
    Private _isFiltered As Boolean = False

    Private _startTime As Date
    Private _endTime As Date
    Private _MasterProdUnit As String
    Private _Location As String

    Private _PR_inout As String
    Private _Team As String
    Private _Shift As String
    Private _Product As String
    Private _ProductCode As String
    Private _ProductGroup As String

    Private _AdjCases As Double
    Private _ActUnits As Double
    Private _ActCases As Double
    Private _AdjUnits As Double
    Private _StatUnits As Double
    Private _ActualRate As Double
    Private _TargetRate As Double
    Private _SchedTime As Double
    Private _UnitsPerCase As Integer
    Private _StatCaseConv As Double

    Public Constraint_Uptime As Double = -1 ' Maple
    Public Number_Of_Constraints As Integer = -1
    Public Number_Of_Constraints_Running As Integer = -1
    Public Line_Sched_Time As Double = -1
    Public Duration As Double = -1
    Public ReadOnly Property RateGainMinutes As Double
        Get
            If _TargetRate = 0 Then
                Return 0 'prevents div by zero error
                ' ElseIf (_ActualRate - TargetRate) < 0 Then 'this is a condition from MAPLE
                '     Return 0
            Else
                Return (_TargetRate - _ActualRate) * _UT / TargetRate
                '  Return (_ActualRate - _TargetRate) * Constraint_Uptime / TargetRate 'original. 5/15/17. st louis pr investigation
            End If
        End Get
    End Property
    Public ReadOnly Property ProductionMinutes As Double
        Get
            If _TargetRate = 0 Then
                Return 0 'prevents div by zero error
            Else
                Return _ActCases / _TargetRate
            End If
        End Get
    End Property


    Public ReadOnly Property ActUnits As Double
        Get
            Return _ActUnits
        End Get
    End Property
    Public ReadOnly Property ActCases As Double
        Get
            Return _ActCases
        End Get
    End Property
    Public ReadOnly Property AdjUnits As Double
        Get
            Return _AdjUnits
        End Get
    End Property
    Public ReadOnly Property AdjCases As Double
        Get
            Return _AdjCases
        End Get
    End Property
    Public ReadOnly Property StatUnits As Double
        Get
            Return _StatUnits
        End Get
    End Property
    Public ReadOnly Property ActualRate As Double
        Get
            Return _ActualRate
        End Get
    End Property
    Public ReadOnly Property TargetRate As Double
        Get
            Return _TargetRate
        End Get
    End Property
    Public ReadOnly Property SchedTime As Double
        Get
            If isExcluded Then Return 0
            Return _SchedTime
        End Get
    End Property
    Public ReadOnly Property UnitsPerCase As Integer
        Get
            Return _UnitsPerCase
        End Get
    End Property
    Public ReadOnly Property StatCaseConv As Double
        Get
            Return _StatCaseConv
        End Get
    End Property
    Public ReadOnly Property Shift As String
        Get
            Return _Shift
        End Get
    End Property


    'MAPPED DATA
    Private _PR As Double
    Private _UT As Double
    Private _Format As String = ""
    Private _Shape As String = ""

    Public Property startTime As Date
        Get
            Return _startTime
        End Get
        Set(value As Date)
            _startTime = value
        End Set
    End Property
    Public Property endTime As Date
        Get
            Return _endTime
        End Get
        Set(value As Date)
            _endTime = value
        End Set
    End Property

    Public ReadOnly Property startTime_24hr As String
        Get
            Return _startTime.ToString("MM/dd/yyyy HH:mm:ss")
        End Get
    End Property
    Public ReadOnly Property endTime_24hr As String
        Get
            Return _endTime.ToString("MM/dd/yyyy HH:mm:ss")
        End Get
    End Property



    Public Property Team As String
        Get
            Return _Team
        End Get
        Set(value As String)
            _Team = value
        End Set
    End Property
    Public ReadOnly Property PR_inout As String
        Get
            Return _PR_inout
        End Get
    End Property
    Public ReadOnly Property ProductCode As String
        Get
            Return _ProductCode
        End Get
    End Property
    Public ReadOnly Property ProductGroup As String
        Get
            Return _ProductGroup
        End Get
    End Property
    Public ReadOnly Property Product As String
        Get
            Return _Product
        End Get
    End Property
    Public ReadOnly Property MasterProductionUnit As String
        Get
            Return _MasterProdUnit
        End Get
    End Property
    Public ReadOnly Property PR As Double
        Get
            If isExcluded Then Return 0
            Return _PR
        End Get
    End Property
    Public ReadOnly Property PR_display As Double
        Get
            Return Math.Round(PR, 2)
        End Get
    End Property
    Public ReadOnly Property UT_display As Double
        Get
            Return Math.Round(UT, 2)
        End Get
    End Property
    Public ReadOnly Property UT As Double
        Get
            If isExcluded Then Return 0
            Return _UT
        End Get
    End Property
    Public ReadOnly Property Rate As Double
        Get
            '   If MAPLE_Rate = -1 Then
            Return Math.Max(_ActualRate, _TargetRate)
            '   Else
            '  Return MAPLE_Rate 
            ' End If
        End Get
    End Property

#End Region

#Region "Construction"

    Public Sub New(startDate As Date)
        _startTime = startDate
        _endTime = startDate
    End Sub

    Public Sub New(parentLineIn As ProdLine, row As Long)
        parentLine = parentLineIn
        rowNum = row
        Select Case parentLine.SQLproductionProcedure
            Case DefaultProficyProductionProcedure.QuickQuery
                initializeFields()
            Case DefaultProficyProductionProcedure.QuickQuery_MOT
                initializeFields()
            Case DefaultProficyProductionProcedure.Maple_New
                initializeFields_MAPLE_NEW() 'this maple new was added when the stored procedure changed
            Case DefaultProficyProductionProcedure.Maple
                initializeFields_MAPLE()
            Case DefaultProficyProductionProcedure.SwingRoad
                initializeFields_SwingRoad()
            Case Else
                Throw New unknownMappingException
        End Select

        If isPRODExcluded() Then
            _isExcluded = True
            _UT = 0
            _PR = 0
        Else
            _isExcluded = False
            If Rate > 0 And _SchedTime > 0 Then
                If _AdjUnits > 0 Then
                    _UT = _AdjUnits / Rate
                ElseIf _ActUnits > 0 Then 'ADDED FOR MAPLE
                    _UT = _ActUnits / Rate
                End If
                _PR = _UT / _SchedTime
            Else
                _PR = 0
                _UT = 0
            End If
        End If


        ' mapMyFormats()
        ' mapMyShapes()

    End Sub
#End Region

#Region "isExcluded / isUnplanned / isPlanned / CO, CIL, etc"


    Private Function isPRODExcluded() As Boolean

        If My.Settings.EnableTimeSpanExclusion Then
            If _startTime.Hour < My.Settings.Exclude_StartHour Or _startTime.Hour > My.Settings.Exclude_EndHour Then
                Return True
            ElseIf _startTime.Hour = My.Settings.Exclude_StartHour Then
                If _startTime.Minute < My.Settings.Exclude_StartMinutes Then
                    Return True
                End If
            ElseIf _startTime.Hour = My.Settings.Exclude_EndHour Then
                If _startTime.Minute > My.Settings.Exclude_EndMinutes Then
                    Return True
                End If
            End If

        End If


        Select Case parentLine.Mapping_DTschedPlannedUnplanned
            Case DTsched_Mapping.Greensboro
                Return isDTExcluded_GBO()
            Case DTsched_Mapping.SkinCare
                Return isDTExcluded_GBO()
            Case DTsched_Mapping.Phenoix
                Return isDTExcluded_PHX()
            Case DTsched_Mapping.HuangpuHC
                Return isDTExcluded_GBO()
            Case DTsched_Mapping.SwingRoad
                Return isDTExcluded_SwingRoad()
            Case DTsched_Mapping.Mariscala
                Return isDTExcluded_GBO()
            Case DTsched_Mapping.Hyderabad
                Return isDTExcluded_GBO()
            Case DTsched_Mapping.Rakona
                Return isDTExcluded_GBO()
            Case Else
                Throw New unknownMappingException
        End Select
    End Function

#Region "SwingRoad"
    'swing road
    Public Function isDTExcluded_SwingRoad() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 5) = "OEE O" Or Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If

        Return False

    End Function

#End Region
#Region "PHX"

    Public Function isDTExcluded_PHX() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If

        Return False

    End Function

#End Region
#Region "GBO"
    'GREENSBORO

    Public Function isDTExcluded_GBO() As Boolean
        If _PR_inout.Equals(BLANK_INDICATOR) Then Return False

        If Left(_PR_inout, 4) = "PR O" Then
            Return True
        ElseIf Left(_ProductCode, 1) = "S" Then
            Return True
        End If

        Return False
    End Function

#End Region
#End Region

#Region "Initialize From Raw Data Array"
    'initialization
    Private Sub initializeFields()
        With parentLine
            If IsDBNull(.rawProficyProductionData(ProductionColumn.StartTime, rowNum)) Then
                _startTime = ""
            Else
                _startTime = .rawProficyProductionData(ProductionColumn.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn.EndTime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyProductionData(ProductionColumn.EndTime, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.ProductionUnit, rowNum)) Then
                _MasterProdUnit = BLANK_INDICATOR   '"0" 'LG Code
            Else
                _MasterProdUnit = .rawProficyProductionData(ProductionColumn.ProductionUnit, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyProductionData(ProductionColumn.Product, rowNum)
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyProductionData(ProductionColumn.ProductCode, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.ProductionStatus, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyProductionData(ProductionColumn.ProductionStatus, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyProductionData(ProductionColumn.Team, rowNum)
                If _Team = "" Then
                    _Team = BLANK_INDICATOR
                End If
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn.Shift, rowNum)) Then
                _Shift = BLANK_INDICATOR '""
            Else
                _Shift = .rawProficyProductionData(ProductionColumn.Shift, rowNum)
                If _Shift = "" Then
                    _Shift = BLANK_INDICATOR
                End If
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.ActualUnits, rowNum)) Then
                _ActUnits = 0
            Else
                _ActUnits = .rawProficyProductionData(ProductionColumn.ActualUnits, rowNum)

            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.ActualCases, rowNum)) Then
                _ActCases = 0
            Else
                _ActCases = .rawProficyProductionData(ProductionColumn.ActualCases, rowNum)
                '    If _ActCases = "" Then
                ' _ActCases = 0
                'End If
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn.AdjustedUnits, rowNum)) Then
                _AdjUnits = 0
            Else
                _AdjUnits = .rawProficyProductionData(ProductionColumn.AdjustedUnits, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.AdjustedCases, rowNum)) Then
                _AdjCases = 0
            Else
                _AdjCases = .rawProficyProductionData(ProductionColumn.AdjustedCases, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.StatUnits, rowNum)) Then
                _StatUnits = 0 '""
            Else
                _StatUnits = .rawProficyProductionData(ProductionColumn.StatUnits, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.ActualRate, rowNum)) Then
                _ActualRate = 0
            Else
                _ActualRate = .rawProficyProductionData(ProductionColumn.ActualRate, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.TargetRate, rowNum)) Then
                _TargetRate = 0
            Else
                _TargetRate = .rawProficyProductionData(ProductionColumn.TargetRate, rowNum)
                '    If _TargetRate = "" Then
                ' _TargetRate = 0
                'End If
            End If


            If IsDBNull(.rawProficyProductionData(ProductionColumn.UnitsPerCase, rowNum)) Then
                _UnitsPerCase = 0
            Else
                _UnitsPerCase = .rawProficyProductionData(ProductionColumn.UnitsPerCase, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.StatConversion, rowNum)) Then
                _UnitsPerCase = 0
            Else
                _UnitsPerCase = .rawProficyProductionData(ProductionColumn.StatConversion, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn.SchedTime, rowNum)) Then
                _SchedTime = 0
            Else
                _SchedTime = .rawProficyProductionData(ProductionColumn.SchedTime, rowNum)
            End If

        End With
    End Sub

    Private Sub initializeFields_SwingRoad()
        With parentLine
            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.StartTime, rowNum)) Then
                _startTime = ""
            Else
                _startTime = .rawProficyProductionData(ProductionColumn_SwingRoad.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.EndTime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyProductionData(ProductionColumn_SwingRoad.EndTime, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.ProductionUnit, rowNum)) Then
                _MasterProdUnit = BLANK_INDICATOR   '"0" 'LG Code
            Else
                _MasterProdUnit = .rawProficyProductionData(ProductionColumn_SwingRoad.ProductionUnit, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyProductionData(ProductionColumn_SwingRoad.Product, rowNum)
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyProductionData(ProductionColumn_SwingRoad.ProductCode, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.ProductionStatus, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyProductionData(ProductionColumn_SwingRoad.ProductionStatus, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyProductionData(ProductionColumn_SwingRoad.Team, rowNum)
                If _Team = "" Then
                    _Team = BLANK_INDICATOR
                End If
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.Shift, rowNum)) Then
                _Shift = BLANK_INDICATOR '""
            Else
                _Shift = .rawProficyProductionData(ProductionColumn_SwingRoad.Shift, rowNum)
                If _Shift = "" Then
                    _Shift = BLANK_INDICATOR
                End If
            End If

            ' If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.ActualUnits, rowNum)) Then
            _ActUnits = 0
            ' Else
            '     _ActUnits = .rawProficyProductionData(ProductionColumn_SwingRoad.ActualUnits, rowNum)

            ' End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.ActualCases, rowNum)) Then
                _ActCases = 0
            Else
                _ActCases = .rawProficyProductionData(ProductionColumn_SwingRoad.ActualCases, rowNum)
                '    If _ActCases = "" Then
                ' _ActCases = 0
                'End If
            End If
            '   If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.AdjustedUnits, rowNum)) Then
            ' _AdjUnits = 0
            '  Else
            '  _AdjUnits = .rawProficyProductionData(ProductionColumn_SwingRoad.AdjustedUnits, rowNum)
            '  End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.AdjustedCases, rowNum)) Then
                _AdjCases = 0
            Else
                _AdjCases = .rawProficyProductionData(ProductionColumn_SwingRoad.AdjustedCases, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.StatUnits, rowNum)) Then
                _StatUnits = 0 '""
            Else
                _StatUnits = .rawProficyProductionData(ProductionColumn_SwingRoad.StatUnits, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.ActualRate, rowNum)) Then
                _ActualRate = 0
            Else
                _ActualRate = .rawProficyProductionData(ProductionColumn_SwingRoad.ActualRate, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.TargetRate, rowNum)) Then
                _TargetRate = 0
            Else
                _TargetRate = .rawProficyProductionData(ProductionColumn_SwingRoad.TargetRate, rowNum)
                '    If _TargetRate = "" Then
                ' _TargetRate = 0
                'End If
            End If


            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.UnitsPerCase, rowNum)) Then
                _UnitsPerCase = 0
            Else
                _UnitsPerCase = .rawProficyProductionData(ProductionColumn_SwingRoad.UnitsPerCase, rowNum)
            End If

            '  If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.StatConversion, rowNum)) Then
            '_UnitsPerCase = 0
            ' Else
            ' _UnitsPerCase = .rawProficyProductionData(ProductionColumn_SwingRoad.StatConversion, rowNum)
            ' End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_SwingRoad.SchedTime, rowNum)) Then
                _SchedTime = 0
            Else
                _SchedTime = .rawProficyProductionData(ProductionColumn_SwingRoad.SchedTime, rowNum)
            End If

            _AdjUnits = _UnitsPerCase * _ActCases
        End With
    End Sub

    Private Sub initializeFields_MAPLE_NEW()
        With parentLine
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Start_Time, rowNum)) Then
                _startTime = ""
            Else
                _startTime = .rawProficyProductionData(ProductionColumn_Maple_New.Start_Time, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.End_Time, rowNum)) Then
                _endTime = _startTime
            Else
                _endTime = .rawProficyProductionData(ProductionColumn_Maple_New.End_Time, rowNum)
            End If

            'DURATION
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Duration, rowNum)) Then
                Duration = 0
            Else
                Duration = .rawProficyProductionData(ProductionColumn_Maple_New.Duration, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Production_Unit, rowNum)) Then
                _MasterProdUnit = BLANK_INDICATOR
            Else
                _MasterProdUnit = .rawProficyProductionData(ProductionColumn_Maple_New.Production_Unit, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Product_Code, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyProductionData(ProductionColumn_Maple_New.Product_Code, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Product, rowNum)) Then
                _Product = BLANK_INDICATOR
            Else
                _Product = .rawProficyProductionData(ProductionColumn_Maple_New.Product, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.LineStatus, rowNum)) Then
                _PR_inout = BLANK_INDICATOR
            Else
                _PR_inout = .rawProficyProductionData(ProductionColumn_Maple_New.LineStatus, rowNum)
            End If

            'LINE SUBSTATE

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.TEAM, rowNum)) Then
                _Team = BLANK_INDICATOR
            Else
                _Team = .rawProficyProductionData(ProductionColumn_Maple_New.TEAM, rowNum)
                If _Team = "" Then
                    _Team = BLANK_INDICATOR
                End If
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.SHIFT, rowNum)) Then
                _Shift = BLANK_INDICATOR
            Else
                _Shift = .rawProficyProductionData(ProductionColumn_Maple_New.SHIFT, rowNum)
                If _Shift = "" Then
                    _Shift = BLANK_INDICATOR
                End If
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Actual_Units, rowNum)) Then
                _ActUnits = 0
            Else
                _ActUnits = .rawProficyProductionData(ProductionColumn_Maple_New.Actual_Units, rowNum)

            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Actual_Cases, rowNum)) Then
                _ActCases = 0
            Else
                _ActCases = .rawProficyProductionData(ProductionColumn_Maple_New.Actual_Cases, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Adjusted_Units, rowNum)) Then
                _AdjUnits = 0
            Else
                _AdjUnits = .rawProficyProductionData(ProductionColumn_Maple_New.Adjusted_Units, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Adjusted_Cases, rowNum)) Then
                _AdjCases = 0
            Else
                _AdjCases = .rawProficyProductionData(ProductionColumn_Maple_New.Adjusted_Cases, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Stat_Units, rowNum)) Then
                _StatUnits = 0
            Else
                _StatUnits = .rawProficyProductionData(ProductionColumn_Maple_New.Stat_Units, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Actual_Rate, rowNum)) Then
                _ActualRate = 0
            Else
                _ActualRate = .rawProficyProductionData(ProductionColumn_Maple_New.Actual_Rate, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Target_Rate, rowNum)) Then
                _TargetRate = 0
            Else
                _TargetRate = .rawProficyProductionData(ProductionColumn_Maple_New.Target_Rate, rowNum)
            End If

            'LINE SCHED TIME
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Line_Scheduled_Time, rowNum)) Then
                Line_Sched_Time = 0
            Else
                Line_Sched_Time = .rawProficyProductionData(ProductionColumn_Maple_New.Line_Scheduled_Time, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Constraint_Scheduled_Time, rowNum)) Then
                _SchedTime = 0
            Else
                _SchedTime = .rawProficyProductionData(ProductionColumn_Maple_New.Constraint_Scheduled_Time, rowNum)
            End If

            'CONSTRAINT UPTIME - duration when pr in = true - constraint downtime
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Constraint_Uptime, rowNum)) Then
                Constraint_Uptime = 0
            Else
                Constraint_Uptime = .rawProficyProductionData(ProductionColumn_Maple_New.Constraint_Uptime, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Units_Per_Case, rowNum)) Then
                _UnitsPerCase = 0
            Else
                _UnitsPerCase = .rawProficyProductionData(ProductionColumn_Maple_New.Units_Per_Case, rowNum)
            End If

            'NUMBER OF CONSTRAINTS
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Number_Of_Constraints, rowNum)) Then
                Number_Of_Constraints = 0
            Else
                Number_Of_Constraints = .rawProficyProductionData(ProductionColumn_Maple_New.Number_Of_Constraints, rowNum)
            End If
            Try
                'CONSTRAINTS RUNNING
                If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple_New.Constraints_Running, rowNum)) Then
                    Number_Of_Constraints_Running = 0
                Else
                    Number_Of_Constraints_Running = .rawProficyProductionData(ProductionColumn_Maple_New.Constraints_Running, rowNum)
                End If
            Catch
                Number_Of_Constraints_Running = 0
            End Try

        End With
    End Sub

    Private Sub initializeFields_MAPLE()
        With parentLine
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.StartTime, rowNum)) Then
                _startTime = ""
            Else
                _startTime = .rawProficyProductionData(ProductionColumn_Maple.StartTime, rowNum)
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.EndTime, rowNum)) Then
                _endTime = _startTime  '"0" 'LG Code
            Else
                _endTime = .rawProficyProductionData(ProductionColumn_Maple.EndTime, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.ProductionUnit, rowNum)) Then
                _MasterProdUnit = BLANK_INDICATOR   '"0" 'LG Code
            Else
                _MasterProdUnit = .rawProficyProductionData(ProductionColumn_Maple.ProductionUnit, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.Product, rowNum)) Then
                _Product = BLANK_INDICATOR '""
            Else
                _Product = .rawProficyProductionData(ProductionColumn_Maple.Product, rowNum)
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.ProductCode, rowNum)) Then
                _ProductCode = BLANK_INDICATOR '""
            Else
                _ProductCode = .rawProficyProductionData(ProductionColumn_Maple.ProductCode, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.ProductionStatus, rowNum)) Then
                _PR_inout = BLANK_INDICATOR '""
            Else
                _PR_inout = .rawProficyProductionData(ProductionColumn_Maple.ProductionStatus, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.Team, rowNum)) Then
                _Team = BLANK_INDICATOR '""
            Else
                _Team = .rawProficyProductionData(ProductionColumn_Maple.Team, rowNum)
                If _Team = "" Then
                    _Team = BLANK_INDICATOR
                End If
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.Shift, rowNum)) Then
                _Shift = BLANK_INDICATOR '""
            Else
                _Shift = .rawProficyProductionData(ProductionColumn_Maple.Shift, rowNum)
                If _Shift = "" Then
                    _Shift = BLANK_INDICATOR
                End If
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.ActualUnits, rowNum)) Then
                _ActUnits = 0
            Else
                _ActUnits = .rawProficyProductionData(ProductionColumn_Maple.ActualUnits, rowNum)

            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.ActualCases, rowNum)) Then
                _ActCases = 0
            Else
                _ActCases = .rawProficyProductionData(ProductionColumn_Maple.ActualCases, rowNum)
                '    If _ActCases = "" Then
                ' _ActCases = 0
                'End If
            End If
            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.AdjustedUnits, rowNum)) Then
                _AdjUnits = 0
            Else
                _AdjUnits = .rawProficyProductionData(ProductionColumn_Maple.AdjustedUnits, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.AdjustedCases, rowNum)) Then
                _AdjCases = 0
            Else
                _AdjCases = .rawProficyProductionData(ProductionColumn_Maple.AdjustedCases, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.StatUnits, rowNum)) Then
                _StatUnits = 0 '""
            Else
                _StatUnits = .rawProficyProductionData(ProductionColumn_Maple.StatUnits, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.ActualRate, rowNum)) Then
                _ActualRate = 0
            Else
                _ActualRate = .rawProficyProductionData(ProductionColumn_Maple.ActualRate, rowNum)
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.TargetRate, rowNum)) Then
                _TargetRate = 0
            Else
                _TargetRate = .rawProficyProductionData(ProductionColumn_Maple.TargetRate, rowNum)
                '    If _TargetRate = "" Then
                ' _TargetRate = 0
                'End If
            End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.UnitsPerCase, rowNum)) Then
                _UnitsPerCase = 0
            Else
                _UnitsPerCase = .rawProficyProductionData(ProductionColumn_Maple.UnitsPerCase, rowNum)
            End If

            '   If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.StatConversion, rowNum)) Then
            '   _UnitsPerCase = 0
            '   Else
            '   _UnitsPerCase = .rawProficyProductionData(ProductionColumn_Maple.StatConversion, rowNum)
            '  End If

            If IsDBNull(.rawProficyProductionData(ProductionColumn_Maple.SchedTime, rowNum)) Then
                _SchedTime = 0
            Else
                _SchedTime = .rawProficyProductionData(ProductionColumn_Maple.SchedTime, rowNum)
            End If
        End With
    End Sub

#End Region

#Region "Other"

    Private Sub mapMyShapes()
        Select Case parentLine.shapeMapping
            Case MappingByShape.SkinCare
                '   getSkinCareShapeMapping(Me)
            Case MappingByShape.NoMapping
            Case Else
                Throw New unknownMappingException(parentLine.shapeMapping)
        End Select
    End Sub




#End Region

#Region "Imlementation of Sortable and Equitable"
    'IMPLEMENTATION OF SORTABLE AND EQUITABLE
    Public Function CompareTo(ByVal Other As ProductionEvent) As Integer Implements System.IComparable(Of ProductionEvent).CompareTo
        Return Me._startTime.CompareTo(Other.startTime)
    End Function
    'equality - find index of
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As ProductionEvent = TryCast(obj, ProductionEvent)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As ProductionEvent) As Boolean _
        Implements IEquatable(Of ProductionEvent).Equals
        If other.startTime >= startTime And other.startTime <= _endTime Then
            Return True
        ElseIf DateDiff(DateInterval.Second, other.startTime, startTime) = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region
End Class

Public Class RateLossEvent
    Implements IComparable(Of RateLossEvent)
    Implements IEquatable(Of RateLossEvent)

#Region "Raw Data Mapping Constants"
    'proficy downtime setup
    Private Enum RateLossColumn
        StartTime = 0
        Endtime = 1
        DT = 2
        UT = 3
        MasterProdUnit = 4
        Fault = 5
        Reason1 = 6
        Reason2 = 7
        Reason3 = 8
        Reason4 = 9
        Rate_Target = 11 ' LG Code
        Rate_Actual = 12
        Team = 13
        Shift = 14
        ProductGroup = 15
        SKU = 16
        Location = 17

        PlannedUnplanned = 20

        PR_InOut = 21
        Comment = 22

    End Enum

#End Region


    Public Function getFieldFromInteger(TargetField As Integer) As String
        Select Case TargetField
            Case DowntimeField.Location
                Return _Location
            Case DowntimeField.Fault
                Return _Fault
            Case DowntimeField.Reason1
                Return _Reason1
            Case DowntimeField.Reason2
                Return _Reason2
            Case DowntimeField.Reason3
                Return _Reason3
            Case DowntimeField.Reason4
                Return _Reason4
                ' Case DowntimeField.PR_inout
            Case DowntimeField.Team
                Return _Team
                ' Case DowntimeField.PlannedUnplanned

                ' Case DowntimeField.Product
            Case DowntimeField.ProductCode
                Return _ProductCode
                ' Case DowntimeField.Comment 'NOPE
                ' 'MAPPED DATA
            Case DowntimeField.Tier1
                Return _Tier1
            Case DowntimeField.Tier2
                Return _Tier2
            Case DowntimeField.Tier3
                Return _Tier3

            Case DowntimeField.ProductGroup
                Return _ProductGroup
            Case Else
                Throw New unknownMappingException
        End Select
    End Function


    Public Overrides Function toString() As String
        Return "S/E: " & _startTime & "/" & _endTime & "  DT/UT:  " & _DT & "/" & _UT
    End Function


#Region "Variables & Properties"

    Public rawMasterData As Object(,)

    Friend _isExcluded As Boolean = False
    Friend _isFiltered As Boolean = False

    Public ReadOnly Property isExcluded As Boolean
        Get
            Return (_isExcluded Or _isFiltered)
        End Get
    End Property



    'RAW DATA
    Friend _startTime As Date
    Friend _endTime As Date
    Friend _DT As Double
    Friend _UT As Double
    Friend _MasterProdUnit As String
    Friend _Location As String
    Friend _Fault As String
    Friend _Reason1 As String
    Friend _Reason2 As String
    Friend _Reason3 As String
    Friend _Reason4 As String

    Public ReadOnly Property TargetRate_Display As Double
        Get
            Return TargetRate
        End Get
    End Property
    Public ReadOnly Property ActualRate_Display As Double
        Get
            Return ActualRate
        End Get
    End Property


    Friend TargetRate As Double
    Friend ActualRate As Double
    Friend _PR_inout As String
    Friend _Team As String
    Friend _PlannedUnplanned As String

    ' Friend _Product As String
    Friend _ProductCode As String
    Friend _ProductGroup As String
    Friend _Comment As String

    'MAPPED DATA
    Friend _Tier1 As String = BLANK_INDICATOR
    Friend _Tier2 As String = BLANK_INDICATOR
    Friend _Tier3 As String = BLANK_INDICATOR

    Private isCrux As Boolean = False

    'date / time fields
    Public ReadOnly Property startTime_UT As Date
        Get
            Return DateAdd(DateInterval.Second, -60 * _UT, _startTime)
        End Get
    End Property

    Public ReadOnly Property RatePCT As Double
        Get
            If TargetRate = 0 Then Return 0
            Return Math.Round(ActualRate * 100 / TargetRate)
        End Get
    End Property

    Public Property startTime As Date
        Get
            Return _startTime
        End Get
        Set(value As Date)
            _startTime = value
        End Set
    End Property
    Public Property endTime As Date
        Get
            Return _endTime
        End Get
        Set(value As Date)
            _endTime = value
        End Set
    End Property

    Public ReadOnly Property startTime_24hr As String
        Get
            Return _startTime.ToString("MM/dd/yyyy HH:mm:ss")
        End Get
    End Property
    Public ReadOnly Property endTime_24hr As String
        Get
            Return _endTime.ToString("MM/dd/yyyy HH:mm:ss")
        End Get
    End Property


    Public Property DT As Double
        Get
            Return _DT
        End Get
        Set(value As Double)
            _DT = value
        End Set
    End Property
    Public ReadOnly Property DT_display As Double
        Get
            Return Math.Round(_DT, 2)
        End Get
    End Property
    Public ReadOnly Property UT_display As Double
        Get
            Return Math.Round(_UT, 2)
        End Get
    End Property
    Public ReadOnly Property UT As Double
        Get
            Return _UT
        End Get
    End Property

    'top level leds fields
    Public ReadOnly Property Location As String
        Get
            Return _Location
        End Get
    End Property
    Public ReadOnly Property Fault As String
        Get
            Return _Fault
        End Get
    End Property

    'tree level fields
    Public ReadOnly Property Reason1 As String
        Get
            Return _Reason1
        End Get
    End Property
    Public ReadOnly Property Reason2 As String
        Get
            Return _Reason2
        End Get
    End Property
    Public ReadOnly Property Reason3 As String
        Get
            Return _Reason3
        End Get
    End Property
    Public ReadOnly Property Reason4 As String
        Get
            Return _Reason4
        End Get
    End Property

    'comment
    Public ReadOnly Property Comment As String
        Get
            Return _Comment
        End Get
    End Property

    Public ReadOnly Property PlannedUnplanned As String
        Get
            Return _PlannedUnplanned
        End Get
    End Property
    Public Property Team As String
        Get
            Return _Team
        End Get
        Set(value As String)
            _Team = value
        End Set
    End Property
    Public ReadOnly Property PR_inout As String
        Get
            Return _PR_inout
        End Get
    End Property
    Public Property ProductCode As String
        Get
            Return _ProductCode
        End Get
        Set(value As String)
            _ProductCode = value
        End Set
    End Property
    Public ReadOnly Property ProductGroup As String
        Get
            Return _ProductGroup
        End Get
    End Property

    Public ReadOnly Property MasterProductionUnit As String
        Get
            Return _MasterProdUnit
        End Get
    End Property

    '''''''THESE ARE MAPPED DATA FIELDS
    Public Property Tier1 As String
        Get
            Return _Tier1
        End Get
        Set(value As String)
            _Tier1 = value
        End Set
    End Property
    Public Property Tier2 As String
        Get
            Return _Tier2
        End Get
        Set(value As String)
            _Tier2 = value
        End Set
    End Property
    Public Property Tier3 As String
        Get
            Return _Tier3
        End Get
        Set(value As String)
            _Tier3 = value
        End Set
    End Property

#End Region

    Friend Sub mapEvent1()
        If isCrux Then
            _Tier1 = _Reason1
            _Tier2 = _Reason2
            _Tier3 = _Reason3
        Else
            If _Fault.Contains("Parent Roll") Or _Reason2.Contains("Parent Roll") Then
                _Tier1 = "Parent Roll Change"
            ElseIf _Reason1.Contains("Speed") Or _Reason1.Contains("Blocked") Then
                _Tier1 = "Speed Bottleneck"
            ElseIf _Reason1.Contains("Operational") Or _Reason1.Contains("Process") Then
                _Tier1 = "Operational"
            ElseIf _Reason1.Contains("Starved") Then
                _Tier1 = "Starved"
            Else
                _Tier1 = "Uncoded Rateloss"
            End If

            mapEvent2()
        End If



    End Sub

    Private Sub mapEvent2()
        If Not isCrux Then
            If RatePCT < 25 Then
                _Tier2 = "Less Than 25% Rate"
            ElseIf RatePCT <= 50 Then
                _Tier2 = "Between 25%-50% Rate"
            ElseIf RatePCT <= 75 Then
                _Tier2 = "Betwenn 50%-75% Rate"
            Else
                _Tier2 = "Greater Than 75% Rate"
            End If
        End If
    End Sub

    Public Function getSingleEventFromMaster(index As Integer) As RateLossEvent
        If AllProdLines(lineIndex).parentSite.Name = SITE_CRUX Then
            isCrux = True
        End If
        Dim tmpEvent As New RateLossEvent()
        With tmpEvent
            If isCrux Then
                If IsDBNull(rawMasterData(DowntimeField.startTime, index)) Then
                    .startTime = ""
                Else
                    .startTime = rawMasterData(DowntimeField.startTime, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.endTime, index)) Then
                    .endTime = .startTime  '"0" 'LG Code
                Else
                    .endTime = rawMasterData(DowntimeField.endTime, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.DT, index)) Then
                    .DT = 0
                Else
                    .DT = rawMasterData(DowntimeField.DT, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.UT, index)) Then
                    ._UT = 0
                Else
                    ._UT = rawMasterData(DowntimeField.UT, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.Location, index)) Then
                    ._Location = BLANK_INDICATOR
                Else
                    ._Location = rawMasterData(DowntimeField.Location, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.MasterProdUnit, index)) Then
                    ._MasterProdUnit = BLANK_INDICATOR
                Else
                    ._MasterProdUnit = rawMasterData(DowntimeField.MasterProdUnit, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.Fault, index)) Then
                    ._Fault = BLANK_INDICATOR '""
                Else
                    ._Fault = rawMasterData(DowntimeField.Fault, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.Reason1, index)) Then
                    ._Reason1 = BLANK_INDICATOR '""
                Else
                    ._Reason1 = rawMasterData(DowntimeField.Reason1, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.Reason2, index)) Then
                    ._Reason2 = BLANK_INDICATOR '""
                Else
                    ._Reason2 = rawMasterData(DowntimeField.Reason2, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.Reason3, index)) Then
                    ._Reason3 = BLANK_INDICATOR '""
                Else
                    ._Reason3 = rawMasterData(DowntimeField.Reason3, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.Reason4, index)) Then
                    ._Reason4 = BLANK_INDICATOR '""
                Else
                    ._Reason4 = rawMasterData(DowntimeField.Reason4, index)
                End If


                If IsDBNull(rawMasterData(DowntimeField.PR_inout, index)) Then
                    ._PR_inout = BLANK_INDICATOR '""
                Else
                    ._PR_inout = rawMasterData(DowntimeField.PR_inout, index)
                End If
                If IsDBNull(rawMasterData(DowntimeField.Team, index)) Then
                    .Team = BLANK_INDICATOR '""
                Else
                    .Team = rawMasterData(DowntimeField.Team, index)
                    If .Team = "" Then
                        .Team = BLANK_INDICATOR
                    End If
                End If
                If IsDBNull(rawMasterData(DowntimeField.PlannedUnplanned, index)) Then
                    ._PlannedUnplanned = BLANK_INDICATOR '""
                Else
                    ._PlannedUnplanned = rawMasterData(DowntimeField.PlannedUnplanned, index)
                End If

                'If IsDBNull(rawMasterData(downtimefield.SKU, index)) Then
                ._ProductCode = BLANK_INDICATOR '""
                '  Else
                '     ._ProductCode = rawMasterData(downtimefield.SKU, index)
                ' End If
                'If IsDBNull(rawMasterData(downtimefield.Rate_Actual, index)) Then
                .ActualRate = 0 '""
                'Else
                '   .ActualRate = rawMasterData(downtimefield.Rate_Actual, index)
                'End If
                ' If IsDBNull(rawMasterData(downtimefield.Rate_Target, index)) Then
                .TargetRate = 0 '""
                '  Else
                '       .TargetRate = rawMasterData(downtimefield.Rate_Target, index)
                '    End If
                If IsDBNull(rawMasterData(DowntimeField.ProductGroup, index)) Then
                    ._ProductGroup = BLANK_INDICATOR '""
                Else
                    ._ProductGroup = rawMasterData(DowntimeField.ProductGroup, index)
                End If
                Try
                    If IsDBNull(rawMasterData(DowntimeField.Comment, index)) Then
                        ._Comment = ""
                    Else
                        ._Comment = rawMasterData(DowntimeField.Comment, index)
                    End If
                Catch e As Exception
                    ._Comment = ""
                End Try

            Else


                If IsDBNull(rawMasterData(RateLossColumn.StartTime, index)) Then
                    .startTime = ""
                Else
                    .startTime = rawMasterData(RateLossColumn.StartTime, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Endtime, index)) Then
                    .endTime = .startTime  '"0" 'LG Code
                Else
                    .endTime = rawMasterData(RateLossColumn.Endtime, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.DT, index)) Then
                    .DT = 0
                Else
                    .DT = rawMasterData(RateLossColumn.DT, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.UT, index)) Then
                    ._UT = 0
                Else
                    ._UT = rawMasterData(RateLossColumn.UT, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Location, index)) Then
                    ._Location = BLANK_INDICATOR
                Else
                    ._Location = rawMasterData(RateLossColumn.Location, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.MasterProdUnit, index)) Then
                    ._MasterProdUnit = BLANK_INDICATOR
                Else
                    ._MasterProdUnit = rawMasterData(RateLossColumn.MasterProdUnit, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Fault, index)) Then
                    ._Fault = BLANK_INDICATOR '""
                Else
                    ._Fault = rawMasterData(RateLossColumn.Fault, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Reason1, index)) Then
                    ._Reason1 = BLANK_INDICATOR '""
                Else
                    ._Reason1 = rawMasterData(RateLossColumn.Reason1, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Reason2, index)) Then
                    ._Reason2 = BLANK_INDICATOR '""
                Else
                    ._Reason2 = rawMasterData(RateLossColumn.Reason2, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Reason3, index)) Then
                    ._Reason3 = BLANK_INDICATOR '""
                Else
                    ._Reason3 = rawMasterData(RateLossColumn.Reason3, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Reason4, index)) Then
                    ._Reason4 = BLANK_INDICATOR '""
                Else
                    ._Reason4 = rawMasterData(RateLossColumn.Reason4, index)
                End If


                If IsDBNull(rawMasterData(RateLossColumn.PR_InOut, index)) Then
                    ._PR_inout = BLANK_INDICATOR '""
                Else
                    ._PR_inout = rawMasterData(RateLossColumn.PR_InOut, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Team, index)) Then
                    .Team = BLANK_INDICATOR '""
                Else
                    .Team = rawMasterData(RateLossColumn.Team, index)
                    If .Team = "" Then
                        .Team = BLANK_INDICATOR
                    End If
                End If
                If IsDBNull(rawMasterData(RateLossColumn.PlannedUnplanned, index)) Then
                    ._PlannedUnplanned = BLANK_INDICATOR '""
                Else
                    ._PlannedUnplanned = rawMasterData(RateLossColumn.PlannedUnplanned, index)
                End If

                If IsDBNull(rawMasterData(RateLossColumn.SKU, index)) Then
                    ._ProductCode = BLANK_INDICATOR '""
                Else
                    ._ProductCode = rawMasterData(RateLossColumn.SKU, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Rate_Actual, index)) Then
                    .ActualRate = 0 '""
                Else
                    .ActualRate = rawMasterData(RateLossColumn.Rate_Actual, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.Rate_Target, index)) Then
                    .TargetRate = 0 '""
                Else
                    .TargetRate = rawMasterData(RateLossColumn.Rate_Target, index)
                End If
                If IsDBNull(rawMasterData(RateLossColumn.ProductGroup, index)) Then
                    ._ProductGroup = BLANK_INDICATOR '""
                Else
                    ._ProductGroup = rawMasterData(RateLossColumn.ProductGroup, index)
                End If
                Try
                    If IsDBNull(rawMasterData(RateLossColumn.Comment, index)) Then
                        ._Comment = ""
                    Else
                        ._Comment = rawMasterData(RateLossColumn.Comment, index)
                    End If
                Catch e As Exception
                    ._Comment = ""
                End Try
            End If
            .mapEvent1()
        End With
        Return tmpEvent
    End Function

    Public Function getAllEventsFromMaster() As List(Of RateLossEvent)
        Dim tmpList As New List(Of RateLossEvent)
        For i As Integer = 0 To rawMasterData.GetLength(1) - 1
            tmpList.Add(getSingleEventFromMaster(i))
        Next i
        Return tmpList
    End Function


#Region "Construction"
    Private lineIndex As Integer = -1
    Public Sub New(masterData As Object(,), lineI As Integer)
        rawMasterData = masterData
        lineIndex = lineI
    End Sub

    Public Sub New()

    End Sub


#End Region

#Region "Imlementation of Sortable and Equitable"
#Region "Custom Sorting Params"
    Private _isStandardSort = True
    Private _sortField As Integer
    Public Sub setSortParam(dtField As Integer)
        If dtField = DowntimeField.startTime Then
            _isStandardSort = True
        Else
            _sortField = dtField
            _isStandardSort = False
        End If
    End Sub
#End Region
    'IMPLEMENTATION OF SORTABLE AND EQUITABLE
    Public Function CompareTo(ByVal Other As RateLossEvent) As Integer Implements IComparable(Of RateLossEvent).CompareTo
        If _isStandardSort Then
            Return Me._startTime.CompareTo(Other.startTime)
        Else
            Select Case _sortField
                Case DowntimeField.DT
                    Return Me.DT.CompareTo(Other.DT)
                Case DowntimeField.UT
                    Return Me.UT.CompareTo(Other.UT)
                Case DowntimeField.Reason1
                    Return Me.Reason1.CompareTo(Other.Reason1)
                Case DowntimeField.Reason2
                    Return Me.Reason2.CompareTo(Other.Reason2)
                Case DowntimeField.Reason3
                    Return Me.Reason3.CompareTo(Other.Reason3)
                Case DowntimeField.Reason4
                    Return Me.Reason4.CompareTo(Other.Reason4)
                Case DowntimeField.endTime
                    Return Me.endTime.CompareTo(Other.endTime)
                Case DowntimeField.Fault
                    Return Me.Fault.CompareTo(Other.Fault)
                Case DowntimeField.ProductCode
                    Return Me.ProductCode.CompareTo(Other.ProductCode)
                Case DowntimeField.ProductGroup
                    Return Me.ProductGroup.CompareTo(Other.ProductGroup)
                Case Else
                    Throw New Exception("Unknown DT Field. CLS DataInterface")
            End Select
        End If
    End Function
    'equality - find index of
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As RateLossEvent = TryCast(obj, RateLossEvent)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As RateLossEvent) As Boolean _
        Implements IEquatable(Of RateLossEvent).Equals
        If other.startTime >= startTime_UT And other.startTime <= _endTime Then
            Return True
        ElseIf DateDiff(DateInterval.Second, other.startTime, startTime_UT) = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region
End Class
