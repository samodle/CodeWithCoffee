Imports System.Threading

Module prStoryEnums
    Public Const PRSTORY_NotMappedString As String = "x"
    Public Const OTHERS_STRING As String = "Others"

    Public Enum prStoryCard
        Unplanned = 1 '6 fields
        Planned = 2 '9 fields
        Equipment = 3 '5 fields
        Equipment_One = 4 '3 each
        Equipment_Two = 5
        Equipment_Three = 6
        Materials = 22 '2
        Bulk = 21 '2
        Stops = 31 '15
        TopThree = 32 '3
        Changeover = 41
    End Enum

    Public Enum prStoryCardFields
        Unplanned = 6 '6 fields
        Planned = 9 '9 fields
        Equipment = 6 '5 fields
        Equipment_One = 3 '3 each
        Equipment_Two = 3
        Equipment_Three = 3
        Materials = 2 '2
        Bulk = 2 '2
        Stops = 15 '15
        TopThree = 3 '3
        Changeover = 7
    End Enum
End Module

Public Class prStoryMainPageReport
    Implements IEquatable(Of prStoryMainPageReport)
    Implements IComparable(Of prStoryMainPageReport)

#Region "Variables & Properties"
    Friend RawRateLossData As Object(,)

    Friend UnplannedList As New List(Of DTevent)
    Friend PlannedList As New List(Of DTevent)
    '  Friend PlannedReportsList As New List(Of PDTeventReport)

    Private T1PlannedList As New List(Of DTevent)
    Friend T1UnplannedList As New List(Of DTevent)
    Private ChangeoverList As New List(Of DTevent)
    Private EquipMainList As New List(Of DTevent) ' Unplanned Tier 2
    Private EquipOneList As New List(Of DTevent) 'Tier 3 A
    Private EquipTwoList As New List(Of DTevent) 'Tier 3 B
    Private EquipThreeList As New List(Of DTevent)

    Friend MaterialsList As New List(Of DTevent)
    Friend BulkList As New List(Of DTevent)
    Friend TopThreeList As New List(Of DTevent)

    Public TopStopsList As New List(Of DTevent)

    Private _bargraphReportWindow As bargraphreportwindow
    Public Sub setBargraphReportWindow(parentWindow As bargraphreportwindow)
        _bargraphReportWindow = parentWindow
    End Sub


    Friend _ParentLine As Integer
    Protected _prstoryMapping As Integer
    Protected _startTime As Date
    Protected _endTime As Date
    Private _eventMaxDT As Double

    Private _eventMaxUPDTsim As Double
    Private _eventMaxPDTsim As Double
    Public ReadOnly Property AvSys_Sim As Double
        Get
            Return _AVsystem
        End Get
    End Property
    Public ReadOnly Property EventMaxDTpctUnplannedSim As Double
        Get
            Return _eventMaxUPDTsim
        End Get
    End Property
    Public ReadOnly Property EventMaxDTpctplannedSim As Double
        Get
            Return _eventMaxPDTsim
        End Get
    End Property

    Protected _ColumnToMap As Integer
    Public MainLEDSReport As SummaryReport

    Public ReadOnly Property MSU As Double
        Get
            If My.Settings.AdvancedSettings_isAvailabilityMode Then Return 0
            Return MainLEDSReport.PROD_Report.UnitsStat
        End Get
    End Property

    Public ReadOnly Property StartDate As Date
        Get
            Return _startTime
        End Get
    End Property
    Public ReadOnly Property EndDate As Date
        Get
            Return _endTime
        End Get
    End Property
    Public ReadOnly Property ParentLineInt As Integer
        Get
            Return _ParentLine
        End Get
    End Property

    'properties
    Public ReadOnly Property schedTime
        Get
            Return MainLEDSReport.schedTime
        End Get
    End Property
    Public ReadOnly Property rateLoss
        Get
            Return MainLEDSReport.RateLossPct
        End Get
    End Property
    Public ReadOnly Property PR
        Get
            Return MainLEDSReport.PR
        End Get
    End Property
    Public ReadOnly Property UPDT As Double
        Get
            Return MainLEDSReport.UPDTpct
        End Get
    End Property
    Public ReadOnly Property PDT As Double
        Get
            Return MainLEDSReport.PDTpct
        End Get
    End Property
    Public ReadOnly Property CasesActual
        Get
            Return MainLEDSReport.ActualCases
        End Get
    End Property
    Public ReadOnly Property CasesAdjusted
        Get
            Return MainLEDSReport.AdjustedCases
        End Get
    End Property
    Public ReadOnly Property StopsPerDay
        Get
            Return MainLEDSReport.SPD
        End Get
    End Property
    Public ReadOnly Property ActualStops
        Get
            Return MainLEDSReport.Stops
        End Get
    End Property
    Public ReadOnly Property MTBF
        Get
            If MainLEDSReport.Stops = 0 Then Return 0

            If Not My.Settings.AdvancedSettings_isAvailabilityMode And (MainLEDSReport.ParentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple Or MainLEDSReport.ParentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple_New) Then
                Return MainLEDSReport.PROD_Report._uptimeCalc / MainLEDSReport.DT_Report.Stops
            Else
                Return MainLEDSReport.UT_DT / MainLEDSReport.Stops
            End If
        End Get
    End Property

    Public ReadOnly Property MTTR
        Get
            If MainLEDSReport.Stops = 0 Then Return 0

            Return MainLEDSReport.UPDT / MainLEDSReport.Stops

        End Get
    End Property

    Public ReadOnly Property EventMaxDTpct As Double
        Get
            If MainLEDSReport.schedTime = 0 Then Return 0
            Return _eventMaxDT / MainLEDSReport.schedTime  '15 'SRONEW
            '   If MainLEDSReport.DT_Report.EventMaxDT = 0 Then Return 0
            '  Return MainLEDSReport.DT_Report.EventMaxDT / MainLEDSReport.schedTime
        End Get
    End Property
    Public ReadOnly Property StartTime As Date
        Get
            Return _startTime
        End Get
    End Property
    Public ReadOnly Property EndTime As Date
        Get
            Return _endTime
        End Get
    End Property
#End Region


#Region "Time Period Comparison"
    Private prstoryReport_secondary As prStoryMainPageReport

    Private Sub initializeSecondaryComparison()
        prstoryReport_secondary = New prStoryMainPageReport(ParentLineInt, _startTime, _endTime)
        prstoryReport_secondary.initializeSimulationMode()
    End Sub

    Public Sub Simulation_SwapTimePeriods()
        For Each DTevent In UnplannedList
            DTevent.swapSecondaryValues()
        Next
        For Each DTevent In PlannedList
            DTevent.swapSecondaryValues()
        Next
        For Each DTevent In Complete_Unplanned_Tier2
            DTevent.swapSecondaryValues()
        Next
        For Each DTevent In Complete_Planned_Tier2
            DTevent.swapSecondaryValues()
        Next
        For Each DTevent In Complete_Unplanned_Tier3
            DTevent.swapSecondaryValues()
        Next
        pushSimulatedEventsToActiveLists()
        pushSimulatedEventsToActiveLists_T3()
    End Sub

#End Region

#Region "Loss Allocation"
    Friend Complete_Unplanned_Tier2 As New List(Of DTevent)
    Friend Complete_Unplanned_Tier3 As New List(Of DTevent)
    Friend Complete_Planned_Tier2 As New List(Of DTevent)
    Private _AVsystem As Double

    'initialize all the master lists the first time through
    Public Sub initializeSimulationMode()
        Dim i As Integer
        'check if we really wanna do this!
        If Complete_Unplanned_Tier2.Count = 0 Then

            For i = 0 To UnplannedList.Count - 1
                UnplannedList(i).initializeSimulation()
            Next
            For i = 0 To PlannedList.Count - 1
                PlannedList(i).UT = MainLEDSReport.UT
                PlannedList(i).initializeSimulation()
            Next
            Simulation_createMasterLists()
        End If
    End Sub
    Private Sub Simulation_createMasterLists()
        Dim Tier1Inc As Integer, Tier2Inc As Integer, Tier3Inc As Integer, T1name As String, T2name As String
        Dim tmp2List As List(Of DTevent), tmp3List As List(Of DTevent)
        'Unplanned
        With MainLEDSReport.DT_Report
            For Tier1Inc = 1 To UnplannedList.Count - 1 'do not include TOTALS
                T1name = UnplannedList(Tier1Inc).Name
                tmp2List = .getTier2Directory(T1name)
                For Tier2Inc = 0 To tmp2List.Count - 1
                    T2name = tmp2List(Tier2Inc).Name

                    tmp2List(Tier2Inc).DTpct = schedTime
                    tmp2List(Tier2Inc).MTBF = MainLEDSReport.UT '_DT

                    tmp2List(Tier2Inc).initializeSimulation()
                    tmp2List(Tier2Inc).ParentEvent = T1name
                    Complete_Unplanned_Tier2.Add(tmp2List(Tier2Inc))
                    'TIER 3 TIME
                    tmp3List = .getTier3Directory(T1name, T2name)
                    For Tier3Inc = 0 To tmp3List.Count - 1
                        tmp3List(Tier3Inc).DTpct = schedTime
                        tmp3List(Tier3Inc).MTBF = MainLEDSReport.UT '_DT

                        tmp3List(Tier3Inc).initializeSimulation()
                        tmp3List(Tier3Inc).ParentEvent = T2name
                        Complete_Unplanned_Tier3.Add(tmp3List(Tier3Inc))
                    Next
                Next
            Next
            'planned

            For Tier1Inc = 1 To PlannedList.Count - 1 'do not include TOTALS
                T1name = PlannedList(Tier1Inc).Name
                tmp2List = .getPlannedTier2Directory(T1name)
                For Tier2Inc = 0 To tmp2List.Count - 1
                    T2name = tmp2List(Tier2Inc).Name
                    tmp2List(Tier2Inc).DTpct = schedTime
                    tmp2List(Tier2Inc).MTBF = MainLEDSReport.UT '_DT
                    tmp2List(Tier2Inc).initializeSimulation()
                    tmp2List(Tier2Inc).ParentEvent = T1name
                    Complete_Planned_Tier2.Add(tmp2List(Tier2Inc))
                Next
            Next
        End With
    End Sub

    'adjust the loss allocations /dtsims on user input
    Public Sub GenerateNewLossAllocation(CardNumber As Integer, FieldName As String, MTTRscaleFactor As Double, MTBFscaleFactor As Double, Optional Tier1Name As String = "", Optional Tier2Name As String = "")
        Dim isInInitialState As Boolean = False

        If MTTRscaleFactor = 1 And MTBFscaleFactor = 1 Then
            If UnplannedList(0).LossAllocation = 0 Then
                isInInitialState = True
            Else
                adjustAllMTTRsMTBFs(CardNumber, FieldName, MTTRscaleFactor, MTBFscaleFactor, Tier1Name) ', Tier2Name)
                isInInitialState = True
                For Each DTevent In UnplannedList
                    If DTevent.MTTR_userScaleFactor <> 1 Or DTevent.MTBF_userScaleFactor <> 1 Then isInInitialState = False
                Next
                If isInInitialState Then
                    For Each DTevent In PlannedList
                        If DTevent.MTTR_userScaleFactor <> 1 Or DTevent.MTBF_userScaleFactor <> 1 Then isInInitialState = False
                    Next
                    If isInInitialState Then
                        For Each DTevent In Complete_Planned_Tier2
                            If DTevent.MTTR_userScaleFactor <> 1 Or DTevent.MTBF_userScaleFactor <> 1 Then isInInitialState = False
                        Next
                        If isInInitialState Then
                            For Each DTevent In Complete_Unplanned_Tier2
                                If DTevent.MTTR_userScaleFactor <> 1 Or DTevent.MTBF_userScaleFactor <> 1 Then isInInitialState = False
                            Next
                            If isInInitialState Then
                                For Each DTevent In Complete_Unplanned_Tier3
                                    If DTevent.MTTR_userScaleFactor <> 1 Or DTevent.MTBF_userScaleFactor <> 1 Then isInInitialState = False
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If isInInitialState Then

            For Each DTevent In UnplannedList
                DTevent.LossAllocation = DTevent.DTpct
            Next
            For Each DTevent In PlannedList
                DTevent.LossAllocation = DTevent.DTpct
            Next
            For Each DTevent In Complete_Planned_Tier2
                DTevent.LossAllocation = DTevent.DTpct
            Next
            For Each DTevent In Complete_Unplanned_Tier2
                DTevent.LossAllocation = DTevent.DTpct
            Next
            For Each DTevent In Complete_Unplanned_Tier3
                DTevent.LossAllocation = DTevent.DTpct
            Next
            _bargraphReportWindow.pr_labelsim.Content = _bargraphReportWindow.pr_label.Content
            pushSimulatedEventsToActiveLists()
        Else
            adjustAllMTTRsMTBFs(CardNumber, FieldName, MTTRscaleFactor, MTBFscaleFactor, Tier1Name) ', Tier2Name)
            reCalculateSimulatedAvailabilities()

            pushSimulatedEventsToActiveLists()
            If My.Settings.AdvancedSettings_isAvailabilityMode Then
                _bargraphReportWindow.pr_labelsim.Content = FormatPercent(AvSys_Sim, 1) & " Av."
            Else
                If AvSys_Sim = 0 Then
                    _bargraphReportWindow.pr_labelsim.Content = _bargraphReportWindow.pr_label.Content
                Else
                    _bargraphReportWindow.pr_labelsim.Content = FormatPercent(AvSys_Sim - rateLoss, 1) & " PR"
                End If
            End If
        End If

    End Sub

    Private Sub adjustAllMTTRsMTBFs(CardNumber As Integer, FieldName As String, MTTRscaleFactor As Double, MTBFscaleFactor As Double, Optional Tier1Name As String = "") ', Optional Tier2Name As String = "")
        Dim targetIndex As Integer, i As Integer, j As Integer, Tier2Name As String
        Dim initialSimDT As Double, DTdifference As Double
        Select Case CardNumber
            Case prStoryCard.Unplanned
                targetIndex = UnplannedList.IndexOf(New DTevent(FieldName, 0))
                With UnplannedList(targetIndex)
                    .MTTR_userScaleFactor = MTTRscaleFactor
                    .MTBF_userScaleFactor = MTBFscaleFactor

                    .DTsim = .DT / .MTBF_netScaleFactor
                    .StopsSim = .Stops / .MTBF_netScaleFactor
                    .DTsim = .DTsim * .MTTR_netScaleFactor


                End With

                For i = 0 To Complete_Unplanned_Tier2.Count - 1
                    With Complete_Unplanned_Tier2(i)
                        If .ParentEvent.Equals(FieldName) Then
                            .MTTR_parentScaleFactor = MTTRscaleFactor
                            .MTBF_parentScaleFactor = MTBFscaleFactor


                            .DTsim = .DT / .MTBF_netScaleFactor
                            .StopsSim = .Stops / .MTBF_netScaleFactor
                            .DTsim = .DTsim * .MTTR_netScaleFactor

                            For j = 0 To Complete_Unplanned_Tier3.Count - 1
                                If Complete_Unplanned_Tier3(j).ParentEvent.Equals(.Name) Then
                                    Complete_Unplanned_Tier3(j).MTTR_parentTwoScaleFactor = MTTRscaleFactor
                                    Complete_Unplanned_Tier3(j).MTBF_parentTwoScaleFactor = MTBFscaleFactor


                                    Complete_Unplanned_Tier3(j).DTsim = Complete_Unplanned_Tier3(j).DT / Complete_Unplanned_Tier3(j).MTBF_netScaleFactor
                                    Complete_Unplanned_Tier3(j).StopsSim = Complete_Unplanned_Tier3(j).Stops / Complete_Unplanned_Tier3(j).MTBF_netScaleFactor
                                    Complete_Unplanned_Tier3(j).DTsim = Complete_Unplanned_Tier3(j).DTsim * Complete_Unplanned_Tier3(j).MTTR_netScaleFactor
                                End If
                            Next
                        End If
                    End With
                Next
            Case prStoryCard.Equipment
                For i = 0 To Complete_Unplanned_Tier2.Count - 1
                    With Complete_Unplanned_Tier2(i)
                        If .Name.Equals(FieldName) And .ParentEvent.Equals(Tier1Name) Then

                            .MTTR_userScaleFactor = MTTRscaleFactor
                            .MTBF_userScaleFactor = MTBFscaleFactor

                            .DTsim = .DT / .MTBF_netScaleFactor
                            .StopsSim = .Stops / .MTBF_netScaleFactor
                            .DTsim = .DTsim * .MTTR_netScaleFactor



                            For j = 0 To Complete_Unplanned_Tier3.Count - 1
                                If Complete_Unplanned_Tier3(j).ParentEvent.Equals(.Name) Then
                                    Complete_Unplanned_Tier3(j).MTTR_parentScaleFactor = MTTRscaleFactor
                                    Complete_Unplanned_Tier3(j).MTBF_parentScaleFactor = MTBFscaleFactor

                                    Complete_Unplanned_Tier3(j).DTsim = Complete_Unplanned_Tier3(j).DT / Complete_Unplanned_Tier3(j).MTBF_netScaleFactor
                                    Complete_Unplanned_Tier3(j).StopsSim = Complete_Unplanned_Tier3(j).Stops / Complete_Unplanned_Tier3(j).MTBF_netScaleFactor
                                    Complete_Unplanned_Tier3(j).DTsim = Complete_Unplanned_Tier3(j).DTsim * Complete_Unplanned_Tier3(j).MTTR_netScaleFactor

                                End If
                            Next
                        End If
                    End With
                Next
            Case prStoryCard.Planned
                targetIndex = PlannedList.IndexOf(New DTevent(FieldName, 0))
                With PlannedList(targetIndex)

                    .MTTR_userScaleFactor = MTTRscaleFactor
                    .MTBF_userScaleFactor = MTBFscaleFactor

                    .DTsim = .DT / MTBFscaleFactor
                    .StopsSim = .Stops / MTBFscaleFactor
                    .DTsim = .DTsim * MTTRscaleFactor

                End With

                For i = 0 To Complete_Planned_Tier2.Count - 1
                    With Complete_Planned_Tier2(i)
                        If .ParentEvent.Equals(FieldName) Then
                            .MTTR_parentScaleFactor = MTTRscaleFactor
                            .MTBF_parentScaleFactor = MTBFscaleFactor

                            .DTsim = .DT / .MTBF_netScaleFactor
                            .StopsSim = .Stops / .MTBF_netScaleFactor
                            .DTsim = .DTsim * .MTTR_netScaleFactor


                        End If
                    End With
                Next
            Case prStoryCard.Changeover
                For i = 0 To Complete_Planned_Tier2.Count - 1
                    With Complete_Planned_Tier2(i)
                        If .Name.Equals(FieldName) And .ParentEvent.Equals(Tier1Name) Then

                            .MTTR_userScaleFactor = MTTRscaleFactor
                            .MTBF_userScaleFactor = MTBFscaleFactor

                            .DTsim = .DT / .MTBF_netScaleFactor
                            .StopsSim = .Stops / .MTBF_netScaleFactor
                            .DTsim = .DTsim * .MTTR_netScaleFactor

                        End If
                    End With
                Next
            Case Else 'assumes this is all tier 3 unplanned
                If CardNumber = prStoryCard.Equipment_One Then
                    Tier2Name = _bargraphReportWindow.Card4Header.Content
                ElseIf CardNumber = prStoryCard.Equipment_Two Then
                    Tier2Name = _bargraphReportWindow.Card5Header.Content
                Else 'this better be equipment three (card 6)...
                    Tier2Name = _bargraphReportWindow.Card6Header.Content
                End If


                For i = 0 To Complete_Unplanned_Tier3.Count - 1
                    With Complete_Unplanned_Tier3(i)
                        If .Name.Equals(FieldName) And .ParentEvent.Equals(Tier2Name) Then
                            initialSimDT = .DTsim
                            .MTTR_userScaleFactor = MTTRscaleFactor
                            .MTBF_userScaleFactor = MTBFscaleFactor

                            .DTsim = .DT / .MTBF_netScaleFactor
                            .StopsSim = .Stops / .MTBF_netScaleFactor
                            .DTsim = .DTsim * .MTTR_netScaleFactor




                            DTdifference = initialSimDT - .DTsim
                        End If
                    End With
                Next
                For i = 0 To Complete_Unplanned_Tier2.Count - 1
                    With Complete_Unplanned_Tier2(i)
                        If .Name.Equals(Tier2Name) And .ParentEvent.Equals(Tier1Name) Then
                            .DTsim = .DTsim - DTdifference
                        End If
                    End With
                Next
        End Select

    End Sub


    Private Sub reCalculateSimulatedAvailabilities()
        'System Availability = 1 / 1 - Sum[(1-A)/A]  

        Dim tmpAvCalcList As New List(Of Double)

        Dim T1i As Integer, T2i As Integer ', T3i As Integer
        Dim tmpAv As Double, tmpCalcA As Double, tmpCalcB As Double
        Dim netPlannedAv As Double = 0, netUnplannedAv As Double = 0
        Dim i As Integer

        _AVsystem = 0
        'STEP 1: Find System Av
        For i = 1 To UnplannedList.Count - 1 'UNPLANNED
            With UnplannedList(i)
                tmpAv = 0
                For T2i = 0 To Complete_Unplanned_Tier2.Count - 1
                    If Complete_Unplanned_Tier2(T2i).ParentEvent.Equals(.Name) Then tmpAv += Complete_Unplanned_Tier2(T2i).AvailSim_System
                Next
                tmpCalcA = 1 / (1 + tmpAv)
                tmpCalcB = (1 - tmpCalcA) / tmpCalcA

                tmpAvCalcList.Add(tmpCalcB)
            End With
        Next
        For i = 1 To PlannedList.Count - 1 'PLANNED
            With PlannedList(i)
                tmpAv = 0
                For T2i = 0 To Complete_Planned_Tier2.Count - 1
                    If Complete_Planned_Tier2(T2i).ParentEvent.Equals(.Name) Then tmpAv += Complete_Planned_Tier2(T2i).AvailSim_System
                Next
                tmpCalcA = 1 / (1 + tmpAv)
                tmpCalcB = (1 - tmpCalcA) / tmpCalcA

                tmpAvCalcList.Add(tmpCalcB)
            End With
        Next
        'NOW SUM THE INDIVIDUAL TOP TIER FAILURE MODES TO FIND NEW/SIMULATED SYSTEM AVAILABILITY
        tmpCalcA = 0
        For i = 0 To tmpAvCalcList.Count - 1
            tmpCalcA += tmpAvCalcList(i)
        Next
        _AVsystem = 1 / (1 + tmpCalcA)

        'STEP 2: Get Individual Failure Mode Simulations



        For Each DTevent In Complete_Unplanned_Tier2
            DTevent.LossAllocation = _AVsystem * DTevent.DTsim / MainLEDSReport.UT '_DT
        Next
        For Each DTevent In Complete_Unplanned_Tier3
            DTevent.LossAllocation = _AVsystem * DTevent.DTsim / MainLEDSReport.UT '_DT
        Next
        For Each DTevent In Complete_Planned_Tier2
            DTevent.LossAllocation = _AVsystem * DTevent.DTsim / MainLEDSReport.UT '_DT
        Next

        Dim tmpLA As Double

        For T1i = 1 To UnplannedList.Count - 1
            tmpLA = 0
            For T2i = 0 To Complete_Unplanned_Tier2.Count - 1
                If Complete_Unplanned_Tier2(T2i).ParentEvent.Equals(UnplannedList(T1i).Name) Then tmpLA += Complete_Unplanned_Tier2(T2i).LossAllocation
            Next
            UnplannedList(T1i).LossAllocation = tmpLA
            netUnplannedAv += tmpLA
        Next
        UnplannedList(0).LossAllocation = netUnplannedAv

        For T1i = 1 To PlannedList.Count - 1
            tmpLA = 0
            For T2i = 0 To Complete_Planned_Tier2.Count - 1
                If Complete_Planned_Tier2(T2i).ParentEvent.Equals(PlannedList(T1i).Name) Then tmpLA += Complete_Planned_Tier2(T2i).LossAllocation
            Next
            PlannedList(T1i).LossAllocation = tmpLA
            netPlannedAv += tmpLA
        Next
        PlannedList(0).LossAllocation = netPlannedAv

        _eventMaxUPDTsim = netUnplannedAv
        _eventMaxPDTsim = netPlannedAv
    End Sub

    Private Sub reCalculateSimulatedAvailabilities_OLD()
        'System Availability = 1 / 1 - Sum[(1-A)/A]  <~~ math

        Dim tmpAvCalcList As New List(Of Double)

        Dim T1i As Integer, T2i As Integer ', T3i As Integer
        Dim tmpAv As Double, tmpCalcA As Double, tmpCalcB As Double
        Dim netPlannedAv As Double = 0, netUnplannedAv As Double = 0
        Dim i As Integer

        _AVsystem = 0
        'STEP 1: Figure It Out
        For i = 1 To UnplannedList.Count - 1 'UNPLANNED
            With UnplannedList(i)
                tmpAv = 0
                For T2i = 0 To Complete_Unplanned_Tier2.Count - 1
                    If Complete_Unplanned_Tier2(T2i).ParentEvent.Equals(.Name) Then tmpAv += Complete_Unplanned_Tier2(T2i).AvailSim_System
                Next
                tmpCalcA = 1 / (1 + tmpAv)
                tmpCalcB = (1 - tmpCalcA) / tmpCalcA

                tmpAvCalcList.Add(tmpCalcB)
            End With
        Next
        For i = 1 To PlannedList.Count - 1 'PLANNED
            With PlannedList(i)
                tmpAv = 0
                For T2i = 0 To Complete_Planned_Tier2.Count - 1
                    If Complete_Planned_Tier2(T2i).ParentEvent.Equals(.Name) Then tmpAv += Complete_Planned_Tier2(T2i).AvailSim_System
                Next
                tmpCalcA = 1 / (1 + tmpAv)
                tmpCalcB = (1 - tmpCalcA) / tmpCalcA

                tmpAvCalcList.Add(tmpCalcB)
            End With
        Next
        'NOW SUM THE INDIVIDUAL TOP TIER FAILURE MODES TO FIND NEW/SIMULATED SYSTEM AVAILABILITY
        tmpCalcA = 0
        For i = 0 To tmpAvCalcList.Count - 1
            tmpCalcA += tmpAvCalcList(i)
        Next
        _AVsystem = 1 / (1 + tmpCalcA)

        'STEP 2: Get Individual Failure Mode Simulations



        For Each DTevent In Complete_Unplanned_Tier2
            DTevent.LossAllocation = _AVsystem * DTevent.DTsim / MainLEDSReport.UT '_DT
        Next
        For Each DTevent In Complete_Unplanned_Tier3
            DTevent.LossAllocation = _AVsystem * DTevent.DTsim / MainLEDSReport.UT '_DT
        Next
        For Each DTevent In Complete_Planned_Tier2
            DTevent.LossAllocation = _AVsystem * DTevent.DTsim / MainLEDSReport.UT '_DT
        Next

        Dim tmpLA As Double

        For T1i = 1 To UnplannedList.Count - 1
            tmpLA = 0
            For T2i = 0 To Complete_Unplanned_Tier2.Count - 1
                If Complete_Unplanned_Tier2(T2i).ParentEvent.Equals(UnplannedList(T1i).Name) Then tmpLA += Complete_Unplanned_Tier2(T2i).LossAllocation
            Next
            UnplannedList(T1i).LossAllocation = tmpLA
            netUnplannedAv += tmpLA
        Next
        UnplannedList(0).LossAllocation = netUnplannedAv

        For T1i = 1 To PlannedList.Count - 1
            tmpLA = 0
            For T2i = 0 To Complete_Planned_Tier2.Count - 1
                If Complete_Planned_Tier2(T2i).ParentEvent.Equals(PlannedList(T1i).Name) Then tmpLA += Complete_Planned_Tier2(T2i).LossAllocation
            Next
            PlannedList(T1i).LossAllocation = tmpLA
            netPlannedAv += tmpLA
        Next
        PlannedList(0).LossAllocation = netPlannedAv

        _eventMaxUPDTsim = netUnplannedAv
        _eventMaxPDTsim = netPlannedAv
    End Sub



    Private Sub pushSimulatedEventsToActiveLists() 'Tier1Name As String) ', Tier2Name As String)

        'added to try to make top bars appear in loss allocation
        Dim ix As Integer = 0
        For i As Integer = 0 To UnplannedList.Count - 1
            ix += UnplannedList(i).DTpctSim
            ix += UnplannedList(i).LossAllocation
        Next

        '''''''''''''''''''''''''''


        Dim T2i As Integer, targetIndex As Integer  ', Tier1Name As String
        'Unplanned First
        For T2i = 0 To Complete_Unplanned_Tier2.Count - 1
            With Complete_Unplanned_Tier2(T2i)
                If .ParentEvent.Equals(_bargraphReportWindow.Tier1Clicked_Unplanned) Then '_bargraphReportWindow.Card3Header.Content) Then
                    targetIndex = EquipMainList.IndexOf(New DTevent(.Name, 0))
                    If targetIndex > -1 Then
                        EquipMainList(targetIndex).LossAllocation = .LossAllocation
                        EquipMainList(targetIndex).DTsim = .DTsim
                        EquipMainList(targetIndex).StopsSim = .StopsSim
                    End If
                End If
            End With
        Next

        'Planned
        For T2i = 0 To Complete_Planned_Tier2.Count - 1
            With Complete_Planned_Tier2(T2i)
                If .ParentEvent = _bargraphReportWindow.Tier1Clicked_planned Then '_bargraphReportWindow.Card41Header.Content Then
                    targetIndex = ChangeoverList.IndexOf(New DTevent(.Name, 0))
                    If targetIndex > -1 Then
                        ChangeoverList(targetIndex).UT = MainLEDSReport.UT
                        ChangeoverList(targetIndex).LossAllocation = .LossAllocation
                        ChangeoverList(targetIndex).DTsim = .DTsim
                        ChangeoverList(targetIndex).StopsSim = .StopsSim
                    End If
                End If
            End With
        Next




        '     For T3i = 0 To Complete_Unplanned_Tier3.Count - 1
        ' With Complete_Unplanned_Tier3(T3i)
        ' If .ParentEvent.Equals(_bargraphReportWindow.Card4Header.Content) Then
        ' targetIndex = EquipOneList.IndexOf(New DTevent(.Name, 0))
        '' If targetIndex > -1 Then
        '' EquipOneList(targetIndex).LossAllocation = .LossAllocation
        ' EquipOneList(targetIndex).StopsSim = .StopsSim
        ' End If
        '
        '        ElseIf .ParentEvent.Equals(_bargraphReportWindow.Card5Header.Content) Then
        '        targetIndex = EquipTwoList.IndexOf(New DTevent(.Name, 0))
        '        If targetIndex > -1 Then
        ' EquipTwoList(targetIndex).LossAllocation = .LossAllocation
        ' EquipTwoList(targetIndex).StopsSim = .StopsSim
        ' End If
        '
        '        ElseIf .ParentEvent.Equals(_bargraphReportWindow.Card6Header.Content) Then
        '        targetIndex = EquipThreeList.IndexOf(New DTevent(.Name, 0))
        '        If targetIndex > -1 Then
        ' EquipThreeList(targetIndex).LossAllocation = .LossAllocation
        ' EquipThreeList(targetIndex).StopsSim = .StopsSim
        ' End If
        '
        '        ElseIf .ParentEvent.Equals(_bargraphReportWindow.Tier2Clicked_Unplanned) Then '_bargraphReportWindow.Card6Header.Content) Then
        '        targetIndex = EquipThreeList.IndexOf(New DTevent(.Name, 0))
        '        If targetIndex > -1 Then
        ' EquipThreeList(targetIndex).LossAllocation = .LossAllocation
        ' EquipThreeList(targetIndex).StopsSim = .StopsSim
        ' End If

        '        End If
        '        End With
        '        Next



    End Sub
    Private Sub pushSimulatedEventsToActiveLists_T3() 'Tier1Name As String) ', Tier2Name As String)
        Dim j As Integer, IndexOne As Integer, IndexTwo As Integer, IndexThree As Integer, WeDone As Boolean
        For j = 0 To Complete_Unplanned_Tier3.Count - 1
            With Complete_Unplanned_Tier3(j)
                WeDone = False
                IndexOne = EquipOneList.IndexOf(Complete_Unplanned_Tier3(j))
                IndexTwo = EquipTwoList.IndexOf(Complete_Unplanned_Tier3(j))
                IndexThree = EquipThreeList.IndexOf(Complete_Unplanned_Tier3(j))

                If IndexOne >= 0 Then
                    If EquipOneList(IndexOne).DT = .DT Then
                        WeDone = True
                        EquipOneList(IndexOne).LossAllocation = .LossAllocation
                        EquipOneList(IndexOne).DTsim = .DTsim
                        EquipOneList(IndexOne).StopsSim = .StopsSim
                        EquipOneList(IndexOne).UT = MainLEDSReport.UT
                        ' EquipOneList(IndexOne).DTpct = schedTime
                    End If
                End If

                If IndexTwo >= 0 And Not WeDone Then
                    If EquipTwoList(IndexTwo).DT = .DT Then
                        WeDone = True
                        EquipTwoList(IndexTwo).LossAllocation = .LossAllocation
                        EquipTwoList(IndexTwo).DTsim = .DTsim
                        EquipTwoList(IndexTwo).StopsSim = .StopsSim
                        EquipTwoList(IndexTwo).UT = MainLEDSReport.UT
                    End If
                End If

                If IndexThree >= 0 And Not WeDone Then
                    If EquipThreeList(IndexThree).DT = .DT Then
                        WeDone = True
                        EquipThreeList(IndexThree).LossAllocation = .LossAllocation
                        EquipThreeList(IndexThree).DTsim = .DTsim
                        EquipThreeList(IndexThree).StopsSim = .StopsSim
                        EquipThreeList(IndexThree).UT = MainLEDSReport.UT
                    End If
                End If

            End With
        Next
    End Sub

#End Region


#Region "Updating Card Lists"

    'top stops list
    Public Sub updateCardList_Stops(ScrollOffset As Integer)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Stops) - 1
        With _bargraphReportWindow.Card_TopStops
            '  If Not TopStopsList.Count - 1 < cardEventFields + ScrollOffset Then 'else we dont care...
            .Clear()
            For i As Integer = ScrollOffset To ScrollOffset + cardEventFields
                If i >= TopStopsList.Count Then
                    .Add(New DTevent("", 0))
                Else
                    .Add(TopStopsList(i))
                End If
            Next
            '  End If
        End With
    End Sub

    'Unplanned / Tier 1 Card
    Public Sub updateCardList_Unplanned_Tier1(ScrollOffset As Integer)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Unplanned) - 1

        Dim totalEvent As New DTevent("Total", 0)
        Dim masterList As New List(Of String)
        masterList = AllProdLines(_ParentLine).UnplannedT1List
        With _bargraphReportWindow.Card_Unplanned_T1

            .Clear()
            T1UnplannedList = MainLEDSReport.DT_Report.getTier1Directory()
            'now we need to update the sim values
            For i As Integer = 0 To T1UnplannedList.Count - 1
                Dim unplannedListIndex As Integer = -1
                For j As Integer = 0 To UnplannedList.Count - 1
                    If UnplannedList(j).Name = T1UnplannedList(i).Name Then
                        unplannedListIndex = j
                    End If
                Next
                If unplannedListIndex > -1 Then
                    T1UnplannedList(i) = UnplannedList(unplannedListIndex)
                End If

            Next
            'end sim update


            For i As Integer = 0 To masterList.Count - 1
                If T1UnplannedList.IndexOf(New DTevent(masterList(i), 0)) = -1 And masterList(i) <> totalEvent.Name Then
                    T1UnplannedList.Add(New DTevent(masterList(i), 0))
                End If
            Next

            For i As Integer = 0 To T1UnplannedList.Count - 1
                T1UnplannedList(i).MTBF = MainLEDSReport.UT_DT
                T1UnplannedList(i).DTpct = schedTime

                totalEvent.DT += T1UnplannedList(i).DT
                totalEvent.Stops += T1UnplannedList(i).Stops

                totalEvent.DTsim += T1UnplannedList(i).DTsim
                totalEvent.StopsSim += T1UnplannedList(i).StopsSim
                totalEvent.LossAllocation += T1UnplannedList(i).LossAllocation

                totalEvent.RawRows.AddRange(T1UnplannedList(i).RawRows)
            Next
            totalEvent.MTBF = MainLEDSReport.UT_DT
            totalEvent.DTpct = schedTime
            totalEvent.RawRows.Sort()
            T1UnplannedList.Add(totalEvent)

            sortEventList_ByDT(T1UnplannedList)

            'For i As Integer = 0 To cardEventFields
            For i As Integer = ScrollOffset To ScrollOffset + cardEventFields ' LG Code
                If i > T1UnplannedList.Count - 1 Then
                    .Add(New DTevent("", 0))
                Else
                    .Add(T1UnplannedList(i))
                End If
            Next
        End With
    End Sub
    'Unplanned / Tier 2 Card
    Public Sub updateCardList_Unplanned_Tier2(ScrollOffset As Integer, selectedTier1 As String)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Equipment) - 1
        With _bargraphReportWindow.Card_Unplanned_T2

            .Clear()
            EquipMainList = MainLEDSReport.DT_Report.getTier2Directory(selectedTier1)
            sortEventList_ByDT(EquipMainList)

            For i As Integer = 0 To EquipMainList.Count - 1
                EquipMainList(i).MTBF = MainLEDSReport.UT_DT
                EquipMainList(i).DTpct = schedTime
            Next

            'IF SIMULATION MODE, UPDATE THE STUFF HERE!!!
            If IsSimulationMode Then pushSimulatedEventsToActiveLists()
            '//////////////////////////////////////////////


            'For i As Integer = 0 To cardEventFields
            For i As Integer = ScrollOffset To ScrollOffset + cardEventFields ' LG Code
                If i > EquipMainList.Count - 1 Then
                    .Add(New DTevent("", 0))
                Else
                    .Add(EquipMainList(i))
                End If
            Next
        End With
        EquipOneList.Clear()
        EquipTwoList.Clear()
        EquipThreeList.Clear()
        'AND NOW WE NEED TO UPDATE THE REASON 3s!!!
        With MainLEDSReport.DT_Report
            If EquipMainList.Count > 0 Then
                EquipOneList = .getTier3Directory(selectedTier1, EquipMainList(0).Name)
                'EquipOneName = EquipMainList(0).Name
                sortEventList_ByDT(EquipOneList)
                If EquipMainList.Count > 1 Then
                    EquipTwoList = .getTier3Directory(selectedTier1, EquipMainList(1).Name)
                    'EquipTwoName = EquipMainList(1).Name
                    sortEventList_ByDT(EquipTwoList)
                End If
                If EquipMainList.Count > 2 Then
                    EquipThreeList = .getTier3Directory(selectedTier1, EquipMainList(2).Name)
                    '  EquipThreeName = EquipMainList(2).Name
                    sortEventList_ByDT(EquipThreeList)
                End If
            End If
        End With

        'IF SIMULATION MODE, UPDATE THE STUFF HERE!!!
        ' If IsSimulationMode Then pushSimulatedEventsToActiveLists()
        '//////////////////////////////////////////////
        If IsSimulationMode Then pushSimulatedEventsToActiveLists_T3()





        With _bargraphReportWindow
            .Card_Unplanned_T3A = EquipOneList
            .Card_Unplanned_T3B = EquipTwoList
            .Card_Unplanned_T3C = EquipThreeList
        End With
    End Sub

    Public Sub updateCardList_Unplanned_Tier3A(ScrollOffset As Integer, selectedTier1 As String, selectedTier2 As String)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Equipment_One) - 1
        EquipOneList.Clear()
        If selectedTier2 <> "" Then
            With MainLEDSReport.DT_Report
                If EquipMainList.Count > 0 Then
                    EquipOneList = .getTier3Directory(selectedTier1, selectedTier2)
                    sortEventList_ByDT(EquipOneList)
                End If
            End With

            With _bargraphReportWindow.Card_Unplanned_T3A
                .Clear()
                'set the scheduled times for % DT
                For i As Integer = 0 To EquipOneList.Count - 1
                    EquipOneList(i).DTpct = schedTime
                    EquipOneList(i).MTBF = MainLEDSReport.UT_DT
                Next
                'make sure we're at 3 items
                If .Count < cardEventFields + 1 Then
                    For i = .Count To cardEventFields + 1
                        '  .Add(New DTevent("", 0))
                    Next
                End If

                For i As Integer = ScrollOffset To ScrollOffset + cardEventFields ' LG Code
                    If i > EquipOneList.Count - 1 Then
                        .Add(New DTevent("", 0))
                    Else
                        .Add(EquipOneList(i))
                    End If
                Next


            End With

            'SIMULATION - LOSS ALLOCATION
            If IsSimulationMode Then pushSimulatedEventsToActiveLists_T3()


            With _bargraphReportWindow
                ' .Card_Unplanned_T3A = EquipOneList
            End With
        Else
            _bargraphReportWindow.Card_Unplanned_T3A.Clear()
            For i = 0 To cardEventFields
                _bargraphReportWindow.Card_Unplanned_T3A.Add(New DTevent("", 0))
            Next
        End If
    End Sub
    Public Sub updateCardList_Unplanned_Tier3B(ScrollOffset As Integer, selectedTier1 As String, selectedTier2 As String)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Equipment_Two) - 1
        EquipTwoList.Clear()
        With MainLEDSReport.DT_Report
            If EquipMainList.Count > 1 Then
                EquipTwoList = .getTier3Directory(selectedTier1, selectedTier2)
                sortEventList_ByDT(EquipTwoList)
            End If
        End With
        With _bargraphReportWindow.Card_Unplanned_T3B
            .Clear()
            'set the scheduled times for % DT
            For i As Integer = 0 To EquipTwoList.Count - 1
                EquipTwoList(i).DTpct = schedTime
                EquipTwoList(i).MTBF = MainLEDSReport.UT_DT
            Next
            'make sure we're at 3 items
            If EquipTwoList.Count < prStoryCardFields.Equipment_Two Then
                For i = EquipTwoList.Count To prStoryCardFields.Equipment_Two
                    '   EquipTwoList.Add(New DTevent("", 0))
                Next
            End If
            For i As Integer = ScrollOffset To ScrollOffset + cardEventFields ' LG Code
                If i > EquipTwoList.Count - 1 Then
                    .Add(New DTevent("", 0))
                Else
                    .Add(EquipTwoList(i))
                End If
            Next

        End With
        'SIMULATION - LOSS ALLOCATION
        If IsSimulationMode Then pushSimulatedEventsToActiveLists_T3()



        With _bargraphReportWindow
            ' .Card_Unplanned_T3B = EquipTwoList
        End With
    End Sub
    Public Sub updateCardList_Unplanned_Tier3C(ScrollOffset As Integer, selectedTier1 As String, selectedTier2 As String)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Equipment_Three) - 1
        EquipThreeList.Clear()
        With MainLEDSReport.DT_Report
            If EquipMainList.Count > 2 Then
                EquipThreeList = .getTier3Directory(selectedTier1, selectedTier2)
                '  EquipThreeName = selectedTier2
                sortEventList_ByDT(EquipThreeList)
            End If
        End With
        With _bargraphReportWindow.Card_Unplanned_T3C
            .Clear()

            'set the scheduled times for % DT
            For i As Integer = 0 To EquipThreeList.Count - 1
                EquipThreeList(i).DTpct = schedTime
                EquipThreeList(i).MTBF = MainLEDSReport.UT_DT
            Next
            'make sure we're at 3 items
            If EquipThreeList.Count < prStoryCardFields.Equipment_Three Then
                For i = EquipThreeList.Count To prStoryCardFields.Equipment_Three
                    '    EquipThreeList.Add(New DTevent("", 0))
                Next
            End If
            For i As Integer = ScrollOffset To ScrollOffset + cardEventFields ' LG Code
                If i > EquipThreeList.Count - 1 Then
                    .Add(New DTevent("", 0))
                Else
                    .Add(EquipThreeList(i))
                End If
            Next
        End With
        'SIMULATION - LOSS ALLOCATION
        If IsSimulationMode Then pushSimulatedEventsToActiveLists_T3()


        With _bargraphReportWindow
            '.Card_Unplanned_T3C = EquipThreeList
        End With
    End Sub

    'Planned / Tier 1 Card
    Public Sub updateCardList_Planned_Tier1(ScrollOffset As Integer)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Planned) - 1
        ' ' Dim searchIndex As Integer
        '  With _bargraphReportWindow.Card_Planned_T1
        '
        '        .Clear()
        '        For i As Integer = ScrollOffset To ScrollOffset + cardEventFields
        '        '  If i < PlannedList.Count Then
        '        searchIndex = PlannedList.IndexOf(New DTevent(getCardFieldName(prStoryCard.Planned, i), 0))
        '        If searchIndex = -1 Then
        '        .Add(New DTevent(getCardFieldName(prStoryCard.Planned, i), 0))
        '        Else
        '        .Add(PlannedList(searchIndex))
        '        End If
        'Else
        ' .Add(New DTevent(getCardFieldName(prStoryCard.Planned, i), 0))
        ' End If
        '        Next
        '        .Sort()
        '        End With


        Dim totalEvent As New DTevent("Total", 0)
        Dim masterList As New List(Of String)
        masterList = AllProdLines(_ParentLine).PlannedT1List
        With _bargraphReportWindow.Card_Planned_T1

            .Clear()
            T1PlannedList = MainLEDSReport.DT_Report.getPlannedTier1Directory()

            'now we need to update the sim values
            For i As Integer = 0 To T1PlannedList.Count - 1
                Dim unplannedListIndex As Integer = -1
                For j As Integer = 0 To PlannedList.Count - 1
                    If PlannedList(j).Name = T1PlannedList(i).Name Then
                        unplannedListIndex = j
                    End If
                Next
                If unplannedListIndex > -1 Then
                    T1PlannedList(i) = PlannedList(unplannedListIndex)
                End If

            Next
            'end sim update



            For i As Integer = 0 To masterList.Count - 1
                If T1PlannedList.IndexOf(New DTevent(masterList(i), 0)) = -1 And masterList(i) <> totalEvent.Name Then
                    T1PlannedList.Add(New DTevent(masterList(i), 0))
                End If
            Next

            For i As Integer = 0 To T1PlannedList.Count - 1
                T1PlannedList(i).MTBF = MainLEDSReport.UT_DT
                T1PlannedList(i).DTpct = schedTime

                totalEvent.DT += T1PlannedList(i).DT
                totalEvent.Stops += T1PlannedList(i).Stops

                totalEvent.DTsim += T1PlannedList(i).DTsim
                totalEvent.StopsSim += T1PlannedList(i).StopsSim
                totalEvent.LossAllocation += T1PlannedList(i).LossAllocation

                totalEvent.RawRows.AddRange(T1PlannedList(i).RawRows)
            Next
            totalEvent.MTBF = MainLEDSReport.UT_DT
            totalEvent.DTpct = schedTime

            totalEvent.RawRows.Sort()

            T1PlannedList.Add(totalEvent)

            sortEventList_ByDT(T1PlannedList)

            'For i As Integer = 0 To cardEventFields
            For i As Integer = ScrollOffset To ScrollOffset + cardEventFields ' LG Code
                If i > T1PlannedList.Count - 1 Then
                    .Add(New DTevent("", 0))
                Else
                    .Add(T1PlannedList(i))
                End If
            Next
        End With
    End Sub
    'Planned / Tier 2 Card
    Public Sub updateCardList_Planned_Tier2(ScrollOffset As Integer, selectedTier1 As String)
        Dim cardEventFields As Integer = getCardEventFields(prStoryCard.Changeover) - 1
        With _bargraphReportWindow.Card_Planned_T2

            .Clear()
            ChangeoverList = MainLEDSReport.DT_Report.getPlannedTier2Directory(selectedTier1)
            sortEventList_ByDT(ChangeoverList)

            For i As Integer = 0 To ChangeoverList.Count - 1
                ChangeoverList(i).DTpct = schedTime
            Next


            'IF SIMULATION MODE, UPDATE THE STUFF HERE!!!
            If IsSimulationMode Then pushSimulatedEventsToActiveLists()
            '//////////////////////////////////////////////

            For i As Integer = 0 To cardEventFields
                If i > ChangeoverList.Count - 1 Then
                    .Add(New DTevent("", 0))
                Else
                    .Add(ChangeoverList(i))
                End If
            Next
        End With

        'create PDT HTML stuff!
        'Dim Plandlesticks As Thread
        'Plandlesticks = New Thread(AddressOf PDT_candlesticks_Thread)
        'Plandlesticks.Start()
    End Sub


    Public Sub setTopNReasonsBforReasonA(ReasonAName As String, ReasonBField As Integer)
        TopThreeList.Clear()
        TopThreeList = MainLEDSReport.DT_Report.getMappedSubdirectory(ReasonAName, ReasonBField)
        sortEventList_ByStops(TopThreeList)
        For j As Integer = 0 To TopThreeList.Count - 1
            TopThreeList(j).DTpct = MainLEDSReport.schedTime * 0.01
        Next
        With _bargraphReportWindow.Card_TopThreeStops
            .Clear()
            For i As Integer = 0 To 2
                If i <= TopThreeList.Count - 1 Then
                    .Add(TopThreeList(i))
                Else
                    .Add(New DTevent("", 0))
                End If
            Next
        End With
    End Sub

#End Region

#Region "Card / Field Specific Information Functions"
    Public Function getCardEventInfo(cardNumber As Integer, listIndex As Integer) As DTevent
        Dim tmpDTevent As DTevent
        tmpDTevent = New DTevent(" ", 0) 'this is the blank event we return when we're being ...
        tmpDTevent.Stops = 0 '...asked for an event number thats outside the actual idex
        Select Case cardNumber
            'for fixed field cards the caller is responsible for handling 'blanks'
            Case prStoryCard.Unplanned
                If listIndex > getCardEventNumber(cardNumber) - 1 Then
                    Return tmpDTevent
                Else
                    UnplannedList(listIndex).DTpct = schedTime
                    UnplannedList(listIndex).MTBF = MainLEDSReport.UT_DT
                End If
                Return UnplannedList(listIndex)
            Case prStoryCard.Planned
                If listIndex > getCardEventNumber(cardNumber) - 1 Then
                    Return tmpDTevent
                Else
                    PlannedList(listIndex).DTpct = schedTime
                End If
                Return PlannedList(listIndex)
                'for sorted field cards the handler (i.e. this function) handles the blanks
            Case prStoryCard.Equipment
                If listIndex > getCardEventNumber(cardNumber) - 1 Then
                    Return tmpDTevent
                Else
                    EquipMainList(listIndex).DTpct = schedTime
                    Return EquipMainList(listIndex)
                End If
            Case prStoryCard.Changeover
                If listIndex > getCardEventNumber(cardNumber) - 1 Then
                    Return tmpDTevent
                Else
                    ChangeoverList(listIndex).DTpct = schedTime
                    Return ChangeoverList(listIndex)
                End If
            Case prStoryCard.Equipment_One
                Debugger.Break()
                Return getEquipmentEventForSingleUnitOp(0, listIndex)
            Case prStoryCard.Equipment_Two
                Debugger.Break()
                Return getEquipmentEventForSingleUnitOp(1, listIndex)
            Case prStoryCard.Equipment_Three
                Debugger.Break()
                Return getEquipmentEventForSingleUnitOp(2, listIndex)

            Case prStoryCard.Bulk
                If listIndex > getCardEventNumber(cardNumber) - 1 Then
                    Return tmpDTevent
                Else
                    BulkList(listIndex).DTpct = schedTime
                    Return BulkList(listIndex)
                End If
            Case prStoryCard.Materials
                If listIndex > getCardEventNumber(cardNumber) - 1 Then
                    Return tmpDTevent
                Else
                    MaterialsList(listIndex).DTpct = schedTime
                    Return MaterialsList(listIndex)
                End If
            Case prStoryCard.Stops
                If listIndex > getCardEventNumber(cardNumber) - 1 Then
                    Return tmpDTevent
                Else
                    TopStopsList(listIndex).DTpct = schedTime
                    TopStopsList(listIndex).MTBF = MainLEDSReport.UT_DT
                    Return TopStopsList(listIndex)
                End If
            Case Else
                Throw New unknownprstoryCardException("Card Number: " & cardNumber)
        End Select
    End Function

    Private Function getEquipmentEventForSingleUnitOp(equipNo As Integer, listIndex As Integer) As DTevent
        Dim tmpDTevent As DTevent
        Dim testName As String
        tmpDTevent = New DTevent("", 0) 'this is the blank event we return when we're being ...
        tmpDTevent.Stops = 0 '...asked for an event number thats outside the actual idex
        If equipNo > EquipMainList.Count - 1 Then Return tmpDTevent 'EquipTargetList.Count - 1 Then Return tmpDTevent
        testName = EquipMainList(equipNo).Name ' EquipTargetList(equipNo).Name
        '  If testName.Equals(EquipOneName) Then
        'check if its there
        If listIndex > EquipOneList.Count - 1 Then
            Return tmpDTevent
        Else 'find it
            EquipOneList(listIndex).DTpct = schedTime
            Return EquipOneList(listIndex)
        End If
        '   ElseIf testName.Equals(EquipTwoName) Then
        'check if its there
        If listIndex > EquipTwoList.Count - 1 Then
            Return tmpDTevent
        Else 'find it
            EquipTwoList(listIndex).DTpct = schedTime
            Return EquipTwoList(listIndex)
        End If
        '  ElseIf testName.Equals(EquipThreeName) Then
        'check if its there
        If listIndex > EquipThreeList.Count - 1 Then
            Return tmpDTevent
        Else 'find it
            EquipThreeList(listIndex).DTpct = schedTime
            Return EquipThreeList(listIndex)
        End If

        '  Else
        Return tmpDTevent
        '  End If
    End Function

    '   Public Function getListIndexFromName(cardNumber As Integer, fieldName As String)
    '   Select Case cardNumber
    '  Case prStoryCard.Unplanned
    '   Return UnplannedList.IndexOf(New DTevent(fieldName, 0))
    ' '  Case prStoryCard.Planned
    '   Return PlannedList.IndexOf(New DTevent(fieldName, 0))
    '   Case prStoryCard.Equipment
    '   Return EquipMainList.IndexOf(New DTevent(fieldName, 0))
    '  Case prStoryCard.Equipment_One
    '  Return EquipOneList.IndexOf(New DTevent(fieldName, 0))
    ' Case prStoryCard.Equipment_Two
    ' Return EquipTwoList.IndexOf(New DTevent(fieldName, 0))
    ' Case prStoryCard.Equipment_Three
    ' Return EquipThreeList.IndexOf(New DTevent(fieldName, 0))
    ' Case prStoryCard.Bulk
    ' Return BulkList.IndexOf(New DTevent(fieldName, 0))
    ' Case prStoryCard.Materials
    ' Return MaterialsList.IndexOf(New DTevent(fieldName, 0))
    ' Case prStoryCard.Changeover
    ' Return ChangeoverList.IndexOf(New DTevent(fieldName, 0))
    ' Case Else
    ' Throw New unknownprstoryCardException("Card Number: " & cardNumber)
    ' End Select
    ' End Function

    Public Function getCardEventFields(cardNumber As Integer) As Integer
        Select Case cardNumber
            Case prStoryCard.Unplanned
                Return prStoryCardFields.Unplanned
            Case prStoryCard.Planned
                Return prStoryCardFields.Planned
            Case prStoryCard.Equipment
                Return prStoryCardFields.Equipment
            Case prStoryCard.Equipment_One
                Return prStoryCardFields.Equipment_One
            Case prStoryCard.Equipment_Two
                Return prStoryCardFields.Equipment_Two
            Case prStoryCard.Equipment_Three
                Return prStoryCardFields.Equipment_Three
            Case prStoryCard.Bulk
                Return prStoryCardFields.Bulk
            Case prStoryCard.Materials
                Return prStoryCardFields.Materials
            Case prStoryCard.Stops
                Return prStoryCardFields.Stops
            Case prStoryCard.TopThree
                Return prStoryCardFields.TopThree
            Case prStoryCard.Changeover
                Return prStoryCardFields.Changeover
            Case Else
                Throw New unknownprstoryCardException("Card Number: " & cardNumber)
        End Select
    End Function

    Public Function getCardEventNumber(cardNumber As Integer) As Integer
        Select Case cardNumber
            Case prStoryCard.Unplanned
                Return AllProdLines(_ParentLine).UnplannedT1List.Count'getprStoryCard_NumberOfFixedFields(AllProductionLines(_ParentLine).prStoryMapping, prStoryCard.Unplanned)
                'Return UnplannedList.Count
            Case prStoryCard.Planned
                Return 9
                ' Return PlannedList.Count
            Case prStoryCard.Equipment
                Return EquipMainList.Count
            Case prStoryCard.Equipment_One
                Return EquipOneList.Count
            Case prStoryCard.Equipment_Two
                Return EquipTwoList.Count
            Case prStoryCard.Equipment_Three
                Return EquipThreeList.Count
            Case prStoryCard.Bulk
                Return BulkList.Count
            Case prStoryCard.Materials
                Return MaterialsList.Count
            Case prStoryCard.Stops 'top 25 stops
                Return TopStopsList.Count
            Case prStoryCard.TopThree
                Return TopThreeList.Count
            Case prStoryCard.Changeover
                Return ChangeoverList.Count
            Case Else
                Throw New unknownprstoryCardException
        End Select
    End Function

    Public Function getCardName(cardNumber As Integer) As String
        Select Case cardNumber
            Case prStoryCard.Unplanned
                Return "UNPLANNED LOSSES"
            Case prStoryCard.Planned
                Return "PLANNED LOSSES"
            Case prStoryCard.Equipment
                Return "Equipment"
            Case prStoryCard.Equipment_One
                '  Return EquipOneName
                If EquipMainList.Count > 0 Then
                    Return EquipMainList(0).Name
                Else
                    Return ""
                End If
            Case prStoryCard.Equipment_Two
                'Return EquipTwoName
                If EquipMainList.Count > 1 Then
                    Return EquipMainList(1).Name
                Else
                    Return ""
                End If
            Case prStoryCard.Equipment_Three
                'Return EquipThreeName
                If EquipMainList.Count > 2 Then
                    Return EquipMainList(2).Name
                Else
                    Return ""
                End If
            Case prStoryCard.Bulk
                Return "PASTE"
            Case prStoryCard.Materials
                Return "PACKING MATERIALS"
            Case prStoryCard.Changeover
                Return "CHANGEOVERS"
            Case Else
                Throw New unknownprstoryCardException
        End Select
    End Function

    '  Public Function getCardFieldName(cardNumber As Integer, fieldNumber As Integer) As String
    '  Return getprStoryCardField(AllProductionLines(_ParentLine).prStoryMapping, cardNumber, fieldNumber)
    ' End Function
#End Region

#Region "Construction / Reinitialization"
    'constructor
    Public Sub New(ParentLineIndex As Integer, ByRef startTime As Date, ByRef endTime As Date)
        _ParentLine = ParentLineIndex
        _ColumnToMap = My.Settings.defaultDownTimeField
        _startTime = startTime
        _endTime = endTime
        MainLEDSReport = New SummaryReport(AllProdLines(_ParentLine), _startTime, _endTime)
        _prstoryMapping = AllProdLines(_ParentLine).prStoryMapping

        startTime = _startTime
        endTime = _endTime
        ProductList = MainLEDSReport.DT_Report.getFilterList(DowntimeField.Product) 'MainLEDSReport.DT_Report.ActiveProducts
        GCASlist = MainLEDSReport.DT_Report.ActiveGCAS
    End Sub

    Public Sub New(ParentLineIndex As Integer, ByRef startTime As Date, ByRef endTime As Date, useDoverData As Boolean)
        _ParentLine = ParentLineIndex
        _ColumnToMap = My.Settings.defaultDownTimeField
        _startTime = startTime
        _endTime = endTime
        MainLEDSReport = New SummaryReport(AllProdLines(_ParentLine), _startTime, _endTime, True)
        _prstoryMapping = AllProdLines(_ParentLine).prStoryMapping

        startTime = _startTime
        endTime = _endTime
        ProductList = New List(Of String)
        GCASlist = New List(Of String)

        For Each x As DowntimeReport In MainLEDSReport.DT_Reports
            ProductList.AddRange(x.getFilterList(DowntimeField.Product))
        Next
        For Each x As DowntimeReport In MainLEDSReport.DT_Reports
            GCASlist.AddRange(x.ActiveGCAS)
        Next

        ProductList = ProductList.Distinct.ToList()
        GCASlist = GCASlist.Distinct.ToList()
    End Sub



    Public Sub reMapReport()
        '  clearAllLists()
        TopStopsList.Clear()
        TopThreeList.Clear()
        MainLEDSReport.reMapDowntime(My.Settings.defaultDownTimeField, My.Settings.defaultDownTimeField_Secondary)
        With MainLEDSReport.DT_Report
            TopStopsList = .UnplannedEventDirectory
            _eventMaxDT = 0
            For i = 0 To TopStopsList.Count - 1
                TopStopsList(i).DTpct = schedTime
                TopStopsList(i).MTBF = MainLEDSReport.UT_DT
                If TopStopsList(i).DT > _eventMaxDT Then _eventMaxDT = TopStopsList(i).DT
            Next
            sortEventList_ByStops(TopStopsList)
            updateCardList_Stops(0)
        End With
    End Sub

    Public ProductList As New List(Of String)
    Public GCASlist As New List(Of String)

    Public Sub updateProductList(FilterSelection As Integer)
        ProductList.Clear()
        ProductList = MainLEDSReport.DT_Report.getFilterList(FilterSelection)
    End Sub


    Public Sub reFilterData_SKU(inclustionList As List(Of String))
        clearAllLists()
        MainLEDSReport.reFilterDowntime_SKU(inclustionList)
        reFilterData_Finalize()
    End Sub
    Public Sub reFilterData_Team(inclustionList As List(Of String), Optional Isitforteamanalysis As Boolean = False)
        clearAllLists()
        MainLEDSReport.reFilterDowntime_Team(inclustionList)
        If Isitforteamanalysis = False Then reFilterData_Finalize()
    End Sub
    Public Sub reFilterData_Format(inclustionList As List(Of String))
        clearAllLists()
        MainLEDSReport.reFilterDowntime_Format(inclustionList)
        reFilterData_Finalize()
    End Sub
    Public Sub reFilterData_Shape(inclustionList As List(Of String))
        clearAllLists()
        MainLEDSReport.reFilterDowntime_Shape(inclustionList)
        reFilterData_Finalize()
    End Sub
    Public Sub reFilterData_ProductGroup(inclustionList As List(Of String))
        clearAllLists()
        MainLEDSReport.reFilterDowntime_ProductGroup(inclustionList)
        reFilterData_Finalize()
    End Sub
    Public Sub reFilterData_ClearAllFilters()
        clearAllLists()
        MainLEDSReport.reFilterDowntime_ClearAllFilters()
        reFilterData_Finalize()
    End Sub
    Private Sub reFilterData_Finalize()
        InitializeBargraphWindowConnection_createAllLists()
    End Sub

    Private Sub clearAllLists()
        UnplannedList.Clear()
        PlannedList.Clear()
        ChangeoverList.Clear()
        EquipMainList.Clear()
        EquipOneList.Clear()
        EquipTwoList.Clear()
        EquipThreeList.Clear()
        TopThreeList.Clear()
    End Sub

    Public Sub InitializeBargraphWindowConnection_createAllLists()
        clearAllLists()
        With MainLEDSReport.DT_Report
            UnplannedList = .getUnplannedEventDirectory(DowntimeField.Tier1, False)
            PlannedList = .getPlannedEventDirectory(DowntimeField.Tier1, False)
            TopStopsList = .UnplannedEventDirectory

            _eventMaxDT = 0
            For i = 0 To TopStopsList.Count - 1
                TopStopsList(i).DTpct = schedTime
                TopStopsList(i).MTBF = MainLEDSReport.UT_DT
                If TopStopsList(i).DT > _eventMaxDT Then _eventMaxDT = TopStopsList(i).DT
            Next
            sortEventList_ByStops(TopStopsList)
        End With

        'add the totals
        Dim totalUPDT As New DTevent("Total", 0, -1, _ColumnToMap)
        Dim totalPDT As New DTevent("Total", 0, -1, _ColumnToMap)
        For i As Integer = 0 To MainLEDSReport.DT_Report.rawDTdata.UnplannedData.Count - 1
            totalUPDT.addStopWithRow(MainLEDSReport.DT_Report.rawDTdata.UnplannedData(i).DT, i) ' += UnplannedList(i).DT
        Next
        For i As Integer = 0 To MainLEDSReport.DT_Report.rawDTdata.PlannedData.Count - 1 'PlannedList.Count - 1
            totalPDT.addStopWithRow(MainLEDSReport.DT_Report.rawDTdata.PlannedData(i).DT, i) ' PlannedList(i).DT, i)
        Next

        'SORTING!!!
        '  UnplannedList.Sort()
        '  PlannedList.Sort()

        UnplannedList.Insert(0, totalUPDT)
        PlannedList.Insert(0, totalPDT)

        For i = 0 To UnplannedList.Count - 1
            UnplannedList(i).DTpct = schedTime
            UnplannedList(i).MTBF = MainLEDSReport.UT_DT
        Next
        For i = 0 To PlannedList.Count - 1
            PlannedList(i).DTpct = schedTime
        Next

        'send to the window!
        updateCardList_Unplanned_Tier1(0)
        updateCardList_Planned_Tier1(0)
        updateCardList_Stops(0)
    End Sub

    Public Sub PlannedVariance_GenerateHTML()
        'create PDT HTML stuff!
        Dim Plandlesticks As Thread
        Plandlesticks = New Thread(AddressOf PDT_candlesticks_Thread)
        Plandlesticks.Start()
    End Sub
    Private Sub PDT_candlesticks_Thread()
        Dim i As Integer
        Dim PlannedReports As New List(Of PDTeventReport)
        For i = 1 To ChangeoverList.Count - 1
            PlannedReports.Add(New PDTeventReport(ChangeoverList(i), AllProdLines(ParentLineInt).rawDowntimeData.PlannedData))
        Next
        HTML_exportPDTcandlesticks(PlannedReports, 1)
    End Sub
#End Region

#Region "Equality & Comparable Overriders"

    Public Function CompareTo(ByVal Other As prStoryMainPageReport) As Integer Implements System.IComparable(Of prStoryMainPageReport).CompareTo
        Return Me._ParentLine.CompareTo(Other._ParentLine)
    End Function



    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As prStoryMainPageReport = TryCast(obj, prStoryMainPageReport)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As prStoryMainPageReport) As Boolean _
        Implements IEquatable(Of prStoryMainPageReport).Equals
        If other Is Nothing Then
            Return False
        End If
        If (Me.ParentLineInt.Equals(other.ParentLineInt)) Then
            If Me.StartDate.Equals(other.StartDate) And Me.EndDate.Equals(other.EndDate) Then
                Return True
            End If
        End If
        Return False
    End Function
#End Region

End Class







Public Class prStoryStopsWatch_24
    Implements IEquatable(Of prStoryStopsWatch_24)
    Private _DayStart As Date
    Private _CurrentMappingCol As Integer
    Private parentLine As ProdLine
    Private ActiveModeStops As New List(Of Integer)
    Friend HourlyReports As New List(Of DowntimeReport)

    Private LastDataTime As Date

    Public ReadOnly Property CurrentDay As Date
        Get
            Return _DayStart
        End Get
    End Property

    Public Function isChangeover(targetHour As Integer) As Boolean
        If HourlyReports(targetHour - 1).ChangeoversNum > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function isCIL(targetHour As Integer) As Boolean
        If HourlyReports(targetHour - 1).PDT = 0 Then Return False
        If HourlyReports(targetHour - 1).CILsNum > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function getSkus(targetHour As Integer) As String
        Dim i As Integer, tmpString As String = ""
        With HourlyReports(targetHour - 1)
            If .ActiveProducts.Count = 0 Then
                Return "No Production Listed"
            ElseIf .ActiveProducts.Count = 1 Then
                Return .ActiveProducts(0)
            Else
                For i = 0 To .ActiveProducts.Count - 2
                    tmpString = tmpString & .ActiveProducts(i) & ", "
                Next
                tmpString = tmpString & .ActiveProducts(.ActiveProducts.Count - 1)
            End If
        End With
        Return tmpString
    End Function
    Public Function getAvailability(targetHour As Integer) As Double
        Return HourlyReports(targetHour - 1).Availability
    End Function
    Public Function getUPDT(targetHour As Integer) As Double
        Dim tmpSchedTime As Double
        tmpSchedTime = HourlyReports(targetHour - 1).schedTime
        If tmpSchedTime > 0 Then
            Return HourlyReports(targetHour - 1).UPDT / tmpSchedTime
        Else
            Return 0
        End If
    End Function
    Public Function getPDT(targetHour As Integer) As Double
        Dim tmpSchedTime As Double
        tmpSchedTime = HourlyReports(targetHour - 1).schedTime
        If tmpSchedTime > 0 Then
            Return HourlyReports(targetHour - 1).PDT / tmpSchedTime
        Else
            Return 0
        End If
    End Function
    Public Function isPRout(targethour As Integer) As Boolean
        If HourlyReports(targethour - 1).schedTime < 4 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function getNetStops(targetHour As Integer) As Integer
        Return HourlyReports(targetHour - 1).Stops
    End Function
    Public Function getModeStops(targetHour As Integer) As Integer
        Return ActiveModeStops(targetHour - 1)
    End Function
    Public ReadOnly Property MaxHourStops As Integer
        Get
            Return ActiveModeStops.Max
        End Get
    End Property

    Public Sub setCurrentFailureMode(modeName As String, Optional modeCol As Integer = DowntimeField.DTGroup)
        If Not _CurrentMappingCol = modeCol Then reMapStopsWatch(modeCol)
        Dim hourIncrementer As Integer
        ActiveModeStops.Clear()

        If modeName.Equals("All failure modes") Then
            For hourIncrementer = 1 To 24
                ActiveModeStops.Add(HourlyReports(hourIncrementer - 1).Stops)
            Next
        Else
            For hourIncrementer = 1 To 24
                ActiveModeStops.Add(getEventByMode(modeName, hourIncrementer, modeCol).Stops)
            Next
        End If

    End Sub
    Public Function getEventByMode(modeName As String, targetHour As Integer, Optional modeCol As Integer = DowntimeField.DTGroup) As DTevent
        Dim i As Integer
        Dim tmpEvent As DTevent
        If Not _CurrentMappingCol = modeCol Then reMapStopsWatch(modeCol)
        i = HourlyReports(targetHour - 1).UnplannedEventDirectory.IndexOf(New DTevent(modeName, 0)) ' HourlyReports(targetHour - 1).isInDTList(modeName)  'SRONEW  
        If i > -1 Then
            Return HourlyReports(targetHour - 1).UnplannedEventDirectory(i)
        Else
            tmpEvent = New DTevent(modeName, 0, 0, modeCol)
            tmpEvent.Stops = 0
            Return tmpEvent
        End If
    End Function

    Public Sub New(parentLineIndex As Integer, targetDay As Date, Optional targetMappingColumn As Integer = DowntimeField.DTGroup, Optional HrsToAnalyze As Integer = 24)
        Dim dayIncrementer As Integer
        parentLine = AllProdLines(parentLineIndex)
        _CurrentMappingCol = targetMappingColumn
        'find midnight on the target day
        LastDataTime = parentLine.rawDowntimeData.rawConstraintData(parentLine.rawDowntimeData.rawConstraintData.Count - 1).endTime


        _DayStart = DateAdd(DateInterval.Second, -Second(targetDay), targetDay)
        _DayStart = DateAdd(DateInterval.Minute, -Minute(_DayStart), _DayStart)
        _DayStart = DateAdd(DateInterval.Hour, -Hour(_DayStart), _DayStart)
        For dayIncrementer = 0 To HrsToAnalyze - 1 'the number of hours in a day...
            If DateAdd(DateInterval.Hour, dayIncrementer, _DayStart) >= LastDataTime Then
                HourlyReports.Add(New DowntimeReport(parentLine, True))
            Else
                HourlyReports.Add(New DowntimeReport(parentLine, DateAdd(DateInterval.Hour, dayIncrementer, _DayStart), DateAdd(DateInterval.Hour, dayIncrementer + 1, _DayStart)))
            End If
        Next
    End Sub

    'analyze the data by a different failure mode type
    Public Sub reMapStopsWatch(newMappingColumn As Integer)
        Dim i As Integer
        _CurrentMappingCol = newMappingColumn
        For i = 0 To HourlyReports.Count - 1
            'SRONEW HourlyReports(i).reAnalyzeDirectories(_CurrentMappingCol)
        Next
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As prStoryStopsWatch_24 = TryCast(obj, prStoryStopsWatch_24)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As prStoryStopsWatch_24) As Boolean _
        Implements IEquatable(Of prStoryStopsWatch_24).Equals
        If other Is Nothing Then
            Return False
        End If
        If DateDiff(DateInterval.Minute, Me.CurrentDay, other.CurrentDay) < 10 And DateDiff(DateInterval.Minute, Me.CurrentDay, other.CurrentDay) > -1 Then
            Return 0
        ElseIf DateDiff(DateInterval.Minute, Me.CurrentDay, other.CurrentDay) < 10 Then
            Return -1
        Else : Return 1
        End If
    End Function
End Class