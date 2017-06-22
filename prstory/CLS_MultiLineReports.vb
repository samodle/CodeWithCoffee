Imports DigitalFactory

Public Class CLS_MultiLineReports
    Implements IEquatable(Of CLS_MultiLineReports)

#Region "Variables and Properties"

    Private _temp_msu_or_schedtime As Double = 0.0

    Private _AvgPR As Double = 0.0
    Private _AvgUPDT As Double = 0.0
    Private _AvgPDT As Double = 0.0
    Private _AvgSPD As Double = 0.0
    Private _AvgMTBF As Double = 0.0
    Private _AvgMTTR As Double = 0.0
    Private _AvgMSU As Double = 0.0
    Private _AvgActualStops As Double = 0.0
    Private _Avgschedtime As Double = 0.0

    Private _Tier1ListofLossTreesofeachline As New List(Of List(Of DTevent))
    Private _Tier2ListofLossTreesofeachline As New List(Of List(Of DTevent))
    Private _Tier3ListofLossTreesofeachline As New List(Of List(Of DTevent))
    Private _DTGroupListofLossTreesofeachline As New List(Of List(Of DTevent))

    Private _Tier1ListofLossTreesofeachlineplanned As New List(Of List(Of DTevent))
    Private _Tier2ListofLossTreesofeachlineplanned As New List(Of List(Of DTevent))
    Private _Tier3ListofLossTreesofeachlineplanned As New List(Of List(Of DTevent))
    Private _DTGroupListofLossTreesofeachlineplanned As New List(Of List(Of DTevent))

    Private _RollupTier1Loss As New List(Of DTevent)
    Private _RollupTier2Loss As New List(Of DTevent)
    Private _RollupTier3Loss As New List(Of DTevent)
    Private _RollupDTGroupLoss As New List(Of DTevent)

    Private _RollupTier1Lossplanned As New List(Of DTevent)
    Private _RollupTier2Lossplanned As New List(Of DTevent)
    Private _RollupTier3Lossplanned As New List(Of DTevent)
    Private _RollupDTGroupLossplanned As New List(Of DTevent)

    public UptimeList as new list(of double)
    Public _ListofPR As New List(Of Double)
    Public _ListofUPDT As New List(Of Double)
    Public _ListofPDT As New List(Of Double)
    Public _ListofSPD As New List(Of Double)
    Public _ListofActualStops As New List(Of Double)
    Public _ListofMTBF As New List(Of Double)
    Public _ListofMTTR As New List(Of Double)
    Public _ListofMSU As New List(Of Double)
    Public _ListofCases As New List(Of Double)
    Public _ListofRateLoss As New List(Of Double)
    Public _ListofLineNames As New List(Of String)
    Public _ListofSchedTime As New List(Of Double)

    Public ListOfAdjustedUnits As New List(Of Double)

    Public _ListofFailurenamesAllLines As New List(Of String)
    Public _ListofFailureDTpctAllLines As New List(Of Double)
    Public _ListofFailureSPDAllLines As New List(Of Double)
    Public _ListofFailureActualstopsAllLines As New List(Of Double)
    Public _ListofFailureMTBFAllLines As New List(Of Double)
    Public _ListofFailureMTTRAllLines As New List(Of Double)

    Public _ListofEndDates As New List(Of DateTime)

    Friend FaultDirectory As New List(Of DTevent)
    Friend Reason1Directory As New List(Of DTevent)


    Friend LocationDirectory As New List(Of DTevent)

    Friend Reason2Directory As New List(Of DTevent)
    Friend Reason3Directory As New List(Of DTevent)
    Friend Reason4Directory As New List(Of DTevent)

    Friend TeamDirectory As New List(Of DTevent)
    Friend SKUDirectory As New List(Of DTevent)
    Friend Tier1Directory As New List(Of DTevent)
    Friend Tier2Directory As New List(Of DTevent)
    Friend Tier3Directory As New List(Of DTevent)
    Friend DTgroupDirectory As New List(Of DTevent)

    Friend Tier1Directoryplanned As New List(Of DTevent)
    Friend Tier2Directoryplanned As New List(Of DTevent)
    Friend Tier3Directoryplanned As New List(Of DTevent)
    Friend DTgroupDirectoryplanned As New List(Of DTevent)

    Friend ExportList As New List(Of DTevent)

    Public _prstoryrawreportlist As New List(Of prStoryMainPageReport)
    Public _multilineindeceslists As New List(Of Integer)


    Public ReadOnly Property AvgPR As Double
        Get
            Return _AvgPR
        End Get
    End Property
    Public ReadOnly Property AvgUPDT As Double
        Get
            Return _AvgUPDT
        End Get
    End Property
    Public ReadOnly Property AvgPDT As Double
        Get
            Return _AvgPDT
        End Get
    End Property
    Public ReadOnly Property AvgSPD As Double
        Get
            Return _AvgSPD
        End Get
    End Property
    Public ReadOnly Property AvgActualStops As Double
        Get
            Return _AvgActualStops
        End Get
    End Property
    Public ReadOnly Property AvgMSU As Double
        Get
            Return _AvgMSU
        End Get
    End Property
    Public ReadOnly Property AvgMTBF As Double
        Get
            Return _AvgMTBF
        End Get
    End Property
    Public ReadOnly Property AvgMTTR As Double
        Get
            Return _AvgMTTR
        End Get
    End Property

    Public ReadOnly Property AvgSchedTime As Double
        Get
            Return _Avgschedtime
        End Get
    End Property
    Public ReadOnly Property PRList As List(Of Double)
        Get
            Return _ListofPR
        End Get
    End Property
    Public ReadOnly Property UPDTList As List(Of Double)
        Get
            Return _ListofUPDT
        End Get
    End Property
    Public ReadOnly Property PDTList As List(Of Double)
        Get
            Return _ListofPDT
        End Get
    End Property
    Public ReadOnly Property SPDList As List(Of Double)
        Get
            Return _ListofSPD
        End Get
    End Property
    Public ReadOnly Property ActualStopsList As List(Of Double)
        Get
            Return _ListofActualStops
        End Get
    End Property
    Public ReadOnly Property MTBFList As List(Of Double)
        Get
            Return _ListofMTBF
        End Get
    End Property
    Public ReadOnly Property MTTRList As List(Of Double)
        Get
            Return _ListofMTTR
        End Get
    End Property
    Public ReadOnly Property MSUList As List(Of Double)
        Get
            Return _ListofMSU
        End Get
    End Property
    Public ReadOnly Property CasesList As List(Of Double)
        Get
            Return _ListofCases
        End Get
    End Property
    Public ReadOnly Property RateLossList As List(Of Double)
        Get
            Return _ListofRateLoss
        End Get
    End Property
    Public ReadOnly Property LineNamesList As List(Of String)
        Get
            Return _ListofLineNames
        End Get
    End Property
    Public ReadOnly Property SchedTimeList As List(Of Double)
        Get
            Return _ListofSchedTime
        End Get
    End Property
    Public ReadOnly Property EndDateList As List(Of DateTime)
        Get
            Return _ListofEndDates
        End Get
    End Property


    Public ReadOnly Property RollupTier1 As List(Of DTevent)
        Get
            Return _RollupTier1Loss
        End Get

    End Property
    Public ReadOnly Property RollupTier2 As List(Of DTevent)
        Get
            Return _RollupTier2Loss
        End Get

    End Property
    Public ReadOnly Property RollupTier3 As List(Of DTevent)
        Get
            Return _RollupTier3Loss
        End Get

    End Property
    Public ReadOnly Property RollupDTGroup As List(Of DTevent)
        Get
            Return _RollupDTGroupLoss
        End Get

    End Property
    Public ReadOnly Property RollupTier1planned As List(Of DTevent)
        Get
            Return _RollupTier1Lossplanned
        End Get

    End Property
    Public ReadOnly Property RollupTier2planned As List(Of DTevent)
        Get
            Return _RollupTier2Lossplanned
        End Get

    End Property
    Public ReadOnly Property RollupTier3planned As List(Of DTevent)
        Get
            Return _RollupTier3Lossplanned
        End Get

    End Property
    Public ReadOnly Property RollupDTGroupplanned As List(Of DTevent)
        Get
            Return _RollupDTGroupLossplanned
        End Get

    End Property


    Public ReadOnly Property Tier1ListofLossTreeosfEachLine As List(Of List(Of DTevent))
        Get
            Return _Tier1ListofLossTreesofeachline
        End Get

    End Property
    Public ReadOnly Property Tier2ListofLossTreeosfEachLine As List(Of List(Of DTevent))
        Get
            Return _Tier2ListofLossTreesofeachline
        End Get

    End Property
    Public ReadOnly Property Tier3ListofLossTreeosfEachLine As List(Of List(Of DTevent))
        Get
            Return _Tier3ListofLossTreesofeachline
        End Get

    End Property
    Public ReadOnly Property DTGroupListofLossTreeosfEachLine As List(Of List(Of DTevent))
        Get
            Return _DTGroupListofLossTreesofeachline
        End Get

    End Property

    Public ReadOnly Property Tier1ListofLossTreeosfEachLineplanned As List(Of List(Of DTevent))
        Get
            Return _Tier1ListofLossTreesofeachlineplanned
        End Get

    End Property
    Public ReadOnly Property Tier2ListofLossTreeosfEachLineplanned As List(Of List(Of DTevent))
        Get
            Return _Tier2ListofLossTreesofeachlineplanned
        End Get

    End Property
    Public ReadOnly Property Tier3ListofLossTreeosfEachLineplanned As List(Of List(Of DTevent))
        Get
            Return _Tier3ListofLossTreesofeachlineplanned
        End Get

    End Property
    Public ReadOnly Property DTGroupListofLossTreeosfEachLineplanned As List(Of List(Of DTevent))
        Get
            Return _DTGroupListofLossTreesofeachlineplanned
        End Get

    End Property

    Public ReadOnly Property AllLinesFailureModesNames As List(Of String)
        Get

            Return _ListofFailurenamesAllLines
        End Get

    End Property
    Public ReadOnly Property AllLinesFailureDTpct As List(Of Double)
        Get

            Return _ListofFailureDTpctAllLines
        End Get

    End Property
    Public ReadOnly Property AllLinesFailureSPD As List(Of Double)
        Get

            Return _ListofFailureSPDAllLines
        End Get

    End Property
    Public ReadOnly Property AllLinesFailureActualStops As List(Of Double)
        Get

            Return _ListofFailureActualstopsAllLines
        End Get

    End Property
    Public ReadOnly Property AllLinesFailureMTBF As List(Of Double)
        Get

            Return _ListofFailureMTBFAllLines
        End Get

    End Property
    Public ReadOnly Property AllLinesFailureMTTR As List(Of Double)
        Get

            Return _ListofFailureMTTRAllLines
        End Get

    End Property
#End Region
    Public Sub New(ByRef multiline_rawreports As List(Of prStoryMainPageReport), ByRef multilineindeces As List(Of Integer))
        ' multiline_rawreports.Sort() 'ADDED TO SORT TEAMS BY LETTER
        _prstoryrawreportlist = Nothing
        _multilineindeceslists = Nothing
        _prstoryrawreportlist = multiline_rawreports
        _multilineindeceslists = multilineindeces
#If DEBUG Then
        GenerateKPIs()
#Else
          

        Try
           
        Catch ex As Exception
            MessageBox.Show("We're sorry, prstory has encountered the following error: " & ex.Message, "Error")
        End Try
#End If
    End Sub
    Private Sub GenerateKPIs()

        Dim i As Integer
        Dim j As Integer = 0
        Dim tmpindex As Integer = -1
        Dim tempprstoryreport As prStoryMainPageReport
        Dim tempAvgPR_numerator As Double = 0
        Dim tempAvgUPDT_numerator As Double = 0
        Dim tempAvgPDT_numerator As Double = 0
        Dim tempAvgSPD_numerator As Double = 0
        Dim tempAvgSPD_denominator As Double = 0
        Dim tempAvgActualStops_numerator As Double = 0
        Dim tempAvgMTBF_numerator As Double = 0
        Dim tempAvgMTTR_numerator As Double = 0
        Dim tempAvgMTBF_denominator As Double = 0
        Dim tempAvgPR_denominator As Double = 0
        Dim tempAvgschedtime_numerator As Double = 0
        Dim endtimeofreport As DateTime

        Dim updtMinList As New List(Of Double)
        Dim stopList As New List(Of Double)

        For i = 0 To _multilineindeceslists.Count - 1
            tempprstoryreport = _prstoryrawreportlist(i)

            endtimeofreport = tempprstoryreport.EndDate  ' this is where we get the latest revised adjusted date after report is generated

            If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                _temp_msu_or_schedtime = tempprstoryreport.MSU
            Else
                _temp_msu_or_schedtime = tempprstoryreport.schedTime
            End If

            ' tempAvgPR_numerator = tempAvgPR_numerator + (tempprstoryreport.PR * _temp_msu_or_schedtime) 'original
            tempAvgPR_numerator += _temp_msu_or_schedtime
            'tempAvgUPDT_numerator = tempAvgUPDT_numerator + (tempprstoryreport.UPDT * _temp_msu_or_schedtime) 'original
            If tempprstoryreport.PR > 0 Then
                tempAvgUPDT_numerator = tempAvgUPDT_numerator + ((tempprstoryreport.UPDT * _temp_msu_or_schedtime) / tempprstoryreport.PR)
                'tempAvgPDT_numerator = tempAvgPDT_numerator + (tempprstoryreport.PDT * _temp_msu_or_schedtime) 'original
                tempAvgPDT_numerator = tempAvgPDT_numerator + ((tempprstoryreport.PDT * _temp_msu_or_schedtime) / tempprstoryreport.PR)

                'tempAvgSPD_numerator = tempAvgSPD_numerator + (tempprstoryreport.StopsPerDay * _temp_msu_or_schedtime) 'original
                'tempAvgSPD_numerator = tempAvgSPD_numerator + ((tempprstoryreport.StopsPerDay * _temp_msu_or_schedtime) / tempprstoryreport.PR) 'new volume weighted
                tempAvgSPD_numerator = tempAvgSPD_numerator + tempprstoryreport.StopsPerDay   'not volume weighted
                tempAvgSPD_denominator += 1

            End If

            'tempAvgMTBF_numerator = tempAvgMTBF_numerator + (tempprstoryreport.MTBF * _temp_msu_or_schedtime) 'original
            tempAvgMTBF_numerator = tempAvgMTBF_numerator + tempprstoryreport.MainLEDSReport.UT
            tempAvgMTBF_denominator += tempprstoryreport.ActualStops


            tempAvgschedtime_numerator = tempAvgschedtime_numerator + tempprstoryreport.schedTime
            tempAvgActualStops_numerator = tempAvgActualStops_numerator + (tempprstoryreport.ActualStops * _temp_msu_or_schedtime)
            _ListofPR.Add(tempprstoryreport.PR)
            _ListofUPDT.Add(tempprstoryreport.UPDT)
            _ListofPDT.Add(tempprstoryreport.PDT)
            _ListofSPD.Add(tempprstoryreport.StopsPerDay)
            _ListofActualStops.Add(tempprstoryreport.ActualStops)
            _ListofMTBF.Add(tempprstoryreport.MTBF)
            UptimeList.Add(tempprstoryreport.mainledsreport.UT)
            _ListofMTTR.Add(tempprstoryreport.MTTR)
            _ListofMSU.Add(_temp_msu_or_schedtime)
            _ListofRateLoss.Add(tempprstoryreport.rateLoss)
            _ListofCases.Add(tempprstoryreport.CasesAdjusted)
            _ListofLineNames.Add(tempprstoryreport.MainLEDSReport.ParentLineName)
            _ListofSchedTime.Add(tempprstoryreport.schedTime)
            ListOfAdjustedUnits.Add(tempprstoryreport.MainLEDSReport.AdjustedUnits)
            _ListofEndDates.Add(endtimeofreport)

            updtMinList.Add(tempprstoryreport.MainLEDSReport.UPDT)
            stopList.Add(tempprstoryreport.MainLEDSReport.Stops)

            ' tempAvgPR_denominator = tempAvgPR_denominator + _temp_msu_or_schedtime 'original
            If tempprstoryreport.PR > 0 Then
                tempAvgPR_denominator += (_temp_msu_or_schedtime / tempprstoryreport.PR)
            End If

            Tier1Directory.Clear()
            Tier2Directory.Clear()
            Tier3Directory.Clear()
            DTgroupDirectory.Clear()

            Tier1Directoryplanned.Clear()
            Tier2Directoryplanned.Clear()
            Tier3Directoryplanned.Clear()
            DTgroupDirectoryplanned.Clear()

            'Getting the list of failure mode names for each rawprstory report
            'unplanned
            For j = 0 To tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData.Count - 1

                With tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData(j)


                    'tier1
                    tmpindex = Tier1Directory.IndexOf(New DTevent(.Tier1, 0))
                    If tmpindex = -1 Then
                        Tier1Directory.Add(New DTevent(.Tier1, .DT, j))
                    Else
                        Tier1Directory(tmpindex).addStopWithRow(.DT, j)
                    End If
                    'tier2
                    tmpindex = Tier2Directory.IndexOf(New DTevent(.Tier2, 0))
                    If tmpindex = -1 Then
                        Tier2Directory.Add(New DTevent(.Tier2, .DT, j, tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData(j).Tier1))
                    Else
                        Tier2Directory(tmpindex).addStopWithRow(.DT, j)
                    End If
                    'tier3
                    tmpindex = Tier3Directory.IndexOf(New DTevent(.Tier3, 0))
                    If tmpindex = -1 Then
                        'Tier3Directory.Add(New DTevent(.Tier3, .DT, j))
                        Tier3Directory.Add(New DTevent(.Tier3, .DT, j, tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData(j).Tier2))
                    Else
                        Tier3Directory(tmpindex).addStopWithRow(.DT, j)
                    End If
                    'DTGroup
                    tmpindex = DTgroupDirectory.IndexOf(New DTevent(.DTGroup, 0))
                    If tmpindex = -1 Then
                        DTgroupDirectory.Add(New DTevent(.DTGroup, .DT, j))
                    Else
                        DTgroupDirectory(tmpindex).addStopWithRow(.DT, j)
                    End If

                End With

            Next

            tmpindex = -1
            'planned
            For j = 0 To tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.PlannedData.Count - 1

                With tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.PlannedData(j)
                    'tier1
                    tmpindex = Tier1Directoryplanned.IndexOf(New DTevent(.Tier1, 0))
                    If tmpindex = -1 Then
                        Tier1Directoryplanned.Add(New DTevent(.Tier1, .DT, j))
                    Else
                        Tier1Directoryplanned(tmpindex).addStopWithRow(.DT, j)
                    End If
                    'tier2
                    tmpindex = Tier2Directoryplanned.IndexOf(New DTevent(.Tier2, 0))
                    If tmpindex = -1 Then
                        Tier2Directoryplanned.Add(New DTevent(.Tier2, .DT, j, tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.PlannedData(j).Tier1))
                    Else
                        Tier2Directoryplanned(tmpindex).addStopWithRow(.DT, j)
                    End If
                    'tier3
                    tmpindex = Tier3Directoryplanned.IndexOf(New DTevent(.Tier3, 0))
                    If tmpindex = -1 Then
                        Tier3Directoryplanned.Add(New DTevent(.Tier3, .DT, j, tempprstoryreport.MainLEDSReport.DT_Report.rawDTdata.PlannedData(j).Tier2))
                    Else
                        Tier3Directoryplanned(tmpindex).addStopWithRow(.DT, j)
                    End If
                    'DTGroup
                    tmpindex = DTgroupDirectoryplanned.IndexOf(New DTevent(.DTGroup, 0))
                    If tmpindex = -1 Then
                        DTgroupDirectoryplanned.Add(New DTevent(.DTGroup, .DT, j))
                    Else
                        DTgroupDirectoryplanned(tmpindex).addStopWithRow(.DT, j)
                    End If

                End With

            Next

            Dim k As Integer
            Dim l As Integer
            Dim founditem As Boolean = False
            '''''''UNPLANNED'''''''''''''''''''Tier 1 2 3 List Exports and Dtpct and UT calculations'''''''''''''''''''''


            Dim ExportList1 As New List(Of DTevent)
            Dim ExportList2 As New List(Of DTevent)
            Dim ExportList3 As New List(Of DTevent)
            Dim ExportList4 As New List(Of DTevent)
            Dim ExportList1temp As New List(Of DTevent)
            Dim ExportList2temp As New List(Of DTevent)
            Dim ExportList3temp As New List(Of DTevent)
            Dim ExportList4temp As New List(Of DTevent)

            'Tier 1 Export
            For k = 0 To Tier1Directory.Count - 1
                ExportList1.Add(Tier1Directory(k))
            Next
            ExportList1.Sort()
            For k = 0 To ExportList1.Count - 1
                ExportList1(k).DTpct = tempprstoryreport.schedTime
                ExportList1(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))

                ExportList1(k).DTspecial = ExportList1(k).DT
                ExportList1(k).Stopsspecial = ExportList1(k).Stops
            Next
            _Tier1ListofLossTreesofeachline.Add(ExportList1)

            'Rolling up for Tier1
            'tempAvgUPDT_numerator = tempAvgUPDT_numerator + ((tempprstoryreport.UPDT * _temp_msu_or_schedtime) / tempprstoryreport.PR) 
            For k = 0 To ExportList1.Count - 1
                ExportList1temp.Add(ExportList1(k))
                founditem = False
                For l = 0 To _RollupTier1Loss.Count - 1
                    If ExportList1temp(k).Name = _RollupTier1Loss(l).Name Then  ' Did find the item
                        _RollupTier1Loss(l).DTpctspecial = _RollupTier1Loss(l).DTpctspecial + ((ExportList1temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)
                        _RollupTier1Loss(l).SPDspecial = _RollupTier1Loss(l).SPDspecial + (ExportList1temp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupTier1Loss(l).DTspecial += ((ExportList1temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)
                        _RollupTier1Loss(l).Stopsspecial += (ExportList1temp(k).Stopsspecial * _temp_msu_or_schedtime)

                        founditem = True
                        Exit For
                    End If
                Next l

                If founditem = False Then
                    _RollupTier1Loss.Add(ExportList1temp(k))
                    _RollupTier1Loss(_RollupTier1Loss.Count - 1).DTpctspecial = _RollupTier1Loss(_RollupTier1Loss.Count - 1).DTpctspecial + ((ExportList1temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)  'need to divide it by total MSU at the end
                    _RollupTier1Loss(_RollupTier1Loss.Count - 1).SPDspecial = _RollupTier1Loss(_RollupTier1Loss.Count - 1).SPDspecial + (ExportList1temp(k).SPDspecial * _temp_msu_or_schedtime)

                    _RollupTier1Loss(_RollupTier1Loss.Count - 1).DTspecial += ((ExportList1temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)  'need to divide it by total MSU at the end
                    _RollupTier1Loss(_RollupTier1Loss.Count - 1).Stopsspecial += (ExportList1temp(k).Stopsspecial * _temp_msu_or_schedtime)

                End If
            Next k



            founditem = False

            'Tier 2 Export
            For k = 0 To Tier2Directory.Count - 1
                ExportList2.Add(Tier2Directory(k))
            Next
            ExportList2.Sort()
            For k = 0 To ExportList2.Count - 1
                ExportList2(k).DTpct = tempprstoryreport.schedTime
                ExportList2(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))

                ExportList2(k).DTspecial = ExportList2(k).DT
                ExportList2(k).Stopsspecial = ExportList2(k).Stops
            Next
            _Tier2ListofLossTreesofeachline.Add(ExportList2)

            'Rolling up for Tier2
            For k = 0 To ExportList2.Count - 1
                ExportList2temp.Add(ExportList2(k))
                founditem = False
                For l = 0 To _RollupTier2Loss.Count - 1
                    If ExportList2temp(k).Name = _RollupTier2Loss(l).Name Then  ' Did find the item
                        _RollupTier2Loss(l).DTpctspecial += ((ExportList2temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier2Loss(l).SPDspecial += (ExportList2temp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupTier2Loss(l).DTspecial += ((ExportList2temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier2Loss(l).Stopsspecial += (ExportList2temp(k).Stopsspecial * _temp_msu_or_schedtime)
                        founditem = True
                        Exit For
                    End If
                Next l

                If founditem = False Then
                    _RollupTier2Loss.Add(ExportList2temp(k))
                    _RollupTier2Loss(_RollupTier2Loss.Count - 1).DTpctspecial += ((ExportList2temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier2Loss(_RollupTier2Loss.Count - 1).SPDspecial += (ExportList2temp(k).SPDspecial * _temp_msu_or_schedtime)

                    _RollupTier2Loss(_RollupTier2Loss.Count - 1).DTspecial += ((ExportList2temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier2Loss(_RollupTier2Loss.Count - 1).Stopsspecial += (ExportList2temp(k).Stopsspecial * _temp_msu_or_schedtime)
                End If
            Next k




            founditem = False
            'Tier 3 Export
            For k = 0 To Tier3Directory.Count - 1
                ExportList3.Add(Tier3Directory(k))
            Next
            ExportList3.Sort()
            For k = 0 To ExportList3.Count - 1
                ExportList3(k).DTpct = tempprstoryreport.schedTime
                ExportList3(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))

                ExportList3(k).DTspecial = ExportList3(k).DT
                ExportList3(k).Stopsspecial = ExportList3(k).Stops

            Next
            _Tier3ListofLossTreesofeachline.Add(ExportList3)


            'Rolling up for Tier3
            For k = 0 To ExportList3.Count - 1
                ExportList3temp.Add(ExportList3(k))
                founditem = False
                For l = 0 To _RollupTier3Loss.Count - 1
                    If ExportList3temp(k).Name = _RollupTier3Loss(l).Name Then  ' Did find the item
                        _RollupTier3Loss(l).DTpctspecial += ((ExportList3temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier3Loss(l).SPDspecial += (ExportList3temp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupTier3Loss(l).DTspecial += ((ExportList3temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier3Loss(l).Stopsspecial += (ExportList3temp(k).Stopsspecial * _temp_msu_or_schedtime)
                        founditem = True
                        Exit For
                    End If
                Next l

                If founditem = False Then
                    _RollupTier3Loss.Add(ExportList3temp(k))
                    _RollupTier3Loss(_RollupTier3Loss.Count - 1).DTpctspecial += ((ExportList3temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier3Loss(_RollupTier3Loss.Count - 1).SPDspecial += (ExportList3temp(k).SPDspecial * _temp_msu_or_schedtime)

                    _RollupTier3Loss(_RollupTier3Loss.Count - 1).DTspecial += ((ExportList3temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier3Loss(_RollupTier3Loss.Count - 1).Stopsspecial += (ExportList3temp(k).Stopsspecial * _temp_msu_or_schedtime)
                End If
            Next k


            founditem = False
            'DTGroup Export
            For k = 0 To DTgroupDirectory.Count - 1
                ExportList4.Add(DTgroupDirectory(k))

                ExportList4(k).DTspecial = ExportList4(k).DT
                ExportList4(k).Stopsspecial = ExportList4(k).Stops
            Next
            ExportList4.Sort()
            For k = 0 To DTgroupDirectory.Count - 1
                DTgroupDirectory(k).DTpct = tempprstoryreport.schedTime
                DTgroupDirectory(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))
            Next
            _DTGroupListofLossTreesofeachline.Add(ExportList4)


            'Rolling up for DTGroup
            For k = 0 To ExportList4.Count - 1
                ExportList4temp.Add(ExportList4(k))
                founditem = False
                For l = 0 To _RollupDTGroupLoss.Count - 1
                    If ExportList4temp(k).Name = _RollupDTGroupLoss(l).Name Then  ' Did find the item
                        _RollupDTGroupLoss(l).DTpctspecial = _RollupDTGroupLoss(l).DTpctspecial + ((ExportList4temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupDTGroupLoss(l).SPDspecial = _RollupDTGroupLoss(l).SPDspecial + (ExportList4temp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupDTGroupLoss(l).DTspecial += ((ExportList4temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupDTGroupLoss(l).Stopsspecial += (ExportList4temp(k).Stopsspecial * _temp_msu_or_schedtime)
                        founditem = True
                        Exit For
                    End If
                Next l

                If founditem = False Then
                    _RollupDTGroupLoss.Add(ExportList4temp(k))
                    _RollupDTGroupLoss(_RollupDTGroupLoss.Count - 1).DTpctspecial += ((ExportList4temp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)  'need to divide it by total MSU at the end
                    _RollupDTGroupLoss(_RollupDTGroupLoss.Count - 1).SPDspecial += (ExportList4temp(k).SPDspecial * _temp_msu_or_schedtime)

                    _RollupDTGroupLoss(_RollupDTGroupLoss.Count - 1).DTspecial += ((ExportList4temp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)  'need to divide it by total MSU at the end
                    _RollupDTGroupLoss(_RollupDTGroupLoss.Count - 1).Stopsspecial += (ExportList4temp(k).Stopsspecial * _temp_msu_or_schedtime)
                End If
            Next k

            '''''''PLANNED'''''''''''''''''''Tier 1 2 3 List Exports and Dtpct and UT calculations'''''''''''''''''''''


            Dim ExportList1planned As New List(Of DTevent)
            Dim ExportList2planned As New List(Of DTevent)
            Dim ExportList3planned As New List(Of DTevent)
            Dim ExportList4planned As New List(Of DTevent)
            Dim ExportList1plannedtemp As New List(Of DTevent)
            Dim ExportList2plannedtemp As New List(Of DTevent)
            Dim ExportList3plannedtemp As New List(Of DTevent)
            Dim ExportList4plannedtemp As New List(Of DTevent)



            founditem = False
            'Tier 1 Export planned
            For k = 0 To Tier1Directoryplanned.Count - 1
                ExportList1planned.Add(Tier1Directoryplanned(k))
            Next
            ExportList1planned.Sort()
            For k = 0 To ExportList1planned.Count - 1
                ExportList1planned(k).DTpct = tempprstoryreport.schedTime
                ExportList1planned(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))

                ExportList1planned(k).DTspecial = ExportList1planned(k).DT
                ExportList1planned(k).Stopsspecial = ExportList1planned(k).Stops
            Next
            _Tier1ListofLossTreesofeachlineplanned.Add(ExportList1planned)

            'Rolling up for Tier1 planned
            For k = 0 To ExportList1planned.Count - 1
                ExportList1plannedtemp.Add(ExportList1planned(k))
                founditem = False
                For l = 0 To _RollupTier1Lossplanned.Count - 1
                    If ExportList1plannedtemp(k).Name = _RollupTier1Lossplanned(l).Name Then  ' Did find the item
                        _RollupTier1Lossplanned(l).DTpctspecial = _RollupTier1Lossplanned(l).DTpctspecial + ((ExportList1plannedtemp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier1Lossplanned(l).SPDspecial = _RollupTier1Lossplanned(l).SPDspecial + (ExportList1plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupTier1Lossplanned(l).Stopsspecial += (ExportList1plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)
                        _RollupTier1Lossplanned(l).DTspecial += (ExportList1plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)
                        founditem = True
                        Exit For
                    End If


                Next l

                If founditem = False Then
                    _RollupTier1Lossplanned.Add(ExportList1plannedtemp(k))
                    _RollupTier1Lossplanned(_RollupTier1Lossplanned.Count - 1).DTpctspecial = _RollupTier1Lossplanned(_RollupTier1Lossplanned.Count - 1).DTpctspecial + ((ExportList1plannedtemp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier1Lossplanned(_RollupTier1Lossplanned.Count - 1).SPDspecial = _RollupTier1Lossplanned(_RollupTier1Lossplanned.Count - 1).SPDspecial + (ExportList1plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)

                    _RollupTier1Lossplanned(_RollupTier1Lossplanned.Count - 1).DTspecial += ((ExportList1plannedtemp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier1Lossplanned(_RollupTier1Lossplanned.Count - 1).Stopsspecial += (ExportList1plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)

                End If
            Next k


            founditem = False
            'Tier 2 Export planned
            For k = 0 To Tier2Directoryplanned.Count - 1
                ExportList2planned.Add(Tier2Directoryplanned(k))
            Next
            ExportList2planned.Sort()
            For k = 0 To ExportList2planned.Count - 1
                ExportList2planned(k).DTpct = tempprstoryreport.schedTime
                ExportList2planned(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))


                ExportList2planned(k).DTspecial = ExportList2planned(k).DT
                ExportList2planned(k).Stopsspecial = ExportList2planned(k).Stops
            Next
            _Tier2ListofLossTreesofeachlineplanned.Add(ExportList2planned)

            'Rolling up for Tier2 planned
            For k = 0 To ExportList2planned.Count - 1
                ExportList2plannedtemp.Add(ExportList2planned(k))
                founditem = False
                For l = 0 To _RollupTier2Lossplanned.Count - 1
                    If ExportList2plannedtemp(k).Name = _RollupTier2Lossplanned(l).Name Then  ' Did find the item
                        _RollupTier2Lossplanned(l).DTpctspecial = _RollupTier2Lossplanned(l).DTpctspecial + ((ExportList2plannedtemp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier2Lossplanned(l).SPDspecial = _RollupTier2Lossplanned(l).SPDspecial + (ExportList2plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupTier2Lossplanned(l).DTspecial += ((ExportList2plannedtemp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier2Lossplanned(l).Stopsspecial += (ExportList2plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)
                        founditem = True
                        Exit For
                    End If


                Next l

                If founditem = False Then
                    _RollupTier2Lossplanned.Add(ExportList2plannedtemp(k))
                    _RollupTier2Lossplanned(_RollupTier2Lossplanned.Count - 1).DTpctspecial = _RollupTier2Lossplanned(_RollupTier2Lossplanned.Count - 1).DTpctspecial + ((ExportList2plannedtemp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier2Lossplanned(_RollupTier2Lossplanned.Count - 1).SPDspecial = _RollupTier2Lossplanned(_RollupTier2Lossplanned.Count - 1).SPDspecial + (ExportList2plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)

                    _RollupTier2Lossplanned(_RollupTier2Lossplanned.Count - 1).DTspecial += ((ExportList2plannedtemp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier2Lossplanned(_RollupTier2Lossplanned.Count - 1).Stopsspecial += (ExportList2plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)

                End If
            Next k



            founditem = False
            'Tier 3 Export planned
            For k = 0 To Tier3Directoryplanned.Count - 1
                ExportList3planned.Add(Tier3Directoryplanned(k))
            Next
            ExportList3planned.Sort()
            For k = 0 To ExportList3planned.Count - 1
                ExportList3planned(k).DTpct = tempprstoryreport.schedTime
                ExportList3planned(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))


                ExportList3planned(k).DTspecial = ExportList3planned(k).DT
                ExportList3planned(k).Stopsspecial = ExportList3planned(k).Stops
            Next
            _Tier3ListofLossTreesofeachlineplanned.Add(ExportList3planned)
            'Rolling up for Tier3 planned
            For k = 0 To ExportList3planned.Count - 1
                ExportList3plannedtemp.Add(ExportList3planned(k))
                founditem = False
                For l = 0 To _RollupTier3Lossplanned.Count - 1
                    If ExportList3plannedtemp(k).Name = _RollupTier3Lossplanned(l).Name Then  ' Did find the item
                        _RollupTier3Lossplanned(l).DTpctspecial = _RollupTier3Lossplanned(l).DTpctspecial + ((ExportList3plannedtemp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier3Lossplanned(l).SPDspecial = _RollupTier3Lossplanned(l).SPDspecial + (ExportList3plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupTier3Lossplanned(l).DTspecial += ((ExportList3plannedtemp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                        _RollupTier3Lossplanned(l).Stopsspecial += (ExportList3plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)

                        founditem = True
                        Exit For
                    End If


                Next l

                If founditem = False Then
                    _RollupTier3Lossplanned.Add(ExportList3plannedtemp(k))
                    _RollupTier3Lossplanned(_RollupTier3Lossplanned.Count - 1).DTpctspecial = _RollupTier3Lossplanned(_RollupTier3Lossplanned.Count - 1).DTpctspecial + ((ExportList3plannedtemp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier3Lossplanned(_RollupTier3Lossplanned.Count - 1).SPDspecial = _RollupTier3Lossplanned(_RollupTier3Lossplanned.Count - 1).SPDspecial + (ExportList3plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)

                    _RollupTier3Lossplanned(_RollupTier3Lossplanned.Count - 1).DTspecial += ((ExportList3plannedtemp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)   'need to divide it by total MSU at the end
                    _RollupTier3Lossplanned(_RollupTier3Lossplanned.Count - 1).Stopsspecial += (ExportList3plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)

                End If
            Next k


            founditem = False
            'DTGroup Export planned
            For k = 0 To DTgroupDirectoryplanned.Count - 1
                ExportList4planned.Add(DTgroupDirectoryplanned(k))


                ExportList4planned(k).DTspecial = ExportList4planned(k).DT
                ExportList4planned(k).Stopsspecial = ExportList4planned(k).Stops
            Next
            ExportList4planned.Sort()
            For k = 0 To DTgroupDirectoryplanned.Count - 1
                DTgroupDirectoryplanned(k).DTpct = tempprstoryreport.schedTime
                DTgroupDirectoryplanned(k).UT = (tempprstoryreport.schedTime - (tempprstoryreport.UPDT * tempprstoryreport.schedTime) - (tempprstoryreport.PDT * tempprstoryreport.schedTime) - (tempprstoryreport.rateLoss * tempprstoryreport.schedTime))
            Next
            _DTGroupListofLossTreesofeachlineplanned.Add(ExportList4planned)


            'Rolling up for DTGroup planned
            For k = 0 To ExportList4planned.Count - 1
                ExportList4plannedtemp.Add(ExportList4planned(k))
                founditem = False
                For l = 0 To _RollupDTGroupLossplanned.Count - 1
                    If ExportList4plannedtemp(k).Name = _RollupDTGroupLossplanned(l).Name Then  ' Did find the item
                        _RollupDTGroupLossplanned(l).DTpctspecial = _RollupDTGroupLossplanned(l).DTpctspecial + ((ExportList4plannedtemp(k).DTpctspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)  'need to divide it by total MSU at the end
                        _RollupDTGroupLossplanned(l).SPDspecial = _RollupDTGroupLossplanned(l).SPDspecial + (ExportList4plannedtemp(k).SPDspecial * _temp_msu_or_schedtime)

                        _RollupDTGroupLossplanned(l).DTpctspecial += ((ExportList4plannedtemp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR)  'need to divide it by total MSU at the end
                        _RollupDTGroupLossplanned(l).SPDspecial += (ExportList4plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)

                        founditem = True
                        Exit For
                    End If
                Next l

                If founditem = False Then
                    _RollupDTGroupLossplanned.Add(ExportList4plannedtemp(k))
                    _RollupDTGroupLossplanned(_RollupDTGroupLossplanned.Count - 1).DTspecial += ((ExportList4plannedtemp(k).DTspecial * _temp_msu_or_schedtime) / tempprstoryreport.PR) 'need to divide it by total MSU at the end
                    _RollupDTGroupLossplanned(_RollupDTGroupLossplanned.Count - 1).Stopsspecial += (ExportList4plannedtemp(k).Stopsspecial * _temp_msu_or_schedtime)


                End If
            Next k

        Next

        If tempAvgPR_denominator <> 0 Then
            _AvgPR = tempAvgPR_numerator / tempAvgPR_denominator
            _AvgUPDT = tempAvgUPDT_numerator / tempAvgPR_denominator
            _AvgPDT = tempAvgPDT_numerator / tempAvgPR_denominator
            _AvgSPD = tempAvgSPD_numerator / tempAvgSPD_denominator 'tempAvgPR_denominator
            '  _AvgMTBF = tempAvgMTBF_numerator / tempAvgPR_denominator 'original
            _AvgMTBF = tempAvgMTBF_numerator / tempAvgMTBF_denominator
            _AvgActualStops = tempAvgActualStops_numerator / tempAvgPR_denominator
            _Avgschedtime = tempAvgschedtime_numerator
            _AvgMSU = tempAvgPR_numerator

            'unplanned
            For l = 0 To _RollupTier1Loss.Count - 1
                _RollupTier1Loss(l).DTpctspecial = _RollupTier1Loss(l).DTpctspecial / tempAvgPR_denominator
                _RollupTier1Loss(l).SPDspecial = _RollupTier1Loss(l).SPDspecial / tempAvgPR_denominator

                _RollupTier1Loss(l).DTspecial = _RollupTier1Loss(l).DTspecial / tempAvgPR_denominator
                _RollupTier1Loss(l).Stopsspecial = _RollupTier1Loss(l).Stopsspecial / tempAvgPR_denominator
            Next
            For l = 0 To _RollupTier2Loss.Count - 1
                _RollupTier2Loss(l).DTpctspecial = _RollupTier2Loss(l).DTpctspecial / tempAvgPR_denominator
                _RollupTier2Loss(l).SPDspecial = _RollupTier2Loss(l).SPDspecial / tempAvgPR_denominator

                _RollupTier2Loss(l).DTspecial = _RollupTier2Loss(l).DTspecial / tempAvgPR_denominator
                _RollupTier2Loss(l).Stopsspecial = _RollupTier2Loss(l).Stopsspecial / tempAvgPR_denominator
            Next

            For l = 0 To _RollupTier3Loss.Count - 1
                _RollupTier3Loss(l).DTpctspecial = _RollupTier3Loss(l).DTpctspecial / tempAvgPR_denominator
                _RollupTier3Loss(l).SPDspecial = _RollupTier3Loss(l).SPDspecial / tempAvgPR_denominator

                _RollupTier3Loss(l).DTspecial = _RollupTier3Loss(l).DTspecial / tempAvgPR_denominator
                _RollupTier3Loss(l).Stopsspecial = _RollupTier3Loss(l).Stopsspecial / tempAvgPR_denominator
            Next

            For l = 0 To _RollupDTGroupLoss.Count - 1
                _RollupDTGroupLoss(l).DTpctspecial = _RollupDTGroupLoss(l).DTpctspecial / tempAvgPR_denominator
                _RollupDTGroupLoss(l).SPDspecial = _RollupDTGroupLoss(l).SPDspecial / tempAvgPR_denominator

                _RollupDTGroupLoss(l).DTspecial = _RollupDTGroupLoss(l).DTspecial / tempAvgPR_denominator
                _RollupDTGroupLoss(l).Stopsspecial = _RollupDTGroupLoss(l).Stopsspecial / tempAvgPR_denominator
            Next



            'planned
            For l = 0 To _RollupTier1Lossplanned.Count - 1
                _RollupTier1Lossplanned(l).DTpctspecial = _RollupTier1Lossplanned(l).DTpctspecial / tempAvgPR_denominator
                _RollupTier1Lossplanned(l).SPDspecial = _RollupTier1Lossplanned(l).SPDspecial / tempAvgPR_denominator

                _RollupTier1Lossplanned(l).DTspecial = _RollupTier1Lossplanned(l).DTspecial / tempAvgPR_denominator
                _RollupTier1Lossplanned(l).Stopsspecial = _RollupTier1Lossplanned(l).Stopsspecial / tempAvgPR_denominator
            Next
            For l = 0 To _RollupTier2Lossplanned.Count - 1
                _RollupTier2Lossplanned(l).DTpctspecial = _RollupTier2Lossplanned(l).DTpctspecial / tempAvgPR_denominator
                _RollupTier2Lossplanned(l).SPDspecial = _RollupTier2Lossplanned(l).SPDspecial / tempAvgPR_denominator

                _RollupTier2Lossplanned(l).DTspecial = _RollupTier2Lossplanned(l).DTspecial / tempAvgPR_denominator
                _RollupTier2Lossplanned(l).Stopsspecial = _RollupTier2Lossplanned(l).Stopsspecial / tempAvgPR_denominator
            Next

            For l = 0 To _RollupTier3Lossplanned.Count - 1
                _RollupTier3Lossplanned(l).DTpctspecial = _RollupTier3Lossplanned(l).DTpctspecial / tempAvgPR_denominator
                _RollupTier3Lossplanned(l).SPDspecial = _RollupTier3Lossplanned(l).SPDspecial / tempAvgPR_denominator

                _RollupTier3Lossplanned(l).DTspecial = _RollupTier3Lossplanned(l).DTspecial / tempAvgPR_denominator
                _RollupTier3Lossplanned(l).Stopsspecial = _RollupTier3Lossplanned(l).Stopsspecial / tempAvgPR_denominator
            Next
            For l = 0 To _RollupDTGroupLossplanned.Count - 1
                _RollupDTGroupLossplanned(l).DTpctspecial = _RollupDTGroupLossplanned(l).DTpctspecial / tempAvgPR_denominator
                _RollupDTGroupLossplanned(l).SPDspecial = _RollupDTGroupLossplanned(l).SPDspecial / tempAvgPR_denominator

                _RollupDTGroupLossplanned(l).DTspecial = _RollupDTGroupLossplanned(l).DTspecial / tempAvgPR_denominator
                _RollupDTGroupLossplanned(l).Stopsspecial = _RollupDTGroupLossplanned(l).Stopsspecial / tempAvgPR_denominator
            Next


        Else

        End If

        Dim tUPDTmin As Double = 0.0
        Dim tStops As Double = 0.0
        For ix As Integer = 0 To updtMinList.Count - 1
            tUPDTmin += updtMinList(ix)
            tStops += stopList(ix)
        Next ix
        If tStops > 0 Then
            _AvgMTTR = tUPDTmin / tStops
        End If

    End Sub

    Public Overloads Function Equals(other As CLS_MultiLineReports) As Boolean Implements IEquatable(Of CLS_MultiLineReports).Equals
        Throw New NotImplementedException()
    End Function
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As CLS_MultiLineReports = TryCast(obj, CLS_MultiLineReports)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
End Class
