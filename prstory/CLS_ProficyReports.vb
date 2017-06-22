Public Class SummaryReport

#Region "Variables & Properties"
    Friend DT_Reports As New List(Of DowntimeReport)
    Friend doIUseReports As Boolean = False 'this is only in here to support multiple mpu usage
    Friend DT_Report As DowntimeReport
    Friend PROD_Report As ProductionReport

    Friend ParentLine As ProdLine
    Private ParentIndex As Integer

    Public property startTime As Date
    Public property endTime As Date

    'these are the filtered criteria for this report
    Public Brandcodes As List(Of String)
    Public Teams As List(Of String)
    Public Shifts As List(Of String)
    Public Products As List(Of String)

    'all known fields
    Public BrandCodeReport As New List(Of String) 'production based
    Public ShiftReport As New List(Of String) 'dt based
    Public TeamReport As New List(Of String) 'dt based
    Public ProductReport As New List(Of String) 'dt based

    public readonly property rawData as list(of downtimeevent)
    get
            return dt_report.rawdtdata.rawconstraintdata
    End Get
    End Property

    'properties
    Public ReadOnly Property PR
        Get
            If doIUseReports Then
                Return 100 - UPDTpct - PDTpct
            Else
                If My.Settings.AdvancedSettings_isAvailabilityMode Then Return UT / schedTime
                Return PROD_Report.PR
            End If

        End Get
    End Property
    Public ReadOnly Property schedTime
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) x.schedTime)
            Else
                If My.Settings.AdvancedSettings_isAvailabilityMode Then Return DT_Report.schedTime
                Return PROD_Report.schedTime
            End If
        End Get
    End Property
    Public ReadOnly Property schedTimeDT
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) x.schedTime)
            Else
                Return DT_Report.schedTime
            End If
        End Get
    End Property
    Public ReadOnly Property UT
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) x.UT)
            Else
                If My.Settings.AdvancedSettings_isAvailabilityMode Then Return DT_Report.UT
                Return PROD_Report.UT
            End If
        End Get
    End Property
    Public ReadOnly Property UT_DT
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) x.UT)
            Else
                Return DT_Report.UT
            End If
        End Get
    End Property
    Public ReadOnly Property UPDT
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) x.UPDT)
            Else
                Return DT_Report.UPDT
            End If
        End Get
    End Property
    Public ReadOnly Property PDT
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) x.PDT)
            Else
                Return DT_Report.PDT
            End If
        End Get
    End Property
    Public ReadOnly Property UPDTpct
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) (x.UPDT / x.schedTime))
            Else
                If schedTime = 0 Then Return 0
                Return DT_Report.UPDT / schedTime
            End If
        End Get
    End Property
    Public ReadOnly Property PDTpct
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) (x.PDT / x.schedTime))
            Else
                If schedTime = 0 Then Return 0
                Return DT_Report.PDT / schedTime
            End If
        End Get
    End Property

    Public ReadOnly Property MTBF
        Get
            If doIUseReports Then
                If Stops = 0 Then
                    Return 0
                Else
                    Return UT / Stops
                End If
            Else
                If Not My.Settings.AdvancedSettings_isAvailabilityMode And (ParentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple Or ParentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple_New) Then
                    If DT_Report.Stops = 0 Then
                        Return 0
                    Else
                        Return PROD_Report._uptimeCalc / DT_Report.Stops
                    End If
                Else
                    Return DT_Report.MTBF
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property RateLoss
        Get
            'If DT_Report.UT - PROD_Report.UT > schedTime Then  ' LG Code - commented out Sam's code
            'Return 0 ' LG Code - commented out Sam's code
            ' Else ' LG Code - commented out Sam's code
            'Return DT_Report.UT - PROD_Report.UT
            Return 1.0 - PR - UPDTpct - PDTpct ' LG Code retaining Sam's code ' rateloss 
            ' End If ' LG Code - commented out Sam's code
        End Get
    End Property
    Public ReadOnly Property RateLossPct
        Get
            If schedTime = 0 Then
                Return 0
            Else
                Return RateLoss '/ schedTime ' LG Code
            End If
        End Get
    End Property
    Public ReadOnly Property Stops
        Get
            If doIUseReports Then
                Return DT_Reports.Sum(Function(x) x.Stops)
            Else
                ' If ParentLine._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode = MultiConstraintAnalysis.NoRateLossStops Then
                'Return DT_Report.Stops - ParentLine.getRateLossEventsInTimePeriod(startTime, endTime)
                ' End If
                Return DT_Report.Stops
            End If
        End Get
    End Property
    Public ReadOnly Property ActualCases
        Get
            If My.Settings.AdvancedSettings_isAvailabilityMode Then Return 0
            Return PROD_Report.CasesActual
        End Get
    End Property
    Public ReadOnly Property AdjustedCases
        Get
            If My.Settings.AdvancedSettings_isAvailabilityMode Then Return 0
            Return PROD_Report.CasesAdjusted
        End Get
    End Property
    Public ReadOnly Property AdjustedUnits
        Get
            If My.Settings.AdvancedSettings_isAvailabilityMode Then Return 0
            Return PROD_Report.UnitsAdjusted
        End Get
    End Property
    Public ReadOnly Property SPD
        Get
            If doIUseReports Then
                If schedTime = 0 Then Return 0
                Return Stops / schedTime * 1440
            Else
                If schedTime = 0 Then Return 0
                Return Stops / schedTime * 1440
            End If
        End Get
    End Property
    Public ReadOnly Property ParentLineName
        Get
            Return ParentLine.Name
        End Get
    End Property
#End Region

#Region "Construction / Reinitialization"
    'constructors
    Public Sub New(InparentLine As ProdLine, ByRef reportStartTime As Date, ByRef reportEndTime As Date)
        ParentLine = InparentLine

        startTime = reportStartTime
        endTime = reportEndTime

        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
            PROD_Report = New ProductionReport(ParentLine, startTime, endTime)
        End If

        DT_Report = New DowntimeReport(ParentLine, startTime, endTime)

        reportStartTime = startTime
        reportEndTime = endTime
    End Sub

    Public Sub New(InparentLine As ProdLine, ByRef reportStartTime As Date, ByRef reportEndTime As Date, isDover As Boolean)
        doIUseReports = isDover

        ParentLine = InparentLine
        startTime = reportStartTime
        endTime = reportEndTime

        For Each x As DowntimeDataset In InparentLine.BabyWipesPRSTORYData
            DT_Reports.Add(New DowntimeReport(ParentLine, startTime, endTime))
        Next

        reportStartTime = startTime
        reportEndTime = endTime
    End Sub

    Public Sub reMapDowntime(mappingA as integer, mappingB as integer)
        If not doIUseReports Then
           ' DT_Report.reMapDataSet(My.Settings.defaultDownTimeField, My.Settings.defaultDownTimeField_Secondary)
             DT_Report.reMapDataSet(mappingA, mappingB)
        End If
    End Sub

    Public Sub reFilterDowntime_ProductGroup(inclusionList As List(Of String))
        DT_Report.reFilterData_ProductGroup(inclusionList)
    End Sub
    Public Sub reFilterDowntime_SKU(inclusionList As List(Of String))
        If Not My.Settings.AdvancedSettings_isAvailabilityMode = True Then PROD_Report = New ProductionReport(ParentLine, startTime, endTime, inclusionList, DowntimeField.Product)
        DT_Report.reFilterData_SKU(inclusionList)
    End Sub
    Public Sub reFilterDowntime_Team(inclusionList As List(Of String))
        If Not My.Settings.AdvancedSettings_isAvailabilityMode = True Then PROD_Report = New ProductionReport(ParentLine, startTime, endTime, inclusionList, DowntimeField.Team)
        DT_Report.reFilterData_Team(inclusionList)
    End Sub
    Public Sub reFilterDowntime_Format(inclusionList As List(Of String))
        If Not My.Settings.AdvancedSettings_isAvailabilityMode = True Then PROD_Report = New ProductionReport(ParentLine, startTime, endTime, inclusionList, DowntimeField.Format)
        DT_Report.reFilterData_Format(inclusionList)
    End Sub
    Public Sub reFilterDowntime_Shape(inclusionList As List(Of String))
        If Not My.Settings.AdvancedSettings_isAvailabilityMode = True Then PROD_Report = New ProductionReport(ParentLine, startTime, endTime, inclusionList, DowntimeField.Shape)
        DT_Report.reFilterData_Shape(inclusionList)
    End Sub
    Public Sub reFilterDowntime_ClearAllFilters()
        If Not My.Settings.AdvancedSettings_isAvailabilityMode = True Then PROD_Report = New ProductionReport(ParentLine, startTime, endTime)
        DT_Report.reFilterData_ClearAllFilters()
    End Sub
#End Region

    public function getSubset(byval startTime as date, byval endTime as date) as summaryreport
        return new SummaryReport(parentline, starttime, endtime)
    End function

    Public Overrides Function ToString() As String
        Return "S/E: " & startTime & "/" & endTime & " PR/Stops: " & PR & "/" & Stops
    End Function
End Class

public Class ProductionReport
#Region "Properties / Variables"
    Private _startTime As Date
    Private _endTime As Date

    Private parentLine As ProdLine
    Private _rawProdData As Array
    Private _rawProductionData As ProductionDataset

    Private _schedTime As Double ' = 0
    Public _uptimeCalc As Double ' = 0
    Private _PR As Double = 0 'DO NOT USE THIS!!!!!
    Private _actCases As Long ' = 0
    Private _adjCases As Long ' = 0
    Private _statUnits As Double ' = 0
    Private _adjUnits As Double ' = 0

    'sorting parameters
    Private _Brandcodes As List(Of String)
    Private _Products As List(Of String)
    Private isFilter As Boolean = False
    Private isFilterProducts As Boolean = False
    Private isFilterBrandcodes As Boolean = False


    Public ReadOnly Property isNothing_rawProdData As Boolean
        Get
            Return IsNothing(_rawProductionData)
        End Get
    End Property

    Public ReadOnly Property RawData As List(Of ProductionEvent)
        Get
            Return _rawProductionData.rawProductionData
        End Get
    End Property
    'properties
    Public ReadOnly Property schedTime As Double
        Get
            Return _schedTime
        End Get
    End Property
    Public ReadOnly Property UT As Double
        Get
            Return _uptimeCalc
        End Get
    End Property
    Public ReadOnly Property PR As Double
        Get
            If _schedTime = 0 Then Return 0
            Return _uptimeCalc / _schedTime
        End Get
    End Property
    Public ReadOnly Property CasesActual As Double
        Get
            Return _actCases
        End Get
    End Property
    Public ReadOnly Property CasesAdjusted As Double
        Get
            Return _adjCases
        End Get
    End Property
    Public ReadOnly Property UnitsAdjusted As Double
        Get
            Return _adjUnits
        End Get
    End Property
    Public ReadOnly Property UnitsStat As Double
        Get
            Return _statUnits
        End Get
    End Property
#End Region

#Region "Construction"
    ' Constructor Implementing Filtering By Brandcode & Products
    Public Sub New(inParentLine As ProdLine, startTime As Date, endTime As Date, Products As List(Of String), FilterField As Integer)
        _rawProdData = inParentLine.rawProficyProductionData
        _startTime = startTime
        _endTime = endTime
        parentLine = inParentLine

        _Products = Products

        isFilter = True
        isFilterProducts = True
        _Products = Products
        ' End If

        Select Case FilterField
            Case DowntimeField.Product
                Call executeProductionReport_FilteredByProduct()
            Case DowntimeField.Team
                Call executeProductionReport_FilteredByTeam()
            Case DowntimeField.Shape
                Call executeProductionReport_FilteredByShape()
            Case DowntimeField.Format
                Call executeProductionReport_FilteredByFormat()
            Case Else
                Throw New unknownMappingException
        End Select
    End Sub

    'Constructor w/ No Filtering
    Public Sub New(iParentLine As ProdLine, ByRef startTime As Date, ByRef endTime As Date)
        _rawProdData = iParentLine.rawProficyProductionData
        _startTime = startTime
        _endTime = endTime
        parentLine = iParentLine

        Try
            _rawProductionData = parentLine.rawProductionData.getSubset(startTime, endTime)
        Catch ex As columnStubbingException

        End Try

        Call getAllProductionMetrics()
        ' Call executeProductionReport()
    End Sub
#End Region


    Private Sub getAllProductionMetrics()
        _uptimeCalc = 0
        _schedTime = 0
        _actCases = 0
        _adjCases = 0
        _adjUnits = 0
        _statUnits = 0

        Dim NetProductionMinutes As Double = 0
        Dim NetRateGainMinutes As Double = 0

        If Not IsNothing(_rawProductionData) Then
            With _rawProductionData
                If .rawProductionData.Count > 0 Then
                    For i As Integer = 0 To .rawProductionData.Count - 1
                        If Not .rawProductionData(i).isExcluded Then
                            _uptimeCalc += .rawProductionData(i).UT
                            _schedTime += .rawProductionData(i).SchedTime
                            _actCases += .rawProductionData(i).ActCases
                            _adjCases += .rawProductionData(i).AdjCases
                            _adjUnits += .rawProductionData(i).AdjUnits
                            _statUnits += .rawProductionData(i).StatUnits

                            If parentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple Or parentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple_New Then
                                NetProductionMinutes += .rawProductionData(i).ProductionMinutes
                                NetRateGainMinutes += .rawProductionData(i).RateGainMinutes
                            End If
                            '  ElseIf parentLine.prStoryMapping = prStoryMapping.FamilyCareUnitOP_ModPACK Then

                        End If
                    Next
                End If
            End With
        End If

        If parentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple Or parentLine.SQLproductionProcedure = DefaultProficyProductionProcedure.Maple_New Then
             _uptimeCalc = NetProductionMinutes + Math.Min(NetRateGainMinutes, 0)
           ' _uptimeCalc = NetProductionMinutes - Math.Max(NetRateGainMinutes, 0) 'original, before 5/15/17. st louis pr investigation
        End If
    End Sub


#Region "Filtering"
    Public Sub executeProductionReport_FilteredByFormat()
        Dim startRow As Integer, testFormat As String
        Dim endRow As Integer, r As Integer
        Dim errTest As Double
        startRow = parentLine.getProdRowFromTime(_startTime)

        If startRow > -1 Then
            endRow = parentLine.getProdRowFromTime(_endTime, True)
            If endRow > startRow Or endRow = startRow Then
                For r = startRow To endRow
                    If Not parentLine.isProductionExcluded(r) Then
                        If Not IsDBNull(_rawProdData(ProductionColumn.Product, r)) Then
                            Select Case parentLine.shapeMapping
                                Case MappingByShape.SkinCare
                                    testFormat = getSkinCareFormatFromSku(_rawProdData(ProductionColumn.Product, r))
                                Case Else
                                    Throw New unknownMappingException
                            End Select
                            If _Products.IndexOf(testFormat) > -1 Then
                                _schedTime = _schedTime + _rawProdData(ProductionColumn.SchedTime, r)
                                _uptimeCalc = _uptimeCalc + getUptimeForPO(r) '_rawProdData(ProductionColumn.UT, r)
                                'PR = 0
                                If IsDBNull(_rawProdData(ProductionColumn.ActualCases, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.ActualCases, r)
                                End If
                                _actCases = _actCases + errTest '_rawProdData(ProductionColumn.ActualCases, r)
                                If IsDBNull(_rawProdData(ProductionColumn.AdjustedCases, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.AdjustedCases, r)
                                End If
                                _adjCases = _adjCases + errTest
                                If IsDBNull(_rawProdData(ProductionColumn.AdjustedUnits, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.AdjustedUnits, r)
                                End If
                                _adjUnits = _adjUnits + errTest
                                If IsDBNull(_rawProdData(ProductionColumn.StatUnits, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.StatUnits, r)
                                End If
                                _statUnits = _statUnits + errTest
                                'rateLossMin = rateLossMin
                            End If
                        End If
                    End If
                Next r
                If _schedTime <> 0 Then _PR = _uptimeCalc / _schedTime
            Else

                MsgBox("Error End Of Production Time", vbCritical, "ERR #14564")
                Debugger.Break()
            End If
        Else
            MsgBox("ERROR: Invalid Production Start Time", vbCritical, "error - invalid time")

            Debugger.Break()
        End If
    End Sub

    Public Sub executeProductionReport_FilteredByShape()
        Dim startRow As Integer, testFormat As String
        Dim endRow As Integer, r As Integer
        Dim errTest As Double
        startRow = parentLine.getProdRowFromTime(_startTime)

        If startRow > -1 Then
            endRow = parentLine.getProdRowFromTime(_endTime, True)
            If endRow > startRow Or endRow = startRow Then
                For r = startRow To endRow
                    If Not parentLine.isProductionExcluded(r) Then
                        If Not IsDBNull(_rawProdData(ProductionColumn.Product, r)) Then
                            Select Case parentLine.shapeMapping
                                Case MappingByShape.SkinCare
                                    testFormat = getSkinCareShapeFromSku(_rawProdData(ProductionColumn.Product, r))
                                Case Else
                                    Throw New unknownMappingException
                            End Select
                            If _Products.IndexOf(testFormat) > -1 Then
                                _schedTime = _schedTime + _rawProdData(ProductionColumn.SchedTime, r)
                                _uptimeCalc = _uptimeCalc + getUptimeForPO(r) '_rawProdData(ProductionColumn.UT, r)
                                'PR = 0
                                If IsDBNull(_rawProdData(ProductionColumn.ActualCases, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.ActualCases, r)
                                End If
                                _actCases = _actCases + errTest '_rawProdData(ProductionColumn.ActualCases, r)
                                If IsDBNull(_rawProdData(ProductionColumn.AdjustedCases, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.AdjustedCases, r)
                                End If
                                _adjCases = _adjCases + errTest
                                If IsDBNull(_rawProdData(ProductionColumn.AdjustedUnits, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.AdjustedUnits, r)
                                End If
                                _adjUnits = _adjUnits + errTest
                                If IsDBNull(_rawProdData(ProductionColumn.StatUnits, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.StatUnits, r)
                                End If
                                _statUnits = _statUnits + errTest
                                'rateLossMin = rateLossMin
                            End If
                        End If
                    End If
                Next r
                If _schedTime <> 0 Then _PR = _uptimeCalc / _schedTime
            Else

                MsgBox("Error End Of Production Time", vbCritical, "ERR #14564")
                Debugger.Break()
            End If
        Else
            MsgBox("ERROR: Invalid Production Start Time", vbCritical, "error - invalid time")

            Debugger.Break()
        End If
    End Sub

    Public Sub executeProductionReport_FilteredByProduct()
        Dim startRow As Integer
        Dim endRow As Integer, r As Integer
        Dim errTest As Double
        startRow = parentLine.getProdRowFromTime(_startTime)

        If startRow > -1 Then
            endRow = parentLine.getProdRowFromTime(_endTime, True)
            If endRow > startRow Or endRow = startRow Then
                For r = startRow To endRow
                    If Not parentLine.isProductionExcluded(r) Then
                        If Not IsDBNull(_rawProdData(ProductionColumn.Product, r)) Then
                            If _Products.IndexOf(_rawProdData(ProductionColumn.Product, r)) > -1 Then
                                _schedTime = _schedTime + _rawProdData(ProductionColumn.SchedTime, r)
                                _uptimeCalc = _uptimeCalc + getUptimeForPO(r) '_rawProdData(ProductionColumn.UT, r)
                                'PR = 0
                                If IsDBNull(_rawProdData(ProductionColumn.ActualCases, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.ActualCases, r)
                                End If
                                _actCases = _actCases + errTest '_rawProdData(ProductionColumn.ActualCases, r)
                                If IsDBNull(_rawProdData(ProductionColumn.AdjustedCases, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.AdjustedCases, r)
                                End If
                                _adjCases = _adjCases + errTest
                                If IsDBNull(_rawProdData(ProductionColumn.AdjustedUnits, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.AdjustedUnits, r)
                                End If
                                _adjUnits = _adjUnits + errTest
                                If IsDBNull(_rawProdData(ProductionColumn.StatUnits, r)) Then
                                    errTest = 0
                                Else
                                    errTest = _rawProdData(ProductionColumn.StatUnits, r)
                                End If
                                _statUnits = _statUnits + errTest
                                'rateLossMin = rateLossMin
                            End If
                        End If
                    End If
                Next r
                If _schedTime <> 0 Then _PR = _uptimeCalc / _schedTime
            Else

                MsgBox("Error End Of Production Time", vbCritical, "ERR #14564")
                Debugger.Break()
            End If
        Else
            MsgBox("ERROR: Invalid Production Start Time", vbCritical, "error - invalid time")

            Debugger.Break()
        End If
    End Sub

    Public Sub executeProductionReport_FilteredByTeam()
        Dim startRow As Integer
        Dim endRow As Integer, r As Integer
        Dim errTest As Double
        startRow = parentLine.getProdRowFromTime(_startTime)

        If startRow > -1 Then
            endRow = parentLine.getProdRowFromTime(_endTime, True)
            If endRow > startRow Or endRow = startRow Then



                'PROFICY / not maple
                If parentLine.SQLproductionProcedure <> DefaultProficyProductionProcedure.Maple_New Then



                    For r = startRow To endRow
                        If Not parentLine.isProductionExcluded(r) Then
                            If Not IsDBNull(_rawProdData(ProductionColumn.Team, r)) Then
                                If _Products.IndexOf(_rawProdData(ProductionColumn.Team, r)) > -1 Then
                                    _schedTime = _schedTime + _rawProdData(ProductionColumn.SchedTime, r)
                                    _uptimeCalc = _uptimeCalc + getUptimeForPO(r) '_rawProdData(ProductionColumn.UT, r)
                                    'PR = 0
                                    If IsDBNull(_rawProdData(ProductionColumn.ActualCases, r)) Then
                                        errTest = 0
                                    Else
                                        errTest = _rawProdData(ProductionColumn.ActualCases, r)
                                    End If
                                    _actCases = _actCases + errTest '_rawProdData(ProductionColumn.ActualCases, r)
                                    If IsDBNull(_rawProdData(ProductionColumn.AdjustedCases, r)) Then
                                        errTest = 0
                                    Else
                                        errTest = _rawProdData(ProductionColumn.AdjustedCases, r)
                                    End If
                                    _adjCases = _adjCases + errTest
                                    If IsDBNull(_rawProdData(ProductionColumn.AdjustedUnits, r)) Then
                                        errTest = 0
                                    Else
                                        errTest = _rawProdData(ProductionColumn.AdjustedUnits, r)
                                    End If
                                    _adjUnits = _adjUnits + errTest
                                    If IsDBNull(_rawProdData(ProductionColumn.StatUnits, r)) Then
                                        errTest = 0
                                    Else
                                        errTest = _rawProdData(ProductionColumn.StatUnits, r)
                                    End If
                                    _statUnits = _statUnits + errTest
                                    'rateLossMin = rateLossMin
                                End If
                            End If
                        End If
                    Next r
                    If _schedTime <> 0 Then _PR = _uptimeCalc / _schedTime


                    'MAPLE

                Else

                    For r = startRow To endRow
                        If Not parentLine.isProductionExcluded(r, True) Then
                            _schedTime = _schedTime + _rawProdData(ProductionColumn_Maple_New.Line_Scheduled_Time, r)
                            _uptimeCalc = _uptimeCalc + getUptimeForPO_MAPLE(r) '_rawProdData(productioncolumn_maple_new.UT, r)
                            'PR = 0
                            If IsDBNull(_rawProdData(ProductionColumn_Maple_New.Actual_Cases, r)) Then
                                errTest = 0
                            Else
                                errTest = _rawProdData(ProductionColumn_Maple_New.Actual_Cases, r)
                            End If
                            _actCases = _actCases + errTest '_rawProdData(productioncolumn_maple_new.ActualCases, r)
                            If IsDBNull(_rawProdData(ProductionColumn_Maple_New.Adjusted_Cases, r)) Then
                                errTest = 0
                            Else
                                errTest = _rawProdData(ProductionColumn_Maple_New.Adjusted_Cases, r)
                            End If
                            _adjCases = _adjCases + errTest
                            If IsDBNull(_rawProdData(ProductionColumn_Maple_New.Adjusted_Units, r)) Then
                                errTest = 0
                            Else
                                errTest = _rawProdData(ProductionColumn_Maple_New.Adjusted_Units, r)
                            End If
                            _adjUnits = _adjUnits + errTest
                            If IsDBNull(_rawProdData(ProductionColumn_Maple_New.Stat_Units, r)) Then
                                errTest = 0
                            Else
                                errTest = _rawProdData(ProductionColumn_Maple_New.Stat_Units, r)
                            End If
                            _statUnits = _statUnits + errTest
                            'rateLossMin = rateLossMin
                        End If
                    Next r




                End If






            Else

                MsgBox("Error End Of Production Time", vbCritical, "ERR #14564")
                Debugger.Break()
            End If
        Else
            MsgBox("ERROR: Invalid Production Start Time", vbCritical, "error - invalid time")

            Debugger.Break()
        End If
    End Sub
#End Region
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


    Private Function getUptimeForPO(poIncrementer As Integer) As Double
        Dim tmpRate As Integer
        If Not IsDBNull(_rawProdData(ProductionColumn.ProductionStatus, poIncrementer)) And Not IsDBNull(_rawProdData(ProductionColumn.AdjustedUnits, poIncrementer)) Then
            'check if this is PR In
            If Not Left(_rawProdData(ProductionColumn.ProductionStatus, poIncrementer), 4) = "PR O" Then
                If IsDBNull(_rawProdData(ProductionColumn.TargetRate, poIncrementer)) Then
                    If Not IsDBNull(_rawProdData(ProductionColumn.ActualRate, poIncrementer)) Then
                        tmpRate = _rawProdData(ProductionColumn.ActualRate, poIncrementer)
                    Else
                        Debugger.Break()
                    End If
                ElseIf IsDBNull(_rawProdData(ProductionColumn.ActualRate, poIncrementer)) Then
                    tmpRate = _rawProdData(ProductionColumn.TargetRate, poIncrementer)
                Else
                    tmpRate = Math.Max(_rawProdData(ProductionColumn.TargetRate, poIncrementer), _rawProdData(ProductionColumn.ActualRate, poIncrementer))
                End If
                Return _rawProdData(ProductionColumn.AdjustedUnits, poIncrementer) / tmpRate
            Else
                Return 0
            End If
            'this means we had an important null somewhere...
        Else
            Return 0
        End If
    End Function

    Private Function getUptimeForPO_MAPLE(poIncrementer As Integer) As Double
        Dim tmpRate As Integer
        If Not IsDBNull(_rawProdData(ProductionColumn_Maple_New.LineStatus, poIncrementer)) And Not IsDBNull(_rawProdData(ProductionColumn_Maple_New.Adjusted_Units, poIncrementer)) Then
            'check if this is PR In
            If Not Left(_rawProdData(ProductionColumn_Maple_New.LineStatus, poIncrementer), 4) = "PR O" Then
                If IsDBNull(_rawProdData(ProductionColumn_Maple_New.Target_Rate, poIncrementer)) Then
                    If Not IsDBNull(_rawProdData(ProductionColumn_Maple_New.Actual_Rate, poIncrementer)) Then
                        tmpRate = _rawProdData(ProductionColumn_Maple_New.Actual_Rate, poIncrementer)
                    Else
                        Debugger.Break()
                    End If
                ElseIf IsDBNull(_rawProdData(ProductionColumn_Maple_New.Actual_Rate, poIncrementer)) Then
                    tmpRate = _rawProdData(ProductionColumn_Maple_New.Target_Rate, poIncrementer)
                Else
                    tmpRate = Math.Max(_rawProdData(ProductionColumn_Maple_New.Target_Rate, poIncrementer), _rawProdData(ProductionColumn_Maple_New.Actual_Rate, poIncrementer))
                End If
                Return _rawProdData(ProductionColumn_Maple_New.Adjusted_Units, poIncrementer) / tmpRate
            Else
                Return 0
            End If
        Else
            Return 0
        End If
    End Function
End Class

Public Class DowntimeReport

#Region "Variables & Properties"
    Private _UPDT As Double = 0
    Private _PDT As Double = 0
    Private _UT As Double = 0
    Private _DT As Double = 0
    Private _excludedTime As Double = 0
    Private _schedTime As Double = 0


    Private _rateLossEvents As Long = 0
    Private _excludedStops As Long = 0
    Private _parentLine As ProdLine

    Private _rawDTData As DowntimeDataset
    public ReadOnly Property rawDTdata As DowntimeDataset
        Get
            Return _rawDTData
        End Get
    End Property

    Public ReadOnly Property UnplannedEventDirectory As List(Of DTevent)
        Get
            Dim tmpList As New List(Of DTevent)
            For i As Integer = 0 To MappedDirectory.Count - 1
                tmpList.Add(MappedDirectory(i))
            Next
            sortEventList_ByStops(tmpList)
            Return tmpList 'MappedDirectory
        End Get
    End Property


    'UNPLANNED - raw data 
    Friend FaultDirectory As New List(Of DTevent)
    Friend Reason1Directory As New List(Of DTevent)

    Friend DTgroupDirectory As New List(Of DTevent)
    Friend LocationDirectory As New List(Of DTevent)

    Friend Reason2Directory As New List(Of DTevent)
    Friend Reason3Directory As New List(Of DTevent)
    Friend Reason4Directory As New List(Of DTevent)

    'mapped top levels
    Friend Tier1Directory As New List(Of DTevent)
    Friend OneClickDirectory As New List(Of DTevent)

    'PLANNED
    Friend PlannedTier1Directory As New List(Of DTevent)

    Friend MappedDirectory As New List(Of DTevent)


    'Fields We Filter By
    Friend ActiveGCAS As New List(Of String)
    Friend ActiveProducts As New List(Of String)
    Friend ActiveTeams As New List(Of String)
    Friend ActiveShapes As New List(Of String)
    Friend ActiveFormats As New List(Of String)
    Friend ActiveProductGroups As New List(Of String)

    Public Function getFilterList(FilterField As Integer) As List(Of String)
        Dim tmpList As New List(Of String), i As Integer
        Select Case FilterField
            Case DowntimeField.Product
                For i = 0 To ActiveProducts.Count - 1
                    tmpList.Add(ActiveProducts(i))
                Next
            Case DowntimeField.Shape
                For i = 0 To ActiveShapes.Count - 1
                    tmpList.Add(ActiveShapes(i))
                Next
            Case DowntimeField.Format
                For i = 0 To ActiveFormats.Count - 1
                    tmpList.Add(ActiveFormats(i))
                Next
            Case DowntimeField.Team
                For i = 0 To ActiveTeams.Count - 1
                    tmpList.Add(ActiveTeams(i))
                Next
            Case DowntimeField.ProductGroup
                For i = 0 To ActiveProductGroups.Count - 1
                    tmpList.Add(ActiveProductGroups(i))
                Next
            Case Else
                Throw New unknownMappingException
        End Select
        Return tmpList
    End Function


    'Properties
    Public ReadOnly Property StartTime As Date
        Get
            Return _rawDTData.StartDate
        End Get
    End Property
    Public ReadOnly Property EndTime As Date
        Get
            Return _rawDTData.EndDate
        End Get
    End Property
    Public ReadOnly Property Stops
        Get
            If _parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.Maple Then
                Dim x As Integer = 0
                For i = 0 To _rawDTData.UnplannedData.Count - 1
                    If _rawDTData.UnplannedData(i).isSplitEvent Then
                        x = x + 1
                    End If
                Next
                Return _rawDTData.Stops - _rateLossEvents - _excludedStops - x
            Else
                Return _rawDTData.Stops - _rateLossEvents - _excludedStops
            End If
        End Get
    End Property
    Public ReadOnly Property PlannedStops
        Get
            Return _rawDTData.PlannedData.Count - 1
        End Get
    End Property
    Public ReadOnly Property ChangeoversNum
        Get
            Return _rawDTData.numChangeovers
        End Get
    End Property
    Public ReadOnly Property CILsNum
        Get
            Return _rawDTData.numCILs
        End Get
    End Property
    Public ReadOnly Property schedTime
        Get
            Return _schedTime
        End Get
    End Property
    Public ReadOnly Property PDT
        Get
            Return _PDT
        End Get
    End Property
    Public ReadOnly Property UPDT
        Get
            Return _UPDT
        End Get
    End Property
    Public ReadOnly Property UT
        Get
            Return _UT
        End Get
    End Property
    Public ReadOnly Property MTBF
        Get
            If Stops = 0 Then Return 0
            Return _UT / Stops
        End Get
    End Property
    Public ReadOnly Property MTTR
        Get
            If Stops = 0 Then Return 0
            Return _UPDT / Stops
        End Get
    End Property
    Public ReadOnly Property Availability As Double
        Get
            If schedTime = 0 Then
                Return 0
            Else
                Return _UT / _schedTime
            End If
        End Get
    End Property
#End Region

    Public Overrides Function ToString() As String
        Return "Sched: " & _schedTime & " Stops: " & Stops
    End Function

#Region "Get Subdirectories For Mapped Reason Level"
    Public Function getMappedSubdirectory(MappedName As String, targetField As Integer) As List(Of DTevent)
        Select Case targetField
            Case DowntimeField.Reason1
                Return getMappedReason1Directory(MappedName)
            Case DowntimeField.Reason2
                Return getMappedReason2Directory(MappedName)
            Case DowntimeField.Reason3
                Return getMappedReason3Directory(MappedName)
            Case DowntimeField.Reason4
                Return getMappedReason4Directory(MappedName)
            Case DowntimeField.Fault
                Return getMappedFaultDirectory(MappedName)
            Case Else
                Throw New unknownMappingException
        End Select
    End Function

    Private Function getMappedFaultDirectory(MappedName As String) As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                If .MappedField = MappedName Then
                    tmpIndex = tmpList.IndexOf(New DTevent(.Fault, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Fault, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End If
            End With
        Next
        Return tmpList
    End Function
    Private Function getMappedReason1Directory(MappedName As String) As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                If .MappedField = MappedName Then
                    tmpIndex = tmpList.IndexOf(New DTevent(.Reason1, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Reason1, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End If
            End With
        Next
        Return tmpList
    End Function
    Private Function getMappedReason4Directory(MappedName As String) As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                If .MappedField = MappedName Then
                    tmpIndex = tmpList.IndexOf(New DTevent(.Reason4, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Reason4, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End If
            End With
        Next
        Return tmpList
    End Function
    Private Function getMappedReason3Directory(MappedName As String) As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                If .MappedField = MappedName Then
                    tmpIndex = tmpList.IndexOf(New DTevent(.Reason3, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Reason3, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End If
            End With
        Next
        Return tmpList
    End Function
    Private Function getMappedReason2Directory(MappedName As String) As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                If .MappedField = MappedName Then
                    tmpIndex = tmpList.IndexOf(New DTevent(.Reason2, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Reason2, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End If
            End With
        Next
        Return tmpList
    End Function
#End Region

#Region "Get Tier 1-3 Directories"
    Public Function getUnplannedEventDirectory(targetDtField As Integer, Optional isByStops As Boolean = True) As List(Of DTevent)
        Dim tmpList As New List(Of DTevent), i As Integer
        Select Case targetDtField
            Case DowntimeField.Reason1
                If isByStops Then
                    sortEventList_ByStops(Reason1Directory)
                Else
                    sortEventList_ByDT(Reason1Directory)
                End If
                For i = 0 To Reason1Directory.Count - 1
                    tmpList.Add(Reason1Directory(i))
                Next
            Case DowntimeField.Reason2
                If isByStops Then
                    sortEventList_ByStops(Reason2Directory)
                Else
                    sortEventList_ByDT(Reason2Directory)
                End If
                For i = 0 To Reason2Directory.Count - 1
                    tmpList.Add(Reason2Directory(i))
                Next
            Case DowntimeField.Reason3
                If isByStops Then
                    sortEventList_ByStops(Reason3Directory)
                Else
                    sortEventList_ByDT(Reason3Directory)
                End If
                For i = 0 To Reason3Directory.Count - 1
                    tmpList.Add(Reason3Directory(i))
                Next
            Case DowntimeField.Reason4
                If isByStops Then
                    sortEventList_ByStops(Reason4Directory)
                Else
                    sortEventList_ByDT(Reason4Directory)
                End If
                For i = 0 To Reason4Directory.Count - 1
                    tmpList.Add(Reason4Directory(i))
                Next
            Case DowntimeField.Fault
                If isByStops Then
                    sortEventList_ByStops(FaultDirectory)
                Else
                    sortEventList_ByDT(FaultDirectory)
                End If
                tmpList = FaultDirectory
                For i = 0 To FaultDirectory.Count - 1
                    tmpList.Add(FaultDirectory(i))
                Next
            Case DowntimeField.Location
                If isByStops Then
                    sortEventList_ByStops(LocationDirectory)
                Else
                    sortEventList_ByDT(LocationDirectory)
                End If
                tmpList = LocationDirectory
                For i = 0 To LocationDirectory.Count - 1
                    tmpList.Add(LocationDirectory(i))
                Next
            Case DowntimeField.Tier1
                If isByStops Then
                    sortEventList_ByStops(Tier1Directory)
                Else
                    sortEventList_ByDT(Tier1Directory)
                End If
                For i = 0 To Tier1Directory.Count - 1
                    tmpList.Add(Tier1Directory(i))
                Next
            Case DowntimeField.OneClick
                If isByStops Then
                    sortEventList_ByStops(OneClickDirectory)
                Else
                    sortEventList_ByDT(OneClickDirectory)
                End If
                For i = 0 To OneClickDirectory.Count - 1
                    tmpList.Add(OneClickDirectory(i))
                Next
            Case Else
                Throw New unknownMappingException
        End Select
        Return tmpList
    End Function


    Public Function getTier1Directory() As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                tmpIndex = tmpList.IndexOf(New DTevent(.Tier1, 0))
                If tmpIndex = -1 Then
                    tmpList.Add(New DTevent(.Tier1, .DT, i))
                Else
                    tmpList(tmpIndex).addStopWithRow(.DT, i)
                End If
            End With
        Next
        Return tmpList
    End Function

    Public Function getPlannedTier1Directory() As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.PlannedData.Count - 1
            With _rawDTData.PlannedData(i)
                tmpIndex = tmpList.IndexOf(New DTevent(.Tier1, 0))
                If tmpIndex = -1 Then
                    tmpList.Add(New DTevent(.Tier1, .DT, i))
                Else
                    tmpList(tmpIndex).addStopWithRow(.DT, i)
                End If
            End With
        Next
        Return tmpList
    End Function

    Public Function getTier2Directory(Optional Tier1Name As String = "") As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        If Tier1Name = "" Then
            For i = 0 To _rawDTData.UnplannedData.Count - 1
                With _rawDTData.UnplannedData(i)
                    tmpIndex = tmpList.IndexOf(New DTevent(.Tier2, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Tier2, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End With
            Next
        Else
            For i = 0 To _rawDTData.UnplannedData.Count - 1
                With _rawDTData.UnplannedData(i)
                    If .Tier1 = Tier1Name And Len(.Tier2) > 1 Then
                        tmpIndex = tmpList.IndexOf(New DTevent(.Tier2, 0))
                        If tmpIndex = -1 Then
                            tmpList.Add(New DTevent(.Tier2, .DT, i))
                        Else
                            tmpList(tmpIndex).addStopWithRow(.DT, i)
                        End If
                    End If
                End With
            Next
        End If
        Return tmpList
    End Function
    Public Function getTier3Directory(Optional Tier1Name As String = "", Optional Tier2Name As String = "") As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        If Tier1Name = "" Then
            For i = 0 To _rawDTData.UnplannedData.Count - 1
                With _rawDTData.UnplannedData(i)
                    tmpIndex = tmpList.IndexOf(New DTevent(.Tier3, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Tier3, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End With
            Next
        Else
            For i = 0 To _rawDTData.UnplannedData.Count - 1
                With _rawDTData.UnplannedData(i)
                    If .Tier1 = Tier1Name And .Tier2 = Tier2Name And Len(.Tier3) > 1 Then
                        tmpIndex = tmpList.IndexOf(New DTevent(.Tier3, 0))
                        If tmpIndex = -1 Then
                            tmpList.Add(New DTevent(.Tier3, .DT, i))
                        Else
                            tmpList(tmpIndex).addStopWithRow(.DT, i)
                        End If
                    End If
                End With
            Next
        End If
        Return tmpList
    End Function

    Public Function getPlannedEventDirectory(targetDtfield As Integer, Optional isByStops As Boolean = True) As List(Of DTevent)
        Dim tmpList As New List(Of DTevent), i As Integer
        Select Case targetDtfield
            Case DowntimeField.Tier1
                If isByStops Then
                    sortEventList_ByStops(PlannedTier1Directory)
                Else
                    sortEventList_ByDT(Tier1Directory)
                End If
                For i = 0 To PlannedTier1Directory.Count - 1
                    tmpList.Add(PlannedTier1Directory(i))
                Next
            Case Else
                Throw New unknownMappingException
        End Select
        Return tmpList
    End Function

    Public Function getPlannedTier2Directory(Optional PlannedTier1Name As String = "") As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        If PlannedTier1Name = "" Then
        Else
            For i = 0 To _rawDTData.PlannedData.Count - 1
                With _rawDTData.PlannedData(i)
                    If .Tier1 = PlannedTier1Name Then
                        tmpIndex = tmpList.IndexOf(New DTevent(.Tier2, 0))
                        If tmpIndex = -1 Then
                            tmpList.Add(New DTevent(.Tier2, .DT, i))
                        Else
                            tmpList(tmpIndex).addStopWithRow(.DT, i)
                        End If
                    End If
                End With
            Next
        End If
        Return tmpList
    End Function
    Public Function getPlannedTier3Directory(Optional PlannedTier1Name As String = "", Optional PlannedTier2Name As String = "") As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        If PlannedTier1Name = "" Then
            For i = 0 To _rawDTData.PlannedData.Count - 1
                With _rawDTData.PlannedData(i)
                    tmpIndex = tmpList.IndexOf(New DTevent(.Tier3, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Tier3, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End With
            Next
        Else
            For i = 0 To _rawDTData.PlannedData.Count - 1
                With _rawDTData.PlannedData(i)
                    If .Tier1 = PlannedTier1Name And .Tier2 = PlannedTier2Name Then
                        tmpIndex = tmpList.IndexOf(New DTevent(.Tier3, 0))
                        If tmpIndex = -1 Then
                            tmpList.Add(New DTevent(.Tier3, .DT, i))
                        Else
                            tmpList(tmpIndex).addStopWithRow(.DT, i)
                        End If
                    End If
                End With
            Next
        End If
        Return tmpList
    End Function

    Public Function getReason2SubDirectory(Reason1Name As String) As List(Of DTevent)
        Dim i As Integer, tmpIndex As Integer, tmpList As New List(Of DTevent)
        For i = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                If .Reason1 = Reason1Name Then
                    tmpIndex = tmpList.IndexOf(New DTevent(.Reason2, 0))
                    If tmpIndex = -1 Then
                        tmpList.Add(New DTevent(.Reason2, .DT, i))
                    Else
                        tmpList(tmpIndex).addStopWithRow(.DT, i)
                    End If
                End If
            End With
        Next
        Return tmpList
    End Function
#End Region

#Region "Construction / Reinitialization"
    Public Sub New(newParentline As ProdLine, isPastLastDataPoint As Boolean)
        _parentLine = newParentline
        _rawDTData = New DowntimeDataset(newParentline)
        If isPastLastDataPoint Then
            _DT = 0
            _UT = 0
            _schedTime = 0
            _PDT = 0
        End If
    End Sub
    'constructors
    Public Sub New(newParentline As ProdLine, startTime As Date, endTime As Date)
        'get our raw data
        _parentLine = newParentline
        _rawDTData = newParentline.rawDowntimeData.getSubset(startTime, endTime)
        initializeFilterAnalysis()
        initializeTacticalAnalysis()
        initializeTacticalAnalysis_Planned()
        initializeTacticalAnalysis_Curtailment()

        'check for that last event
        With _rawDTData
            If .rawConstraintData.Count > 1 Then
                If .rawConstraintData(.rawConstraintData.Count - 1).DT = 0 Then
                    If .rawConstraintData(.rawConstraintData.Count - 2).isExcluded = False Then
                        _UT = _UT + .rawConstraintData(.rawConstraintData.Count - 1).UT
                    End If
                End If
            End If
        End With

        If _parentLine.doIincludeAllUptime Then initializeUT_IncludeAll()

        _DT = _PDT + _UPDT
        _schedTime = _UT + _DT

        If newParentline._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode = MultiConstraintAnalysis.NoRateLossStops Then
            For listIncrementer As Integer = 0 To _rawDTData.UnplannedData.Count - 1
                If _rawDTData.UnplannedData(listIncrementer).MasterProductionUnit.Contains("Rate") Then _rateLossEvents += 1
            Next
        End If
    End Sub

    Public Sub reMapDataSet(mappingA as integer, mappingB as integer)
        Dim tmpIndex As Integer
        'change the mapped field
     '   _rawDTData.reMapData(My.Settings.defaultDownTimeField, My.Settings.defaultDownTimeField_Secondary)
        _rawDTData.reMapData(mappingA, mappingB)
        'clear existing directory
        MappedDirectory.Clear()
        'recreate the directory
        For i As Integer = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                'mapped stuff
                tmpIndex = MappedDirectory.IndexOf(New DTevent(.MappedField, 0))
                If tmpIndex = -1 Then
                    MappedDirectory.Add(New DTevent(.MappedField, .DT, i))
                Else
                    MappedDirectory(tmpIndex).addStopWithRow(.DT, i)
                End If
            End With
        Next
    End Sub

    Public Sub reFilterData_SKU(inclusionList As List(Of String))
        reFilter_Initialize()
        rawDTdata.reFilterData_SKU(inclusionList)
        reFilter_Finalize()
    End Sub
    Public Sub reFilterData_Team(inclusionList As List(Of String))
        reFilter_Initialize()
        rawDTdata.reFilterData_Team(inclusionList)
        reFilter_Finalize()
    End Sub
    Public Sub reFilterData_Shape(inclusionList As List(Of String))
        reFilter_Initialize()
        rawDTdata.reFilterData_Shape(inclusionList)
        reFilter_Finalize()
    End Sub
    Public Sub reFilterData_Format(inclusionList As List(Of String))
        reFilter_Initialize()
        rawDTdata.reFilterData_Format(inclusionList)
        reFilter_Finalize()
    End Sub
    Public Sub reFilterData_ProductGroup(inclusionList As List(Of String))
        reFilter_Initialize()
        rawDTdata.reFilterData_ProductGroup(inclusionList)
        reFilter_Finalize()
    End Sub
    Public Sub reFilterData_ClearAllFilters()
        reFilter_Initialize()
        rawDTdata.reFilterData_ClearAllFilters()
        reFilter_Finalize()
    End Sub
    Private Sub reFilter_Initialize()
        _UPDT = 0
        _PDT = 0
        _UT = 0
        _DT = 0
        _excludedTime = 0
        _schedTime = 0

        _rateLossEvents = 0

        'UNPLANNED - raw data 
        FaultDirectory.Clear()
        Reason1Directory.Clear()

        DTgroupDirectory.Clear()
        LocationDirectory.Clear()

        Reason2Directory.Clear()
        Reason3Directory.Clear()
        Reason4Directory.Clear()

        'mapped top levels
        Tier1Directory.Clear()
        OneClickDirectory.Clear()

        'PLANNED
        PlannedTier1Directory.Clear()

        MappedDirectory.Clear()
    End Sub
    Private Sub reFilter_Finalize()
        initializeTacticalAnalysis()
        initializeTacticalAnalysis_Planned()
        initializeTacticalAnalysis_Curtailment()
        If _parentLine.doIincludeAllUptime Then initializeUT_IncludeAll()

        _DT = _PDT + _UPDT
        _schedTime = _UT + _DT

        If _parentLine._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode = MultiConstraintAnalysis.NoRateLossStops Then
            For listIncrementer As Integer = 0 To _rawDTData.UnplannedData.Count - 1
                If _rawDTData.UnplannedData(listIncrementer).MasterProductionUnit.Contains("Rate") Then _rateLossEvents += 1
            Next
        End If
    End Sub


    Private Sub initializeFilterAnalysis()
        'set up our sorting fields
        Dim tmpGCAS As New List(Of String)
        Dim tmpProducts As New List(Of String)
        Dim tmpTeams As New List(Of String)
        Dim tmpShapes As New List(Of String)
        Dim tmpFormats As New List(Of String)
        Dim tmpProductGroups As New List(Of String)
        'look at all the unplanned data
        For i As Integer = 0 To _rawDTData.rawConstraintData.Count - 1
            With _rawDTData.rawConstraintData(i)
                If Not .isExcluded Then
                    'sorting fields
                    tmpFormats.Add(.Format)
                    tmpProducts.Add(.Product)
                    tmpGCAS.Add(.ProductCode)
                    tmpShapes.Add(.Shape)
                    tmpTeams.Add(.Team)
                    tmpProductGroups.Add(.ProductGroup)
                End If
            End With
        Next
        'finalize our sorting fields
        ActiveGCAS = tmpGCAS.Distinct().ToList
        ActiveProducts = tmpProducts.Distinct().ToList
        ActiveTeams = tmpTeams.Distinct().ToList
        ActiveShapes = tmpShapes.Distinct().ToList
        ActiveFormats = tmpFormats.Distinct().ToList
        ActiveProductGroups = tmpProductGroups.Distinct().ToList
    End Sub
    Private Sub initializeTacticalAnalysis()
        Dim tmpIndex As Integer
        'look at all the unplanned data
        For i As Integer = 0 To _rawDTData.UnplannedData.Count - 1
            With _rawDTData.UnplannedData(i)
                _UPDT += .DT
                _UT += .UT

                If _UT = 0 Then _excludedStops += 1
                'data fields

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
                'onceclick
                tmpIndex = OneClickDirectory.IndexOf(New DTevent(.OneClick, 0))
                If tmpIndex = -1 Then
                    OneClickDirectory.Add(New DTevent(.OneClick, .DT, i))
                Else
                    OneClickDirectory(tmpIndex).addStopWithRow(.DT, i)
                End If
                'mapped stuff
                tmpIndex = MappedDirectory.IndexOf(New DTevent(.MappedField, 0))
                If tmpIndex = -1 Then
                    MappedDirectory.Add(New DTevent(.MappedField, .DT, i))
                Else
                    MappedDirectory(tmpIndex).addStopWithRow(.DT, i)
                End If

            End With
        Next

    End Sub

    Private Sub initializeTacticalAnalysis_Planned()
        Dim tmpIndex As Integer
        'look at all the unplanned data
        For i As Integer = 0 To _rawDTData.PlannedData.Count - 1
            With _rawDTData.PlannedData(i)
                _PDT += .DT
                _UT += .UT
                'data fields

                'tier1
                tmpIndex = PlannedTier1Directory.IndexOf(New DTevent(.Tier1, 0))
                If tmpIndex = -1 Then
                    PlannedTier1Directory.Add(New DTevent(.Tier1, .DT, i))
                Else
                    PlannedTier1Directory(tmpIndex).addStopWithRow(.DT, i)
                End If
            End With
        Next
    End Sub

    Private Sub initializeTacticalAnalysis_Curtailment()
        For i As Integer = 0 To _rawDTData.CurtailmentData.Count - 1
            With _rawDTData.CurtailmentData(i)
                _UT += .UT
            End With
        Next
    End Sub

    Private Sub initializeUT_IncludeAll()
        _UT = 0 'reset this, need all UT
        For i As Integer = 0 To _rawDTData.rawConstraintData.Count - 1
            _UT += _rawDTData.rawConstraintData(i).UT
        Next
    End Sub


#End Region

End Class

#Region "Downtime Events"
Public Class DTeventFields
    Private _startTime As Date
    Private _DT As Double
    Private _priorUT As Double
    Private _Comments As String
    Private _RowNumber As Long 'position in raw data array
    Private _ProductDescription As String
    Private _Team As String

    Public Overrides Function ToString() As String
        If Len(_Comments) < 2 Then
            Return "Start Time: " & _startTime & " Duration: " & " min" & vbCrLf & "Comments: " & _Comments
        Else
            Return "Start Time: " & _startTime & " Duration: " & " min" & vbCrLf & "No Comment " & _Comments
        End If
    End Function

    Public Sub New(tmpEvent As DowntimeEvent)
        Me.New(tmpEvent.DT, tmpEvent.UT, tmpEvent.Comment, tmpEvent.startTime, tmpEvent.Team, tmpEvent.Product)
    End Sub

    Public Sub New(Downtime As Double, Uptime As Double, Comment As String, StartTime As Date, Optional RowNumber As Long = -1)
        _startTime = StartTime
        _Comments = Comment
        _DT = Math.Round(Downtime, 1)
        _priorUT = Uptime
        _RowNumber = RowNumber
    End Sub
    Public Sub New(Downtime As Double, Uptime As Double, Comment As String, StartTime As Date, Team As String, Product As String)
        Me.New(Downtime, Uptime, Comment, StartTime)
        _Team = Team
        _ProductDescription = Product
    End Sub

    Public ReadOnly Property DT
        Get
            Return _DT
        End Get
    End Property
    Public ReadOnly Property UT
        Get
            Return _priorUT
        End Get
    End Property
    Public ReadOnly Property Comment
        Get
            Return _Comments
        End Get
    End Property
    Public ReadOnly Property StartTime
        Get
            Return _startTime
        End Get
    End Property
    Public ReadOnly Property Row
        Get
            Return _RowNumber
        End Get
    End Property
    Public ReadOnly Property Product As String
        Get
            Return _ProductDescription
        End Get
    End Property
    Public ReadOnly Property Team As String
        Get
            Return _Team
        End Get
    End Property
End Class

Public MustInherit Class simpleDTevent
    Protected _Name As String
    Protected _DTColumn As Integer
    Protected _parentname As String
    Protected _Stops As Long
    Protected _DT As Double
    Protected _DTpct As Double = 0
    Protected _SPD As Double = 0

    'Added For Loss Allocation
    Protected _schedTime As Double = 0

    Protected _DTsim As Double = 0
    Protected _StopsSim As Double = 0
    Protected _UT As Double

    Protected isSortByDT As Boolean = True

    'public properties
    Public ReadOnly Property Column As Integer
        Get
            Return _DTColumn
        End Get
    End Property
    Public Property DT As Double
        Get
            Return _DT
        End Get
        Set(Downtime As Double)
            _DT = Downtime
        End Set
    End Property
    Public Property UT As Double
        Get
            Return _UT
        End Get
        Set(value As Double)
            _UT = value
        End Set
    End Property

    Public ReadOnly Property DT_Display As Double
        Get
            Return Math.Round(_DT, 2)
        End Get
    End Property
    Public Property Stops
        Get
            Return _Stops
        End Get
        Set(stopNum)
            _Stops = stopNum
        End Set
    End Property
    Public ReadOnly Property Name As String
        Get
            Return _Name
        End Get
    End Property
    Public ReadOnly Property ParentName As String
        Get
            Return _parentname
        End Get
    End Property
    Public Property StopsSim As Double
        Get
            Return _StopsSim
        End Get
        Set(value As Double)
            _StopsSim = value
        End Set
    End Property

    'constructors
    Public Sub New(eventName As String, firstDT As Double, Optional dtCol As Integer = -1)
        _Name = eventName
        _DTColumn = dtCol
        _DT = firstDT
        _Stops = 1
        isSortByDT = True
    End Sub
    Public Sub New(eventName As String, firstDT As Double, Optional dtCol As Integer = -1, Optional parentnamevalue As String = "")
        _Name = eventName
        _DTColumn = dtCol
        _DT = firstDT
        _Stops = 1
        isSortByDT = True
        _parentname = parentnamevalue
    End Sub


    'basic functionality
    Public Sub addStop(downTime As Double)
        _Stops = _Stops + 1
        _DT = _DT + downTime
    End Sub

#Region "Sort Parameters"
    Public ReadOnly Property SortParam
        Get
            If isSortByDT Then
                Return _DT
            Else
                Return _Stops
            End If
        End Get
    End Property
    Public ReadOnly Property SortParamSecondary
        Get
            If Not isSortByDT Then
                Return _DT
            Else
                Return _Stops
            End If
        End Get
    End Property


    Public Sub sortBy_DT()
        isSortByDT = True
    End Sub
    Public Sub sortBy_Stops()
        isSortByDT = False
    End Sub
#End Region
    Public ReadOnly Property SPDspecialrounded
        Get
            Return Math.Round(SPDtemp, 1)
        End Get
    End Property
    Public Property SPDspecial As Double
        Get
            Return SPDtemp

        End Get

        Set(ByVal value As Double)
            SPDtemp = value
        End Set

    End Property
    Public ReadOnly Property DTpctspecialrounded
        Get
            Return Math.Round(100 * DTpcttemp, 1)
        End Get
    End Property

    Private DTpcttemp As Double
    Private SPDtemp As Double
    Private DTtemp As Double = 0
    Private Stopstemp As Double = 0

    Public ReadOnly Property MTTRspecial As Double
        Get
            If Stopsspecial = 0 Then
                Return 0
            Else
                Return Math.Round(DTspecial / Stopsspecial,1)
            End If
        End Get
    End Property

    Public Property DTspecial As Double
        Get
            Return DTtemp

        End Get

        Set(ByVal value As Double)
            DTtemp = value
        End Set

    End Property

    Public Property Stopsspecial As Double
        Get
            Return Stopstemp

        End Get

        Set(ByVal value As Double)
            Stopstemp = value
        End Set

    End Property

    Public Property DTpctspecial As Double
        Get
            Return DTpcttemp

        End Get

        Set(ByVal value As Double)
            DTpcttemp = value
        End Set

    End Property
    Public ReadOnly Property DTpctrounded
        Get
            Return Math.Round(100 * _DTpct, 1)
        End Get

    End Property
    Public ReadOnly Property SPDrounded
        Get
            Return Math.Round(_SPD, 1)
        End Get

    End Property
    Public Property DTpct
        Get
            Return _DTpct
        End Get
        Set(schedTime)
            If _DTpct = 0 Then
                If schedTime = 0 Then
                    _DTpct = 0
                    DTpcttemp = 0
                Else
                    _schedTime = schedTime
                    _DTpct = (_DT / schedTime)
                    _SPD = (_Stops / schedTime) * 1440
                    DTpcttemp = _DTpct
                    SPDtemp = _SPD
                End If
            End If
        End Set
    End Property
    Public Property SPD
        Get
            Return _SPD
        End Get
        Set(schedTime)
            If _SPD = 0 Then
                _DTpct = (_DT / schedTime)
                _SPD = (_Stops / schedTime) * 1440
                SPDtemp = _SPD
            End If
        End Set
    End Property

End Class

Public Class DTevent
    Inherits simpleDTevent
    Implements IComparable(Of DTevent)
    Implements IEquatable(Of DTevent)

    Friend RawInfo As New List(Of DTeventFields)
    Friend RawRows As New List(Of Integer)

    public property SchedTime as double

#Region "MTD / Secondary"
    Private _secondaryStops As Double
    Private _secondaryDT As Double

    Private WriteOnly Property SecondaryStops As Double
        Set(value As Double)
            _secondaryStops = value
        End Set
    End Property
    Private WriteOnly Property SecondaryDT As Double
        Set(value As Double)
            _secondaryDT = value
        End Set
    End Property

    Public Sub swapSecondaryValues()
        Dim tmpStop As Double
        Dim tmpDT As Double

        tmpStop = _StopsSim
        tmpDT = _DTsim

        _StopsSim = _secondaryStops
        _DTsim = _secondaryDT

        _secondaryStops = tmpStop
        _secondaryDT = tmpDT

    End Sub
#End Region


#Region "LOSS ALLOCATION"
    Private _ParentEventName As String
    Private _LossAllocation As Double ' = 0

    Private _MTTR_userScaleFactor As Double = 1
    Private _MTBF_userScaleFactor As Double = 1
    Private _MTBF_parentScaleFactor As Double = 1
    Private _MTTR_parentScaleFactor As Double = 1
    Private _MTBF_parentTwoScaleFactor As Double = 1
    Private _MTTR_parentTwoScaleFactor As Double = 1

    Public ReadOnly Property MTBF_netScaleFactor As Double
        Get
            Return _MTBF_parentScaleFactor * _MTBF_parentTwoScaleFactor * _MTBF_userScaleFactor
        End Get
    End Property
    Public ReadOnly Property MTTR_netScaleFactor As Double
        Get
            Return _MTTR_parentScaleFactor * _MTTR_parentTwoScaleFactor * _MTTR_userScaleFactor
        End Get
    End Property

    Public Property MTTR_parentScaleFactor As Double
        Get
            Return _MTTR_parentScaleFactor
        End Get
        Set(value As Double)
            _MTTR_parentScaleFactor = value
        End Set
    End Property
    Public Property MTBF_parentScaleFactor As Double
        Get
            Return _MTBF_parentScaleFactor
        End Get
        Set(value As Double)
            _MTBF_parentScaleFactor = value
        End Set
    End Property
    Public Property MTTR_parentTwoScaleFactor As Double
        Get
            Return _MTTR_parentTwoScaleFactor
        End Get
        Set(value As Double)
            _MTTR_parentTwoScaleFactor = value
        End Set
    End Property
    Public Property MTBF_parentTwoScaleFactor As Double
        Get
            Return _MTBF_parentTwoScaleFactor
        End Get
        Set(value As Double)
            _MTBF_parentTwoScaleFactor = value
        End Set
    End Property

    Public Property MTTR_userScaleFactor As Double
        Get
            Return _MTTR_userScaleFactor
        End Get
        Set(value As Double)
            _MTTR_userScaleFactor = value
        End Set
    End Property
    Public Property MTBF_userScaleFactor As Double
        Get
            Return _MTBF_userScaleFactor
        End Get
        Set(value As Double)
            _MTBF_userScaleFactor = value
        End Set
    End Property

    Public Property LossAllocation
        Get
            Return _LossAllocation
        End Get
        Set(value)
            _LossAllocation = value
        End Set
    End Property

    Public Property ParentEvent As String
        Get
            Return _ParentEventName
        End Get
        Set(value As String)
            _ParentEventName = value
        End Set
    End Property

    Public Sub initializeSimulation()
        _StopsSim = _Stops
        _DTsim = _DT
        _LossAllocation = 0
    End Sub

    Public ReadOnly Property MTTRsim As Double
        Get
            If _StopsSim = 0 Then Return 0
            Return _DTsim / _StopsSim
        End Get
    End Property
    Public ReadOnly Property MTBFsim As Double
        Get
            If _StopsSim = 0 Then Return 0
            Return _UT / _StopsSim
        End Get
    End Property

    Public ReadOnly Property SPDsim As Double
        Get
            If _schedTime = 0 Then Return 0
            Return _StopsSim * 1440 / _schedTime
        End Get
    End Property

    Public Property DTsim As Double
        Get
            Return _DTsim
        End Get
        Set(value As Double)
            _DTsim = value
        End Set
    End Property

    Public ReadOnly Property DTpctSim As Double
        Get
            '   If _schedTime = 0 Then Return 0
            '  Return DTsim / _schedTime
            Return _LossAllocation 'DTsim
        End Get
    End Property

    Public ReadOnly Property Avail As Double
        Get
            ' Return MTBF / (MTBF + MTTR)
            Return _UT / (_DT + _UT)
        End Get
    End Property
    Public ReadOnly Property AvailSim As Double
        Get
            '  Return MTBF / (MTBF + MTTR)
            If _UT = 0 Then Return 0
            Return _UT / (_DTsim + _UT)
        End Get
    End Property
    Public ReadOnly Property AvailSim_System As Double
        Get
            If AvailSim = 0 Then Return 0
            Return (1 - AvailSim) / AvailSim
        End Get
    End Property
#End Region

    Public Overrides Function ToString() As String
        Return _Name & " DT: " & Math.Round(_DT, 1) & " / " & Math.Round(_DTsim, 1) & " LA: " & Math.Round(_LossAllocation, 3)
    End Function

    Public ReadOnly Property MTTR As Double
        Get
            If _Stops = 0 Then Return 0
            Return _DT / _Stops
        End Get
    End Property

    Public ReadOnly Property MTTRrounded As Double
        Get
            Return Math.Round(MTTR, 1)
        End Get
    End Property

    Public Property MTBF As Double
        Get
            Return _UT / Stops
        End Get
        Set(upTime As Double)
            _UT = upTime
        End Set
    End Property



    Public ReadOnly Property MTBFrounded As Double

        Get
            Return Math.Round((_UT / Stops), 1)

        End Get
    End Property

#Region "Construction"
    Public Sub New(eventName As String, firstDT As Double)
        MyBase.New(eventName, firstDT, -1)
        If eventName = "" Then
            If firstDT = 0 Then
                _Stops = 0
            End If
        End If
    End Sub
    Public Sub New(eventName As String, firstDT As Double, parentname As String)
        MyBase.New(eventName, firstDT, -1, parentname)
        If eventName = "" Then
            If firstDT = 0 Then
                _Stops = 0
            End If
        End If
    End Sub
    Public Sub New(eventName As String, firstDT As Double, rowNum As Integer)
        Me.New(eventName, firstDT)
        If rowNum > -1 Then RawRows.Add(rowNum)
    End Sub
    Public Sub New(eventName As String, firstDT As Double, rowNum As Integer, Optional parentnameval As String = "")
        Me.New(eventName, firstDT, parentnameval)
        If rowNum > -1 Then RawRows.Add(rowNum)
    End Sub

    Public Sub New(eventName As String, firstDT As Double, rowNum As Integer, dtCol As Integer)
        MyBase.New(eventName, firstDT, dtCol)
        If rowNum > -1 Then RawRows.Add(rowNum)
    End Sub
    Public Sub New(eventName As String, firstDT As Double, upTime As Double, Comment As String, startTime As Date, rowNum As Integer, dtCol As Integer)
        Me.New(eventName, firstDT, rowNum, dtCol)
        RawInfo.Add(New DTeventFields(firstDT, upTime, Comment, startTime))
    End Sub



#End Region

    'adding a new event w/ or w/o fields
    Public Sub addStopWithFields(downTime As Double, upTime As Double, Comment As String, startTime As Date)
        addStop(downTime)
        RawInfo.Add(New DTeventFields(downTime, upTime, Comment, startTime))
    End Sub
    Public Sub addStopWithRow(downTime As Double, rowNum As Long)
        MyBase.addStop(downTime)
        RawRows.Add(rowNum)
    End Sub

#Region "Implements Equitable & Comparable"
    'properties
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As DTevent = TryCast(obj, DTevent)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As DTevent) As Boolean _
        Implements IEquatable(Of DTevent).Equals
        If other Is Nothing Then
            Return False
        End If
        Return (Me.Name.Equals(other.Name))
    End Function

    'sortable
    ''' <summary>
    ''' Sorts DTEvents by Stops Or Downtime
    ''' </summary>
    ''' <param name="Other">DTEvent to be compared</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function CompareTo(ByVal Other As DTevent) As Integer Implements System.IComparable(Of DTevent).CompareTo
        If Other.SortParam = SortParam Then
            If Other.SortParamSecondary = SortParamSecondary Then
                Return _Name.Equals(Other.Name)
            ElseIf Other.SortParamSecondary > SortParamSecondary Then
                Return 1
            Else
                Return -1
            End If
        ElseIf Other.SortParam > SortParam Then
            Return 1
        Else 'Other.SortParam < _Sortable
            Return -1
        End If
    End Function
#End Region
End Class

#End Region
