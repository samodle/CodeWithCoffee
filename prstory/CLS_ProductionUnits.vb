
Imports Newtonsoft.Json

Public Class ProdLine
    Implements IEquatable(Of ProdLine)

#Region "Variables & Properties"
    Public IS_MOT As Boolean = False
    <JsonIgnore()>
    Public RawRateLossDataArray As Object(,)
    Public RateLoss_Mode As Integer = -1
    Public IsStartupMode As Boolean = False

    Public MultilineGroupName As String = ""

    Public ReadOnly Property MultilineGroup As String
        Get
            If MultilineGroupName = "" Then
                Return parentModule.MultilineGroup
            Else
                Return MultilineGroupName
            End If
        End Get
    End Property

    '  Public ServerDatabase As String = "GBDB"
    Public ReadOnly Property ServerDatabase As String
        Get
            Return parentSite.ServerDatabase
        End Get
    End Property

    Public _UnplannedT1List As New List(Of String)
    Public _PlannedT1List As New List(Of String)
    <JsonIgnore()>
    Public ReadOnly Property UnplannedT1List As List(Of String)
        Get
            'Return _UnplannedT1List
            If _UnplannedT1List.Count = 0 And parentModule.UnplannedT1List.Count = 0 Then
                Return getUnplannedT1List()
            ElseIf parentModule.UnplannedT1List.Count > 0 Then
                Return parentModule.UnplannedT1List
            Else
                Return _UnplannedT1List
            End If
        End Get
    End Property
    Private Function getUnplannedT1List() As List(Of String)
        Dim T1inc As Integer = 1
        Dim tmpName As String
        Dim tmpList As New List(Of String)

        tmpName = getprStoryCardField(prStoryMapping, prStoryCard.Unplanned, T1inc)
        While tmpName.Length > 1 And tmpList.Count < 25 '<25 added as error handler
            tmpList.Add(tmpName)
            T1inc += 1
            tmpName = getprStoryCardField(prStoryMapping, prStoryCard.Unplanned, T1inc)
        End While
        Return tmpList
    End Function
    <JsonIgnore()>
    Public ReadOnly Property PlannedT1List As List(Of String)
        Get
            Return _PlannedT1List
            If _PlannedT1List.Count = 0 And parentModule.PlannedT1List.Count = 0 Then
                Return getPlannedT1List()
            ElseIf parentModule.PlannedT1List.Count > 0 Then
                Return parentModule.PlannedT1List
            Else
                Return _PlannedT1List
            End If
        End Get
    End Property
    Private Function getPlannedT1List() As List(Of String)
        Dim T1inc As Integer = 1
        Dim tmpName As String
        Dim tmpList As New List(Of String)

        tmpName = getprStoryCardField(prStoryMapping, prStoryCard.Planned, T1inc)
        While tmpName.Length > 1
            tmpList.Add(tmpName)
            T1inc += 1
            tmpName = getprStoryCardField(prStoryMapping, prStoryCard.Planned, T1inc)
        End While
        Return tmpList
    End Function

    'parent site/line
    ' <JsonIgnore()>
    Public parentModule As ProdModule
    ' <JsonIgnore()>
    Public parentSite As ProdSite


    'targets
    <JsonIgnore()>
    Private DTtargets As DTPct_Targets
    <JsonIgnore()>
    Private _doIhaveTargets As Boolean = False
    Public doIincludeAllUptime As Boolean = False
    <JsonIgnore()>
    Public ReadOnly Property doIhaveTargets As Boolean
        Get
            Return _doIhaveTargets And My.Settings.AdvancedSettings_isTargetsEnabled
        End Get
    End Property
    <JsonIgnore()>
    Public Property DowntimePercentTargets As DTPct_Targets
        Get
            Return DTtargets
        End Get
        Set(value As DTPct_Targets)
            DTtargets = value
            _doIhaveTargets = True
        End Set
    End Property


    'static fields
    Protected _lineName As String = ""
    Private _lineSAPname As String = ""
    Private _lineMAPLEname As String = ""
    Private _defaultDowntimeFieldA As Integer = -1
    Private _defaultDowntimeFieldB As Integer = -1

    Private _customServer As String = ""
    Private useCustomServer As Boolean = False
    Private _customUser As String = ""
    Private _customPassword As String = ""
    Private useCustomLogin As Boolean = False

    'dual constraint stuff
    Friend _isDualConstraint As Boolean
    Public mainProdUnits As List(Of String) = New List(Of String)
    Protected _mainProdUnit As String
    Protected _mainProfProd As String
    Friend _rateLossDisplay As String
    Protected _rateLossEvents As Long
    <JsonIgnore()>
    Private _rawRateLossData As Array

    Protected _prStoryMapping As Integer
    Public formatMapping As Integer = MappingByFormat.NoMapping
    Public shapeMapping As Integer = MappingByShape.NoMapping
    Public OneClickMapping As Integer = MappingOneClick.NoMapping

    'shift configuration, critical for production data
    Protected _numberOfShifts As Integer = 2
    Protected _shiftDurationHrs As Double = 8
    Public _DayStartTimeHrs As Double

    Protected _FirstShitStartHrs As Double = 0
    Protected _SecondShiftStartHrs As Double = 0
    Protected _ThirdShiftStartHrs As Double = 0

    Public ReadOnly Property ShiftStartFirst_Hr
        Get
            Return _DayStartTimeHrs
        End Get
    End Property
    Public ReadOnly Property ShiftStartSecond_Hr
        Get
            Return _SecondShiftStartHrs
        End Get
    End Property
    Public ReadOnly Property ShiftStartThird_Hr
        Get
            If _numberOfShifts = 3 Then
                Return _ThirdShiftStartHrs
            Else
                'MsgBox("Only " & _numberOfShifts & " Shifts! You asked for 3! Will Return -1. Best of Luck...")
                Return -1
            End If
        End Get
    End Property
    Public ReadOnly Property NumberOfShifts
        Get
            Return _numberOfShifts
        End Get
    End Property


    'raw data
    Private _rawProfStartTime As Date
    Private _rawProfEndTime As Date

    <JsonIgnore()>
    Friend rawDowntimeData As DowntimeDataset
    <JsonIgnore()>
    Friend rawProductionData As ProductionDataset
    <JsonIgnore()>
    Private _rawProficyData As Array
    <JsonIgnore()>
    Private _rawProficyProductionData As Array
    <JsonIgnore()>
    Public Property rawRateLossData As Array
        Get
            Return _rawRateLossData
        End Get
        Set(value As Array)
            _rawRateLossData = value
        End Set
    End Property
    <JsonIgnore()>
    Public Property rawProficyData As Array
        Get
            Return _rawProficyData
        End Get
        Set(value As Array)
            _rawProficyData = value
            '  rawDowntimeEvents.Clear()
            BrandCodesWeWant.Clear()
            BrandCodeReport.Clear()
            ' populateReportLists()
        End Set
    End Property
    <JsonIgnore()>
    Public Property rawProficyProductionData As Array
        Get
            Return _rawProficyProductionData
        End Get
        Set(value As Array)
            _rawProficyProductionData = value
        End Set
    End Property
    <JsonIgnore()>
    Public Property rawProfStartTime As Date
        Get
            Return _rawProfStartTime
        End Get
        Set(value As Date)
            _rawProfStartTime = value
        End Set
    End Property
    <JsonIgnore()>
    Public Property rawProfEndTime As Date
        Get
            Return _rawProfEndTime
        End Get
        Set(value As Date)
            _rawProfEndTime = value
        End Set
    End Property
    <JsonIgnore()>
    Public Property MappingLevelA As Integer
        Get
            If _defaultDowntimeFieldA = -1 Then
                Return parentModule.MappingLevelA
            Else
                Return _defaultDowntimeFieldA
            End If
        End Get
        Set(value As Integer)
            _defaultDowntimeFieldA = value
        End Set
    End Property
    <JsonIgnore()>
    Public Property MappingLevelB As Integer
        Get
            If _defaultDowntimeFieldB = -1 Then
                Return parentModule.MappingLevelB
            Else
                Return _defaultDowntimeFieldB
            End If
        End Get
        Set(value As Integer)
            _defaultDowntimeFieldB = value
        End Set
    End Property
    <JsonIgnore()>
    Public ReadOnly Property Mapping_DTschedPlannedUnplanned As Integer
        Get
            Return parentModule.DTschedMap
        End Get
    End Property

    Private Enum DownTimeColumn
        StartTime = 0
        Endtime = 1
        PlannedUnplanned = 15
    End Enum

    Public isFilterByBrandcode As Boolean = False
    <JsonIgnore()>
    Public BrandCodesWeWant As New List(Of String)
    <JsonIgnore()>
    Public BrandCodeReport As New List(Of String) 'production based
    <JsonIgnore()>
    Public ShiftReport As New List(Of String) 'dt based
    <JsonIgnore()>
    Public TeamReport As New List(Of String) 'dt based
    <JsonIgnore()>
    Public ProductReport As New List(Of String) 'dt based

    Public BabyWipesData As New List(Of Object(,))
    Public BabyWipesPRSTORYData As New List(Of DowntimeDataset)

    'properties
    <JsonIgnore()>
    Public ReadOnly Property SQLdowntimeProcedure As Integer
        Get
            Return parentModule.SQLprocedure
        End Get
    End Property
    <JsonIgnore()>
    Public ReadOnly Property SQLproductionProcedure As Integer
        Get
            Return parentModule.SQLprocedurePROD
        End Get
    End Property
    <JsonIgnore()>
    Public ReadOnly Property Sector As String
        Get
            Return parentModule.Sector
        End Get
    End Property
    <JsonIgnore()>
    Public ReadOnly Property SiteName As String
        Get
            Return parentSite.Name
        End Get
    End Property


    Public ReadOnly Property mainProdUnit As String
        Get
            Return _mainProdUnit
        End Get
    End Property
    Public ReadOnly Property mainProfProd As String
        Get
            Return _mainProfProd
        End Get
    End Property
    Public ReadOnly Property Name As String
        Get
            Return _lineName
        End Get
    End Property
    Public Property Name_MAPLE As String
        Get
            Return _lineMAPLEname
        End Get
        Set(value As String)
            _lineMAPLEname = value
        End Set
    End Property
    Public ReadOnly Property prStoryMapping As Integer
        Get
            Return _prStoryMapping
        End Get
    End Property
    <JsonIgnore()>
    Public ReadOnly Property Reason1Name As String
        Get
            Return parentModule._Reason1Name
        End Get
    End Property
    <JsonIgnore()>
    Public ReadOnly Property Reason2Name As String
        Get
            Return parentModule._Reason2Name
        End Get
    End Property
    <JsonIgnore()>
    Public ReadOnly Property Reason3Name As String
        Get
            Return parentModule._Reason3Name
        End Get
    End Property
    <JsonIgnore()>
    Public ReadOnly Property Reason4Name As String
        Get
            Return parentModule._Reason4Name
        End Get
    End Property
    Public ReadOnly Property FaultCodeName As String
        Get
            Return "Fault Code"
        End Get
    End Property
    Public ReadOnly Property DTgroupName As String
        Get
            Return "DT Group"
        End Get
    End Property

    <JsonIgnore()>
    Public Property ProficyServer_Name
        Get
            If useCustomServer Then
                Return _customServer
            Else
                Return parentSite.ProficyServer
            End If

        End Get
        Set(value)
            useCustomServer = True
            _customServer = value
        End Set
    End Property
    <JsonIgnore()>
    Public Property ProficyServer_Username
        Get
            If useCustomLogin Then
                Return _customUser
            Else
                Return parentSite.ProficyServer_Username
            End If
        End Get
        Set(value)
            useCustomLogin = True
            _customUser = value
        End Set
    End Property
    Public Property ProficyServer_Password
        Get
            If useCustomLogin Then
                Return _customPassword
            Else
                Return parentSite.ProficyServer_Password
            End If

        End Get
        Set(value)
            useCustomLogin = True
            _customPassword = value
        End Set
    End Property

    'showing fields
    Private _doIuseProductGroup As Boolean = False
    Public Property FieldCheck_ProductGroup As Boolean
        Get
            Return _doIuseProductGroup
        End Get
        Set(value As Boolean)
            _doIuseProductGroup = value
        End Set
    End Property

#End Region

    Public Overrides Function ToString() As String
        Return parentSite.Name & " " & parentModule.Name & " " & _lineName
    End Function

#Region "Construction / Mapping / Filtering"
    'constructor
    Public Sub New(ByVal lineName As String, ParentSiteName As String)
        _lineName = lineName
        _lineSAPname = lineName
        parentSite = AllProductionSites(getSiteIndexFromName(ParentSiteName))
    End Sub
    Public Sub New(ByVal lineName As String, SAPname As String, ByVal ParentSiteName As String, ByVal parentModuleID As Guid, ByVal numShifts As Integer, ByVal shiftLenHrs As Double, ByVal DailyStartTimeHr As Double, ByVal MainProfDT As String, mainProfProd As String, Optional specialPRSTORYmapping As Integer = -1, Optional ByVal dualConstraint As Boolean = False, Optional RateLossDisplay As String = "")
        Me.New(lineName, ParentSiteName)

        Dim parentModuleIndex As Integer = getModuleIndexFromID(parentModuleID)

        parentModule = AllProdModules(parentModuleIndex)
        AllProdModules(parentModuleIndex).LinesList.Add(Me)
        _mainProdUnit = MainProfDT
        _mainProfProd = mainProfProd
        _isDualConstraint = dualConstraint
        _numberOfShifts = numShifts
        _shiftDurationHrs = shiftLenHrs
        _DayStartTimeHrs = DailyStartTimeHr
        _lineSAPname = SAPname
        If specialPRSTORYmapping = -1 Then
            _prStoryMapping = AllProdModules(parentModuleIndex).prStory_Mapping
        Else
            _prStoryMapping = specialPRSTORYmapping
        End If
        _rateLossDisplay = RateLossDisplay
        Select Case numShifts
            Case 2
                _SecondShiftStartHrs = _DayStartTimeHrs + _shiftDurationHrs
            Case 3
                _SecondShiftStartHrs = _DayStartTimeHrs + _shiftDurationHrs
                _ThirdShiftStartHrs = _SecondShiftStartHrs + _shiftDurationHrs
            Case Else
                MsgBox("Error: " & numShifts & " shifts not a recognized schedule format.", vbCritical, "Shift # Err")
                Debugger.Break()
        End Select

    End Sub

    Public Sub New(ByVal lineName As String, SAPname As String, ByVal ParentSiteName As String, ByVal parentModuleID As Guid, ByVal numShifts As Integer, ByVal shiftLenHrs As Double, ByVal DailyStartTimeHr As Double, ByVal MainProfDT As String, mainProfProd As String, specialPRSTORYmapping As Integer, ByVal dualConstraint As Boolean, RateLossDisplay As List(Of String))
        Me.New(lineName, ParentSiteName)

        Dim parentModuleIndex As Integer = getModuleIndexFromID(parentModuleID)

        parentModule = AllProdModules(parentModuleIndex)
        AllProdModules(parentModuleIndex).LinesList.Add(Me)
        _mainProdUnit = MainProfDT
        _mainProfProd = mainProfProd
        _isDualConstraint = dualConstraint
        _numberOfShifts = numShifts
        _shiftDurationHrs = shiftLenHrs
        _DayStartTimeHrs = DailyStartTimeHr
        _lineSAPname = SAPname
        If specialPRSTORYmapping = -1 Then
            _prStoryMapping = AllProdModules(parentModuleIndex).prStory_Mapping
        Else
            _prStoryMapping = specialPRSTORYmapping
        End If
        mainProdUnits = RateLossDisplay
        Select Case numShifts
            Case 2
                _SecondShiftStartHrs = _DayStartTimeHrs + _shiftDurationHrs
            Case 3
                _SecondShiftStartHrs = _DayStartTimeHrs + _shiftDurationHrs
                _ThirdShiftStartHrs = _SecondShiftStartHrs + _shiftDurationHrs
            Case Else
                MsgBox("Error: " & numShifts & " shifts not a recognized schedule format.", vbCritical, "Shift # Err")
                Debugger.Break()
        End Select

    End Sub


    Public Sub New(ByVal lineName As String, ByVal ParentSiteName As String, ByVal parentModuleID As Guid, ByVal numShifts As Integer, ByVal shiftLenHrs As Double, ByVal DailyStartTimeHr As Double, ByVal MainProfDT As String, mainProfProd As String, Optional specialPRSTORYmapping As Integer = -1, Optional ByVal dualConstraint As Boolean = False, Optional RateLossDisplay As String = "", Optional SecondaryDataProfile As Integer = -1)
        Me.New(lineName, ParentSiteName)

        Dim parentModuleIndex As Integer = getModuleIndexFromID(parentModuleID)

        parentModule = AllProdModules(parentModuleIndex)
        AllProdModules(parentModuleIndex).LinesList.Add(Me)
        _mainProdUnit = MainProfDT
        _mainProfProd = mainProfProd
        _isDualConstraint = dualConstraint
        _numberOfShifts = numShifts
        _shiftDurationHrs = shiftLenHrs
        _DayStartTimeHrs = DailyStartTimeHr

        RateLoss_Mode = SecondaryDataProfile

        If specialPRSTORYmapping = -1 Then
            _prStoryMapping = AllProdModules(parentModuleIndex).prStory_Mapping
        Else
            _prStoryMapping = specialPRSTORYmapping
        End If
        _rateLossDisplay = RateLossDisplay
        Select Case numShifts
            Case 2
                _SecondShiftStartHrs = _DayStartTimeHrs + _shiftDurationHrs
            Case 3
                _SecondShiftStartHrs = _DayStartTimeHrs + _shiftDurationHrs
                _ThirdShiftStartHrs = _SecondShiftStartHrs + _shiftDurationHrs
            Case Else
                MsgBox("Error: " & numShifts & " shifts not a recognized schedule format.", vbCritical, "Shift # Err")
                Debugger.Break()
        End Select

    End Sub


    Public Sub New(ByVal lineName As String, ByVal parentModuleID As Guid, ByVal MainProfDT As String, Optional mainProfProd As String = "")
        _lineName = lineName
        _lineSAPname = lineName

        'set site and parent module
        Dim parentModuleIndex As Integer = getModuleIndexFromID(parentModuleID)

        parentModule = AllProdModules(parentModuleIndex)
        parentSite = parentModule.parentSite

        'set everything else
        AllProdModules(parentModuleIndex).LinesList.Add(Me)
        _mainProdUnit = MainProfDT
        _mainProfProd = mainProfProd
        _isDualConstraint = False 'dualConstraint


        RateLoss_Mode = -1

        _prStoryMapping = AllProdModules(parentModuleIndex).prStory_Mapping

        _rateLossDisplay = ""

        'populate shift info from parent module
                _numberOfShifts = parentModule.numShifts
        _shiftDurationHrs = parentModule.shiftdurationhrs
        _DayStartTimeHrs = parentModule.ShiftStartFirst_Hr

        Select Case parentModule.numShifts
            Case 2
                _SecondShiftStartHrs =  parentModule._DayStartTimeHrs +  _shiftDurationHrs
            Case 3
                _SecondShiftStartHrs =  parentModule._DayStartTimeHrs +  _shiftDurationHrs
                _ThirdShiftStartHrs =  _SecondShiftStartHrs +  _shiftDurationHrs
            Case Else
                MsgBox("Error: " &  parentModule.numShifts & " shifts not a recognized schedule format. Will assume 2 shifts", vbCritical, "Shift # Err")
                _SecondShiftStartHrs = _DayStartTimeHrs + _shiftDurationHrs
        End Select
    End Sub

    Public Sub reMapRawData()
        rawDowntimeData.reMapData(My.Settings.defaultDownTimeField, My.Settings.defaultDownTimeField_Secondary)
    End Sub

    Public Sub reFilterData_SKU(inclustionList As List(Of String))

        rawDowntimeData.reFilterData_SKU(inclustionList)

    End Sub
    Public Sub reFilterData_Team(inclustionList As List(Of String))

        rawDowntimeData.reFilterData_Team(inclustionList)

    End Sub
    Public Sub reFilterData_ProductGroup(inclustionList As List(Of String))

        rawDowntimeData.reFilterData_ProductGroup(inclustionList)

    End Sub
    Public Sub reFilterData_Format(inclustionList As List(Of String))

        rawDowntimeData.reFilterData_Format(inclustionList)

    End Sub
    Public Sub reFilterData_Shape(inclustionList As List(Of String))

        rawDowntimeData.reFilterData_Shape(inclustionList)

    End Sub
    Public Sub reFilterData_ClearAllFilters(inclustionList As List(Of String))

        rawDowntimeData.reFilterData_ClearAllFilters()

    End Sub
#End Region

    'due to column stubbing needs to be exact match
    Public Function getProdRowFromTime(tmpTime As Date, Optional isEndTime As Boolean = False) As Integer
        Dim r As Integer, testTime As Date ', i As Integer, j As Double
        getProdRowFromTime = -1
        For r = 0 To rawProficyProductionData.GetLength(1) - 1
            testTime = rawProficyProductionData(ProductionColumn.StartTime, r)

            If tmpTime >= rawProficyProductionData(ProductionColumn.StartTime, r) And tmpTime < rawProficyProductionData(ProductionColumn.EndTime, r) Then
                Return r
            End If


            If DateDiff(DateInterval.Second, tmpTime, testTime) = 0 Then 'testTime = tmpTime Then
                getProdRowFromTime = r
                GoTo FoundIt
            ElseIf r <> 0 And r <> rawProficyProductionData.GetLength(1) - 1 Then 'LG Code
                If tmpTime > rawProficyProductionData(ProductionColumn.StartTime, r) And tmpTime < rawProficyProductionData(ProductionColumn.StartTime, r + 1) Then 'LG Code
                    getProdRowFromTime = r + 1 'LG Code
                    GoTo FoundIt ' LG Code
                End If
            ElseIf r = rawProficyProductionData.GetLength(1) - 1 And tmpTime >= rawProficyProductionData(ProductionColumn.StartTime, r) Then ' LG Code
                getProdRowFromTime = r 'LG Code
                GoTo FoundIt ' LG Code
            ElseIf r = 0 Then
                If tmpTime < rawProficyProductionData(ProductionColumn.StartTime, r) Then 'LG Code
                    getProdRowFromTime = r 'LG Code
                    GoTo FoundIt 'LG Code
                End If
            End If
        Next r
FoundIt:
        If getProdRowFromTime = -1 And isEndTime Then
            If DateDiff(DateInterval.Second, tmpTime, rawProficyProductionData(ProductionColumn.EndTime, rawProficyProductionData.GetLength(1) - 1)) = 0 Then
                getProdRowFromTime = rawProficyProductionData.GetLength(1) - 1
            Else 'LG code
                getProdRowFromTime = rawProficyProductionData.GetLength(1) - 1 'LG code
            End If 'LG code
            If tmpTime = rawProficyProductionData(ProductionColumn.EndTime, rawProficyProductionData.GetLength(1) - 1) Then getProdRowFromTime = rawProficyProductionData.GetLength(1) - 1
        ElseIf isEndTime Then
            getProdRowFromTime = getProdRowFromTime + 0 ' - 1 'account for the fact that we're only searching for start times ' lg code
        End If

    End Function

    Public Function getNearestProdRowFromTime(tmpTime As Date) As Integer
        Dim r As Integer, testEndTime As Date ', i As Integer, j As Double
        For r = 0 To _rawProficyProductionData.GetLength(1) - 1
            ' testStartTime = rawProficyProductionData(ProductionColumn.StartTime, r)
            testEndTime = _rawProficyProductionData(ProductionColumn.EndTime, r)
            If tmpTime < testEndTime Then
                Return r
            ElseIf DateDiff(DateInterval.Second, tmpTime, testEndTime) = 0 Then
                Return r
            ElseIf tmpTime > _rawProficyProductionData(ProductionColumn.EndTime, _rawProficyProductionData.GetLength(1) - 1) Then
                Return _rawProficyProductionData.GetLength(1) - 1
            End If
        Next r
        Return -1
    End Function

    Public Function getNearestDTRowFromTime(tmpTime As Date) As Integer
        Dim r As Integer, testEndTime As Date ', i As Integer, j As Double
        For r = 0 To _rawProficyData.GetLength(1) - 1

            testEndTime = _rawProficyData(1, r)


            Try

                If tmpTime < testEndTime Then
                    Return r
                ElseIf DateDiff(DateInterval.Second, tmpTime, testEndTime) = 0 Then
                    Return r
                ElseIf tmpTime > _rawProficyData(1, _rawProficyData.GetLength(1) - 1) Then
                    Return _rawProficyData.GetLength(1) - 1
                End If

            Catch e As Exception
                Return _rawProficyData.GetLength(1) - 1
            End Try
        Next r
        Return -1
    End Function


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


    Public Function isProductionExcluded(r As Integer, Optional isMaple As Boolean = False) As Boolean
        If isMaple Then
            If IsDBNull(_rawProficyProductionData(ProductionColumn_Maple_New.LineStatus, r)) Then Return True
            If Left(_rawProficyProductionData(ProductionColumn_Maple_New.LineStatus, r), 4) = "PR O" Then Return True
            Return False
        Else
            If IsDBNull(_rawProficyProductionData(ProductionColumn.ProductionStatus, r)) Then Return True
            If Left(_rawProficyProductionData(ProductionColumn.ProductionStatus, r), 4) = "PR O" Then Return True
            Return False
        End If

    End Function

#Region "Sortable & Equitable"
    'THIS WHOLE REGION NEEDS TO BE ADDED BACK 
    'implementation of ISEQUITABLE
    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As ProdLine = TryCast(obj, ProdLine)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As ProdLine) As Boolean _
        Implements IEquatable(Of ProdLine).Equals
        If other Is Nothing Then
            Return False
        End If
        If Me.parentSite.ThreeLetterID.Equals(other.parentSite.ThreeLetterID) Then
            If Me.Name.Equals(other.Name) Then
                Return True
            End If
        End If
        Return (Me.Name.Equals(other.Name) And Me.SiteName.Equals(other.SiteName))
    End Function
#End Region

End Class


#Region "Module / Site / Sector"
Public Class ProdModule
    Protected _moduleName As String
    <JsonIgnore()>
    Public Property parentSite As ProdSite
    <JsonIgnore()>
    Public parentSector As BusinessUnit
    <JsonIgnore()>
    Public LinesList As New List(Of ProdLine)()
    Public prStory_Mapping As Integer

    Public ID As Guid

    Public UnplannedT1List As New List(Of String)
    Public PlannedT1List As New List(Of String)

    Private _SQLprocedure
    Private _SQLprocedureProduction

    Private _defaultDowntimeField As Integer
    Private _defaultDowntimeField_Secondary As Integer

    Private _DTschedPlannedUnplannedMapping As Integer

    Friend _Reason1Name As String
    Friend _Reason2Name As String
    Friend _Reason3Name As String
    Friend _Reason4Name As String

    Public Property MultilineGroup As String = ""

    Public ReadOnly Property MappingLevelA As Integer
        Get
            Return _defaultDowntimeField
        End Get
    End Property
    Public ReadOnly Property MappingLevelB As Integer
        Get
            Return _defaultDowntimeField_Secondary
        End Get
    End Property
    '<JsonIgnore()>
    Public ReadOnly Property Sector
        Get
            Return parentSector.Name
        End Get
    End Property
    Friend ReadOnly Property DTschedMap As Integer
        Get
            Return _DTschedPlannedUnplannedMapping
        End Get
    End Property
    Public ReadOnly Property SQLprocedure As Integer
        Get
            Return _SQLprocedure
        End Get
    End Property
    Public ReadOnly Property SQLprocedurePROD As Integer
        Get
            Return _SQLprocedureProduction
        End Get
    End Property
    Public Overrides Function toString() As String
        Dim tempString
        tempString = ", Lines: " & vbCr
        If LinesList.Count > 0 Then
            For Each productionLine In LinesList
                tempString = tempString & productionLine.Name & vbCr
            Next
        End If
        Return "Module: " & _moduleName & vbCr & tempString
    End Function

    'shift configuration, critical for production data
    Protected _numberOfShifts As Integer = 2
    Public property ShiftDurationHrs As Double = 8
    Public _DayStartTimeHrs As Double

    Protected _FirstShitStartHrs As Double = 0
    Protected _SecondShiftStartHrs As Double = 0
    Protected _ThirdShiftStartHrs As Double = 0

    Public ReadOnly Property ShiftStartFirst_Hr
        Get
            Return _DayStartTimeHrs
        End Get
    End Property
    Public ReadOnly Property ShiftStartSecond_Hr
        Get
            Return _SecondShiftStartHrs
        End Get
    End Property
    Public ReadOnly Property ShiftStartThird_Hr
        Get
            If _numberOfShifts = 3 Then
                Return _ThirdShiftStartHrs
            Else
                'MsgBox("Only " & _numberOfShifts & " Shifts! You asked for 3! Will Return -1. Best of Luck...")
                Return -1
            End If
        End Get
    End Property
    Public ReadOnly Property NumShifts
        Get
            Return _numberOfShifts
        End Get
    End Property

    Public Sub New(ByVal ModuleID As Guid, ByVal moduleName As String, ByVal siteName As String, ByVal BUname As String, ByVal prstoryMappingNo As Integer, ByVal defaultSQLprocedure As Integer, ByVal defaultSQLprocedure_Production As Integer, ByVal R1 As String, ByVal R2 As String, ByVal R3 As String, ByVal R4 As String, ByVal primaryMapping As Integer, ByVal secondaryMapping As Integer, ByVal DTschedMappingConst As Integer)
        _moduleName = moduleName
        ID = ModuleID

        _SQLprocedure = defaultSQLprocedure
        _SQLprocedureProduction = defaultSQLprocedure_Production

        parentSite = AllProductionSites(getSiteIndexFromName(siteName))
        prStory_Mapping = prstoryMappingNo
        AllProductionSites(getSiteIndexFromName(siteName)).ModulesList.Add(Me)

        AllProductionSectors(getBUIndexFromName(BUname)).ModuleList.Add(Me)
        parentSector = AllProductionSectors(getBUIndexFromName(BUname))
        _Reason1Name = R1
        _Reason2Name = R2
        _Reason3Name = R3
        _Reason4Name = R4

        _DTschedPlannedUnplannedMapping = DTschedMappingConst

        _defaultDowntimeField = primaryMapping
        _defaultDowntimeField_Secondary = secondaryMapping
    End Sub

    ' ByVal numShifts As Integer, ByVal shiftLenHrs As Double, ByVal DailyStartTimeHr As Double,
    Public Sub New(ByVal ModuleID As Guid, ByVal moduleName As String, ByVal siteName As String, ByVal BUname As String, ByVal prstoryMappingNo As Integer, ByVal defaultSQLprocedure As Integer, ByVal defaultSQLprocedure_Production As Integer, ByVal R1 As String, ByVal R2 As String, ByVal R3 As String, ByVal R4 As String, ByVal primaryMapping As Integer, ByVal secondaryMapping As Integer, ByVal DTschedMappingConst As Integer, ByVal numShifts As Integer, ByVal shiftLenHrs As Double, ByVal DailyStartTimeHr As Double)
        _moduleName = moduleName
        ID = ModuleID

        _SQLprocedure = defaultSQLprocedure
        _SQLprocedureProduction = defaultSQLprocedure_Production

        parentSite = AllProductionSites(getSiteIndexFromName(siteName))
        prStory_Mapping = prstoryMappingNo
        AllProductionSites(getSiteIndexFromName(siteName)).ModulesList.Add(Me)

        AllProductionSectors(getBUIndexFromName(BUname)).ModuleList.Add(Me)
        parentSector = AllProductionSectors(getBUIndexFromName(BUname))
        _Reason1Name = R1
        _Reason2Name = R2
        _Reason3Name = R3
        _Reason4Name = R4

        _DTschedPlannedUnplannedMapping = DTschedMappingConst

        _defaultDowntimeField = primaryMapping
        _defaultDowntimeField_Secondary = secondaryMapping

        _numberOfShifts = numShifts
        ShiftDurationHrs = shiftLenHrs
        _DayStartTimeHrs = DailyStartTimeHr

        Select Case numShifts
            Case 2
                _SecondShiftStartHrs = _DayStartTimeHrs + ShiftDurationHrs
            Case 3
                _SecondShiftStartHrs = _DayStartTimeHrs + ShiftDurationHrs
                _ThirdShiftStartHrs = _SecondShiftStartHrs + ShiftDurationHrs
            Case Else
                MsgBox("Error: " & numShifts & " shifts not a recognized schedule format. Will assume 2 shifts", vbCritical, "Shift # Err")
                _SecondShiftStartHrs = _DayStartTimeHrs + ShiftDurationHrs
        End Select
    End Sub



    'properties for protected variables
    Public ReadOnly Property Name As String
        Get
            Return _moduleName
        End Get
    End Property
End Class

Public Class ProdSite
    Implements IEquatable(Of ProdSite)
    Implements IComparable(Of ProdSite)



#Region "Variables & Properties"
    Public ServerDatabase As String = "GBDB"

    Protected _siteName As String
    Protected _ProficyServerAddress As String
    Private _ProficyServerUsername As String
    Private _ProficyServerPassword As String

    Protected _HistorianServerAddress As String
    <JsonIgnore()>
    Public ModulesList As New List(Of ProdModule)()

    Private _ThreeLetterID As String

    Public ReadOnly Property ThreeLetterID As String
        Get
            Return _ThreeLetterID
        End Get
    End Property

    'properties for protected variables
    Public ReadOnly Property Name As String
        Get
            Return _siteName
        End Get
    End Property
    Public ReadOnly Property ProficyServer As String
        Get
            Return _ProficyServerAddress
        End Get
    End Property
    Public ReadOnly Property ProficyServer_Password As String
        Get
            Return _ProficyServerPassword
        End Get
    End Property
    Public ReadOnly Property ProficyServer_Username As String
        Get
            Return _ProficyServerUsername
        End Get
    End Property

    Public ReadOnly Property HistorianServer As String
        Get
            Return _HistorianServerAddress
        End Get
    End Property
#End Region

    Public Overrides Function toString() As String
        Return _siteName
    End Function

    Public Sub New(ByVal siteName As String, profServer As String, HistServer As String, profPassword As String, profUsername As String, newThreeLetterID As String) ', ColloquialName As String) ', Optional languageSelected As Integer = Language.English)
        _siteName = siteName
        _ProficyServerAddress = profServer
        _HistorianServerAddress = HistServer
        _ProficyServerPassword = profPassword
        _ProficyServerUsername = profUsername
        _ThreeLetterID = newThreeLetterID
    End Sub

    Public Sub New(ByVal siteName As String, profServer As String, profPassword As String, profUsername As String, newThreeLetterID As String) ', ColloquialName As String) ', Optional languageSelected As Integer = Language.English)
        _siteName = siteName
        _ProficyServerAddress = profServer
        _HistorianServerAddress = ""
        _ProficyServerPassword = profPassword
        _ProficyServerUsername = profUsername
        _ThreeLetterID = newThreeLetterID
    End Sub



    'implementation of ISEQUITABLE


    Public Function CompareTo(ByVal Other As ProdSite) As Integer Implements System.IComparable(Of ProdSite).CompareTo
        Return Me.Name.CompareTo(Other.Name)
    End Function


    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As ProdSite = TryCast(obj, ProdSite)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As ProdSite) As Boolean Implements IEquatable(Of ProdSite).Equals
        If other Is Nothing Then
            Return False
        End If
        Return (Me.Name.Equals(other.Name))
    End Function
End Class

Public Class BusinessUnit
    Implements IEquatable(Of BusinessUnit)
    Protected _BUname As String
    'Public LinesList As List(Of productionLine)()
    <JsonIgnore()>
    Public ModuleList As New List(Of ProdModule)

    Public Overrides Function ToString() As String
        Return _BUname
    End Function
    'CONSTRUCTOR 
    Public Sub New(ByVal BUname As String)
        _BUname = BUname
    End Sub

    'properties for protected variables
    Public ReadOnly Property Name As String
        Get
            Return _BUname
        End Get
    End Property
    Public Function isSectorAtSite(siteName As String)
        For i As Integer = 0 To ModuleList.Count - 1
            If siteName.Equals(ModuleList(i).parentSite) Then Return True
        Next
        Return False
    End Function


    'implementation of ISEQUITABLE



    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As BusinessUnit = TryCast(obj, BusinessUnit)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As BusinessUnit) As Boolean _
        Implements IEquatable(Of BusinessUnit).Equals
        If other Is Nothing Then
            Return False
        End If
        Return (Me.Name.Equals(other.Name))
    End Function
End Class
#End Region
