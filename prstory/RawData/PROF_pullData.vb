Imports System.Threading

Module PROF_pullData
    Public ActiveProf_ConnectionError As Boolean
    'Partial Data Arrays
    Dim RawData_A(,) As Object
    Dim RawData_B(,) As Object
    Dim RawData_C(,) As Object
    Dim RawData_D(,) As Object
    Dim RawData_E(,) As Object
    Dim RawData_F(,) As Object

    Dim RawProd_A(,) As Object
    Dim RawProd_B(,) As Object
    Dim RawProd_C(,) As Object
    Dim RawProd_D(,) As Object
    Dim RawProd_E(,) As Object
    Dim RawProd_F(,) As Object

    Private Const ParentLineArrayIndex As Integer = 0
    Private Const StartTimeArrayIndex As Integer = 1
    Private Const EndTimeArrayIndex As Integer = 2

    ' Dim tmpProdArray As Array
    'data for rate loss
    ' Dim finalRateLossData(,) As Object


    Public Function pullMaxDataForPRODline(ByVal parentLineIndex As Integer, Optional ByVal timePeriodInDays As Integer = 30) As Boolean
        Dim i As Integer, rawDTdataColumns As Integer, rawProdDataColumns As Integer
        Dim _startTime As Date
        Dim _endTime As Date, timeNow As Date
        Dim lineToAnalyze As ProdLine
        Dim netEvents As Long, netProdEvents As Long

        Dim getDTdataThread As New Thread(AddressOf getRawData_A)
        Dim getDTdataThread2 As New Thread(AddressOf getRawData_B)
        Dim getDTdataThread3 As New Thread(AddressOf getRawData_C)
        Dim getDTdataThread4 As New Thread(AddressOf getRawData_D)
        Dim getDTdataThread5 As New Thread(AddressOf getRawData_E)
        Dim getDTdataThread6 As New Thread(AddressOf getRawData_F)

        ActiveProf_ConnectionError = False

        Dim paramObj_One(3) As Object
        Dim paramObj_Two(3) As Object
        Dim paramObj_Three(3) As Object
        Dim paramObj_Four(3) As Object
        Dim paramObj_Five(3) As Object
        Dim paramObj_Six(3) As Object

        timeNow = Now()
        _endTime = DateAdd("s", -Second(timeNow), timeNow)  'current time w 0 seconds

        lineToAnalyze = AllProdLines(parentLineIndex)
       ' lineToAnalyze.verifySystemSettings()

        _startTime = DateAdd(DateInterval.Day, -timePeriodInDays, _endTime)

        paramObj_One(ParentLineArrayIndex) = parentLineIndex
        paramObj_Two(ParentLineArrayIndex) = parentLineIndex
        paramObj_Three(ParentLineArrayIndex) = parentLineIndex
        paramObj_Four(ParentLineArrayIndex) = parentLineIndex
        paramObj_Five(ParentLineArrayIndex) = parentLineIndex
        paramObj_Six(ParentLineArrayIndex) = parentLineIndex

        paramObj_One(StartTimeArrayIndex) = DateAdd(DateInterval.Day, -5 * timePeriodInDays, _startTime)
        paramObj_Two(StartTimeArrayIndex) = DateAdd(DateInterval.Day, -4 * timePeriodInDays, _startTime)
        paramObj_Three(StartTimeArrayIndex) = DateAdd(DateInterval.Day, -3 * timePeriodInDays, _startTime)
        paramObj_Four(StartTimeArrayIndex) = DateAdd(DateInterval.Day, -2 * timePeriodInDays, _startTime)
        paramObj_Five(StartTimeArrayIndex) = DateAdd(DateInterval.Day, -timePeriodInDays, _startTime)
        paramObj_Six(StartTimeArrayIndex) = _startTime

        paramObj_One(EndTimeArrayIndex) = DateAdd(DateInterval.Day, -5 * timePeriodInDays, _endTime)
        paramObj_Two(EndTimeArrayIndex) = DateAdd(DateInterval.Day, -4 * timePeriodInDays, _endTime)
        paramObj_Three(EndTimeArrayIndex) = DateAdd(DateInterval.Day, -3 * timePeriodInDays, _endTime)
        paramObj_Four(EndTimeArrayIndex) = DateAdd(DateInterval.Day, -2 * timePeriodInDays, _endTime)
        paramObj_Five(EndTimeArrayIndex) = DateAdd(DateInterval.Day, -timePeriodInDays, _endTime)
        paramObj_Six(EndTimeArrayIndex) = _endTime

        getDTdataThread.Start(paramObj_One)
        getDTdataThread2.Start(paramObj_Two)
        getDTdataThread3.Start(paramObj_Three)
        getDTdataThread4.Start(paramObj_Four)
        getDTdataThread5.Start(paramObj_Five)
        getDTdataThread6.Start(paramObj_Six)

        While IsNothing(RawData_A) And Not ActiveProf_ConnectionError
            System.Threading.Thread.Sleep(500)
        End While

        While (IsNothing(RawData_C) Or IsNothing(RawData_B)) And Not ActiveProf_ConnectionError
            System.Threading.Thread.Sleep(500)
        End While
        While (IsNothing(RawData_D) Or IsNothing(RawData_E) Or IsNothing(RawData_F)) And Not ActiveProf_ConnectionError
            System.Threading.Thread.Sleep(500)
        End While

        If Not ActiveProf_ConnectionError Then

            'figure out how many columns our raw data has
            Select Case lineToAnalyze.SQLdowntimeProcedure
                Case DefaultProficyDowntimeProcedure.OneClick
                    rawDTdataColumns = 20 'DownTimeColumn_OneClick.Max '20
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    rawDTdataColumns = 29 'DownTimeColumn.Max '29
                Case DefaultProficyDowntimeProcedure.RE_CentralServer
                    rawDTdataColumns = 20 'DownTimeColumn.Max '29
            End Select
            rawProdDataColumns = 17

            'copy into a single array
            netEvents = RawData_A.GetLength(1) + RawData_B.GetLength(1) + RawData_C.GetLength(1) + RawData_D.GetLength(1) + RawData_E.GetLength(1) + RawData_F.GetLength(1)
            Dim completeDTarray(rawDTdataColumns, netEvents - 1) As Object
            For i = 0 To rawDTdataColumns '29
                System.Array.Copy(RawData_A, i * (RawData_A.GetLength(1)), completeDTarray, i * netEvents, RawData_A.GetLength(1))
                System.Array.Copy(RawData_B, i * (RawData_B.GetLength(1)), completeDTarray, i * netEvents + RawData_A.GetLength(1), RawData_B.GetLength(1))
                System.Array.Copy(RawData_C, i * (RawData_C.GetLength(1)), completeDTarray, i * netEvents + RawData_A.GetLength(1) + RawData_B.GetLength(1), RawData_C.GetLength(1))
                System.Array.Copy(RawData_D, i * (RawData_D.GetLength(1)), completeDTarray, i * netEvents + RawData_A.GetLength(1) + RawData_B.GetLength(1) + RawData_C.GetLength(1), RawData_D.GetLength(1))
                System.Array.Copy(RawData_E, i * (RawData_E.GetLength(1)), completeDTarray, i * netEvents + RawData_A.GetLength(1) + RawData_B.GetLength(1) + RawData_C.GetLength(1) + RawData_D.GetLength(1), RawData_E.GetLength(1))
                System.Array.Copy(RawData_F, i * (RawData_F.GetLength(1)), completeDTarray, i * netEvents + RawData_A.GetLength(1) + RawData_B.GetLength(1) + RawData_C.GetLength(1) + RawData_D.GetLength(1) + RawData_E.GetLength(1), RawData_F.GetLength(1))
            Next

            'CHECK FOR DUAL CONSTRAINT
            '   Dim rateLossData(,) As Object
            '   If lineToAnalyze._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode <> MultiConstraintAnalysis.SingleConstraint Then
            'rateLossData = getRawProficyData(_endTime, DateAdd(DateInterval.Day, -90, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
            '   finalRateLossData = PROF_mergeRateLossWithMain(rateLossData, completeDTarray)
            'End If
            ''''''''''''''''''''''''
            While IsNothing(RawProd_A) And Not My.Settings.AdvancedSettings_isAvailabilityMode
                System.Threading.Thread.Sleep(500)
            End While
            While (IsNothing(RawProd_B) Or IsNothing(RawProd_C) Or IsNothing(RawProd_D)) And Not My.Settings.AdvancedSettings_isAvailabilityMode
                System.Threading.Thread.Sleep(500)
            End While
            While (IsNothing(RawProd_E) Or IsNothing(RawProd_F)) And Not My.Settings.AdvancedSettings_isAvailabilityMode
                System.Threading.Thread.Sleep(500)
            End While


            If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
                netProdEvents = RawProd_A.GetLength(1) + RawProd_B.GetLength(1) + RawProd_C.GetLength(1) + RawProd_D.GetLength(1) + RawProd_E.GetLength(1) + RawProd_F.GetLength(1)

                Dim completeProdArray(rawProdDataColumns, netProdEvents - 1) As Object

                For i = 0 To rawProdDataColumns
                    System.Array.Copy(RawProd_A, i * (RawProd_A.GetLength(1)), completeProdArray, i * netProdEvents, RawProd_A.GetLength(1))
                    System.Array.Copy(RawProd_B, i * (RawProd_B.GetLength(1)), completeProdArray, i * netProdEvents + RawProd_A.GetLength(1), RawProd_B.GetLength(1))
                    System.Array.Copy(RawProd_C, i * (RawProd_C.GetLength(1)), completeProdArray, i * netProdEvents + RawProd_A.GetLength(1) + RawProd_B.GetLength(1), RawProd_C.GetLength(1))
                    System.Array.Copy(RawProd_D, i * (RawProd_D.GetLength(1)), completeProdArray, i * netProdEvents + RawProd_A.GetLength(1) + RawProd_B.GetLength(1) + RawProd_C.GetLength(1), RawProd_D.GetLength(1))
                    System.Array.Copy(RawProd_E, i * (RawProd_E.GetLength(1)), completeProdArray, i * netProdEvents + RawProd_A.GetLength(1) + RawProd_B.GetLength(1) + RawProd_C.GetLength(1) + RawProd_D.GetLength(1), RawProd_E.GetLength(1))
                    System.Array.Copy(RawProd_F, i * (RawProd_F.GetLength(1)), completeProdArray, i * netProdEvents + RawProd_A.GetLength(1) + RawProd_B.GetLength(1) + RawProd_C.GetLength(1) + RawProd_D.GetLength(1) + RawProd_E.GetLength(1), RawProd_F.GetLength(1))
                Next
                ' End If

                With AllProdLines(parentLineIndex) 'lineToAnalyze
                    ' If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
                    .rawProficyProductionData = completeProdArray
                    .rawProductionData = New ProductionDataset(AllProdLines(parentLineIndex), True)
                End With
            End If
            With AllProdLines(parentLineIndex)
                .rawProfStartTime = DateAdd(DateInterval.Day, -(6 * timePeriodInDays), _endTime)
                .rawProfEndTime = _endTime

                '     If ._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode <> MultiConstraintAnalysis.SingleConstraint Then
                ' .rawProficyData = finalRateLossData
                ' .rawRateLossData = rateLossData

                '                If Not My.Settings.AdvancedSettings_isAvailabilityMode Then PROF_setDTteamsFromProd(parentLineIndex)
                '                .rawDowntimeData = New prstoryData(AllProductionLines(parentLineIndex), finalRateLossData)
                '                Else
                .rawProficyData = completeDTarray
                If Not My.Settings.AdvancedSettings_isAvailabilityMode Then PROF_setDTteamsFromProd(parentLineIndex)
                .rawDowntimeData = New DowntimeDataset(AllProdLines(parentLineIndex), completeDTarray)
                '               End If


            End With

            RawData_A = Nothing
            RawData_B = Nothing
            RawData_C = Nothing
            RawData_D = Nothing
            RawData_E = Nothing
            RawData_F = Nothing

            RawProd_A = Nothing
            RawProd_B = Nothing
            RawProd_C = Nothing
            RawProd_D = Nothing
            RawProd_E = Nothing
            RawProd_F = Nothing



        Else 'this means there was a proficy connection error!
            'MsgBox("oops! " & AllProductionLines(parentLineIndex).ToString)
        End If

        Return ActiveProf_ConnectionError
    End Function

#Region "Data Thread Fcns"
    Private Sub getRawData_A(ByVal paramObj As Object)
        RawData_A = BeerMe_DT(paramObj)
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then RawProd_A = BeerMe_Prod(paramObj)
    End Sub

    Private Sub getRawData_B(ByVal paramObj As Object)
        RawData_B = BeerMe_DT(paramObj)
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then RawProd_B = BeerMe_Prod(paramObj)
    End Sub

    Private Sub getRawData_C(ByVal paramObj As Object)
        RawData_C = BeerMe_DT(paramObj)
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then RawProd_C = BeerMe_Prod(paramObj)
    End Sub

    Private Sub getRawData_D(ByVal paramObj As Object)
        RawData_D = BeerMe_DT(paramObj)
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then RawProd_D = BeerMe_Prod(paramObj)
    End Sub

    Private Sub getRawData_E(ByVal paramObj As Object)
        RawData_E = BeerMe_DT(paramObj)
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then RawProd_E = BeerMe_Prod(paramObj)
    End Sub

    Private Sub getRawData_F(ByVal paramObj As Object)
        RawData_F = BeerMe_DT(paramObj)
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then RawProd_F = BeerMe_Prod(paramObj)
    End Sub
#End Region

#Region "Extracting Data"
    Private Function BeerMe_DT(ByVal paramObj As Object) As Object(,) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As String

        lineIndex = paramObj(ParentLineArrayIndex)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnit
        End With
        STARTx = paramObj(StartTimeArrayIndex)
        ENDx = paramObj(EndTimeArrayIndex)

        Try
            Select Case preferredDtQuery
                Case DefaultProficyDowntimeProcedure.OneClick
                    Return getRawProficyData_OneClick(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    Return getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
            End Select
        Catch ex As System.Runtime.InteropServices.COMException
            ' Debugger.Break()
            If Not ActiveProf_ConnectionError Then
                ActiveProf_ConnectionError = True

            End If
        Catch ex As Exception
            If Not ActiveProf_ConnectionError Then
                ActiveProf_ConnectionError = True

            End If
        End Try
        '  Return Nothing 'this is bad
    End Function
    Private Function BeerMe_Prod(ByVal paramObj As Object) As Object(,) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date)
        Dim STARTx As Date
        Dim ENDx As Date

        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredProdQuery As Integer, prodUnit As String, databaseName As String

        lineIndex = paramObj(ParentLineArrayIndex)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredProdQuery = .SQLproductionProcedure
            prodUnit = .mainProfProd
            databaseName = .ServerDatabase
        End With

        STARTx = paramObj(StartTimeArrayIndex)
        ENDx = paramObj(EndTimeArrayIndex)

        Try
            Select Case preferredProdQuery
                Case DefaultProficyProductionProcedure.QuickQuery
                    Return getRawProficyProductionData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0)) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0))
  Case DefaultProficyProductionProcedure.Maple
                    Return getMaplePRODData(ENDx, STARTx, prodUnit, prodUnit, serverName, serverUsername, serverPassword, databaseName) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0)) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0))

                Case DefaultProficyProductionProcedure.SwingRoad
                    Return getRawProficyProductionData_SwingRoad(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
            End Select
        Catch ex As System.Runtime.InteropServices.COMException
            If Not ActiveProf_ConnectionError Then
                ActiveProf_ConnectionError = True

            End If
        Catch ex As Exception
            If Not ActiveProf_ConnectionError Then
                ActiveProf_ConnectionError = True

            End If
        End Try
        'Return Nothing 'this is bad
    End Function
#End Region
End Module
