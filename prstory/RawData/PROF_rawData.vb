
Module PROF_rawData
    Public Function getRawProficyData(ByVal endTime As Date, ByVal startTime As Date, ByVal masterProdUnit As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String) As Array
        Dim poCMD As New ADODB.Command
        Dim psConnString As String
        'Check to make sure that we have a database connection before attempting to execute the stored procedure
        Dim gobjConnA = New ADODB.Connection
        If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
            psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=GBDB; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"
            If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
                gobjConnA.ConnectionString = psConnString
                gobjConnA.CommandTimeout = 30
                gobjConnA.Open()
            End If
        End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnA
            .CommandText = "spLocal_PQQ_DowntimeExplorer"
            .Parameters.Refresh()
            .Parameters(1).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(3).Value = masterProdUnit
            .Parameters(4).Value = "All"
            .Parameters(5).Value = "All"
            .Parameters(6).Value = "All"
            .Parameters(7).Value = "All"
            .Parameters(8).Value = "All"
            .Parameters(9).Value = "All"
            .Parameters(10).Value = "All"
            .Parameters(11).Value = "All"
            .Parameters(12).Value = "All"
            .Parameters(13).Value = "All"
            .Parameters(14).Value = "All"
            .Parameters(15).Value = "All"
            .Parameters(16).Value = "Nothing"
            .Parameters(17).Value = "Downtime"
            .Parameters(18).Value = 0
            .Parameters(19).Value = ">"
            .Parameters(20).Value = "Raw Data"
            .Parameters(21).Value = 1440
            .Parameters(22).Value = "All"
            .Parameters(23).Value = "All"
            .Parameters(24).Value = "All"
            .Parameters(25).Value = "All"
            .Parameters(26).Value = 1 'UPTIME
            .Parameters(27).Value = 0

            Try
                getRawProficyData = .Execute.GetRows() 'ORIGINAL CODE
            Catch ex As Exception
                Throw New Exception(ex.Message & "   [QQ]")
            End Try
        End With
        If gobjConnA.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.Close()
        End If
        gobjConnA = Nothing
    End Function
    Public Function getRawProficyData_OneClick(ByVal endTime As Date, ByVal startTime As Date, ByVal masterProdUnit As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String) As Array
        Dim poCMD As New ADODB.Command
        Dim RS As New ADODB.Recordset
        Dim psConnString As String
        'Check to make sure that we have a database connection before attempting to execute the stored procedure
        Dim gobjConnA = New ADODB.Connection
        ' If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
        psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=GBDB; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"

            If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
                gobjConnA.ConnectionString = psConnString
                gobjConnA.CommandTimeout = 30
                gobjConnA.Open()
            End If
            ' End If

            With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnA
            .CommandText = SQL_PROCEDURE_ONECLICK
            .Parameters.Refresh()
            .Parameters(1).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(3).Value = masterProdUnit
            .Parameters(4).Value = "0" '0
            .Parameters(5).Value = "1"
            .Parameters(6).Value = "1"
            .Parameters(7).Value = "1"
            .Parameters(8).Value = "1"
            .Parameters(9).Value = "1"
            .Parameters(10).Value = "1"
            .Parameters(11).Value = "1"
            .Parameters(12).Value = "1"
            .Parameters(13).Value = "1"
            .Parameters(14).Value = "1"
            .Parameters(15).Value = "1"
            .Parameters(16).Value = "0" '0

            Try
                RS = .Execute()
                getRawProficyData_OneClick = RS.GetRows() 'ORIGINAL CODe
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End With
        If gobjConnA.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.Close()
        End If
        gobjConnA = Nothing
    End Function


    Public Function getRawProficyData_OneClick_MultiUnit(ByVal endTime As Date, ByVal startTime As Date, ByVal masterProdUnits As List(Of String), ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String) As Array
        Dim poCMD As New ADODB.Command
        Dim RS As New ADODB.Recordset
        Dim psConnString As String
        'Check to make sure that we have a database connection before attempting to execute the stored procedure
        Dim gobjConnA = New ADODB.Connection
        ' If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
        psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=GBDB; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"

        If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.ConnectionString = psConnString
            gobjConnA.CommandTimeout = 30
            gobjConnA.Open()
        End If
        ' End If

        'now find the master prod unit string
        Dim masterProdUnit As String
        If (masterProdUnits.Count) < 1 Then
            Throw New Exception("No Prod Units Selected!")
        ElseIf (masterProdUnits.count = 1) Then
            masterProdUnit = masterProdUnits(0)
        Else
            masterProdUnit = masterProdUnits(0)
            For i As Integer = 1 To masterProdUnits.Count - 1
                masterProdUnit = masterProdUnit & "," & masterProdUnits(i)
            Next
        End If

        'execute stored procedure
        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnA
            .CommandText = SQL_PROCEDURE_ONECLICK
            .Parameters.Refresh()
            .Parameters(1).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(3).Value = masterProdUnit
            .Parameters(4).Value = "1" '0
            .Parameters(5).Value = "1"
            .Parameters(6).Value = "1"
            .Parameters(7).Value = "1"
            .Parameters(8).Value = "1"
            .Parameters(9).Value = "1"
            .Parameters(10).Value = "1"
            .Parameters(11).Value = "1"
            .Parameters(12).Value = "1"
            .Parameters(13).Value = "1"
            .Parameters(14).Value = "1"
            .Parameters(15).Value = "1"
            .Parameters(16).Value = "1" '0

            Try
                RS = .Execute()
                getRawProficyData_OneClick_MultiUnit = RS.GetRows()
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End With
        If gobjConnA.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.Close()
        End If
        gobjConnA = Nothing
    End Function




    Public Function getRawProficyData_OneClick_v27(ByVal endTime As Date, ByVal startTime As Date, ByVal masterProdUnit As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String) As Array
        Dim poCMD As New ADODB.Command
        Dim RS As New ADODB.Recordset
        Dim psConnString As String
        'Check to make sure that we have a database connection before attempting to execute the stored procedure
        Dim gobjConnA = New ADODB.Connection
        ' If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
        psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=GBDB; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"

        If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.ConnectionString = psConnString
            gobjConnA.CommandTimeout = 30
            gobjConnA.Open()
        End If
        ' End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnA
            .CommandText = SQL_PROCEDURE_ONECLICK_FAMILY
            .Parameters.Refresh()
            .Parameters(1).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(3).Value = masterProdUnit
            .Parameters(4).Value = "0"
            .Parameters(5).Value = "1"
            .Parameters(6).Value = "1"
            .Parameters(7).Value = "1"
            .Parameters(8).Value = "1"
            .Parameters(9).Value = "1"
            .Parameters(10).Value = "1"
            .Parameters(11).Value = "1"
            .Parameters(12).Value = "1"
            .Parameters(13).Value = "1"
            .Parameters(14).Value = "1"
            .Parameters(15).Value = "1"
            .Parameters(16).Value = "0"
            .Parameters(17).Value = "1"
            .Parameters(18).Value = "1"

            Try
                RS = .Execute()


                If RS.EOF <> True And RS.BOF <> True Then
                    getRawProficyData_OneClick_v27 = RS.GetRows() 'ORIGINAL CODE

                Else
                    Throw New Exception("SP27 Error. PROF.RawData")
                End If

            Catch ex As Exception
                Throw
            End Try

        End With
        If gobjConnA.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.Close()
        End If
        gobjConnA = Nothing
    End Function

End Module

Module PROF_CustomSQL

    Public Function getRawV6data() As Array
        Dim poCMD As New ADODB.Command
        Dim psConnString As String
        Dim gobjConnA = New ADODB.Connection
        If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
            psConnString = "Driver={SQL Server}; server=marp-mesdatabe; database=GBDB; UID=" & "MarpLocalTools" & "; PWD=" & "Proftools" & "; network=dbmssocn"

            If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
                gobjConnA.ConnectionString = psConnString
                gobjConnA.CommandTimeout = 30
                gobjConnA.Open()
            End If
        End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnA
            .CommandText = initializeSQLstring()

            '  .Parameters.Refresh()

            Try
                getRawV6data = .Execute.GetRows()
            Catch ex As Exception
                Throw
            End Try
        End With
        If gobjConnA.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.Close()
        End If
        gobjConnA = Nothing
    End Function



End Module


Module ProficyProductionPull

    Function getRawProficyProductionData(ByVal endTime As Date, ByVal startTime As Date, ByVal masterProdUnit As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String) As Array
        Dim tmpRawData As Array ', poIncrementer As Integer, tmpRate As Integer
        Dim psConnString As String
        Dim poCMD As New ADODB.Command
        Dim gobjConnB = New ADODB.Connection
        If gobjConnB.State <> ADODB.ObjectStateEnum.adStateOpen Then
            psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=GBDB; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"

            If psServer = "?" Or psUID = "?" Or psPassword = "?" Then
                'frmOptions.Show vbModal
            Else
                If gobjConnB.State <> ADODB.ObjectStateEnum.adStateOpen Then
                    gobjConnB.ConnectionString = psConnString
                    gobjConnB.CommandTimeout = 30
                    gobjConnB.Open()

                End If
            End If
        End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnB
            .CommandText = "spLocal_PQQ_ProductionRaw"
            .Parameters.Refresh()
            .Parameters(1).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(3).Value = masterProdUnit
            .Parameters(4).Value = "All"
            .Parameters(5).Value = "All"

            Try
                tmpRawData = .Execute.GetRows()
            Catch ex As Exception
                Throw
            End Try

        End With

        Return tmpRawData
        If gobjConnB.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnB.Close()
        End If
        gobjConnB = Nothing
    End Function

    Function getRawProficyProductionData_SwingRoad(ByVal endTime As Date, ByVal startTime As Date, ByVal masterProdUnit As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String) As Array
        Dim tmpRawData As Array ', poIncrementer As Integer, tmpRate As Integer
        Dim psConnString As String
        Dim poCMD As New ADODB.Command
        Dim gobjConnB = New ADODB.Connection
        If gobjConnB.State <> ADODB.ObjectStateEnum.adStateOpen Then
            psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=GBDB; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"

            If psServer = "?" Or psUID = "?" Or psPassword = "?" Then
                'frmOptions.Show vbModal
            Else
                If gobjConnB.State <> ADODB.ObjectStateEnum.adStateOpen Then
                    gobjConnB.ConnectionString = psConnString
                    gobjConnB.CommandTimeout = 30
                    gobjConnB.Open()

                End If
            End If
        End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnB
            .CommandText = "spLocal_GetResultsByTime"
            .Parameters.Refresh()
            .Parameters(1).Value = masterProdUnit ' "Line 5 Production"
            .Parameters(2).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(3).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(4).Value = "1"

            Try
                tmpRawData = .Execute.GetRows()
            Catch ex As Exception
                Throw
            End Try

        End With

        Return tmpRawData
        If gobjConnB.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnB.Close()
        End If
        gobjConnB = Nothing
    End Function


    'MOT Production through QQ SP

End Module



Module PROF_HelperFunctions
    Public Sub PROF_setDTteamsFromProd(targetLineIndex As Integer)
        Dim ProdCol As Long, tmpTeam As String
        With AllProdLines(targetLineIndex)
            For i As Long = 0 To .rawProficyData.GetLength(1) - 1
                tmpTeam = ""
                If Not IsDBNull(.rawProficyData(13, i)) Then
                    tmpTeam = .rawProficyData(13, i)
                End If
                If tmpTeam = "" Then
                    ProdCol = .getNearestProdRowFromTime(.rawProficyData(0, i)) 'DownTimeColumn.StartTime, i))
                    If ProdCol <> -1 Then
                        If Not IsDBNull(.rawProficyProductionData(ProductionColumn.Team, ProdCol)) Then
                            '   .rawProficyData(DownTimeColumn.Team, i) = .rawProficyProductionData(ProductionColumn.Team, ProdCol)
                            .rawProficyData(13, i) = .rawProficyProductionData(ProductionColumn.Team, ProdCol)
                        End If
                    Else
                        Debugger.Break()
                    End If
                End If

            Next
        End With
    End Sub

    Public Sub PROF_setDTGroupInProdFromDT(targetLineIndex As Integer)
        Dim ProdCol As Long, tmpTeam As String
        Dim DTgroupCol As Integer = 21
        Dim PRODskuCol As Integer = 3

        With AllProdLines(targetLineIndex)
            For i As Long = 0 To .rawProficyProductionData.GetLength(1) - 1
                tmpTeam = ""

                ProdCol = .getNearestDTRowFromTime(.rawProficyProductionData(0, i))
                If ProdCol <> -1 Then
                    If Not IsDBNull(.rawProficyData(DTgroupCol, ProdCol)) Then
                        '   .rawProficyData(DownTimeColumn.Team, i) = .rawProficyProductionData(ProductionColumn.Team, ProdCol)
                        Dim testString As String = .rawProficyData(DTgroupCol, ProdCol)
                        '  Dim x2 As String = .rawProficyData(DTgroupCol, ProdCol - 1)
                        .rawProficyProductionData(PRODskuCol, i) = .rawProficyData(DTgroupCol, ProdCol)
                        .rawProficyProductionData(PRODskuCol + 1, i) = .rawProficyData(DTgroupCol, ProdCol)
                    End If
                Else
                    Debugger.Break()
                End If
            Next
        End With
    End Sub



End Module

Module SQL_Procedure
    Public Function initializeSQLstring()
        Dim SQL As String
        SQL = "USE [GBDB] " & vbNewLine
        SQL = SQL & "SET ANSI_NULLS OFF " & vbNewLine


        SQL = SQL & " " & vbNewLine
        SQL = SQL & "SET QUOTED_IDENTIFIER OFF " & vbNewLine
        SQL = SQL & "DECLARE  " & vbNewLine
        SQL = SQL & "                           @InputStartTime                   DateTime = '10/29/2015 06:00:00 AM', " & vbNewLine
        SQL = SQL & "                           @Inputendtime              DateTime = '11/1/2015 06:00:00 AM', " & vbNewLine
        SQL = SQL & "                           @InputMasterProdUnit   nVarChar(4000) = 'HCMR011 Main', " & vbNewLine
        SQL = SQL & "                           @ShowMasterProdUnit        int = '0', " & vbNewLine
        SQL = SQL & "                           @showreason3               int = '1', " & vbNewLine
        SQL = SQL & "                           @showreason4               int = '1', " & vbNewLine
        SQL = SQL & "                           @showTeam                  int = '1', " & vbNewLine
        SQL = SQL & "                           @showshift                 int = '1', " & vbNewLine
        SQL = SQL & "                           @showComment               int = '1', " & vbNewLine
        SQL = SQL & "                           @showProduct               int = '1', " & vbNewLine
        SQL = SQL & "                           @showbrandCode                    int = '1', " & vbNewLine
        SQL = SQL & "                           @showProdGroup                    int = '1', " & vbNewLine
        SQL = SQL & "                           @showCat1                  int = '1', " & vbNewLine
        SQL = SQL & "                           @showCat2                  int = '1', " & vbNewLine
        SQL = SQL & "                           @showCat3                  int = '1', " & vbNewLine
        SQL = SQL & "                           @showCat4                  int = '0' " & vbNewLine
        SQL = SQL & " " & vbNewLine


        SQL = SQL & "SET ANSI_WARNINGS OFF " & vbNewLine
        SQL = SQL & "SET NOCOUNT ON " & vbNewLine




        SQL = SQL & "DECLARE       @PositiON                  int, " & vbNewLine
        SQL = SQL & "              @InputORderByClause  nvarChar(4000), " & vbNewLine
        SQL = SQL & "              @InputGroupByClause  nvarChar(4000), " & vbNewLine
        SQL = SQL & "              @strSQL                    VarChar(4000), " & vbNewLine
        SQL = SQL & "              @current             datetime, " & vbNewLine
        SQL = SQL & "              @tmpStartTime as     datetime, " & vbNewLine
        SQL = SQL & "              @tmpendtime as             datetime, " & vbNewLine
        SQL = SQL & "              @tmpCount as         int, " & vbNewLine
        SQL = SQL & "              @tmpLoopCounter      int, " & vbNewLine
        SQL = SQL & "              @RptProdPUId         int, " & vbNewLine
        SQL = SQL & "              @ShowLineStatus            int, " & vbNewLine
        SQL = SQL & "              @CatPrefix           varchar (30), " & vbNewLine
        SQL = SQL & "                @PUID                    int " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "DROP TABLE #Downtime " & vbNewLine
        SQL = SQL & "DROP TABLE #DTPuIDStartTime " & vbNewLine
        SQL = SQL & "CREATE TABLE #DownTime(     " & vbNewLine
        SQL = SQL & "       DT_ID           int IDENTITY (0, 1) NOT NULL , " & vbNewLine
        SQL = SQL & "       StartTime            datetime, " & vbNewLine
        SQL = SQL & "       endtime                    datetime, " & vbNewLine
        SQL = SQL & "       Uptime               Float, " & vbNewLine
        SQL = SQL & "       Downtime             Float, " & vbNewLine
        SQL = SQL & "       MasterProdUnit               varchar(100), " & vbNewLine
        SQL = SQL & "       location             varchar(50), " & vbNewLine
        SQL = SQL & "       split                varchar(50), " & vbNewLine
        SQL = SQL & "       Fault                varchar(100), " & vbNewLine
        SQL = SQL & "       reason1                    varchar(100), " & vbNewLine
        SQL = SQL & "       reason2                    varchar(100), " & vbNewLine
        SQL = SQL & "       reason3                    varchar(100), " & vbNewLine
        SQL = SQL & "       reason4                    varchar(100), " & vbNewLine
        SQL = SQL & "       Team                 varchar(25), " & vbNewLine
        SQL = SQL & "       shift                varchar(25),  " & vbNewLine
        SQL = SQL & "       ishift               INT,   " & vbNewLine
        SQL = SQL & "     Cat1                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat2                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat3                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat4                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat5                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat6                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat7                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat8                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat9                 varchar(50), " & vbNewLine
        SQL = SQL & "       Cat10                varchar(50), " & vbNewLine
        SQL = SQL & "       Product                    varchar(100), " & vbNewLine
        SQL = SQL & "       ProductGroup  varchar(100), " & vbNewLine
        SQL = SQL & "       brand                varchar(100),  " & vbNewLine
        SQL = SQL & "       ibrand               INT, " & vbNewLine
        SQL = SQL & "       LineStatus           varchar(100), " & vbNewLine
        SQL = SQL & "       Cause_Comment_ID     INT, " & vbNewLine
        SQL = SQL & "       Comments             varchar(2000), " & vbNewLine
        SQL = SQL & "       StartTime_Act         datetime, " & vbNewLine
        SQL = SQL & "       endtime_Act          datetime, " & vbNewLine
        SQL = SQL & "       endtime_Prev          datetime, " & vbNewLine
        SQL = SQL & "       PUID                 INT, " & vbNewLine
        SQL = SQL & "       SourcePUID           INT, " & vbNewLine
        SQL = SQL & "       tedet_id             INT, " & vbNewLine
        SQL = SQL & "       reasonID1            INT, " & vbNewLine
        SQL = SQL & "       reasonID2            INT, " & vbNewLine
        SQL = SQL & "       reasonID3            INT, " & vbNewLine
        SQL = SQL & "       reasonID4            INT, " & vbNewLine
        SQL = SQL & "       ERTD_ID                    int, " & vbNewLine
        SQL = SQL & "       ErtdID1                    INT, " & vbNewLine
        SQL = SQL & "       ErtdID2                    INT, " & vbNewLine
        SQL = SQL & "       ErtdID3                    INT, " & vbNewLine
        SQL = SQL & "       ErtdID4                    INT, " & vbNewLine
        SQL = SQL & "       SBP                  INT, " & vbNewLine
        SQL = SQL & "       EAP                  INT, " & vbNewLine
        SQL = SQL & "       TargetSpeed          float, " & vbNewLine
        SQL = SQL & "       ActualSpeed          float " & vbNewLine
        SQL = SQL & "       ) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "CREATE TABLE #DTPuIDStartTime( " & vbNewLine
        SQL = SQL & "  RowId int IDENTITY, " & vbNewLine
        SQL = SQL & "  PU_Id int, " & vbNewLine
        SQL = SQL & "  StartTime DateTime  " & vbNewLine
        SQL = SQL & "  PRIMARY KEY(PU_Id, StartTime)) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "     " & vbNewLine
        SQL = SQL & "   " & vbNewLine
        SQL = SQL & "DECLARE @schedule_puid TABLE  ( " & vbNewLine
        SQL = SQL & "       pu_id                int,  " & vbNewLine
        SQL = SQL & "       schedule_puid        int,  " & vbNewLine
        SQL = SQL & "       tmp1                 int, " & vbNewLine
        SQL = SQL & "       tmp2                 int, " & vbNewLine
        SQL = SQL & "       info                 varchar(300)) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "DECLARE @TESTS TABLE  ( " & vbNewLine
        SQL = SQL & "       var_id               int, " & vbNewLine
        SQL = SQL & "       result               varchar(100), " & vbNewLine
        SQL = SQL & "       result_ON            datetime, " & vbNewLine
        SQL = SQL & "       extendedinfo  varchar(255)) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "CREATE INDEX  td_PUId_StartTime " & vbNewLine
        SQL = SQL & "       ON     #DownTime (PUId, StartTime) " & vbNewLine
        SQL = SQL & "CREATE INDEX  td_PUId_endtime " & vbNewLine
        SQL = SQL & "       ON     #DownTime (PUId, endtime) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "DECLARE       @ErrorMessages TABLE ( " & vbNewLine
        SQL = SQL & "       ErrMsg               nVarChar(255) ) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF     ISDate(@InputStartTime) <> 1 " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "       INSERT @ErrorMessages (ErrMsg) " & vbNewLine
        SQL = SQL & "              VALUES ('StartTime IS not a Date.') " & vbNewLine
        SQL = SQL & "       GOTO   ErrorMessagesWrite " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & "IF     ISDate(@Inputendtime) <> 1 " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "       INSERT @ErrorMessages (ErrMsg) " & vbNewLine
        SQL = SQL & "              VALUES ('endtime IS not a Date.') " & vbNewLine
        SQL = SQL & "       GOTO   ErrorMessagesWrite " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showreason3=1 AND @showreason4=1  " & vbNewLine
        SQL = SQL & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2,reason3,reasonID3 " & vbNewLine
        SQL = SQL & ",reason4,reasonID4,startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID) " & vbNewLine
        SQL = SQL & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name, " & vbNewLine
        SQL = SQL & "r2.event_reason_name,r3.event_reason_name,ted.reason_level3,r4.event_reason_name,ted.reason_level4, " & vbNewLine
        SQL = SQL & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID,  " & vbNewLine
        SQL = SQL & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID " & vbNewLine
        SQL = SQL & "FROM dbo.timed_event_details AS ted with (nolock)  " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r3 with (nolock) ON (r3.event_reason_id = ted.reason_level3) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r4 with (nolock) ON (r4.event_reason_id = ted.reason_level4) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id) " & vbNewLine
        SQL = SQL & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id) " & vbNewLine
        SQL = SQL & "inner join dbo.prod_units AS pu2 with (nolock) ON (pu2.pu_id = ted.pu_id) " & vbNewLine
        SQL = SQL & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) ) " & vbNewLine
        SQL = SQL & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All')) " & vbNewLine
        SQL = SQL & "ORDER BY ted.start_Time,  ted.pu_ID " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showreason3=0 AND @showreason4=1 " & vbNewLine
        SQL = SQL & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2, " & vbNewLine
        SQL = SQL & "reason4,reasonID4,startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID) " & vbNewLine
        SQL = SQL & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name, " & vbNewLine
        SQL = SQL & "r2.event_reason_name,r4.event_reason_name,ted.reason_level4, " & vbNewLine
        SQL = SQL & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID,  " & vbNewLine
        SQL = SQL & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID " & vbNewLine
        SQL = SQL & "FROM dbo.timed_event_details AS ted with (nolock)  " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r4 with (nolock) ON (r4.event_reason_id = ted.reason_level4) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id) " & vbNewLine
        SQL = SQL & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id) " & vbNewLine
        SQL = SQL & "inner join dbo.prod_units AS pu2 with (nolock) ON (pu2.pu_id = ted.pu_id) " & vbNewLine
        SQL = SQL & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) ) " & vbNewLine
        SQL = SQL & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All')) " & vbNewLine
        SQL = SQL & "ORDER BY ted.start_Time,  ted.pu_ID " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showreason4=0 AND @showreason3=1 " & vbNewLine
        SQL = SQL & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2, " & vbNewLine
        SQL = SQL & "reason3,reasonID3,startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID) " & vbNewLine
        SQL = SQL & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name, " & vbNewLine
        SQL = SQL & "r2.event_reason_name,r3.event_reason_name,ted.reason_level3, " & vbNewLine
        SQL = SQL & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID,  " & vbNewLine
        SQL = SQL & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID " & vbNewLine
        SQL = SQL & "FROM dbo.timed_event_details AS ted with (nolock)  " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r3 with (nolock) ON (r3.event_reason_id = ted.reason_level3) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id) " & vbNewLine
        SQL = SQL & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id) " & vbNewLine
        SQL = SQL & "inner join dbo.prod_units AS pu2 with (nolock) ON (pu2.pu_id = ted.pu_id) " & vbNewLine
        SQL = SQL & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) ) " & vbNewLine
        SQL = SQL & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All')) " & vbNewLine
        SQL = SQL & "ORDER BY ted.start_Time,  ted.pu_ID " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showreason4=0 AND @showreason3=0 " & vbNewLine
        SQL = SQL & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2, " & vbNewLine
        SQL = SQL & "startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID) " & vbNewLine
        SQL = SQL & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name, " & vbNewLine
        SQL = SQL & "r2.event_reason_name, " & vbNewLine
        SQL = SQL & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID,  " & vbNewLine
        SQL = SQL & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID " & vbNewLine
        SQL = SQL & "FROM dbo.timed_event_details AS ted with (nolock)  " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2) " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id) " & vbNewLine
        SQL = SQL & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id) " & vbNewLine
        SQL = SQL & "inner join dbo.prod_units AS  pu2 with (nolock) ON (pu2.pu_id = ted.pu_id) " & vbNewLine
        SQL = SQL & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) ) " & vbNewLine
        SQL = SQL & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All')) " & vbNewLine
        SQL = SQL & "ORDER BY ted.start_Time,  ted.pu_ID " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcomment=1 " & vbNewLine
        SQL = SQL & "UPDATE ted SET " & vbNewLine
        SQL = SQL & "       comments = REPLACE(coalesce(convert(varchar(2000),co.comment_text),''), char(13)+char(10), ' ') " & vbNewLine
        SQL = SQL & "FROM dbo.#Downtime ted with (nolock) " & vbNewLine
        SQL = SQL & "left join dbo.Comments co with (nolock) " & vbNewLine
        SQL = SQL & "ON ted.cause_comment_id = co.comment_id " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET Downtime = datediff(s,starttime,endtime)/60.0 " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET Downtime = datediff(s,starttime,@Inputendtime)/60.0 " & vbNewLine
        SQL = SQL & "       WHERE endtime IS NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showteam=1 OR @showshift=1 " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       INSERT INTO @schedule_puid (pu_id,info) SELECT pu_id,extended_info FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "       WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All') " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE @schedule_puid SET tmp1=charindex('scheduleunit=',info) " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE @schedule_puid SET tmp2=charindex(';',info,tmp1) WHERE tmp1>0 " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE @schedule_puid SET schedule_puid=cast(substring(info,tmp1+13,tmp2-tmp1-13) as int) WHERE tmp1>0 AND tmp2>0 AND not tmp2 IS NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE @schedule_puid SET schedule_puid=cast(substring(info,tmp1+13,len(info)-tmp1-12) as int)WHERE tmp1>0 AND tmp2=0 " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "    IF NOT EXISTS(SELECT schedule_puid FROM @schedule_puid) " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "         UPDATE @schedule_puid SET schedule_puid=(SELECT   TOP 1 Table_Fields_Values.Value " & vbNewLine
        SQL = SQL & "         FROM Table_Fields_Values with (nolock) INNER JOIN  Table_Fields ON Table_Fields_Values.Table_Field_Id = " & vbNewLine
        SQL = SQL & "            Table_Fields.Table_Field_Id  WHERE(Table_Fields.Table_Field_Desc = 'ScheduleUnit')  " & vbNewLine
        SQL = SQL & "         AND (Table_Fields_Values.KeyId in (SELECT  pu_id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "        WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "        WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All'))))) " & vbNewLine
        SQL = SQL & "      END " & vbNewLine
        SQL = SQL & "       UPDATE @schedule_puid SET schedule_puid=pu_id WHERE schedule_puid IS NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showteam=1 " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET team=( SELECT  crew_desc FROM dbo.crew_schedule cs with (nolock)  " & vbNewLine
        SQL = SQL & "join @schedule_puid sp ON cs.pu_id=sp.schedule_puid " & vbNewLine
        SQL = SQL & "       WHERE #downtime.starttime>=cs.start_time AND cs.END_time>#downtime.starttime AND #downtime.puid=sp.pu_id) " & vbNewLine
        SQL = SQL & "IF @showshift=1 " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET shift=( SELECT  shift_desc FROM dbo.crew_schedule cs with (nolock)  " & vbNewLine
        SQL = SQL & "join @schedule_puid sp ON cs.pu_id=sp.schedule_puid " & vbNewLine
        SQL = SQL & "       WHERE #downtime.starttime>=cs.start_time AND cs.END_time>#downtime.starttime AND #downtime.puid=sp.pu_id) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showproduct=1 " & vbNewLine
        SQL = SQL & "UPDATE #DownTime SET product=( " & vbNewLine
        SQL = SQL & "       SELECT p.Prod_Desc FROM dbo.products p with (nolock)  " & vbNewLine
        SQL = SQL & "       join dbo.production_starts ps with (nolock) ON ps.prod_id= p.prod_id  " & vbNewLine
        SQL = SQL & "       join dbo.prod_units with (nolock) ON ps.pu_id=prod_units.pu_id WHERE ps.start_time <= #downtime.starttime " & vbNewLine
        SQL = SQL & "       AND ((#downtime.starttime < ps.END_time) OR (ps.END_time IS NULL)) AND ps.pu_id=#downtime.puid) " & vbNewLine
        SQL = SQL & "IF @showbrandcode=1 or @showbrandcode=2 " & vbNewLine
        SQL = SQL & "UPDATE #DownTime SET brand=( " & vbNewLine
        SQL = SQL & "       SELECT p.prod_code FROM dbo.products p with (nolock) join dbo.production_starts ps with (nolock) ON ps.prod_id= p.prod_id  " & vbNewLine
        SQL = SQL & "       join dbo.prod_units with (nolock) ON ps.pu_id=prod_units.pu_id WHERE ps.start_time <= #downtime.starttime " & vbNewLine
        SQL = SQL & "       AND ((#downtime.starttime < ps.END_time) OR (ps.END_time IS NULL)) AND ps.pu_id=#downtime.puid) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #DownTime SET ProductGroup = ( " & vbNewLine
        SQL = SQL & "       SELECT TOP 1 product_grp_desc FROM product_groups pg " & vbNewLine
        SQL = SQL & "       join product_group_data pgd ON pgd.Product_Grp_Id = pg.Product_Grp_Id " & vbNewLine
        SQL = SQL & "       join products p ON pgd.Prod_Id = p.Prod_Id " & vbNewLine
        SQL = SQL & "       join comments c ON pg.comment_id = c.comment_id " & vbNewLine
        SQL = SQL & "       WHERE p.prod_code = #downtime.brand " & vbNewLine
        SQL = SQL & "       AND c.comment_text Like '%Package Size%') " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #DownTime SET ProductGroup = ( " & vbNewLine
        SQL = SQL & "       SELECT TOP 1 product_grp_desc FROM product_groups pg " & vbNewLine
        SQL = SQL & "       join product_group_data pgd ON pgd.Product_Grp_Id = pg.Product_Grp_Id " & vbNewLine
        SQL = SQL & "       join products p ON pgd.Prod_Id = p.Prod_Id " & vbNewLine
        SQL = SQL & "       WHERE p.prod_code = #downtime.brand) " & vbNewLine
        SQL = SQL & "WHERE ProductGroup is null " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF (SELECT TOP 1 pu_id FROM dbo.prod_units pu with (nolock) " & vbNewLine
        SQL = SQL & "       JOIN dbo.local_pg_line_status ls with (nolock) ON pu.pu_id=ls.unit_id " & vbNewLine
        SQL = SQL & "       WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "                                  WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All')))  " & vbNewLine
        SQL = SQL & "IS NOT NULL " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "  INSERT into #DTPuIDStartTime(PU_Id, StartTime) " & vbNewLine
        SQL = SQL & "   SELECT DISTINCT PUID, StartTime  " & vbNewLine
        SQL = SQL & "   FROM #DOWNTIME " & vbNewLine
        SQL = SQL & "   ORDER BY PUID ASC " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE #DownTime SET LineStatus=( " & vbNewLine
        SQL = SQL & "            SELECT p.phrase_value FROM dbo.phrase p with (nolock) JOIN dbo.local_pg_line_status ls with (nolock) " & vbNewLine
        SQL = SQL & "            ON ls.line_status_id = p.phrase_id JOIN dbo.prod_units pu with (nolock) ON pu.pu_id=ls.unit_id " & vbNewLine
        SQL = SQL & "            join #DTPuIDStartTime pt ON pt.PU_Id=#DownTime.PUID and pt.StartTime = #DownTime.starttime " & vbNewLine
        SQL = SQL & "        WHERE ls.start_datetime <= #DownTime.starttime AND ((#DownTime.starttime < ls.end_datetime) OR  " & vbNewLine
        SQL = SQL & "              (ls.end_datetime IS NULL)) and ls.unit_id = pt.PU_Id) " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF (SELECT TOP 1 pu_id FROM dbo.prod_units pu with (nolock) " & vbNewLine
        SQL = SQL & "       JOIN dbo.local_pg_line_status ls with (nolock) ON pu.pu_id=ls.unit_id " & vbNewLine
        SQL = SQL & "       WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "                                  WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All'))) " & vbNewLine
        SQL = SQL & "IS NULL " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "SELECT @RptProdPUId = (SELECT TOP 1 pu_id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "                                  WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All')) " & vbNewLine
        SQL = SQL & "AND charindex('production=true', extended_info) > 0) " & vbNewLine
        SQL = SQL & "IF @RptProdPUId = null " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "  SELECT @RptProdPUId = (SELECT TOP 1  Prod_Units.PU_Id " & vbNewLine
        SQL = SQL & "  FROM Table_Fields_Values with (nolock) INNER JOIN " & vbNewLine
        SQL = SQL & "     Table_Fields ON Table_Fields_Values.Table_Field_Id = Table_Fields.Table_Field_Id INNER JOIN " & vbNewLine
        SQL = SQL & "     Prod_Units ON Table_Fields_Values.KeyId = Prod_Units.PU_Id  WHERE(Table_Fields.Table_Field_Desc = 'RE-ProductionUnit')  " & vbNewLine
        SQL = SQL & "  AND (Table_Fields_Values.KeyId in (SELECT  pu_id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "  WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock)  " & vbNewLine
        SQL = SQL & "                            WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All'))))) " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & "INSERT @TESTS (var_id, extendedinfo, result_ON, result) " & vbNewLine
        SQL = SQL & "SELECT vv.var_id, vv.extended_info, result_ON, result FROM dbo.tests tt with (nolock) " & vbNewLine
        SQL = SQL & "JOIN dbo.variables vv with (nolock) ON vv.var_id = tt.var_id  " & vbNewLine
        SQL = SQL & "       AND (charindex('rpthook=productionstatus', vv.extended_info)>0)  " & vbNewLine
        SQL = SQL & "JOIN dbo.prod_units pu with (nolock) ON pu.pu_id = vv.pu_id AND pu.pu_id = @RPTProdPUId " & vbNewLine
        SQL = SQL & "WHERE tt.result_ON <= @inputendtime+1 AND tt.result_ON > @inputstarttime-1 " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET LineStatus = ( " & vbNewLine
        SQL = SQL & "       SELECT TOP 1 result FROM @tests tt  " & vbNewLine
        SQL = SQL & "       JOIN dbo.prod_units pu with (nolock) ON pu.pu_id = @RPTProdPUId " & vbNewLine
        SQL = SQL & "       WHERE charindex('rpthook=productionstatus', tt.extendedinfo)>0 " & vbNewLine
        SQL = SQL & "       AND tt.result_ON > #downtime.starttime  " & vbNewLine
        SQL = SQL & "       ORDER BY tt.result_ON) " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & "IF @showcat1=1 " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET ErtdID1=( " & vbNewLine
        SQL = SQL & "SELECT  ertd.Event_reason_tree_data_id FROM  " & vbNewLine
        SQL = SQL & "dbo.Prod_events pe with (nolock) " & vbNewLine
        SQL = SQL & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id  " & vbNewLine
        SQL = SQL & "WHERE ertd.event_reason_level=1 " & vbNewLine
        SQL = SQL & "AND #downtime.SourcePUID=pe.pu_id  " & vbNewLine
        SQL = SQL & "AND ertd.event_reason_id=#downtime.reasonid1 " & vbNewLine
        SQL = SQL & "AND pe.Event_type = 2)WHERE #downtime.reasonid1 IS NOT NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat2=1 " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET ErtdID2=( " & vbNewLine
        SQL = SQL & "SELECT  ertd.Event_reason_tree_data_id FROM  " & vbNewLine
        SQL = SQL & "dbo.Prod_events pe with (nolock) " & vbNewLine
        SQL = SQL & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id  " & vbNewLine
        SQL = SQL & "WHERE ertd.event_reason_level=2 " & vbNewLine
        SQL = SQL & "AND #downtime.SourcePUID=pe.pu_id  " & vbNewLine
        SQL = SQL & "AND ertd.event_reason_id=#downtime.reasonid2 " & vbNewLine
        SQL = SQL & "AND ertd.Parent_Event_reason_id=#downtime.reasonid1 " & vbNewLine
        SQL = SQL & "AND pe.Event_type = 2)WHERE #downtime.reasonid2 IS NOT NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat3=1 " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET ErtdID3=( " & vbNewLine
        SQL = SQL & "SELECT  ertd.Event_reason_tree_data_id FROM  " & vbNewLine
        SQL = SQL & "dbo.Prod_events pe with (nolock) " & vbNewLine
        SQL = SQL & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id  " & vbNewLine
        SQL = SQL & "join dbo.event_reason_tree_data ertd1 with (nolock) ON ertd1.Event_reason_tree_data_id = ertd.Parent_Event_R_Tree_Data_Id " & vbNewLine
        SQL = SQL & "WHERE ertd.event_reason_level=3 " & vbNewLine
        SQL = SQL & "AND #downtime.SourcePUID=pe.pu_id  " & vbNewLine
        SQL = SQL & "AND ertd.event_reason_id=#downtime.reasonid3  " & vbNewLine
        SQL = SQL & "AND ertd.Parent_Event_reason_id=#downtime.reasonid2 " & vbNewLine
        SQL = SQL & "AND ertd1.Parent_Event_reason_id=#downtime.reasonid1 " & vbNewLine
        SQL = SQL & "AND pe.Event_type = 2 " & vbNewLine
        SQL = SQL & ")WHERE #downtime.reasonid3 IS NOT NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat4=1 " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET ErtdID4=( " & vbNewLine
        SQL = SQL & "SELECT  ertd.Event_reason_tree_data_id FROM  " & vbNewLine
        SQL = SQL & "dbo.Prod_events pe with (nolock) " & vbNewLine
        SQL = SQL & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id " & vbNewLine
        SQL = SQL & "join dbo.event_reason_tree_data ertd1 with (nolock) ON ertd1.Event_reason_tree_data_id = ertd.Parent_Event_R_Tree_Data_Id " & vbNewLine
        SQL = SQL & "join dbo.event_reason_tree_data ertd2 with (nolock) ON ertd2.Event_reason_tree_data_id = ertd1.Parent_Event_R_Tree_Data_Id " & vbNewLine
        SQL = SQL & "WHERE ertd.event_reason_level=4 " & vbNewLine
        SQL = SQL & "AND #downtime.SourcePUID=pe.pu_id  " & vbNewLine
        SQL = SQL & "AND ertd.event_reason_id=#downtime.reasonid4 " & vbNewLine
        SQL = SQL & "AND ertd.Parent_Event_reason_id=#downtime.reasonid3 " & vbNewLine
        SQL = SQL & "AND ertd1.Parent_Event_reason_id=#downtime.reasonid2 " & vbNewLine
        SQL = SQL & "AND ertd2.Parent_Event_reason_id=#downtime.reasonid1 " & vbNewLine
        SQL = SQL & "AND pe.Event_type = 2 " & vbNewLine
        SQL = SQL & ")WHERE #downtime.reasonid4 IS NOT NULL " & vbNewLine
        SQL = SQL & "IF (SELECT count(erc_id) FROM dbo.event_reason_catagories with (nolock)  " & vbNewLine
        SQL = SQL & "WHERE charindex('DTSched', erc_desc) > 0 " & vbNewLine
        SQL = SQL & "OR charindex('DTGroup', erc_desc) > 0 " & vbNewLine
        SQL = SQL & "OR charindex('DTMach', erc_desc) > 0) > 5 " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "IF @showcat1=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "       SELECT @CatPrefix='DTSched-' " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat1=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID4 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat1=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID3 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat1=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID2 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat1=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID1 IS NOT NULL " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat2=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "       SELECT @CatPrefix='DTGroup-' " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat2=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID4 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat2=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID3 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat2=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID2 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat2=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID1 IS NOT NULL " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat3=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "       SELECT @CatPrefix='DTMach-' " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat3=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID4 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat3=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID3 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat3=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID2 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat3=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID1 IS NOT NULL " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat4=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "       SELECT @CatPrefix='DTType-' " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat4=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID4 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat4=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID3 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat4=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID2 IS NOT NULL " & vbNewLine
        SQL = SQL & "        " & vbNewLine
        SQL = SQL & "       UPDATE #downtime SET Cat4=( " & vbNewLine
        SQL = SQL & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM  " & vbNewLine
        SQL = SQL & "       dbo.event_reason_catagories erc with (nolock) " & vbNewLine
        SQL = SQL & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id " & vbNewLine
        SQL = SQL & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0)  " & vbNewLine
        SQL = SQL & "       AND ercd.Propegated_FROM_etDid IS NULL " & vbNewLine
        SQL = SQL & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID1 IS NOT NULL " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF (SELECT count(erc_id) FROM dbo.event_reason_catagories with (nolock)  " & vbNewLine
        SQL = SQL & "WHERE charindex('category:', erc_desc) > 0 " & vbNewLine
        SQL = SQL & "OR charindex('Schedule:', erc_desc) > 0 " & vbNewLine
        SQL = SQL & "OR charindex('Subsystem:', erc_desc) > 0 " & vbNewLine
        SQL = SQL & "OR charindex('GroupCause:', erc_desc) > 0) > 5 " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat1=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "       SELECT @CatPrefix='category:' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE td SET " & vbNewLine
        SQL = SQL & "              Cat1 = right(erc_desc,len(erc_desc)-len(@CatPrefix)) " & vbNewLine
        SQL = SQL & "       FROM dbo.#downtime td with (nolock) " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id  " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON ercd.erc_id = erc.erc_id  " & vbNewLine
        SQL = SQL & "       WHERE erc.ERC_Desc LIKE @CatPrefix + '%' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat2=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "   SELECT @CatPrefix='Schedule:' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE td SET " & vbNewLine
        SQL = SQL & "              Cat2 = right(erc_desc,len(erc_desc)-len(@CatPrefix)) " & vbNewLine
        SQL = SQL & "       FROM dbo.#downtime td with (nolock) " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id  " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON ercd.erc_id = erc.erc_id  " & vbNewLine
        SQL = SQL & "       where erc.ERC_Desc LIKE @CatPrefix + '%' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat3=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & "       SELECT @CatPrefix='Subsystem:' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE td SET " & vbNewLine
        SQL = SQL & "              Cat3 = right(erc_desc,len(erc_desc)-len(@CatPrefix)) " & vbNewLine
        SQL = SQL & "       FROM dbo.#downtime td with (nolock) " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id  " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON ercd.erc_id = erc.erc_id  " & vbNewLine
        SQL = SQL & "       where erc.ERC_Desc LIKE @CatPrefix + '%' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showcat4=1 " & vbNewLine
        SQL = SQL & "       BEGIN " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       SELECT @CatPrefix='GroupCause:' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       UPDATE td SET " & vbNewLine
        SQL = SQL & "              Cat4 = right(erc_desc,len(erc_desc)-len(@CatPrefix)) " & vbNewLine
        SQL = SQL & "       FROM dbo.#downtime td with (nolock) " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id  " & vbNewLine
        SQL = SQL & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK)  " & vbNewLine
        SQL = SQL & "       ON ercd.erc_id = erc.erc_id  " & vbNewLine
        SQL = SQL & "       WHERE erc.ERC_Desc LIKE @CatPrefix + '%' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "       END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET uptime=datedIFf(s,( " & vbNewLine
        SQL = SQL & "SELECT TOP 1 dt2.endtime FROM dbo.#downtime dt2 with (nolock) WHERE ((dt2.DT_ID=(#downtime.DT_ID - 1)))),starttime)/60.0 " & vbNewLine
        SQL = SQL & "WHERE #downtime.DT_ID > 0 " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET uptime=datedIFf(s,( " & vbNewLine
        SQL = SQL & "SELECT max(dt2.END_time) FROM dbo.timed_event_details dt2 with (nolock)  " & vbNewLine
        SQL = SQL & "join dbo.prod_units pu2 with (nolock) ON pu2.pu_id=dt2.pu_id " & vbNewLine
        SQL = SQL & "WHERE  END_time<=#downtime.starttime  " & vbNewLine
        SQL = SQL & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All')) " & vbNewLine
        SQL = SQL & "),starttime)/60.0 WHERE uptime IS NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "declare @preEND datetime " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "SELECT @preEND = max(END_time) " & vbNewLine
        SQL = SQL & "FROM dbo.timed_event_details ted with (nolock)  " & vbNewLine
        SQL = SQL & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id) " & vbNewLine
        SQL = SQL & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id) " & vbNewLine
        SQL = SQL & "inner join dbo.prod_units as pu2 with (nolock) ON (pu2.pu_id = ted.pu_id) " & vbNewLine
        SQL = SQL & "WHERE  (END_time <= @InputStartTime)  " & vbNewLine
        SQL = SQL & "AND (CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  ) > 0 OR (@InputMasterProdUnit = 'All')) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #downtime SET uptime=datedIFf(s,@preEND,starttime)/60.0  " & vbNewLine
        SQL = SQL & "WHERE uptime IS NULL " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "DECLARE  " & vbNewLine
        SQL = SQL & "                     @fltDBVersion        FLOAT " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF     (      SELECT        IsNumeric(App_Version) " & vbNewLine
        SQL = SQL & "                     FROM   dbo.AppVersions with (nolock) " & vbNewLine
        SQL = SQL & "                     WHERE  App_Id = 2) = 1 " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "       SELECT        @fltDBVersion = Convert(Float, App_Version) " & vbNewLine
        SQL = SQL & "              FROM   dbo.AppVersions with (nolock) " & vbNewLine
        SQL = SQL & "              WHERE  App_Id = 2 " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & "ELSE " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "       SELECT @fltDBVersion = 1.0 " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "DECLARE  " & vbNewLine
        SQL = SQL & "              @DowntimesystemUserID             INT " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "Select @DowntimesystemUserID =  " & vbNewLine
        SQL = SQL & "       User_ID " & vbNewLine
        SQL = SQL & "       FROM dbo.USERS with (nolock) " & vbNewLine
        SQL = SQL & "       WHERE UserName = 'ReliabilitySystem' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF     @fltDBVersion <= 300172.90  " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & "              UPDATE #Downtime " & vbNewLine
        SQL = SQL & "                           Set Split = 'S' " & vbNewLine
        SQL = SQL & "              FROM dbo.#Downtime tdt with (nolock) " & vbNewLine
        SQL = SQL & "              Join dbo.Timed_Event_Detail_History ted with (nolock) on tdt.TeDet_Id = ted.TEDET_ID  " & vbNewLine
        SQL = SQL & "              and ted.User_ID = '2'  " & vbNewLine
        SQL = SQL & "              WHERE ted.TEDET_ID IS NOT Null " & vbNewLine
        SQL = SQL & "                       And tdt.Split IS Null " & vbNewLine
        SQL = SQL & "               " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & "ELSE  " & vbNewLine
        SQL = SQL & "BEGIN                       " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "              UPDATE #Downtime " & vbNewLine
        SQL = SQL & "                           Set Split = ISNULL((Select Case     WHEN Min(User_Id) < 50 " & vbNewLine
        SQL = SQL & "                                                                               THEN       '' " & vbNewLine
        SQL = SQL & "                                                                               WHEN       Min(User_Id) = @DowntimesystemUserID " & vbNewLine
        SQL = SQL & "                                                                               THEN       '' " & vbNewLine
        SQL = SQL & "                                                                               ELSE       'S' " & vbNewLine
        SQL = SQL & "                                                            END " & vbNewLine
        SQL = SQL & "                                                            FROM dbo.Timed_Event_Detail_History with (nolock)  " & vbNewLine
        SQL = SQL & "                                                            WHERE Tedet_Id = ted.TEDET_ID),'S')  " & vbNewLine
        SQL = SQL & "              FROM dbo.#Downtime tdt with (nolock) Left Join dbo.Timed_Event_Detail_History ted  " & vbNewLine
        SQL = SQL & "                    with (nolock) on tdt.TEDet_ID = ted.TEDET_ID " & vbNewLine
        SQL = SQL & "              Where tdt.Uptime = 0 " & vbNewLine
        SQL = SQL & "               " & vbNewLine
        SQL = SQL & "                     UPDATE #Downtime " & vbNewLine
        SQL = SQL & "                                  Set Split = NuLL " & vbNewLine
        SQL = SQL & "                     FROM dbo.#Downtime tdt with (nolock) " & vbNewLine
        SQL = SQL & "                     Join dbo.Timed_Event_Detail_History tedh with (nolock) on tdt.TEDet_ID = tedh.TEDET_ID " & vbNewLine
        SQL = SQL & "                     WHERE tedh.User_ID = @DowntimeSystemUserID " & vbNewLine
        SQL = SQL & "                                         And tdt.Downtime <> 0 " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF EXISTS (select * from dbo.sysobjects where id = object_id(N'[dbo].[fnLocal_GlblParseInfo]')) " & vbNewLine
        SQL = SQL & "BEGIN " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "      IF EXISTS (SELECT PU_ID from prod_units where charindex('RateLoss', extended_info) > 0) " & vbNewLine
        SQL = SQL & "      BEGIN " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "            UPDATE #downtime SET targetspeed = ( " & vbNewLine
        SQL = SQL & "                  SELECT result FROM tests tt " & vbNewLine
        SQL = SQL & "                  join variables vv ON vv.var_id = tt.var_id " & vbNewLine
        SQL = SQL & "                        and vv.pu_id = #downtime.puid " & vbNewLine

        SQL = SQL & "                        and dbo.fnLocal_GlblParseInfo(vv.Extended_Info, 'GlblDesc=') LIKE '%' + replace('Line Target Speed',' ','') " & vbNewLine
        SQL = SQL & "                  WHERE tt.result_on = #downtime.starttime)   " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "            UPDATE #downtime SET actualspeed = ( " & vbNewLine
        SQL = SQL & "                  SELECT result FROM tests tt " & vbNewLine
        SQL = SQL & "                  join variables vv ON vv.var_id = tt.var_id " & vbNewLine
        SQL = SQL & "                        and vv.pu_id = #downtime.puid " & vbNewLine

        SQL = SQL & "                        and dbo.fnLocal_GlblParseInfo(vv.Extended_Info, 'GlblDesc=') LIKE '%' + replace('Line Actual Speed',' ','') " & vbNewLine
        SQL = SQL & "                  WHERE tt.result_on = #downtime.starttime)   " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine

        SQL = SQL & "           UPDATE #downtime SET downtime = (targetspeed - actualspeed) * downtime / targetspeed--, " & vbNewLine


        SQL = SQL & "                  WHERE targetspeed is not null  " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "      END " & vbNewLine
        SQL = SQL & "END " & vbNewLine
        SQL = SQL & " " & vbNewLine

        SQL = SQL & " " & vbNewLine

        SQL = SQL & "       SELECT @ShowLineStatus = 1 " & vbNewLine

        SQL = SQL & "UPDATE #DownTime " & vbNewLine
        SQL = SQL & "SET iBrand=cast(Brand as INT) " & vbNewLine
        SQL = SQL & "WHERE (IsNumeric(Brand)=1 AND LEN(Brand)=8) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "UPDATE #DownTime " & vbNewLine
        SQL = SQL & "SET iShift=cast(Shift as INT) " & vbNewLine
        SQL = SQL & "WHERE IsNumeric(Shift)=1 " & vbNewLine

        SQL = SQL & " " & vbNewLine

        SQL = SQL & "SELECT @strsql='SELECT starttime,endtime,downtime,uptime,location,fault,reason1,reason2,split' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "IF @showreason3      =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql='SELECT starttime,endtime,downtime,uptime,location,fault,reason1,reason2,reason3,split' " & vbNewLine
        SQL = SQL & "IF @showreason4      =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql='SELECT starttime,endtime,downtime,uptime,location,fault,reason1,reason2,reason3,reason4,split' " & vbNewLine
        SQL = SQL & "IF @showmasterprodunit =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+',MasterProdUnit' " & vbNewLine
        SQL = SQL & "IF @showTeam  =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+',team' " & vbNewLine
        SQL = SQL & "IF @showshift =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',ishift' " & vbNewLine
        SQL = SQL & "IF @showshift =2                        " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',shift'  " & vbNewLine


        SQL = SQL & "IF @showProduct      =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',product' " & vbNewLine
        SQL = SQL & "IF @showbrandCode=1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',ibrand'  " & vbNewLine
        SQL = SQL & "IF @showbrandCode=2                        " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',brand'  " & vbNewLine
        SQL = SQL & "IF @showProdGroup = 1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',ProductGroup' " & vbNewLine
        SQL = SQL & "IF @showCat1  =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',cat1' " & vbNewLine
        SQL = SQL & "IF @showCat2  =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',Cat2' " & vbNewLine
        SQL = SQL & "IF @showCat3  =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',cat3' " & vbNewLine
        SQL = SQL & "IF @showCat4  =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',cat4' " & vbNewLine
        SQL = SQL & "IF @ShowLineStatus   =1 " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',LineStatus' " & vbNewLine
        SQL = SQL & "IF @showComment      =1  " & vbNewLine
        SQL = SQL & "       SELECT @strsql=@strsql+ ',comments'  " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "SELECT @strsql=@strsql+' FROM dbo.#downtime with (nolock) ORder by starttime' " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "print @strsql " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "exec (@strsql) " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "GOTO FinIShed " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "ErrorMessagesWrite: " & vbNewLine
        SQL = SQL & "       SELECT ErrMsg " & vbNewLine
        SQL = SQL & "              FROM   @ErrorMessages " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "Finished: " & vbNewLine
        SQL = SQL & " " & vbNewLine
        SQL = SQL & "DROP TABLE #Downtime " & vbNewLine
        SQL = SQL & "DROP TABLE #DTPuIDStartTime " & vbNewLine

        Return SQL
    End Function

End Module
