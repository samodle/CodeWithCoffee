Public Module PROF_customSQLqueries
    Public SQL_pqq1 As String
    Public SQL_pqq2 As String
    Public SQL_pqq3 As String

    Public SQL_pqqprod1 As String
    Public SQL_pqqprod2 As String
    Public SQL_pqqprod3 As String

    Public SQL_oneclick1 As String
    Public SQL_oneclick2 As String
    Public SQL_oneclick3 As String


    Public Sub setSQLstrings_PQQDowntime(startTime As Date, endTime As Date, masterProdUnit As String)



        ' Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
        ' .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")


        SQL_pqq1 = SQL_pqq1 & "USE [GBDB]" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "SET ANSI_NULLS ON" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "SET QUOTED_IDENTIFIER OFF" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "DECLARE" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputStartTime         DateTime = '" & Format(startTime, "Short Date") & " " & Format(startTime, "Long Time") & "'," & vbNewLine   '  '10/27/2015 06:00:00 AM'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputEndTime           DateTime = '" & Format(endTime, "Short Date") & " " & Format(endTime, "Long Time") & "'," & vbNewLine   '  '11/1/2015 06:00:00 AM'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputMasterProdUnit          nVarChar(4000) = '" & masterProdUnit & "'," & vbNewLine   '  'HCMR011 Main'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputLocations         nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputFaults            nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputReason1s    nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputReason2s    nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputReason3s    nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputReason4s    nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputTeams       nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputShifts            nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputCat1s       nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputCat2s       nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputCat3s       nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputCat4s       nVarChar(4000) = 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputGroupBy           varchar(50) = 'Nothing'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputOrderBy           varchar(50) = 'Downtime'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputDurationLimit     varchar(50) = 0," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputDurationOper      varchar(50) = '>'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputOutPutType  varchar(50) = 'Raw Data'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputInterval          int = 1440," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputProducts          nvarchar(4000)= 'All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputProductGroups     nvarchar(4000)='All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputStopClassifications     nvarchar(4000)='All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputProductionStatus        nvarchar(4000)='All'," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputCalcUptime  int = 1," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "@InputCalcShiftTeam     int = 0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "DECLARE     @Position int," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @PU_Id  int," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @ScheduleUnit_PU_Id int," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @PUScheduleUnitStr  nVarChar(255)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @PartialString  nVarChar(255)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @@ExtendedInfo nvarChar(255)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @InputOrderByClause     nvarChar(4000)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @InputGroupByClause     nvarChar(4000)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @strSQL                 nVarChar(4000)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @current datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @tmpStartTime as datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @tmpEndTime as datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @tmpCount as int," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @tmpLoopCounter  int," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @PLId int," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            @RptProdPUId      int" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        ' SQL_pqq1 = SQL_pqq1 & "DROP TABLE #DownTime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "DECLARE @Tests TABLE(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      var_id                  int," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      result                  varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      result_on         datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      extendedinfo      varchar(255))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Create table #DownTime(    " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      StartTime   datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      EndTime           datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Uptime            Float," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Downtime    Float," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      MasterProdUnit    varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Location    varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Fault       varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason1           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason2           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason3           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason4           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      StopClass   varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ProdStatus  varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Team        varchar(25)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Shift       varchar(25)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat1        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat2        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat3        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat4        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat5        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat6        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat7        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat8        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat9        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat10       varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Product           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ProductCode varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ProductGroup      varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Comments    varchar(2550)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      StartTime_Act     datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      EndTime_Act datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Endtime_Prev      datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      PUID        INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      SourcePUID  INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      tedet_id          INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID1   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID2   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID3   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID4   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID1     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID2     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID3     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID4     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      action_level1     varchar(255)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      SBP   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      EAP   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      CommentId   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare1            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare2            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare3            varchar(255)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare4            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare5            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare6            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare7            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare8            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare9            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare10           varchar(255)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "CREATE      INDEX td_PUId_StartTime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ON    #DownTime (PUId, StartTime)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "CREATE      INDEX td_PUId_EndTime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ON    #DownTime (PUId, EndTime)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Create table #DownTimeFinal(    " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      StartTime   datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      EndTime           datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Uptime            Float," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Downtime    Float," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      MasterProdUnit    varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Location    varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Fault       varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason1           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason2           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason3           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Reason4           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      StopClass   varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ProdStatus  varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Team        varchar(25)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Shift       varchar(25)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat1        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat2        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat3        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat4        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat5        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat6        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat7        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat8        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat9        varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Cat10       varchar(50)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Product           varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ProductCode varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ProductGroup      varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Comments    varchar(2550)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      StartTime_Act     datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      EndTime_Act datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Endtime_Prev      datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      PUID        INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      SourcePUID  INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      tedet_id          INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID1   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID2   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID3   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ReasonID4   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID1     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID2     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID3     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErtdID4     INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      action_level1     varchar(255)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      SBP   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      EAP   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      CommentId   INT," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare1            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare2            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare3            varchar(255)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare4            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare5            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare6            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare7            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare8            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare9            varchar(255), " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Spare10           varchar(255)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "CREATE TABLE #DTSummary ( " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      groupName  varchar(100)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Downtime    Float," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Stops Int" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "CREATE TABLE #DTInterval (    " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      StartTime         datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      EndTime           datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Duration          Float," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      TotalDowntime           Float," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      TotalStops        Int" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "DECLARE @TmpCrew TABLE(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      Crew varchar(30)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      shift varchar(30)," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      starttime datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      endtime datetime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      puid int)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "DECLARE @ErrorMessages TABLE(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ErrMsg                        nVarChar(255) )" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "IF    IsDate(@InputStartTime) <> 1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "BEGIN" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      INSERT      @ErrorMessages (ErrMsg)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            VALUES      ('StartTime is not a Date.')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      GOTO  ErrorMessagesWrite" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "END" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "IF    IsDate(@InputEndTime) <> 1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "BEGIN" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      INSERT      @ErrorMessages (ErrMsg)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            VALUES      ('EndTime is not a Date.')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      GOTO  ErrorMessagesWrite" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "END" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "IF datepart(Year,@inputstarttime) > 2500" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "BEGIN" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select @inputstarttime = DATEADD(Year,-543,@inputstarttime)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select @inputendtime = DATEADD(Year,-543,@inputendtime)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "END" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select @PU_Id = (Select PU_ID  from prod_units pu WITH (NOLOCK) where  pu.pu_desc =  @InputMasterProdUnit)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "SELECT @ScheduleUnit_PU_Id = (" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      SELECT      tfv.Value" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      FROM dbo.Table_Fields_Values tfv   WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            JOIN dbo.Prod_Units                pu    WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "                                                                  ON    pu.PU_Id = tfv.KeyId" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            JOIN  dbo.Table_Fields        tf    WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "                                                                  ON    tf.Table_Field_Id       = tfv.Table_Field_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      WHERE tf.Table_Field_Desc = 'ScheduleUnit'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            AND   tfv.TableId = (SELECT TableId FROM dbo.Tables WITH (NOLOCK) WHERE TableName = 'Prod_Units')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            AND   pu.PU_Id = @PU_Id)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "SELECT @PLId = (select PL_Id from prod_units WITH (NOLOCK) where pu_desc = @InputMasterProdUnit)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "SELECT @RptProdPUId = (" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      SELECT TOP 1 PU_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      FROM dbo.Prod_Units                 pu    WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            JOIN  dbo.Table_Fields_Values tfv   WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "                                                                  ON    pu.PU_Id = tfv.KeyId" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            JOIN  dbo.Table_Fields        tf    WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "                                                                  ON    tf.Table_Field_Id       = tfv.Table_Field_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      WHERE tf.Table_Field_Desc = 'RE-ProductionUnit'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            AND   tfv.TableId = (SELECT   TableId     FROM dbo.Tables WITH (NOLOCK) WHERE      TableName = 'Prod_Units')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            AND   tfv.Value = '1' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            AND pu.pl_id = @PLId)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "INSERT INTO #DownTime (StartTime,EndTime,MasterProdUnit,Fault,Location,PUID,Reason1,Reason2,Reason3,Reason4," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "startTime_act,EndTime_act,tedet_id,REASONID1,REASONID2,REASONID3,REASONID4,SourcePUID, action_level1, CommentId) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc ,  ted.pu_ID, r1.event_reason_name, " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "r2.event_reason_name, r3.event_reason_name, r4.event_reason_name, ted.start_time, ted.end_time, ted.tedet_id," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "ted.reason_level1, ted.reason_level2, ted.reason_level3, ted.reason_level4, ted.Source_PU_ID, action_level1, Cause_comment_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "FROM timed_event_details AS ted WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "LEFT JOIN event_reasons AS r1 WITH (NOLOCK) ON (r1.event_reason_id = ted.reason_level1)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "LEFT JOIN event_reasons AS r2 WITH (NOLOCK) ON (r2.event_reason_id = ted.reason_level2)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "LEFT JOIN event_reasons AS r3 WITH (NOLOCK) ON (r3.event_reason_id = ted.reason_level3)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "LEFT JOIN event_reasons AS r4 WITH (NOLOCK) ON (r4.event_reason_id = ted.reason_level4)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "LEFT JOIN timed_event_fault AS tef WITH (NOLOCK) ON (tef.tefault_id = ted.tefault_id)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "LEFT Join prod_units AS pu WITH (NOLOCK) ON (pu.pu_id = ted.source_PU_Id)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "INNER JOIN prod_units as pu2 WITH (NOLOCK) ON (pu2.pu_id = ted.pu_id)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "WHERE (ted.pu_id = @PU_Id)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and( ((Start_time < =  @InputEndTime) and (end_time > @InputStartTime)) or ((Start_time < =  @InputEndTime )and end_time Is Null) )" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ((CHARINDEX( ','+PU.pu_desc+','  ,  ','+ @InputLocations+','  )) > 0 or (@InputLocations = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ((CHARINDEX( ','+tef.tefault_name+','  ,  ','+ @InputFaults+','  )) > 0 or (@InputFaults = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ((CHARINDEX( ','+r1.event_reason_name+','  ,  ','+ @InputReason1s+','  )) > 0 or (@InputReason1s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ((CHARINDEX( ','+r2.event_reason_name+','  ,  ','+ @InputReason2s+','  )) > 0 or (@InputReason2s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ((CHARINDEX( ','+r3.event_reason_name+','  ,  ','+ @InputReason3s+','  )) > 0 or (@InputReason3s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ((CHARINDEX( ','+r4.event_reason_name+','  ,  ','+ @InputReason4s+','  )) > 0 or (@InputReason4s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "if @InputOutputType = 'raw data' and @InputCalcUptime=1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      update #downtime set uptime = datediff(s,isnull((" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select max(end_time)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      from timed_event_details ted WITH (NOLOCK) where ted.Pu_id = #downtime.puid and ted.start_time<#downtime.starttime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      ),#downtime.starttime), #downtime.starttime)/60.0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "declare @comment_text varchar(2000), @comment_source int" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "declare c_id insensitive cursor for" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select comment_text, dt.CommentId from Comments c WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join #downtime dt on dt.CommentId = c.topofchain_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where c.comment_text is not null" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      order by Modified_On" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "open c_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "FETCH NEXT From c_id into @comment_text, @comment_source" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "while @@fetch_status=0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      begin" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      update #downtime set comments = isnull(comments,'') + ',' + @comment_text where CommentId = @comment_source" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      FETCH NEXT From c_id into @comment_text, @comment_source" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      end" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "deallocate c_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set comments = substring (comments,2,2000)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set starttime=@InputStartTime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            SBP = 1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where starttime<@InputStartTime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set endtime=@InputEndTime," & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            EAP = 1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where endtime>@InputEndTime " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Downtime = datediff(s,starttime,endtime)/60.0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Downtime = datediff(s,starttime,@InputEndTime)/60.0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where Endtime is null" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "insert into @TmpCrew (crew,shift,starttime,endtime,puid)(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Select Crew_Desc,shift_desc,start_time,end_time,pu_id from crew_schedule cs WITH (NOLOCK) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "WHERE (((cs.end_time between @InputStartTime    and @InputEndTime) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "or (cs.start_time between @InputStartTime and @InputEndTime)     )" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "or ( (@InputStartTime   >= cs.start_time) and( @InputEndTime <  cs.end_time)))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and (PU_id = @ScheduleUnit_PU_Id))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "insert @Tests (var_id, extendedinfo, result_on, result)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select vv.var_id, vv.extended_info, result_on, result from tests tt WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join variables vv WITH (NOLOCK) on vv.var_id = tt.var_id and " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "(charindex('rpthook=productionstatus', vv.extended_info)>0 or " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "charindex('rpthook=starttime', vv.extended_info)>0 or " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "charindex('rpthook=team', vv.extended_info)>0 or " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "charindex('rpthook=shift', vv.extended_info)>0)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join prod_units pu WITH (NOLOCK) on pu.pu_id = vv.pu_id and pu.pu_id = @RPTProdPUId" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where tt.result_on <= @inputendtime+1 and tt.result_on > @inputstarttime-1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "if (@inputoutputtype='Raw Data' and @InputCalcShiftTeam = 1) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "or (@inputoutputtype = 'Pareto' and @inputgroupby = 'team')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "or @inputteams<>'all'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      begin" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      update #downtime set team=( select  crew from @TmpCrew tc where #downtime.starttime>=tc.starttime " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and tc.endtime>#downtime.starttime)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      end" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "if (@inputoutputtype='Raw Data' and @InputCalcShiftTeam = 1) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "or (@inputoutputtype = 'Pareto' and @inputgroupby = 'shift')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "or @inputshifts<>'all'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      begin" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      update #downtime set SHIFT = ( select shift from @TmpCrew tc where #downtime.starttime>=tc.starttime " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and tc.endtime>#downtime.starttime)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      end" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #DownTime set product=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select p.Prod_Desc from products p WITH (NOLOCK) join production_starts ps WITH (NOLOCK) on ps.prod_id= p.prod_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join prod_units WITH (NOLOCK) on ps.pu_id=prod_units.pu_id where ps.start_time <= #downtime.starttime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((#downtime.starttime < ps.end_time) or (ps.end_time is null)) and ps.pu_id=@RptProdPUId)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #DownTime set productcode=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select p.Prod_code from products p WITH (NOLOCK) join production_starts ps WITH (NOLOCK) on ps.prod_id= p.prod_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join prod_units WITH (NOLOCK) on ps.pu_id=prod_units.pu_id where ps.start_time <= #downtime.starttime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((#downtime.starttime < ps.end_time) or (ps.end_time is null)) and ps.pu_id=@RptProdPUId)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set productGroup = (" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select top 1 product_grp_desc from product_groups pg WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join product_group_data pgd WITH (NOLOCK) on pgd.Product_Grp_Id = pg.Product_Grp_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join products p WITH (NOLOCK) on pgd.Prod_Id = p.Prod_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join comments c WITH (NOLOCK) ON pg.comment_id = c.comment_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where p.prod_code = #downtime.productcode" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      AND c.comment_text Like '%Package Size%')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set productGroup = (" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select top 1 product_grp_desc from product_groups pg WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join product_group_data pgd WITH (NOLOCK) on pgd.Product_Grp_Id = pg.Product_Grp_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join products p WITH (NOLOCK) on pgd.Prod_Id = p.Prod_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where p.prod_code = #downtime.productcode)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "WHERE ProductGroup is null" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set ErtdID1=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  ertd.Event_reason_tree_data_id from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Prod_events pe WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_tree_data ertd WITH (NOLOCK) on ertd.tree_name_id=pe.Name_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ertd.event_reason_level=1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd.event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and pe.Event_type = 2)where #downtime.reasonid1 is NOT null" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set ErtdID2=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  ertd.Event_reason_tree_data_id from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Prod_events pe WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_tree_data ertd WITH (NOLOCK) on ertd.tree_name_id=pe.Name_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ertd.event_reason_level=2" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd.event_reason_id=#downtime.reasonid2" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd.Parent_Event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and pe.Event_type = 2)where #downtime.reasonid2 is NOT null" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set ErtdID3=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  ertd.Event_reason_tree_data_id from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Prod_events pe WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_tree_data ertd WITH (NOLOCK) on ertd.tree_name_id=pe.Name_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_tree_data ertd1 WITH (NOLOCK) on ertd1.Event_reason_tree_data_id = ertd.Parent_Event_R_Tree_Data_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "                                       " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ertd.event_reason_level=3" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd.event_reason_id=#downtime.reasonid3 " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd.Parent_Event_reason_id=#downtime.reasonid2" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd1.Parent_Event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and pe.Event_type = 2" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.reasonid3 is NOT null" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set ErtdID4=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  ertd.Event_reason_tree_data_id from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Prod_events pe WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_tree_data ertd WITH (NOLOCK) on ertd.tree_name_id=pe.Name_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_tree_data ertd1 WITH (NOLOCK) on ertd1.Event_reason_tree_data_id = ertd.Parent_Event_R_Tree_Data_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_tree_data ertd2 WITH (NOLOCK) on ertd2.Event_reason_tree_data_id = ertd1.Parent_Event_R_Tree_Data_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ertd.event_reason_level=4" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd.event_reason_id=#downtime.reasonid4" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd.Parent_Event_reason_id=#downtime.reasonid3" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd1.Parent_Event_reason_id=#downtime.reasonid2" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ertd2.Parent_Event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and pe.Event_type = 2" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.reasonid4 is NOT null" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "declare @xxx varchar (30)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select @xxx='DTSched-'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat1=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID4 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat1 is null and #downtime.ErtdID4 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat1=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID3 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat1 is null and #downtime.ErtdID3 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat1=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID2 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat1 is null and #downtime.ErtdID2 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat1=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID1 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat1 is null and #downtime.ErtdID1 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select @xxx='DTGroup-'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat2=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID4 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat2 is null and #downtime.ErtdID4 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat2=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID3 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat2 is null and #downtime.ErtdID3 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat2=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID2 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat2 is null and #downtime.ErtdID2 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat2=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID1 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat2 is null and #downtime.ErtdID1 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select @xxx='DTMach-'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat3=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID4 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat3 is null and #downtime.ErtdID4 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat3=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID3 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat3 is null and #downtime.ErtdID3 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat3=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID2 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat3 is null and #downtime.ErtdID2 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat3=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID1 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat3 is null and #downtime.ErtdID1 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select @xxx='DTType-'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat4=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID4 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat4 is null and #downtime.ErtdID4 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat4=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID3 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat4 is null and #downtime.ErtdID3 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat4=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID2 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat4 is null and #downtime.ErtdID2 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set Cat4=(" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select  top 1 right(erc_desc,len(erc_desc)-len(@xxx)) from " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "event_reason_catagories erc WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join event_reason_Category_data ercd WITH (NOLOCK) on ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where ercd.Event_reason_tree_data_id = #downtime.ErtdID1 and (charindex(lower(@xxx),erc_desc)>0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "and ercd.Propegated_from_etDid is NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & ")where #downtime.Cat4 is null and #downtime.ErtdID1 is NOT NULL" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update dt set dt.action_level1 = erc.erc_desc " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "from #downtime dt" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "left join event_reason_tree_data td WITH (NOLOCK) on dt.action_level1 = td.event_reason_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "left JOIN EVENT_REASON_CATEGORY_DATA ERCD WITH (NOLOCK) ON ERCD.Event_Reason_Tree_Data_Id = td.Event_Reason_Tree_Data_Id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "left JOIN EVENT_REASON_CATAGORIES ERC WITH (NOLOCK) ON ERC.ERC_id = ercd.erc_id" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "WHERE charindex('DTClass', erc.erc_desc)>0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "declare @medium float, @MAJOR FLOAT" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select @medium =10.0, @major =60.0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "IF EXISTS (SELECT Prop_id from Product_Properties WHERE prop_desc ='dtbreakdownsettings')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "BEGIN" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select @medium= isnull(min(convert(float,l_reject)),@medium)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      from product_properties pp WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join characteristics cc WITH (NOLOCK) on cc.prop_id = pp.prop_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join active_specs spec WITH (NOLOCK) on spec.char_id = cc.char_id and effective_date<@inputendtime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and (expiration_date>=@inputstarttime or expiration_date is null)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where pp.prop_desc ='dtbreakdownsettings' and isnumeric(l_reject)>0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select @major= isnull(max(convert(float,l_reject)),@major)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      from product_properties pp WITH (NOLOCK)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join characteristics cc WITH (NOLOCK) on cc.prop_id = pp.prop_id " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      join active_specs spec WITH (NOLOCK) on spec.char_id = cc.char_id and effective_date<@inputendtime" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and (expiration_date>=@inputstarttime or expiration_date is null)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where pp.prop_desc ='dtbreakdownsettings' and isnumeric( l_reject)>0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "END" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set stopclass = 'Minor' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where cat1 ='unplanned' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and (cat3 = 'internal' OR cat3 = 'supply')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and (action_level1= 'dtclass-minorstop' or downtime <= 10.0) " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set stopclass = 'Process Failure' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      where cat1 ='unplanned' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and (cat3 = 'internal' OR cat3 = 'supply')" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and (action_level1 = 'dtclass-processfailure' or downtime > 10.0)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set stopclass = 'Minor Breakdown' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where cat1 ='unplanned' and action_level1='dtclass-breakdown' and downtime < @MEDIUM " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set stopclass = 'Medium Breakdown' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where cat1 ='unplanned' and action_level1='dtclass-breakdown' and downtime >= @MEDIUM and downtime < @MAJOR " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set stopclass = 'Major Breakdown' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "where cat1 ='unplanned' and action_level1='dtclass-breakdown' and downtime >= @MAJOR" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "update #downtime set prodstatus = (" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "select top 1 result from @Tests tt " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "join prod_units pu WITH (NOLOCK) on pu.pu_id = @RPTProdPUId" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Where charindex('rpthook=productionstatus', tt.extendedinfo)>0" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "AND tt.result_on > #downtime.starttime " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "order by tt.result_on)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "IF @InputDurationOper <>  '<' and @InputDurationOper <> '>'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "BEGIN" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      INSERT      @ErrorMessages (ErrMsg)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      VALUES      ('Duration Oper Not Valid=' +  @InputDurationOper )" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      GOTO  ErrorMessagesWrite" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "END" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "IF @InputDurationLimit =  '' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "BEGIN" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      INSERT      @ErrorMessages (ErrMsg)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      VALUES      ('Duration Limit Not Valid=' +  @InputDurationLimit )" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      GOTO  ErrorMessagesWrite" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "END" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "If @InputDurationOper = '<'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Begin" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      insert into #DownTimeFinal" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select *" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      FROM #DownTime as DT" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      WHERE  ((CHARINDEX( ','+DT.Shift+','  ,  ','+ @InputShifts+','  )) > 0 or (@InputShifts = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Team+','  ,  ','+ @InputTeams+','  )) > 0 or (@InputTeams = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat1+','  ,  ','+ @InputCat1s+','  )) > 0 or (@InputCat1s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat2+','  ,  ','+ @InputCat2s+','  )) > 0 or (@InputCat2s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat3+','  ,  ','+ @InputCat3s+','  )) > 0 or (@InputCat3s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat4+','  ,  ','+ @InputCat4s+','  )) > 0 or (@InputCat4s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.stopclass+','  ,  ','+ @Inputstopclassifications+','  )) > 0 or (@Inputstopclassifications = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.prodstatus+','  ,  ','+ @Inputproductionstatus+','  )) > 0 or (@Inputproductionstatus = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.productcode+','  ,  ','+ @Inputproducts + ','  )) > 0 or (@Inputproducts = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.productgroup+','  ,  ','+ @Inputproductgroups + ','  )) > 0 or (@Inputproductgroups = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and downtime <  @InputDurationLimit" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "end" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "If @InputDurationOper = '>'" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "Begin" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      insert into #DownTimeFinal" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      select *" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      FROM #DownTime as DT" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      WHERE  ((CHARINDEX( ','+DT.Shift+','  ,  ','+ @InputShifts+','  )) > 0 or (@InputShifts = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Team+','  ,  ','+ @InputTeams+','  )) > 0 or (@InputTeams = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat1+','  ,  ','+ @InputCat1s+','  )) > 0 or (@InputCat1s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat2+','  ,  ','+ @InputCat2s+','  )) > 0 or (@InputCat2s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat3+','  ,  ','+ @InputCat3s+','  )) > 0 or (@InputCat3s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.Cat4+','  ,  ','+ @InputCat4s+','  )) > 0 or (@InputCat4s = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.stopclass+','  ,  ','+ @Inputstopclassifications+','  )) > 0 or (@Inputstopclassifications = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.prodstatus+','  ,  ','+ @Inputproductionstatus+','  )) > 0 or (@Inputproductionstatus = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.productcode+','  ,  ','+ @Inputproducts + ','  )) > 0 or (@Inputproducts = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and ((CHARINDEX( ','+DT.productgroup+','  ,  ','+ @Inputproductgroups + ','  )) > 0 or (@Inputproductgroups = 'All'))" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      and downtime >  @InputDurationLimit" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "end" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "IF  @InputOutPutType <> 'Raw Data' " & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "BEGIN" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      INSERT      @ErrorMessages (ErrMsg)" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      VALUES      ('OutPut Type Not Valid=' + @InputOutPutType )" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      GOTO  ErrorMessagesWrite" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "END" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "ErrorMessagesWrite:" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "      SELECT      ErrMsg" & vbNewLine
        SQL_pqq1 = SQL_pqq1 & "            FROM  @ErrorMessages" & vbNewLine




        SQL_pqq2 = SQL_pqq2 & "SELECT StartTime,EndTime,  Downtime,Uptime, MasterProdUnit, Location,Fault,Reason1,Reason2,Reason3,Reason4, StopClass, ProdStatus,Team,Shift,Cat1,Cat2,Cat3,Cat4,Product, ProductCode, ProductGroup, Comments, StartTime_Act, EndTime_act, SBP, EAP, PUID, SourcePUID, tedet_id FROM #DOWNTIMEFINAL  ORDER BY sTARTTIME      " & vbNewLine


        SQL_pqq3 = SQL_pqq3 & "" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "GOTO Finished" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "Finished:" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "DROP TABLE #downtime" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "DROP TABLE #downtimeFinal" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "DROP TABLE #DTInterval" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & "" & vbNewLine
        SQL_pqq3 = SQL_pqq3 & ""
    End Sub

    Public Sub setSQLstrings_PQQProduction(startTime As Date, endTime As Date, masterProdUnit As String)
        SQL_pqqprod1 = SQL_pqqprod1 & "USE [GBDB]" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "SET ANSI_NULLS OFF" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "SET QUOTED_IDENTIFIER OFF" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "DECLARE" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "@InputStartTime         DateTime = '" & Format(startTime, "Short Date") & " " & Format(startTime, "Long Time") & "'," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "@InputEndTime           DateTime = '" & Format(endTime, "Short Date") & " " & Format(endTime, "Long Time") & "'," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "@InputLine			nVarChar(4000) = '" & masterProdUnit & "'," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "@inputproductcode	nVarChar(4000)," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "@InputShifts		nVarChar(4000)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "DECLARE @ProdUnits TABLE(" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	pu_id			int)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "CREATE TABLE #Production (     	" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	StartTime		datetime," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	EndTime			datetime," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	pu_desc			nvarchar(100)," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	ProductCode		nvarchar(100)," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	Product			nvarchar(100)," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	ProdStatus		nvarchar(100)," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	Shift 			nvarchar(25)," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	Team			nvarchar(25)," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	ActUnits		Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	ActCases		Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	AdjCases		Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	AdjUnits		Float, " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	StatUnits		Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	ActualRate		Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	TargetRate		Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	SchedTime		Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	UnitsPerCase	Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	StatCaseConv	Float," & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	puid			int" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & ")" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "DECLARE	@ErrorMessages TABLE(" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	ErrMsg				nVarChar(255) )" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "IF	IsDate(@InputStartTime) <> 1" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "BEGIN" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	INSERT	@ErrorMessages (ErrMsg)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "		VALUES	('StartTime is not a Date.')" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "END" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "IF	IsDate(@inputendTime) <> 1" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "BEGIN" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	INSERT	@ErrorMessages (ErrMsg)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "		VALUES	('EndTime is not a Date.')" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "END" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "IF datepart(Year,@inputstarttime) > 2500" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "BEGIN" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	select @inputstarttime = DATEADD(Year,-543,@inputstarttime)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	select @inputendtime = DATEADD(Year,-543,@inputendtime)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "END" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "INSERT INTO @ProdUnits (pu_id) " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	(SELECT PU_Id" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	FROM dbo.Prod_Units 			pu	WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "		JOIN	dbo.Table_Fields_Values	tfv	WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "											ON	pu.PU_Id = tfv.KeyId" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "		JOIN	dbo.Table_Fields		tf	WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "											ON	tf.Table_Field_Id 	= tfv.Table_Field_Id" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	WHERE	tf.Table_Field_Desc = 'RE-ProductionUnit'" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "		AND	tfv.TableId = (SELECT	TableId	FROM dbo.Tables WITH (NOLOCK) WHERE	TableName = 'Prod_Units')" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "		AND	tfv.Value = '1' " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "		AND pu.pl_id = (SELECT pl_id from dbo.prod_lines WITH (NOLOCK) WHERE pl_desc = @inputLine))" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "INSERT INTO #Production (starttime,endtime, puid)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	SELECT null,e.timestamp  ,PU.Pu_id" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	FROM events e WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	JOIN @ProdUnits PU on PU.pu_id = e.pu_id " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	WHERE e.timestamp>@inputStartTime-1 and e.timestamp<=@inputendTime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "UPDATE prd SET pu_desc = pu.pu_desc " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "FROM #Production prd" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "JOIN PROD_UNITS pu ON pu.pu_id = prd.puid  " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update prd set shift = tt.result" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from #Production prd" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join variables vv on prd.puid = vv.pu_id and vv.extended_info = 'rpthook=shift'" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = prd.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "if @inputshifts<>'all'" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	delete from #Production where charindex(','+shift+',',','+@inputshifts+',')=0" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update prd set StatCaseConv = tt.result" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from #Production prd" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join variables vv on prd.puid = vv.pu_id and vv.extended_info = 'RptHook=StatCaseConvFactor'" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = prd.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update prd set team = tt.result" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from #Production prd" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join variables vv on prd.puid = vv.pu_id and vv.extended_info = 'rpthook=team'" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = prd.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update prd set unitspercase = tt.result" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from #Production prd" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join variables vv on prd.puid = vv.pu_id and vv.extended_info = 'RptHook=UnitsPercase'" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = prd.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set starttime=(select max(p.endtime) from #Production p where p.endtime<#Production.endtime)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "delete from #Production where (starttime <@inputStartTime or starttime is null)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "UPDATE PP SET product=P.PROD_DESC,productcode = P.PROD_CODE" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "FROM #Production PP" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join production_starts ps on ps.pu_id=PP.puid" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "and ps.start_time<=PP.ENDTIME and (ps.end_time>=PP.ENDTIME or ps.end_time is null)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "LEFT join products p on ps.prod_id= p.prod_id  " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "if @inputproductcode<>'all'" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	delete from #Production where charindex(','+productcode+',',','+@inputproductcode+',')=0" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set targetrate = ( select convert(float,tt.result)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex( 'rpthook=TargetRate',vv.extended_info)>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set actualrate = ( select convert(float,tt.result)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex( 'rpthook=ActualRate',vv.extended_info)>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set adjunits= ( select convert(int,convert(float,tt.result))" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex('rpthook=adjustedunits',vv.extended_info)>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set adjcases= ( select convert(int,convert(float,tt.result))" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex('rpthook=adjustedcases',vv.extended_info)>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set actcases= ( select convert(int,convert(float,tt.result))" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex('rpthook=ActualCases',vv.extended_info )>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set actunits= ( select convert(int,convert(float,tt.result))" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex('rpthook=Actualunits',vv.extended_info )>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set schedtime = ( select convert(float,tt.result)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex('rpthook=scheduledtime',vv.extended_info)>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set statunits= ( select convert(int,convert(float,tt.result))" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex( 'rpthook=statunits',vv.extended_info )>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set prodstatus = (select tt.result" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "from variables vv WITH (NOLOCK)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "join tests tt on tt.var_id = vv.var_id and tt.result_on = #Production.endtime" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where vv.pu_id = #Production.puid and charindex( 'RptHook=ProductionStatus',vv.extended_info)>0)" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "update #Production set actualrate = case when convert(float,schedtime)=0 then 0 " & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "	else isnull(convert(float,actunits) / convert(float,schedtime),0) end" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "where #Production.actualrate is null" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod1 = SQL_pqqprod1 & "" & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "SELECT 	StartTime," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	EndTime	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	PU_DESC AS 'Unit'," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	ProductCode	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	Product		," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	ProdStatus	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	Shift 		," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	Team		," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	ActUnits	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	ActCases	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	AdjCases	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	AdjUnits	, " & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	StatUnits	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	ActualRate	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	TargetRate	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	SchedTime	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	UnitsPerCase	," & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & "	StatCaseConv" & vbNewLine
        SQL_pqqprod2 = SQL_pqqprod2 & " FROM #Production" & vbNewLine

        SQL_pqqprod3 = SQL_pqqprod3 & "GOTO Finished" & vbNewLine
        SQL_pqqprod3 = SQL_pqqprod3 & "" & vbNewLine
        SQL_pqqprod3 = SQL_pqqprod3 & "" & vbNewLine

        SQL_pqqprod3 = SQL_pqqprod3 & "Finished:" & vbNewLine
        SQL_pqqprod3 = SQL_pqqprod3 & "" & vbNewLine
        SQL_pqqprod3 = SQL_pqqprod3 & "drop table #Production" & vbNewLine



    End Sub


    Public Function getSQL_oneclick1string_OneClick() As String

        SQL_oneclick1 = "USE [GBDB]" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "Set ANSI_NULLS OFF" & vbNewLine



        SQL_oneclick1 = SQL_oneclick1 & "Set QUOTED_IDENTIFIER OFF" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "Declare " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @InputStartTime                   DateTime = '10/29/2015 06:00:00 AM'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @Inputendtime              DateTime = '11/1/2015 06:00:00 AM'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @InputMasterProdUnit   nVarChar(4000) = 'HCMR011 Main'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @ShowMasterProdUnit        int = '0'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showreason3               int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showreason4               int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showTeam                  int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showshift                 int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showComment               int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showProduct               int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showbrandCode                    int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showProdGroup                    int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showCat1                  int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showCat2                  int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showCat3                  int = '1'," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           @showCat4                  int = '0'" & vbNewLine



        SQL_oneclick1 = SQL_oneclick1 & "SET ANSI_WARNINGS OFF" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SET NOCOUNT ON" & vbNewLine




        SQL_oneclick1 = SQL_oneclick1 & "DECLARE       @PositiON                  int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @InputORderByClause  nvarChar(4000)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @InputGroupByClause  nvarChar(4000)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @strSQL_oneclick1                    VarChar(4000)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @current             datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @tmpStartTime as     datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @tmpendtime as             datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @tmpCount as         int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @tmpLoopCounter      int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @RptProdPUId         int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @ShowLineStatus            int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @CatPrefix           varchar (30)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                @PUID                    int" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "DROP TABLE #Downtime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "DROP TABLE #DTPuIDStartTime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "CREATE TABLE #DownTime(    " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       DT_ID           int IDENTITY (0, 1) NOT NULL ," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       StartTime            datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       endtime                    datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Uptime               Float," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Downtime             Float," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       MasterProdUnit               varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       location             varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       split                varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Fault                varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reason1                    varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reason2                    varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reason3                    varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reason4                    varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Team                 varchar(25)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       shift                varchar(25), " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ishift               INT,  " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "     Cat1                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat2                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat3                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat4                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat5                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat6                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat7                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat8                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat9                 varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cat10                varchar(50)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Product                    varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ProductGroup  varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       brand                varchar(100), " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ibrand               INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       LineStatus           varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Cause_Comment_ID     INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       Comments             varchar(2000)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       StartTime_Act         datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       endtime_Act          datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       endtime_Prev          datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       PUID                 INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SourcePUID           INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       tedet_id             INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reasonID1            INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reasonID2            INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reasonID3            INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       reasonID4            INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ERTD_ID                    int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ErtdID1                    INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ErtdID2                    INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ErtdID3                    INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ErtdID4                    INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SBP                  INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       EAP                  INT," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       TargetSpeed          float," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ActualSpeed          float" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "CREATE TABLE #DTPuIDStartTime(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  RowId int IDENTITY," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  PU_Id int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  StartTime DateTime " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  PRIMARY KEY(PU_Id, StartTime))" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "    " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "DECLARE @schedule_puid TABLE  (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       pu_id                int, " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       schedule_puid        int, " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       tmp1                 int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       tmp2                 int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       info                 varchar(300))" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "DECLARE @TESTS TABLE  (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       var_id               int," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       result               varchar(100)," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       result_ON            datetime," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       extendedinfo  varchar(255))" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "CREATE INDEX  td_PUId_StartTime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON     #DownTime (PUId, StartTime)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "CREATE INDEX  td_PUId_endtime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON     #DownTime (PUId, endtime)" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "DECLARE       @ErrorMessages TABLE (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ErrMsg               nVarChar(255) )" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "IF     ISDate(@InputStartTime) <> 1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       INSERT @ErrorMessages (ErrMsg)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              VALUES ('StartTime IS not a Date.')" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       GOTO   ErrorMessagesWrite" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF     ISDate(@Inputendtime) <> 1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       INSERT @ErrorMessages (ErrMsg)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              VALUES ('endtime IS not a Date.')" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       GOTO   ErrorMessagesWrite" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showreason3=1 AND @showreason4=1 " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2,reason3,reasonID3" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & ",reason4,reasonID4,startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "r2.event_reason_name,r3.event_reason_name,ted.reason_level3,r4.event_reason_name,ted.reason_level4," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID, " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "FROM dbo.timed_event_details AS ted with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r3 with (nolock) ON (r3.event_reason_id = ted.reason_level3)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r4 with (nolock) ON (r4.event_reason_id = ted.reason_level4)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "inner join dbo.prod_units AS pu2 with (nolock) ON (pu2.pu_id = ted.pu_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) )" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All'))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ORDER BY ted.start_Time,  ted.pu_ID" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showreason3=0 AND @showreason4=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "reason4,reasonID4,startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "r2.event_reason_name,r4.event_reason_name,ted.reason_level4," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID, " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "FROM dbo.timed_event_details AS ted with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r4 with (nolock) ON (r4.event_reason_id = ted.reason_level4)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "inner join dbo.prod_units AS pu2 with (nolock) ON (pu2.pu_id = ted.pu_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) )" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All'))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ORDER BY ted.start_Time,  ted.pu_ID" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showreason4=0 AND @showreason3=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "reason3,reasonID3,startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "r2.event_reason_name,r3.event_reason_name,ted.reason_level3," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID, " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "FROM dbo.timed_event_details AS ted with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r3 with (nolock) ON (r3.event_reason_id = ted.reason_level3)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "inner join dbo.prod_units AS pu2 with (nolock) ON (pu2.pu_id = ted.pu_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) )" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All'))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ORDER BY ted.start_Time,  ted.pu_ID" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showreason4=0 AND @showreason3=0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "INSERT INTO #DownTime (StartTime,endtime,MasterProdUnit,Fault,location,PUID,reason1,reason2," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "startTime_act,endtime_act,tedet_id,reasonID1,reasonID2,SourcePUID, ERTD_ID, Cause_Comment_ID)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT ted.start_time, ted.end_time, Pu2.pu_desc, tef.tefault_name, Pu.pu_desc, ted.pu_ID, r1.event_reason_name," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "r2.event_reason_name," & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.start_time, ted.end_time, ted.tedet_id,ted.reason_level1, ted.reason_level2,   ted.Source_PU_ID, " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ted.event_reason_tree_data_id, ted.Cause_Comment_ID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "FROM dbo.timed_event_details AS ted with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r1 with (nolock) ON (r1.event_reason_id = ted.reason_level1)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.event_reasons AS r2 with (nolock) ON (r2.event_reason_id = ted.reason_level2)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "inner join dbo.prod_units AS  pu2 with (nolock) ON (pu2.pu_id = ted.pu_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE  ( ((Start_time < =  @Inputendtime) AND (end_time > @InputStartTime)) OR ((Start_time < =  @Inputendtime )AND end_time IS NULL) )" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All'))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ORDER BY ted.start_Time,  ted.pu_ID" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcomment=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE ted SET" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       comments = REPLACE(coalesce(convert(varchar(2000),co.comment_text),''), char(13)+char(10), ' ')" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "FROM dbo.#Downtime ted with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "left join dbo.Comments co with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ON ted.cause_comment_id = co.comment_id" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET Downtime = datediff(s,starttime,endtime)/60.0" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET Downtime = datediff(s,starttime,@Inputendtime)/60.0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE endtime IS NULL" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "IF @showteam=1 OR @showshift=1" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       INSERT INTO @schedule_puid (pu_id,info) SELECT pu_id,extended_info FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All')" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE @schedule_puid SET tmp1=charindex('scheduleunit=',info)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE @schedule_puid SET tmp2=charindex(';',info,tmp1) WHERE tmp1>0" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE @schedule_puid SET schedule_puid=cast(substring(info,tmp1+13,tmp2-tmp1-13) as int) WHERE tmp1>0 AND tmp2>0 AND not tmp2 IS NULL" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE @schedule_puid SET schedule_puid=cast(substring(info,tmp1+13,len(info)-tmp1-12) as int)WHERE tmp1>0 AND tmp2=0" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "    IF NOT EXISTS(SELECT schedule_puid FROM @schedule_puid)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "         UPDATE @schedule_puid SET schedule_puid=(SELECT   TOP 1 Table_Fields_Values.Value" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "         FROM Table_Fields_Values with (nolock) INNER JOIN  Table_Fields ON Table_Fields_Values.Table_Field_Id =" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "            Table_Fields.Table_Field_Id  WHERE(Table_Fields.Table_Field_Desc = 'ScheduleUnit') " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "         AND (Table_Fields_Values.KeyId in (SELECT  pu_id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "        WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "        WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All')))))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "      END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE @schedule_puid SET schedule_puid=pu_id WHERE schedule_puid IS NULL" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showteam=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET team=( SELECT  crew_desc FROM dbo.crew_schedule cs with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join @schedule_puid sp ON cs.pu_id=sp.schedule_puid" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE #downtime.starttime>=cs.start_time AND cs.END_time>#downtime.starttime AND #downtime.puid=sp.pu_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showshift=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET shift=( SELECT  shift_desc FROM dbo.crew_schedule cs with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join @schedule_puid sp ON cs.pu_id=sp.schedule_puid" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE #downtime.starttime>=cs.start_time AND cs.END_time>#downtime.starttime AND #downtime.puid=sp.pu_id)" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showproduct=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #DownTime SET product=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT p.Prod_Desc FROM dbo.products p with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.production_starts ps with (nolock) ON ps.prod_id= p.prod_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.prod_units with (nolock) ON ps.pu_id=prod_units.pu_id WHERE ps.start_time <= #downtime.starttime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ((#downtime.starttime < ps.END_time) OR (ps.END_time IS NULL)) AND ps.pu_id=#downtime.puid)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showbrandcode=1 or @showbrandcode=2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #DownTime SET brand=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT p.prod_code FROM dbo.products p with (nolock) join dbo.production_starts ps with (nolock) ON ps.prod_id= p.prod_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.prod_units with (nolock) ON ps.pu_id=prod_units.pu_id WHERE ps.start_time <= #downtime.starttime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ((#downtime.starttime < ps.END_time) OR (ps.END_time IS NULL)) AND ps.pu_id=#downtime.puid)" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #DownTime SET ProductGroup = (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT TOP 1 product_grp_desc FROM product_groups pg" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join product_group_data pgd ON pgd.Product_Grp_Id = pg.Product_Grp_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join products p ON pgd.Prod_Id = p.Prod_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join comments c ON pg.comment_id = c.comment_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE p.prod_code = #downtime.brand" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND c.comment_text Like '%Package Size%')" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #DownTime SET ProductGroup = (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT TOP 1 product_grp_desc FROM product_groups pg" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join product_group_data pgd ON pgd.Product_Grp_Id = pg.Product_Grp_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join products p ON pgd.Prod_Id = p.Prod_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE p.prod_code = #downtime.brand)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE ProductGroup is null" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF (SELECT TOP 1 pu_id FROM dbo.prod_units pu with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.local_pg_line_status ls with (nolock) ON pu.pu_id=ls.unit_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                  WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All'))) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "  INSERT into #DTPuIDStartTime(PU_Id, StartTime)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "   SELECT DISTINCT PUID, StartTime " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "   FROM #DOWNTIME" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "   ORDER BY PUID ASC" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #DownTime SET LineStatus=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "            SELECT p.phrase_value FROM dbo.phrase p with (nolock) JOIN dbo.local_pg_line_status ls with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "            ON ls.line_status_id = p.phrase_id JOIN dbo.prod_units pu with (nolock) ON pu.pu_id=ls.unit_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "            join #DTPuIDStartTime pt ON pt.PU_Id=#DownTime.PUID and pt.StartTime = #DownTime.starttime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "        WHERE ls.start_datetime <= #DownTime.starttime AND ((#DownTime.starttime < ls.end_datetime) OR " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              (ls.end_datetime IS NULL)) and ls.unit_id = pt.PU_Id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF (SELECT TOP 1 pu_id FROM dbo.prod_units pu with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.local_pg_line_status ls with (nolock) ON pu.pu_id=ls.unit_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                  WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All')))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT @RptProdPUId = (SELECT TOP 1 pu_id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                  WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All'))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND charindex('production=true', extended_info) > 0)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @RptProdPUId = null" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  SELECT @RptProdPUId = (SELECT TOP 1  Prod_Units.PU_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  FROM Table_Fields_Values with (nolock) INNER JOIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "     Table_Fields ON Table_Fields_Values.Table_Field_Id = Table_Fields.Table_Field_Id INNER JOIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "     Prod_Units ON Table_Fields_Values.KeyId = Prod_Units.PU_Id  WHERE(Table_Fields.Table_Field_Desc = 'RE-ProductionUnit') " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  AND (Table_Fields_Values.KeyId in (SELECT  pu_id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "  WHERE PL_ID = (SELECT DISTINCT PL_Id FROM dbo.prod_units with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                            WHERE (charindex(','+pu_desc+',',','+@inputmasterprodunit+',')>0 OR @inputmasterprodunit='All')))))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "INSERT @TESTS (var_id, extendedinfo, result_ON, result)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT vv.var_id, vv.extended_info, result_ON, result FROM dbo.tests tt with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "JOIN dbo.variables vv with (nolock) ON vv.var_id = tt.var_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND (charindex('rpthook=productionstatus', vv.extended_info)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "JOIN dbo.prod_units pu with (nolock) ON pu.pu_id = vv.pu_id AND pu.pu_id = @RPTProdPUId" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE tt.result_ON <= @inputendtime+1 AND tt.result_ON > @inputstarttime-1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET LineStatus = (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT TOP 1 result FROM @tests tt " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.prod_units pu with (nolock) ON pu.pu_id = @RPTProdPUId" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE charindex('rpthook=productionstatus', tt.extendedinfo)>0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND tt.result_ON > #downtime.starttime " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ORDER BY tt.result_ON)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat1=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET ErtdID1=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT  ertd.Event_reason_tree_data_id FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "dbo.Prod_events pe with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE ertd.event_reason_level=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd.event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND pe.Event_type = 2)WHERE #downtime.reasonid1 IS NOT NULL" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat2=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET ErtdID2=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT  ertd.Event_reason_tree_data_id FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "dbo.Prod_events pe with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE ertd.event_reason_level=2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd.event_reason_id=#downtime.reasonid2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd.Parent_Event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND pe.Event_type = 2)WHERE #downtime.reasonid2 IS NOT NULL" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat3=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET ErtdID3=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT  ertd.Event_reason_tree_data_id FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "dbo.Prod_events pe with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.event_reason_tree_data ertd1 with (nolock) ON ertd1.Event_reason_tree_data_id = ertd.Parent_Event_R_Tree_Data_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE ertd.event_reason_level=3" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd.event_reason_id=#downtime.reasonid3 " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd.Parent_Event_reason_id=#downtime.reasonid2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd1.Parent_Event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND pe.Event_type = 2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & ")WHERE #downtime.reasonid3 IS NOT NULL" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat4=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET ErtdID4=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT  ertd.Event_reason_tree_data_id FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "dbo.Prod_events pe with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.event_reason_tree_data ertd with (nolock) ON ertd.tree_name_id=pe.Name_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.event_reason_tree_data ertd1 with (nolock) ON ertd1.Event_reason_tree_data_id = ertd.Parent_Event_R_Tree_Data_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.event_reason_tree_data ertd2 with (nolock) ON ertd2.Event_reason_tree_data_id = ertd1.Parent_Event_R_Tree_Data_Id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE ertd.event_reason_level=4" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND #downtime.SourcePUID=pe.pu_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd.event_reason_id=#downtime.reasonid4" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd.Parent_Event_reason_id=#downtime.reasonid3" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd1.Parent_Event_reason_id=#downtime.reasonid2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ertd2.Parent_Event_reason_id=#downtime.reasonid1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND pe.Event_type = 2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & ")WHERE #downtime.reasonid4 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF (SELECT count(erc_id) FROM dbo.event_reason_catagories with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE charindex('DTSched', erc_desc) > 0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "OR charindex('DTGroup', erc_desc) > 0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "OR charindex('DTMach', erc_desc) > 0) > 5" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat1=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @CatPrefix='DTSched-'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat1=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID4 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat1=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID3 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat1=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID2 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat1=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat1 IS NULL AND #downtime.ErtdID1 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat2=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @CatPrefix='DTGroup-'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat2=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID4 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat2=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID3 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat2=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID2 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat2=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat2 IS NULL AND #downtime.ErtdID1 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat3=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @CatPrefix='DTMach-'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat3=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID4 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat3=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID3 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat3=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID2 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat3=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat3 IS NULL AND #downtime.ErtdID1 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat4=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @CatPrefix='DTType-'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat4=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID4 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID4 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat4=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID3 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID3 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat4=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID2 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID2 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE #downtime SET Cat4=(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT  TOP 1 right(erc_desc,len(erc_desc)-len(@CatPrefix)) FROM " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       dbo.event_reason_catagories erc with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       join dbo.event_reason_category_data ercd with (nolock) ON ercd.erc_id =  erc.erc_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE ercd.Event_reason_tree_data_id = #downtime.ErtdID1 AND (charindex(lower(@CatPrefix),erc_desc)>0) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       AND ercd.Propegated_FROM_etDid IS NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       )WHERE #downtime.Cat4 IS NULL AND #downtime.ErtdID1 IS NOT NULL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF (SELECT count(erc_id) FROM dbo.event_reason_catagories with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE charindex('category:', erc_desc) > 0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "OR charindex('Schedule:', erc_desc) > 0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "OR charindex('Subsystem:', erc_desc) > 0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "OR charindex('GroupCause:', erc_desc) > 0) > 5" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat1=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @CatPrefix='category:'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE td SET" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              Cat1 = right(erc_desc,len(erc_desc)-len(@CatPrefix))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       FROM dbo.#downtime td with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON ercd.erc_id = erc.erc_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE erc.ERC_Desc LIKE @CatPrefix + '%'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat2=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "   SELECT @CatPrefix='Schedule:'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE td SET" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              Cat2 = right(erc_desc,len(erc_desc)-len(@CatPrefix))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       FROM dbo.#downtime td with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON ercd.erc_id = erc.erc_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       where erc.ERC_Desc LIKE @CatPrefix + '%'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat3=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @CatPrefix='Subsystem:'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE td SET" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              Cat3 = right(erc_desc,len(erc_desc)-len(@CatPrefix))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       FROM dbo.#downtime td with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON ercd.erc_id = erc.erc_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       where erc.ERC_Desc LIKE @CatPrefix + '%'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "IF @showcat4=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       BEGIN" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @CatPrefix='GroupCause:'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       UPDATE td SET" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              Cat4 = right(erc_desc,len(erc_desc)-len(@CatPrefix))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       FROM dbo.#downtime td with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_category_data ercd WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON TD.ERTD_ID = ercd.event_reason_tree_data_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       JOIN dbo.event_reason_catagories erc WITH (NOLOCK) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       ON ercd.erc_id = erc.erc_id " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE erc.ERC_Desc LIKE @CatPrefix + '%'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "       END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET uptime=datedIFf(s,(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT TOP 1 dt2.endtime FROM dbo.#downtime dt2 with (nolock) WHERE ((dt2.DT_ID=(#downtime.DT_ID - 1)))),starttime)/60.0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE #downtime.DT_ID > 0" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET uptime=datedIFf(s,(" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SELECT max(dt2.END_time) FROM dbo.timed_event_details dt2 with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "join dbo.prod_units pu2 with (nolock) ON pu2.pu_id=dt2.pu_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE  END_time<=#downtime.starttime " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND ((CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  )) > 0 OR (@InputMasterProdUnit = 'All'))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "),starttime)/60.0 WHERE uptime IS NULL" & vbNewLine



        SQL_oneclick1 = SQL_oneclick1 & "declare @preEND datetime" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "SELECT @preEND = max(END_time)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "FROM dbo.timed_event_details ted with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT JOIN dbo.timed_event_fault AS tef with (nolock) ON (tef.tefault_id = ted.tefault_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "LEFT Join dbo.prod_units AS pu with (nolock) ON (pu.pu_id = ted.source_PU_Id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "inner join dbo.prod_units as pu2 with (nolock) ON (pu2.pu_id = ted.pu_id)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE  (END_time <= @InputStartTime) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "AND (CHARINDEX( ','+pu2.pu_desc+','  ,  ','+ @InputMasterProdUnit+','  ) > 0 OR (@InputMasterProdUnit = 'All'))" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #downtime SET uptime=datedIFf(s,@preEND,starttime)/60.0 " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE uptime IS NULL" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "DECLARE " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                     @fltDBVersion        FLOAT" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF     (      SELECT        IsNumeric(App_Version)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                     FROM   dbo.AppVersions with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                     WHERE  App_Id = 2) = 1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT        @fltDBVersion = Convert(Float, App_Version)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              FROM   dbo.AppVersions with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              WHERE  App_Id = 2" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ELSE" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @fltDBVersion = 1.0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "DECLARE " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              @DowntimesystemUserID             INT" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "Select @DowntimesystemUserID = " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       User_ID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       FROM dbo.USERS with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       WHERE UserName = 'ReliabilitySystem'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF     @fltDBVersion <= 300172.90 " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              UPDATE #Downtime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           Set Split = 'S'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              FROM dbo.#Downtime tdt with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              Join dbo.Timed_Event_Detail_History ted with (nolock) on tdt.TeDet_Id = ted.TEDET_ID " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              and ted.User_ID = '2' " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              WHERE ted.TEDET_ID IS NOT Null" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                       And tdt.Split IS Null" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "ELSE " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN                      " & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "              UPDATE #Downtime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                           Set Split = ISNULL((Select Case     WHEN Min(User_Id) < 50" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                                                               THEN       ''" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                                                               WHEN       Min(User_Id) = @DowntimesystemUserID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                                                               THEN       ''" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                                                               ELSE       'S'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                                            END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                                            FROM dbo.Timed_Event_Detail_History with (nolock) " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                                            WHERE Tedet_Id = ted.TEDET_ID),'S') " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              FROM dbo.#Downtime tdt with (nolock) Left Join dbo.Timed_Event_Detail_History ted " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                    with (nolock) on tdt.TEDet_ID = ted.TEDET_ID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              Where tdt.Uptime = 0" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                     UPDATE #Downtime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                  Set Split = NuLL" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                     FROM dbo.#Downtime tdt with (nolock)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                     Join dbo.Timed_Event_Detail_History tedh with (nolock) on tdt.TEDet_ID = tedh.TEDET_ID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                     WHERE tedh.User_ID = @DowntimeSystemUserID" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                                         And tdt.Downtime <> 0" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine



        SQL_oneclick1 = SQL_oneclick1 & "IF EXISTS (select * from dbo.sysobjects where id = object_id(N'[dbo].[fnLocal_GlblParseInfo]'))" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "BEGIN" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "      IF EXISTS (SELECT PU_ID from prod_units where charindex('RateLoss', extended_info) > 0)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "      BEGIN" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "            UPDATE #downtime SET targetspeed = (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                  SELECT result FROM tests tt" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                  join variables vv ON vv.var_id = tt.var_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                        and vv.pu_id = #downtime.puid" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "                        and dbo.fnLocal_GlblParseInfo(vv.Extended_Info, 'GlblDesc=') LIKE '%' + replace('Line Target Speed',' ','')" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                  WHERE tt.result_on = #downtime.starttime)  " & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "            UPDATE #downtime SET actualspeed = (" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                  SELECT result FROM tests tt" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                  join variables vv ON vv.var_id = tt.var_id" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                        and vv.pu_id = #downtime.puid" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "                        and dbo.fnLocal_GlblParseInfo(vv.Extended_Info, 'GlblDesc=') LIKE '%' + replace('Line Actual Speed',' ','')" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "                  WHERE tt.result_on = #downtime.starttime)  " & vbNewLine



        SQL_oneclick1 = SQL_oneclick1 & "           UPDATE #downtime SET downtime = (targetspeed - actualspeed) * downtime / targetspeed--," & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "                  WHERE targetspeed is not null " & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "      END" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "END" & vbNewLine




        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @ShowLineStatus = 1" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #DownTime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SET iBrand=cast(Brand as INT)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE (IsNumeric(Brand)=1 AND LEN(Brand)=8)" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "UPDATE #DownTime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "SET iShift=cast(Shift as INT)" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "WHERE IsNumeric(Shift)=1" & vbNewLine



        SQL_oneclick1 = SQL_oneclick1 & "SELECT @strSQL_oneclick1='SELECT starttime,endtime,downtime,uptime,location,fault,reason1,reason2,split'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "IF @showreason3      =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1='SELECT starttime,endtime,downtime,uptime,location,fault,reason1,reason2,reason3,split'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showreason4      =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1='SELECT starttime,endtime,downtime,uptime,location,fault,reason1,reason2,reason3,reason4,split'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showmasterprodunit =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+',MasterProdUnit'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showTeam  =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+',team'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showshift =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',ishift'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showshift =2                       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',shift' " & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "IF @showProduct      =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',product'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showbrandCode=1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',ibrand' " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showbrandCode=2                       " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',brand' " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showProdGroup = 1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',ProductGroup'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showCat1  =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',cat1'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showCat2  =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',Cat2'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showCat3  =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',cat3'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showCat4  =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',cat4'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @ShowLineStatus   =1" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',LineStatus'" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "IF @showComment      =1 " & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT @strSQL_oneclick1=@strSQL_oneclick1+ ',comments' " & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "SELECT @strSQL_oneclick1=@strSQL_oneclick1+' FROM dbo.#downtime with (nolock) ORder by starttime'" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "print @strSQL_oneclick1" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "exec (@strSQL_oneclick1)" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "GOTO FinIShed" & vbNewLine


        SQL_oneclick1 = SQL_oneclick1 & "ErrorMessagesWrite:" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "       SELECT ErrMsg" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "              FROM   @ErrorMessages" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "Finished:" & vbNewLine

        SQL_oneclick1 = SQL_oneclick1 & "DROP TABLE #Downtime" & vbNewLine
        SQL_oneclick1 = SQL_oneclick1 & "DROP TABLE #DTPuIDStartTime" & vbNewLine

        Return SQL_oneclick1

    End Function



End Module
