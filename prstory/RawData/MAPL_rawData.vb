Module MAPLE_Data
    Public Function getMapleData(ByVal endTime As Date, ByVal startTime As Date, ByVal lineName As String, ByVal masterProdUnit As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String, ByVal databaseName As String) As Array
        Dim poCMD As New ADODB.Command
        Dim psConnString As String
        'Check to make sure that we have a database connection before attempting to execute the stored procedure
        Dim gobjConnA = New ADODB.Connection
        If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
            'psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=RAKMES2; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"
            psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=" & databaseName & "; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"
            If psServer = "?" Or psUID = "?" Or psPassword = "?" Then
                'frmOptions.Show vbModal
            Else
                If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
                    gobjConnA.ConnectionString = psConnString
                    gobjConnA.CommandTimeout = 30
                    gobjConnA.Open()
                End If
            End If
        End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnA
            .CommandText = "One_Click_RetrieveDTData_v001"
            .Parameters.Refresh()
            .Parameters(1).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(3).Value = lineName
            .Parameters(4).Value = masterProdUnit
            .Parameters(5).Value = "0"
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
            .Parameters(16).Value = "1"
            .Parameters(17).Value = "0"

            Try
                getMapleData = .Execute.GetRows() 'ORIGINAL CODE
            Catch ex As Exception
                Throw
            End Try

        End With
        If gobjConnA.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.Close()
        End If
        gobjConnA = Nothing
    End Function

    Public Function getMaplePRODData(ByVal endTime As Date, ByVal startTime As Date, ByVal lineName As String, ByVal masterProdUnit As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String, ByVal databaseName As String) As Array
        Dim poCMD As New ADODB.Command
        Dim psConnString As String
        'Check to make sure that we have a database connection before attempting to execute the stored procedure
        Dim gobjConnA = New ADODB.Connection
        If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
            'psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=RAKMES2; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"
            psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=" & databaseName & "; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"
            If psServer = "?" Or psUID = "?" Or psPassword = "?" Then
                'frmOptions.Show vbModal
            Else
                If gobjConnA.State <> ADODB.ObjectStateEnum.adStateOpen Then
                    gobjConnA.ConnectionString = psConnString
                    gobjConnA.CommandTimeout = 30
                    gobjConnA.Open()
                End If
            End If
        End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnA
            .CommandText = "One_Click_RetrievePDData_v001"
            .Parameters.Refresh()
            .Parameters(1).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(2).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(3).Value = lineName


            Try
                getMaplePRODData = .Execute.GetRows() 'ORIGINAL CODE
            Catch ex As Exception
                Throw
            End Try

        End With
        If gobjConnA.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnA.Close()
        End If
        gobjConnA = Nothing
        If IsNothing(getMaplePRODData) Then
            ' Console.Write("! ")
        End If
    End Function

End Module

Module GLEDS_Data
    Public Function getGLEDSData(ByVal endTime As Date, ByVal startTime As Date, ByVal lineName As String, ByVal psServer As String, ByVal psUID As String, ByVal psPassword As String) As Array
        Dim poCMD As New ADODB.Command
        Dim psConnString As String
        'Check to make sure that we have a database connection before attempting to execute the stored procedure
        Dim gobjConnG = New ADODB.Connection
        If gobjConnG.State <> ADODB.ObjectStateEnum.adStateOpen Then
            'psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=LSUD; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; network=dbmssocn"

            psConnString = "Driver={SQL Server}; server=" & Trim(psServer) & "; database=GLEDS_DATA; UID=" & Trim(psUID) & "; PWD=" & Trim(psPassword) & "; QuotedID=No"
            If psServer = "?" Or psUID = "?" Or psPassword = "?" Then
                'frmOptions.Show vbModal
            Else
                If gobjConnG.State <> ADODB.ObjectStateEnum.adStateOpen Then
                    gobjConnG.ConnectionString = psConnString
                    gobjConnG.CommandTimeout = 30
                    gobjConnG.Open()
                End If
            End If
        End If

        With poCMD
            .CommandTimeout = 600
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .ActiveConnection = gobjConnG
            .CommandText = "GDI_getDowntimeGrouped_PR_Story_V1" '"GDI_getDowntimeGrouped_OneClick_V1" ' '"GDI_getDowntimeGrouped_OneClick_V1" '"GDI_getDowntimeGrouped_oneclick_test"
            .Parameters.Refresh()
            .Parameters(1).Value = "@ALL"
            .Parameters(2).Value = "ENG"
            .Parameters(3).Value = Format(startTime, "Short Date") & " " & Format(startTime, "Long Time")
            .Parameters(4).Value = Format(endTime, "Short Date") & " " & Format(endTime, "Long Time")
            .Parameters(5).Value = lineName
            '.Parameters(8).Value = 1
            .Parameters(9).Value = "D"
            .Parameters(10).Value = 1
            .Parameters(12).Value = 0
            .Parameters(13).Value = "1"
            .Parameters(15).Value = 2
            .Parameters(16).Value = ""

            Try
                getGLEDSData = .Execute.GetRows()
            Catch ex As Exception
                Throw
            End Try

        End With
        If gobjConnG.State = ADODB.ObjectStateEnum.adStateOpen Then
            gobjConnG.Close()
        End If
        gobjConnG = Nothing
    End Function


End Module