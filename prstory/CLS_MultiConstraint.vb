Imports System.Windows.Forms

Module CLS_MultiConstraint
    Private Enum DownTimeColumn
        StartTime = 0
        Endtime = 1
        DT = 2
        UT = 3
        MasterProdUnit = 4
        Location = 5
        Reason1 = 7
        PR_InOut = 12
        PlannedUnplanned = 15
        DTGroup = 16
        Max = 29
    End Enum
    Private Enum DownTimeColumnOneClick
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
        Product = 13 '19
        ProductCode = 14 '20
        PlannedUnplanned = 16 '15
        DTGroup = 17 '16
        PR_InOut = 19 '12
        Comment = 20 '22
        Max = 20
    End Enum



    Public Function PROF_mergeRateLossWithMain_Cairo_OneClick(ByVal rateLoss(,) As Object, ByRef rawDT(,) As Object)
        Dim isRateDone As Boolean = False, isRawDone As Boolean = False
        Dim nRateLoss As Long, nRawDT As Long, netEvents As Long
        Dim rawIncrementer As Long = 0, rateIncrementer As Long = 0, finalIncrementer = 0
        Dim rateTime As Date, rawTime As Date
        nRateLoss = rateLoss.GetLength(1)
        nRawDT = rawDT.GetLength(1)
        netEvents = nRateLoss + nRawDT - 1
        'merge the data sets
        Dim finalData(DownTimeColumnOneClick.Max, netEvents) As Object
        While finalIncrementer <= netEvents
            'see if either is maxed out...
            If isRawDone Then
                rateTime = New Date(0)
                rawTime = New Date(1)
            ElseIf isRateDone Then
                rateTime = New Date(1)
                rawTime = New Date(0)
            Else
                rateTime = rateLoss(DownTimeColumnOneClick.StartTime, rateIncrementer)
                rawTime = rawDT(DownTimeColumnOneClick.StartTime, rawIncrementer)
            End If

            'get down to business
            If rawTime < rateTime Then
                For i As Integer = 0 To DownTimeColumnOneClick.Max
                    finalData(i, finalIncrementer) = rawDT(i, rawIncrementer)
                Next
                rawIncrementer += 1
                If rawIncrementer = rawDT.GetLength(1) Then isRawDone = True
            Else
                For i As Integer = 0 To DownTimeColumnOneClick.Max
                    finalData(i, finalIncrementer) = rateLoss(i, rateIncrementer)
                    '  finalData(DownTimeColumnOneClick.UT, finalIncrementer) = 0
                Next
                rateIncrementer += 1
                If rateIncrementer = rateLoss.GetLength(1) Then isRateDone = True
            End If
            finalIncrementer += 1
        End While

        Return finalData
    End Function


    Public Function PROF_mergeRateLossWithMain_Cairo(ByVal rateLoss(,) As Object, ByRef rawDT(,) As Object)
        Dim isRateDone As Boolean = False, isRawDone As Boolean = False
        Dim nRateLoss As Long, nRawDT As Long, netEvents As Long
        Dim rawIncrementer As Long = 0, rateIncrementer As Long = 0, finalIncrementer = 0
        Dim rateTime As Date, rawTime As Date
        nRateLoss = rateLoss.GetLength(1)
        nRawDT = rawDT.GetLength(1)
        netEvents = nRateLoss + nRawDT - 1
        'merge the data sets
        Dim finalData(DownTimeColumn.Max, netEvents) As Object
        While finalIncrementer <= netEvents
            'see if either is maxed out...
            If isRawDone Then
                rateTime = New Date(0)
                rawTime = New Date(1)
            ElseIf isRateDone Then
                rateTime = New Date(1)
                rawTime = New Date(0)
            Else
                rateTime = rateLoss(DownTimeColumn.StartTime, rateIncrementer)
                rawTime = rawDT(DownTimeColumn.StartTime, rawIncrementer)
            End If

            'get down to business
            If rawTime < rateTime Then
                For i As Integer = 0 To DownTimeColumn.Max
                    finalData(i, finalIncrementer) = rawDT(i, rawIncrementer)
                Next
                rawIncrementer += 1
                If rawIncrementer = rawDT.GetLength(1) Then isRawDone = True
            Else
                For i As Integer = 0 To DownTimeColumn.Max
                    finalData(i, finalIncrementer) = rateLoss(i, rateIncrementer)
                    '  finalData(DownTimeColumn.UT, finalIncrementer) = 0
                Next
                rateIncrementer += 1
                If rateIncrementer = rateLoss.GetLength(1) Then isRateDone = True
            End If
            finalIncrementer += 1
        End While

        Return finalData
    End Function




    Public Function PROF_mergeRateLossWithMain(ByVal rateLoss(,) As Object, ByRef rawDT(,) As Object)
        Dim isRateDone As Boolean = False, isRawDone As Boolean = False, dtCUT As Double
        Dim nRateLoss As Long, nRawDT As Long, netEvents As Long
        Dim rawIncrementer As Long = 0, rateIncrementer As Long = 0, finalIncrementer = 0
        Dim rateTime As Date, rawTime As Date
        nRateLoss = rateLoss.GetLength(1)
        nRawDT = rawDT.GetLength(1)
        netEvents = nRateLoss + nRawDT - 1
        'merge the data sets
        Dim finalData(DownTimeColumn.Max, netEvents) As Object
        While finalIncrementer <= netEvents
            'see if either is maxed out...
            If isRawDone Then
                rateTime = New Date(0)
                rawTime = New Date(1)
            ElseIf isRateDone Then
                rateTime = New Date(1)
                rawTime = New Date(0)
            Else
                rateTime = rateLoss(DownTimeColumn.StartTime, rateIncrementer)
                rawTime = rawDT(DownTimeColumn.StartTime, rawIncrementer)
            End If

            'get down to business
            If rawTime < rateTime Then
                For i As Integer = 0 To DownTimeColumn.Max
                    finalData(i, finalIncrementer) = rawDT(i, rawIncrementer)
                Next
                rawIncrementer += 1
                If rawIncrementer = rawDT.GetLength(1) Then isRawDone = True
            Else
                For i As Integer = 0 To DownTimeColumn.Max
                    finalData(i, finalIncrementer) = rateLoss(i, rateIncrementer)
                    '  finalData(DownTimeColumn.UT, finalIncrementer) = 0
                Next
                rateIncrementer += 1
                If rateIncrementer = rateLoss.GetLength(1) Then isRateDone = True
            End If
            finalIncrementer += 1
        End While


        'adjust the uptimes and downtimes
        finalIncrementer = 0
        While finalIncrementer <= finalData.GetLength(1) - 1 'for each event...
            If InStr(finalData(DownTimeColumn.MasterProdUnit, finalIncrementer), "rate", CompareMethod.Text) > 0 Or InStr(finalData(DownTimeColumn.MasterProdUnit, finalIncrementer), "Perdida de velocidad", CompareMethod.Text) > 0 Then 'if its a rate loss event...
                dtCUT = finalData(DownTimeColumn.DT, finalIncrementer) / 2 'first cut its downtime in half
                finalData(DownTimeColumn.DT, finalIncrementer) -= dtCUT 'first cut its downtime in half
                If IsDBNull(finalData(DownTimeColumn.Endtime, finalIncrementer)) Then  'LG code to handle DBnull
                    finalData(DownTimeColumn.Endtime, finalIncrementer) = finalData(DownTimeColumn.StartTime, finalIncrementer) ' LG code
                Else  'LG code
                    finalData(DownTimeColumn.Endtime, finalIncrementer) = DateAdd(DateInterval.Minute, -dtCUT, finalData(DownTimeColumn.Endtime, finalIncrementer)) 'LG code
                End If
                If finalIncrementer > 0 Then
                    If IsDBNull(finalData(DownTimeColumn.Endtime, finalIncrementer - 1)) Then finalData(DownTimeColumn.Endtime, finalIncrementer - 1) = finalData(DownTimeColumn.StartTime, finalIncrementer - 1)
                End If
                If finalIncrementer > 0 Then finalData(DownTimeColumn.UT, finalIncrementer) = DateDiff(DateInterval.Second, finalData(DownTimeColumn.Endtime, finalIncrementer - 1), finalData(DownTimeColumn.StartTime, finalIncrementer)) / 60 'adjust the prior uptime
                If Not finalIncrementer = finalData.GetLength(1) - 1 Then 'make sure we're not at the end
                    If Not InStr(finalData(DownTimeColumn.MasterProdUnit, finalIncrementer + 1), "rate", CompareMethod.Text) > 0 Or InStr(finalData(DownTimeColumn.MasterProdUnit, finalIncrementer), "Perdida de velocidad", CompareMethod.Text) > 0 Then 'if the next one is not rate loss then we need to fix its uptime too
                        finalIncrementer += 1
                        finalData(DownTimeColumn.UT, finalIncrementer) = DateDiff(DateInterval.Second, finalData(DownTimeColumn.Endtime, finalIncrementer - 1), finalData(DownTimeColumn.StartTime, finalIncrementer)) / 60
                    End If
                End If
            End If
            finalIncrementer += 1
        End While


        Return finalData
    End Function




    Public Function PROF_mergeRateLossWithMain_OneClick(ByVal rateLoss(,) As Object, ByRef rawDT(,) As Object)
        '    Dim finalRateLossData(,) As object
        Dim FamilyNullLocationList As New List(Of Date)

        Dim SecondaryUnitUp As String

        If AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyMaking Then
            SecondaryUnitUp = "Sheetbreak"
        Else
            SecondaryUnitUp = "Blocked/Starved"
        End If

        Dim isRateDone As Boolean = False, isRawDone As Boolean = False, dtCUT As Double
        Dim nRateLoss As Long, nRawDT As Long, netEvents As Long
        Dim rawIncrementer As Long = 0, rateIncrementer As Long = 0, finalIncrementer = 0
        Dim rateTime As Date, rawTime As Date
        nRawDT = rawDT.GetLength(1)

        If IsNothing(rateLoss) Then
            netEvents = nRawDT - 1

        Else
            nRateLoss = rateLoss.GetLength(1)
            netEvents = nRateLoss + nRawDT - 1

        End If

        'merge the data sets
        Dim finalData(DownTimeColumnOneClick.Max, netEvents) As Object
        While finalIncrementer <= netEvents
            'see if either is maxed out...
            If isRawDone Then
                rateTime = New Date(0)
                rawTime = New Date(1)
            ElseIf isRateDone Then
                rateTime = New Date(1)
                rawTime = New Date(0)
            Else

                rawTime = rawDT(DownTimeColumnOneClick.StartTime, rawIncrementer)
                If IsNothing(rateLoss) Then
                    rateTime = rawTime.AddDays(1)  'forcing rate time to be always greater than rawtime so that if loop only allows for addition of rawdata to finaldata.
                Else
                    rateTime = rateLoss(DownTimeColumnOneClick.StartTime, rateIncrementer)

                End If




            End If

            'get down to business
            If rawTime < rateTime Then
                For i As Integer = 0 To DownTimeColumnOneClick.Max
                    finalData(i, finalIncrementer) = rawDT(i, rawIncrementer)
                Next
                rawIncrementer += 1
                If rawIncrementer = rawDT.GetLength(1) Then isRawDone = True
            Else
                If Not IsNothing(rateLoss) Then  'checks for null records for rate loss
                    For i As Integer = 0 To DownTimeColumnOneClick.Max
                        finalData(i, finalIncrementer) = rateLoss(i, rateIncrementer)
                        '  finalData(DowntimeColumnOneClick.UT, finalIncrementer) = 0
                    Next
                    rateIncrementer += 1
                    If rateIncrementer = rateLoss.GetLength(1) Then isRateDone = True
                End If
            End If
            finalIncrementer += 1
        End While

        'FAMILY CARE null location check
        If AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.FamilyMaking Then
            finalIncrementer = 0
            While finalIncrementer < finalData.GetLength(1)
                If IsDBNull(finalData(DownTimeColumnOneClick.Location, finalIncrementer)) Then
                    finalData(DownTimeColumnOneClick.Location, finalIncrementer) = BLANK_INDICATOR
                    If Not IsDBNull(finalData(DownTimeColumnOneClick.StartTime, finalIncrementer)) Then
                        FamilyNullLocationList.Add(finalData(DownTimeColumnOneClick.StartTime, finalIncrementer))
                    End If
                End If
                finalIncrementer = finalIncrementer + 1
            End While
        End If

        'adjust the uptimes and downtimes
        finalIncrementer = 0
        If IsNothing(rateLoss) Then finalIncrementer = finalData.GetLength(1) + 3
        While finalIncrementer <= finalData.GetLength(1) - 1
            Try
                If InStr(finalData(DownTimeColumnOneClick.Location, finalIncrementer), SecondaryUnitUp, CompareMethod.Text) > 0 Then 'if its a rate loss event...
                    dtCUT = 0
                    finalData(DownTimeColumnOneClick.DT, finalIncrementer) -= dtCUT
                    If IsDBNull(finalData(DownTimeColumnOneClick.Endtime, finalIncrementer)) Then
                        finalData(DownTimeColumnOneClick.Endtime, finalIncrementer) = finalData(DownTimeColumnOneClick.StartTime, finalIncrementer) ' LG code
                    Else
                        finalData(DownTimeColumnOneClick.Endtime, finalIncrementer) = DateAdd(DateInterval.Minute, -dtCUT, finalData(DownTimeColumnOneClick.Endtime, finalIncrementer)) 'LG code
                    End If
                    If finalIncrementer > 0 Then
                        If IsDBNull(finalData(DownTimeColumnOneClick.Endtime, finalIncrementer - 1)) Then finalData(DownTimeColumnOneClick.Endtime, finalIncrementer - 1) = finalData(DownTimeColumnOneClick.StartTime, finalIncrementer - 1)
                    End If
                    If finalIncrementer > 0 Then finalData(DownTimeColumnOneClick.UT, finalIncrementer) = DateDiff(DateInterval.Second, finalData(DownTimeColumnOneClick.Endtime, finalIncrementer - 1), finalData(DownTimeColumnOneClick.StartTime, finalIncrementer)) / 60 'adjust the prior uptime
                    If Not finalIncrementer = finalData.GetLength(1) - 1 Then 'make sure we're not at the end

                        If Not InStr(finalData(DownTimeColumnOneClick.Location, finalIncrementer + 1), SecondaryUnitUp, CompareMethod.Text) > 0 Then 'if the next one is not rate loss then we need to fix its uptime too
                            finalIncrementer += 1
                            finalData(DownTimeColumnOneClick.UT, finalIncrementer) = DateDiff(DateInterval.Second, finalData(DownTimeColumnOneClick.Endtime, finalIncrementer - 1), finalData(DownTimeColumnOneClick.StartTime, finalIncrementer)) / 60
                        End If
                    End If
                End If
            Catch
            End Try

            finalIncrementer += 1
        End While


        'FAMILY CARE null location check
        If FamilyNullLocationList.Count > 0 Then 'this assures its family also

            shouldSnakeClose = True

            Dim xString As String = ""
            For i = 0 To FamilyNullLocationList.Count - 1
                xString = xString + FamilyNullLocationList(i).ToString() & vbCrLf
            Next
            Dim xString2 = "Blank Locations Detected At The Following Times:" + vbCrLf + xString
            MessageBox.Show(xString2,
"Blank Location Warning!",
MessageBoxButtons.OK,
MessageBoxIcon.Exclamation,
MessageBoxDefaultButton.Button1)
        End If

        Return finalData
    End Function

End Module
