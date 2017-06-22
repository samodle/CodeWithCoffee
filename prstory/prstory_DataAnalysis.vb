
Option Explicit On
Public Module prstory_DataAnalysis



    Public uptime As Double
    Public downtime As Double
    Public UPDT As Double
    Public UPDTper As Double
    Public scheduledtimetotal As Double
    Public scheduledtime_forUT As Double
    Public DTdata As Array
    Public PRODdata As Array
    Public PR As Double



    Private Sub AnalyzeMain_for_prstoryX()
        '    Dim bargraphreportwindow As New bargraphreportwindow
        Dim windowmain_prstory As New WindowMain_prstory
        'Call AllProductionLines(selectedindexofLine_temp).pullProficyData_Custom(starttimeselected, endtimeselected)
        'MsgBox("Proficy data downloaded for Line " & AllProductionLines(selectedindexofLine_temp).Name)

        'DTdata = AllProductionLines(selectedindexofLine_temp).rawProficyData
        'PRODdata = AllProductionLines(selectedindexofLine_temp).rawProficyProductionData
        'Calculate_ScheduledTime_fromproddata()
        'calculateUPDT()
        ' MsgBox(FormatPercent(PR))
        ' MsgBox(FormatPercent(UPDTper))
        windowmain_prstory.Close()
        '  bargraphreportwindow.Show()

        'CategorizeDT(16)
    End Sub



    Public Sub Calculate_ScheduledTime_fromrawdata()


        Dim i As Integer


        uptime = 0
        downtime = 0
        scheduledtime_forUT = 0

        For i = 0 To DTdata.GetLength(1) - 1



            uptime = uptime + DTdata(3, i)
            downtime = downtime + DTdata(2, i)

        Next i



        If IsNothing(uptime) Then
            scheduledtime_forUT = 0   'return scheduled time as zero if there is no uptime or downtime found
        Else
            scheduledtime_forUT = uptime + downtime
        End If


    End Sub

    Public Sub Calculate_ScheduledTime_fromproddata()


        Dim i As Integer
        Dim targetrate As Double
        Dim actualrate As Double
        Dim casescount As Double
        Dim temptimediff As Double
        Dim scheduledtime As Double
        Dim ccount As Double
        Dim statcaseconv As Double
        Dim actualproduction As Double
        Dim actualproductionTotal As Double
        Dim targetproduction As Double
        Dim targetproductionTotal As Double

        scheduledtime = 0
        scheduledtimetotal = 0
        actualproduction = 0
        actualproductionTotal = 0
        targetproduction = 0
        targetproductionTotal = 0

        For i = 0 To PRODdata.GetLength(1) - 1



            If InStr(PRODdata(5, i), "PR In", CompareMethod.Text) > 0 Then
                temptimediff = DateDiff(DateInterval.Minute, PRODdata(0, i), PRODdata(1, i))


                scheduledtime = temptimediff
                casescount = PRODdata(10, i)     ' cases produced
                actualrate = PRODdata(13, i)
                targetrate = PRODdata(14, i)
                statcaseconv = PRODdata(17, i)
                ccount = PRODdata(16, i)    'actual Case Count ex - (24/case)

                scheduledtimetotal = scheduledtimetotal + scheduledtime

                actualproduction = casescount * statcaseconv
                actualproductionTotal = actualproductionTotal + actualproduction


                targetproduction = scheduledtime * (targetrate / ccount) * statcaseconv
                targetproductionTotal = targetproductionTotal + targetproduction

            End If


        Next

        If targetproductionTotal = 0 Then
            PR = 0
        Else
            PR = actualproductionTotal / targetproductionTotal
        End If

    End Sub

    Public Sub calculateUPDT()

        Dim i As Integer

        UPDT = 0
        For i = 0 To DTdata.GetLength(1) - 1

            If DTdata(15, i).ToString = "Unplanned" Then

                UPDT = UPDT + DTdata(2, i)

            End If

        Next

        UPDTper = UPDT / scheduledtimetotal

    End Sub


    Public Sub CategorizeDT(RL As Integer)

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim v As Object


        Dim alldt(10000) As String
        Dim d As Object
        Dim dtnamelist(1000) As String
        Dim dttime(1000) As Double
        Dim dtstops(1000) As Long
        Dim dtloss(1000) As Double
        Dim stopscount(1) As Long

        stopscount(1) = 1
 
        For k = 0 To DTdata.GetLength(1) - 1
            If IsDBNull(DTdata(RL, k)) Then
                alldt(k) = " "
            Else
                alldt(k) = DTdata(RL, k)
            End If

        Next k


        d = CreateObject("Scripting.Dictionary")
        For Each el In alldt
            d(el) = 1
        Next
        v = d.Keys

        For i = 0 To UBound(v)

            dtnamelist(i) = v(i)
            dtstops(i) = 0
            dttime(i) = 0

        Next



        For j = 0 To dtnamelist.GetLength(0) - 1

            For i = 0 To DTdata.GetLength(1) - 1

                If IsDBNull(DTdata(RL, i)) Then

                Else
                    If DTdata(RL, i) = dtnamelist(j) Then
                        dttime(j) = dttime(j) + DTdata(2, i)
                        dtstops(j) = dtstops(j) + 1
                    End If
                End If
                

            Next

            dtloss(j) = dttime(j) / scheduledtimetotal
        Next

        MsgBox(dtnamelist(6) & " " & dtstops(6) & " " & FormatPercent(dtloss(6)))

    End Sub

End Module




