Imports System.IO
Imports Awesomium.Core
Imports System.ComponentModel
Imports Awesomium.Windows.Controls
Imports System.Globalization
Public Class Window_Weibull

    Public totalnoevents As Integer
    Public totalnoevents_competing As Integer
    Public selectedfailuremodeList As New List(Of String)

    Private prstoryReport_weibull As prStoryMainPageReport

    Public Sub New(ByVal storyReport As prStoryMainPageReport)
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        prstoryReport_weibull = storyReport
    End Sub
    Private Sub weibull_loaded()
        LineName.Content = AllProdLines(selectedindexofLine_temp).ToString
        weibullplot.Visibility = Windows.Visibility.Visible
        hidemenu()
        selectedfailuremodeList.Clear()
        CreateSurvivalTable(My.Settings.maxuptimecutoff)

    End Sub

    Public Sub CreateSurvivalTable(maxuptimecutoff As Double, Optional dtcolumn As Integer = 16, Optional Onlyfailuremodes As Boolean = False)

        Dim uptimegroup As Double = 0.0
        Dim survivalchartstring As String
        Dim maxuptime As Double = 1440
        Dim lowestUptimegroup As Double = 0.0
        Dim uptimecount As Integer = 0
        Dim uptimeresolution As Double = 0.5
        Dim ListIndex As Integer = -1
        Dim k As Integer = 0
        Dim i As Integer
        Dim j As Integer

        ' determining max uptime in the raw downtime data
        maxuptime = 0
        totalnoevents = 0
        totalnoevents_competing = 0

        For i = 0 To AllProdLines(selectedindexofLine_temp).rawDowntimeData.UnplannedData.Count - 1
            With AllProdLines(selectedindexofLine_temp).rawDowntimeData.UnplannedData(i)

                If .MappedField <> "" Then
                    '        If .DTGroup.Contains("Equip") Or .DTGroup.Contains("Quality") Then
                    If .UT < maxuptimecutoff Then
                        maxuptime = Math.Max(maxuptime, .UT)
                    End If
                    totalnoevents_competing = totalnoevents_competing + 1
                    'Else
                    totalnoevents += 1
                    '    End If
                Else
                    If .UT < maxuptimecutoff Then
                        maxuptime = Math.Max(maxuptime, .UT)
                    End If
                    totalnoevents += 1
                    totalnoevents_competing = totalnoevents_competing + 1
                End If
            End With
        Next

        j = 0

        Dim uptimearray(4, totalnoevents) As Object

        ' definition of uptimearray
        ' 0 --> raw uptime for all failure modes with PR in
        ' 1 --> failure mode names or "" if DBnull from the default mapping column
        ' 2 --> computed censored uptimes
        ' 3 --> competing or cumulative flag

        'creating the uptimearray from raw proficy data
        j = 0
        For i = 0 To AllProdLines(selectedindexofLine_temp).rawDowntimeData.UnplannedData.Count - 1
            With AllProdLines(selectedindexofLine_temp).rawDowntimeData.UnplannedData(i)
                ' If .DTGroup <> "" Then
                '      If .DTGroup.Contains("Equip") Or .DTGroup.Contains("Quality") Then
                uptimearray(0, j) = .UT
                uptimearray(2, j) = .UT
                uptimearray(1, j) = .MappedField '.DTGroup
                uptimearray(3, j) = 1  ' is competing
                '  Else
                '  uptimearray(0, j) = .UT
                '  uptimearray(2, j) = .UT
                '  uptimearray(1, j) = .DTGroup
                '  End If
                '     Else
                '    uptimearray(0, j) = .UT
                '    uptimearray(2, j) = .UT
                '   uptimearray(1, j) = ""
                '   uptimearray(3, j) = 1 ' is competing
                '   End If
                j = j + 1
                If j = totalnoevents Then Exit For
            End With
        Next

        'lg code expirement for MTBF 

        ''''

        ' calculating censored uptimes for each failure mode  ' need to add condition for competing cause
        Dim m As Integer
        Dim n As Integer
        Dim failuremode_analyzed As String

        For m = uptimearray.GetLength(1) - 1 To 0 Step -1
            If uptimearray(1, m) <> "" Then
                failuremode_analyzed = uptimearray(1, m)
                If m <> 0 Then
                    If uptimearray(1, m - 1) <> failuremode_analyzed Then
                        For n = m - 1 To 0 Step -1
                            If uptimearray(1, n) = failuremode_analyzed Then
                                Exit For
                            End If

                            If uptimearray(3, n) = 1 Then uptimearray(2, m) = uptimearray(2, m) + uptimearray(0, n) ' add uptimes only if a competing cause

                        Next
                    End If
                End If
            End If
        Next m


        'creating the actual survival table for all failure modes and selected failure mode
        uptimegroup = 0
        uptimeresolution = 0.5
        j = 0

        Dim survivaltable(13, (maxuptime / uptimeresolution)) As Double

        ' survival table definition
        ' 0 --> uptime group list (0, 0.5, 1, 1.5, 2 .....
        ' 1 --> uptime count for all failure modes
        ' 2 ->  CDF for all failure modes
        ' 3 --> CDF for selected failure mode 1
        ' 4 --> CDF for selected failure mode 2
        ' 5 --> CDF for selected failure mode 3
        ' 6 --> CDF for selected failure mode 4
        ' 7 --> CDF for selected failure mode 5
        ' 8 --> CDF for selected failure mode 6
        ' 9 --> CDF for selected failure mode 7
        ' 10 --> CDF for selected failure mode 8
        ' 11 --> CDF for selected failure mode 9
        ' 12 --> CDF for selected failure mode 10



        For uptimegroup = lowestUptimegroup To maxuptime Step uptimeresolution
            uptimecount = 0
            For i = 0 To uptimearray.GetLength(1) - 1
                If uptimearray(0, i) <= uptimegroup Then
                    uptimecount = uptimecount + 1
                End If

                If uptimearray(2, i) <= uptimegroup Then
                    If selectedfailuremodeList.Count <> 0 Then
                        ListIndex = selectedfailuremodeList.FindIndex(Function(value As String)
                                                                          Return value = uptimearray(1, i)
                                                                      End Function)

                        survivaltable(3 + ListIndex, j) = survivaltable(3 + ListIndex, j) + 1

                    End If
                End If
            Next
            survivaltable(0, j) = uptimegroup
            survivaltable(1, j) = uptimecount
            survivaltable(2, j) = Math.Round(1 - survivaltable(1, j) / totalnoevents, 4)
            'For k = 3 To 12
            ' If survivaltable(k, j) <> 0 Then
            ' survivaltable(k, j) = Math.Round(1 - survivaltable(k, j) / totalnoevents_competing, 4)
            ' End If
            'Next


            j = j + 1
        Next

        ' calculating CDF for selected failure modes with censoring
        Dim l As Integer
        If selectedfailuremodeList.Count > 0 Then
            For k = 3 To selectedfailuremodeList.Count + 2
                For l = 0 To j - 1
                    If survivaltable(k, l) = survivaltable(k, j - 1) Then
                        survivaltable(k, l) = 0.0
                    Else
                        survivaltable(k, l) = Math.Round(1 - survivaltable(k, l) / survivaltable(k, j - 1), 4) 'LG code
                        'survivaltable(k, l) = Math.Round(1 - survivaltable(k, l) / totalnoevents_competing, 4) ' LG Code
                    End If
                Next
            Next
        End If





        Dim actualListsize_ofselectedfailuremodeList As Integer
        actualListsize_ofselectedfailuremodeList = selectedfailuremodeList.Count
        If selectedfailuremodeList.Count <> 0 Then
            For i = actualListsize_ofselectedfailuremodeList To 10
                selectedfailuremodeList.Add("")

            Next

        End If

        survivalchartstring = "['Time', 'Total Line Survivability',"

        If selectedfailuremodeList.Count <> 0 Then
            For k = 3 To actualListsize_ofselectedfailuremodeList + 2
                survivalchartstring = survivalchartstring & "'" & selectedfailuremodeList(k - 3) & "', "

            Next
            survivalchartstring = survivalchartstring & "],"
        Else
            survivalchartstring = "['Time', 'Total Line Survivability'],"
        End If


        'first case - do not want 'null'
        i = 0
        If selectedfailuremodeList.Count <> 0 Then
            survivalchartstring = survivalchartstring & "[" & survivaltable(0, i) & "," & survivaltable(2, i)
            For k = 3 To actualListsize_ofselectedfailuremodeList + 2
                survivalchartstring = survivalchartstring & " ," & survivaltable(k, i)
            Next
            survivalchartstring = survivalchartstring & "],"
        Else
            survivalchartstring = survivalchartstring & "[" & survivaltable(0, i) & "," & survivaltable(2, i) & "],"
        End If

        'remaining cases
        For i = 1 To survivaltable.GetLength(1) - 1
            If selectedfailuremodeList.Count <> 0 Then
                survivalchartstring = survivalchartstring & "[" & survivaltable(0, i) & "," & ReturnNullifzero(survivaltable(2, i))
                For k = 3 To actualListsize_ofselectedfailuremodeList + 2
                    survivalchartstring = survivalchartstring & " ," & ReturnNullifzero(survivaltable(k, i))
                Next
                survivalchartstring = survivalchartstring & "],"
            Else
                survivalchartstring = survivalchartstring & "[" & survivaltable(0, i) & "," & ReturnNullifzero(survivaltable(2, i)) & "],"
            End If
        Next

        ' CreateSurvivalPlot(survivalchartstring)

        CreateSurvivalPlot_AMCHarts(survivaltable, actualListsize_ofselectedfailuremodeList, selectedfailuremodeList)
        weibullplot.Reload(ignoreCache:=True)
    End Sub
    
    Private Function CountUptime(uptimegroup As Double) As Integer
        CountUptime = 0
        Dim i As Integer

        For i = 0 To AllProdLines(selectedindexofLine_temp).rawDowntimeData.UnplannedData.Count - 1 'AllProductionLines(selectedindexofLine_temp).rawProficyData.GetLength(1) - 1
            '   If IsDBNull(AllProductionLines(selectedindexofLine_temp).rawProficyData(12, i)) = False Then
            'If InStr(AllProductionLines(selectedindexofLine_temp).rawProficyData(12, i), "PR In", vbTextCompare) > 0 Then
            If AllProdLines(selectedindexofLine_temp).rawDowntimeData.UnplannedData(i).UT <= uptimegroup Then ' If AllProductionLines(selectedindexofLine_temp).rawProficyData(3, i) <= uptimegroup Then
                CountUptime = CountUptime + 1
                'End If
                '  End If
            End If


        Next


        Return CountUptime
    End Function
    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        'sender.Opacity = 0.7
    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        ' sender.Opacity = 1.0
    End Sub

    Private Sub ShowMenu()
        splashweibull.Visibility = Windows.Visibility.Visible
        Failuremodelistbox.Visibility = Windows.Visibility.Visible
        failuremodefilterDonebutton.Visibility = Windows.Visibility.Visible
        failuremodefilterCancelbutton.Visibility = Windows.Visibility.Visible
        failuremodelegendheading.Visibility = Windows.Visibility.Visible
        CreateFailureModeList()
        UseTrack_WeibullMain_failuremodes = True
    End Sub

    Private Sub DoneMenu()
        hidemenu()

        For i As Integer = 0 To Failuremodelistbox.SelectedItems.Count - 1

            If Failuremodelistbox.SelectedItems(i).ToString = "<BLANK>" Then
                MsgBox("<BLANK> failure modes cannot be analyzed. Please select any failure mode(s) excluding <BLANK>")
                selectedfailuremodeList.Clear()
                Exit Sub
            End If
            selectedfailuremodeList.Add(CStr(Failuremodelistbox.SelectedItems(i)))
        Next

        If selectedfailuremodeList.Count > 9 Then
            MsgBox("Not more than 10 failure modes can be selected for survival plots. Try selecting failure modes again.", vbInformation)
            selectedfailuremodeList.Clear()
            Exit Sub
        End If

        If selectedfailuremodeList.Count <> 0 Then
            CreateSurvivalTable(My.Settings.maxuptimecutoff, My.Settings.defaultMappingLevel)
        End If

    End Sub
    Private Sub hidemenu()
        splashweibull.Visibility = Windows.Visibility.Hidden
        Failuremodelistbox.Visibility = Windows.Visibility.Hidden
        failuremodefilterDonebutton.Visibility = Windows.Visibility.Hidden
        failuremodefilterCancelbutton.Visibility = Windows.Visibility.Hidden
        failuremodelegendheading.Visibility = Windows.Visibility.Hidden
    End Sub
    Private Sub CreateFailureModeList()
        Failuremodelistbox.Items.Clear()
        selectedfailuremodeList.Clear()
        Dim tmpdtevent4 As DTevent
        For i = 0 To 14
            tmpdtevent4 = prstoryReport_weibull.getCardEventInfo(31, i)
            Failuremodelistbox.Items.Add(tmpdtevent4.Name)
        Next i
    End Sub
    Private Sub Calculate632percentile(utarray As Object)
        Dim utlist As New List(Of Double)
        Dim i As Integer

        For i = 0 To (utarray.getlength(1) - 1)
            utlist.Add(utarray(0, i))
        Next

        utlist.Sort()

        Dim mtbf As Double = utlist.Item(Int(utlist.Count * 0.632))

        MsgBox(mtbf)

    End Sub
    Sub showfailuremodesonly()
        If selectedfailuremodeList.Count > 0 Then
            CreateSurvivalTable(My.Settings.maxuptimecutoff, My.Settings.defaultMappingLevel, True)
        End If
    End Sub
End Class