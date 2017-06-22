Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports Awesomium.Core
Imports Awesomium.Windows.Controls


Public Class Window_MotionChart

    Public IsSPDActive As Boolean = False
    Public ISDTActive As Boolean = False
    Public ISMTBFActive As Boolean = False
    Public IsLaunchedfromstops_InMOtionChart As Boolean = True
    Public selectedfailuremode_inMotionChart As Integer = 0

    'added for s shape
    Public EventDirectory As List(Of DTevent)
    Public ledsReport As SummaryReport
    Public downtimeData As DowntimeDataset
    Public stopName As String

 


    Private Sub motionchart_loaded()

        SetSourceString()
        Dailybtn.Background = mybrushbrightorange
        Weeklybtn.Background = mybrushdarkgray
        Monthlybtn.Background = mybrushdarkgray
    End Sub



    Public Sub New(IsLaunchedFromTopStops As Boolean, Optional current_selected_failuremodeno As Integer = 0, Optional stopname As String = "")

        InitializeComponent()
        Dim failuremodeno As Integer
        Dim sourcestringS As String
        Dim sourcestringD As String
        If IsLaunchedFromTopStops = True Then
            Me.stopName = stopname 'added for s shape

            failuremodeno = current_selected_failuremodeno

            IsLaunchedfromstops_InMOtionChart = IsLaunchedFromTopStops
            selectedfailuremode_inMotionChart = current_selected_failuremodeno

            sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
            MotionChartS.Source = New Uri(sourcestringS)

            sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D.html"
            MotionChartD.Source = New Uri(sourcestringD)
            losscardnamelabel.Content = stopname & " losses over last 3 months"


            sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S_Weekly.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
            MotionChartS_Weekly.Source = New Uri(sourcestringS)

            sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D_Weekly.html"
            MotionChartD_Weekly.Source = New Uri(sourcestringD)


            sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S_Monthly.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
            MotionChartS_Monthly.Source = New Uri(sourcestringS)

            sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "D_Monthly.html"
            MotionChartD_Monthly.Source = New Uri(sourcestringD)


            sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
            MotionChart_MTBF.Source = New Uri(sourcestringS)

            sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF_Monthly.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
            MotionChart_MTBF_Monthly.Source = New Uri(sourcestringS)

            sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "MTBF_Weekly.html"
            MotionChart_MTBF_Weekly.Source = New Uri(sourcestringS)


            UseTrack_TopStopsTrends = True

        End If

        ' ConnectToDatabase()
    End Sub


    Private Sub SetSourceString()

        Dim sourcestringS As String
        Dim sourcestringD As String

        prclicked()
        Select Case motionchartsource
            Case 31
                'losscardnamelabel.Content = "Top losses in analysis period"
                'losscardnamelabel.Content =

                stopclicked()


                Exit Sub
            Case 0
                losscardnamelabel.Content = "Line Performance"
                ' stopsbutton.Visibility = Windows.Visibility.Hidden
                ' prbutton.Visibility = Windows.Visibility.Hidden
                If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                    prbutton.Content = "PR"

                Else
                    prbutton.Content = "Av."
                End If

                mtbfbutton.Visibility = Windows.Visibility.Hidden
                losscardnamelabel.Content = "Line performance in last 3 months"
                UseTrack_PROverallTrends = True
        End Select


        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "S.html"
        MotionChartS.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "D.html"
        MotionChartD.Source = New Uri(sourcestringD)


        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "S_Weekly.html"
        MotionChartS_Weekly.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "D_Weekly.html"
        MotionChartD_Weekly.Source = New Uri(sourcestringD)

        sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "S_Monthly.html"
        MotionChartS_Monthly.Source = New Uri(sourcestringS)

        sourcestringD = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "D_MOnthly.html"
        MotionChartD_Monthly.Source = New Uri(sourcestringD)




    End Sub

    Private Sub stopclicked()
        stopsbutton.Opacity = 1.0
        prbutton.Opacity = 0.2
        mtbfbutton.Opacity = 0.2
        MotionChartS.Visibility = Windows.Visibility.Visible
        MotionChartD.Visibility = Windows.Visibility.Hidden
        MotionChartD_Weekly.Visibility = Windows.Visibility.Hidden
        MotionChartD_Monthly.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF_Monthly.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF_Weekly.Visibility = Windows.Visibility.Hidden

        IsSPDActive = True
        ISDTActive = False
        ISMTBFActive = False

        DailyClicked()
    End Sub

    Private Sub prclicked()
        stopsbutton.Opacity = 0.2
        prbutton.Opacity = 1.0
        mtbfbutton.Opacity = 0.2
        MotionChartD.Visibility = Windows.Visibility.Visible
        MotionChartS.Visibility = Windows.Visibility.Hidden
        MotionChartS_Weekly.Visibility = Windows.Visibility.Hidden
        MotionChartS_Monthly.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF_Monthly.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF_Weekly.Visibility = Windows.Visibility.Hidden
        IsSPDActive = False
        ISDTActive = True
        ISMTBFActive = False


        DailyClicked()
    End Sub

    Private Sub mtbfclicked()
        stopsbutton.Opacity = 0.2
        prbutton.Opacity = 0.2
        mtbfbutton.Opacity = 1.0
        MotionChartD.Visibility = Windows.Visibility.Hidden
        MotionChartS.Visibility = Windows.Visibility.Hidden
        MotionChartS_Weekly.Visibility = Windows.Visibility.Hidden
        MotionChartS_Monthly.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF.Visibility = Windows.Visibility.Visible
        MotionChart_MTBF_Monthly.Visibility = Windows.Visibility.Hidden
        MotionChart_MTBF_Weekly.Visibility = Windows.Visibility.Hidden
        IsSPDActive = False
        ISDTActive = False
        ISMTBFActive = True


        DailyClicked()
    End Sub

    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        'sender.Opacity = 0.7
    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        ' sender.Opacity = 1.0
    End Sub

    Private Sub DailyClicked()
        Dailybtn.Background = mybrushbrightorange
        Weeklybtn.Background = mybrushdarkgray
        Monthlybtn.Background = mybrushdarkgray


        If IsSPDActive Then

            MotionChartS.Visibility = Windows.Visibility.Visible
            MotionChartS_Weekly.Visibility = Windows.Visibility.Hidden
            MotionChartS_Monthly.Visibility = Windows.Visibility.Hidden

        ElseIf ISDTActive Then
            MotionChartD.Visibility = Windows.Visibility.Visible
            MotionChartD_Weekly.Visibility = Windows.Visibility.Hidden
            MotionChartD_Monthly.Visibility = Windows.Visibility.Hidden
        ElseIf ISMTBFActive Then
            MotionChart_MTBF.Visibility = Windows.Visibility.Visible
            MotionChart_MTBF_Weekly.Visibility = Windows.Visibility.Hidden
            MotionChart_MTBF_Monthly.Visibility = Windows.Visibility.Hidden

        End If

    End Sub

    Private Sub WeeklyClicked(sender As Object, e As RoutedEventArgs)
        Dailybtn.Background = mybrushdarkgray
        Weeklybtn.Background = mybrushbrightorange
        Monthlybtn.Background = mybrushdarkgray

        If IsSPDActive Then

            MotionChartS.Visibility = Windows.Visibility.Hidden
            MotionChartS_Weekly.Visibility = Windows.Visibility.Visible
            MotionChartS_Monthly.Visibility = Windows.Visibility.Hidden


        ElseIf ISDTActive Then
            MotionChartD.Visibility = Windows.Visibility.Hidden
            MotionChartD_Weekly.Visibility = Windows.Visibility.Visible
            MotionChartD_Monthly.Visibility = Windows.Visibility.Hidden
        ElseIf ISMTBFActive Then
            MotionChart_MTBF.Visibility = Windows.Visibility.Hidden
            MotionChart_MTBF_Weekly.Visibility = Windows.Visibility.Visible
            MotionChart_MTBF_Monthly.Visibility = Windows.Visibility.Hidden
        End If


    End Sub
    Private Sub MonthlyClicked(sender As Object, e As RoutedEventArgs)
        Dailybtn.Background = mybrushdarkgray
        Weeklybtn.Background = mybrushdarkgray
        Monthlybtn.Background = mybrushbrightorange

        If IsSPDActive Then

            MotionChartS.Visibility = Windows.Visibility.Hidden
            MotionChartS_Weekly.Visibility = Windows.Visibility.Hidden
            MotionChartS_Monthly.Visibility = Windows.Visibility.Visible


        ElseIf ISDTActive Then
            MotionChartD.Visibility = Windows.Visibility.Hidden
            MotionChartD_Weekly.Visibility = Windows.Visibility.Hidden
            MotionChartD_Monthly.Visibility = Windows.Visibility.Visible
        ElseIf ISMTBFActive Then
            MotionChart_MTBF.Visibility = Windows.Visibility.Hidden
            MotionChart_MTBF_Weekly.Visibility = Windows.Visibility.Hidden
            MotionChart_MTBF_Monthly.Visibility = Windows.Visibility.Visible
        End If

    End Sub

End Class
