Public Class Window_Trends
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
        If IsLaunchedFromTopStops = True Then
            Me.stopName = stopname 'added for s shape

            failuremodeno = current_selected_failuremodeno

            IsLaunchedfromstops_InMOtionChart = IsLaunchedFromTopStops
            selectedfailuremode_inMotionChart = current_selected_failuremodeno

          '  sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "_" & failuremodeno & "S.html"   '"file:///C:/Users/Public/motion" & motionchartsource & "_" & failuremodeno & "S.html"
          '  MotionChartS.Source = New Uri(sourcestringS)

         
            UseTrack_TopStopsTrends = True

        End If

        ' ConnectToDatabase()
    End Sub


    Private Sub SetSourceString()

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


     '   sourcestringS = "file:///" & SERVER_FOLDER_PATH & "motion" & motionchartsource & "S.html"
     '   MotionChartS.Source = New Uri(sourcestringS)

    End Sub

    #region "Show/Hide/Change Color"
    Private Sub stopclicked()
        stopsbutton.Opacity = 1.0
        prbutton.Opacity = 0.2
        mtbfbutton.Opacity = 0.2

        IsSPDActive = True
        ISDTActive = False
        ISMTBFActive = False

        DailyClicked()
    End Sub

    Private Sub prclicked()
        stopsbutton.Opacity = 0.2
        prbutton.Opacity = 1.0
        mtbfbutton.Opacity = 0.2

        IsSPDActive = False
        ISDTActive = True
        ISMTBFActive = False


        DailyClicked()
    End Sub

    Private Sub mtbfclicked()
        stopsbutton.Opacity = 0.2
        prbutton.Opacity = 0.2
        mtbfbutton.Opacity = 1.0

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
    End Sub

    Private Sub WeeklyClicked(sender As Object, e As RoutedEventArgs)
        Dailybtn.Background = mybrushdarkgray
        Weeklybtn.Background = mybrushbrightorange
        Monthlybtn.Background = mybrushdarkgray
    End Sub
    Private Sub MonthlyClicked(sender As Object, e As RoutedEventArgs)
        Dailybtn.Background = mybrushdarkgray
        Weeklybtn.Background = mybrushdarkgray
        Monthlybtn.Background = mybrushbrightorange
    End Sub
    #end region
End Class
