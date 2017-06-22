

Imports MongoDB.Bson
Imports MongoDB.Driver

Public Class Window_Celebrations

    Public Sub New(LineName As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        LinenameLabel.Content = LineName

        Select Case LineName
            Case "X-Wing Assembly Bay 1"
                MainLogo.Source = New BitmapImage(New Uri("april_rebel.jpg", UriKind.Relative))
                KPI1Header.Content = "Process Reliability"
                KPI2Header.Content = "MTBF"
                KPI3Header.Content = "Targeting Computer Status"
                KPI1value.Content = "85%"
                KPI2value.Content = "90"
                KPI3value.Content = "Offline"
            Case "Death Star Reactor Room"
                MainLogo.Source = New BitmapImage(New Uri("april_imperial.png", UriKind.Relative))
                KPI1Header.Content = "Process Reliability"
                KPI2Header.Content = "MTBF"
                KPI3Header.Content = "Explosions"
                KPI1value.Content = "0%"
                KPI2value.Content = "1.5 Movies"
                KPI3value.Content = "2"
            Case "LexCorp Converting 17C"
                MainLogo.Source = New BitmapImage(New Uri("april_lexcorp.jpg", UriKind.Relative))
                KPI1Header.Content = "Process Reliability"
                KPI2Header.Content = "MTBF"
                KPI3Header.Content = "Stops"
                KPI1value.Content = "64%"
                KPI2value.Content = "45"
                KPI3value.Content = "29"
            '   Case "Stark Industries Line 7"
            '       MainLogo.Source = New BitmapImage(New Uri("april_stark.png", UriKind.Relative))
            '       KPI1Header.Content = "Process Reliability"
            '       KPI2Header.Content = "MTBF"
            '       KPI3Header.Content = "Stops"
            '       KPI1value.Content = "Classified"
            '       KPI2value.Content = "Classified"
            '       KPI3value.Content = "Classified"
            Case "Stark Industries Line 9"
                MainLogo.Source = New BitmapImage(New Uri("april_stark.png", UriKind.Relative))
                KPI1Header.Content = "Process Reliability"
                KPI2Header.Content = "MTBF"
                KPI3Header.Content = "Stops"
                KPI1value.Content = "110%"
                KPI2value.Content = "Inf"
                KPI3value.Content = "0"
            Case "Area 51"
                MainLogo.Source = New BitmapImage(New Uri("april_51.jpg", UriKind.Relative))
                KPI1Header.Content = "Process Reliability"
                KPI2Header.Content = "MTBF"
                KPI3Header.Content = "Stops"
                KPI1value.Content = "Classified"
                KPI2value.Content = "Classified"
                KPI3value.Content = "Classified"
            Case "Buy n Large Batteries"
                MainLogo.Source = New BitmapImage(New Uri("april_bnl.png", UriKind.Relative))
                KPI1Header.Content = "Process Reliability (Earth)"
                KPI2Header.Content = "MTBF"
                KPI3Header.Content = "Stops"
                KPI1value.Content = "2%"
                KPI2value.Content = "5"
                KPI3value.Content = "0"
            Case "SPECTRE Line 007"
                MainLogo.Source = New BitmapImage(New Uri("april_spectre.jpg", UriKind.Relative))
                KPI1Header.Content = "PR of Car"
                KPI2Header.Content = "MTBM (martini)"
                KPI3Header.Content = "% Shaken"
                KPI1value.Content = "0%"
                KPI2value.Content = "18 Hrs"
                KPI3value.Content = "100"

            Case "Acme Corp Anvil Assembly"

                MainLogo.Source = New BitmapImage(New Uri("april_acme.jpg", UriKind.Relative))
                KPI1Header.Content = "Success Rate"
                KPI2Header.Content = "MTBF"
                KPI3Header.Content = "Notes"
                KPI1value.Content = "0%"
                KPI2value.Content = "0.01"
                KPI3value.Content = "Beep Beep!"
            Case "Wonka Gobstoppers 13"
                MainLogo.Source = New BitmapImage(New Uri("april_wonka.jpg", UriKind.Relative))
                KPI1Header.Content = "Golden Tickets"
                KPI2Header.Content = "Tour Survival Rate"
                KPI3Header.Content = "Flavor Expiry"
                KPI1value.Content = "5"
                KPI2value.Content = "20%"
                KPI3value.Content = "#DIV0"
            Case "Platform 9 3/4"
                MainLogo.Source = New BitmapImage(New Uri("april_platform.png", UriKind.Relative))
                KPI1Header.Content = "Station"
                KPI2Header.Content = "Ticket Sales"
                KPI3Header.Content = "Muggle Reject Rate"
                KPI1value.Content = "Kings Cross"
                KPI2value.Content = "4,328"
                KPI3value.Content = "100%"
            Case "Wayne Enterprises BioTech"
                MainLogo.Source = New BitmapImage(New Uri("april_wayne.jpg", UriKind.Relative))
                KPI1Header.Content = "Marketing Budget ($)"
                KPI2Header.Content = "Operational Expense ($)"
                KPI3Header.Content = "'Research' ($)"
                KPI1value.Content = "0.2 MM"
                KPI2value.Content = "1.6 MM"
                KPI3value.Content = "193 MM"
            Case "Globex Corp Windows"
                MainLogo.Source = New BitmapImage(New Uri("april_globex.jpg", UriKind.Relative))
                KPI1Header.Content = "% Walls"
                KPI2Header.Content = "% Windows"
                KPI3Header.Content = "Power Source"
                KPI1value.Content = "0"
                KPI2value.Content = "100"
                KPI3value.Content = "Nuclear"
        End Select

        APRIL1SendUserAnalyticsDatatoServer()
    End Sub


    Private Sub Celebrations_loaded()
        '  celebrationsprlabel.Content = FormatPercent(prStoryReport.PR, 0)

    End Sub
    Private Sub closecelebrationswindow()


        Me.Close()
    End Sub

    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        sender.Opacity = 0.7
    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        sender.Opacity = 1.0
    End Sub

    Private Sub APRIL1SendUserAnalyticsDatatoServer()
        Try
            If My.Computer.Network.Ping("prstory.pg.com") Then
                Dim client As MongoClient
                Dim server As MongoServer
                Dim db As MongoDatabase
                Dim col1 As MongoCollection
                client = New MongoClient("mongodb://prstory.pg.com/MongoServer")

                server = client.GetServer()
                db = server.GetDatabase("prstory")
                '  col1 = db.GetCollection(Of BsonDocument)("UseTrack")
                col1 = db("UseTrack")

                Dim currentloginname As String

                currentloginname = Environment.UserName
                Dim NewInfoBson As BsonDocument = New BsonDocument() _
                                                   .Add("who", String.Format(currentloginname)) _
                                                     .Add("when", String.Format(Now(), "MM dd yyyy hh:mm")) _
                                                    .Add("Line", String.Format(LinenameLabel.Content)) _
                                                    .Add("A", UseTrack_UPDTview) _
                                                    .Add("B", UseTrack_PDTview) _
                                                    .Add("C", UseTrack_PROverallTrends) _
                                                    .Add("D", UseTrack_RawDatawindow_Main) _
                                                    .Add("E", UseTrack_RawDatawindow_Paretos) _
                                                    .Add("F", UseTrack_RawDatawindow_Variance) _
                                                    .Add("G", UseTrack_WeibullMain) _
                                                    .Add("H", UseTrack_WeibullMain_failuremodes) _
                                                    .Add("I", UseTrack_IncontrolMain) _
                                                    .Add("J", UseTrack_IncontrolControlChart) _
                                                    .Add("K", UseTrack_IncontrolControlShift) _
                                                    .Add("L", UseTrack_TopStopsMain) _
                                                    .Add("M", UseTrack_StopsWatchMain) _
                                                    .Add("N", UseTrack_TopStopsTrends) _
                                                    .Add("O", UseTrack_ChangeMapping) _
                                                    .Add("P", UseTrack_Filter) _
                                                    .Add("Q", UseTrack_ExportLossTree) _
                                                    .Add("R", UseTrack_ExportDowntime) _
                                                    .Add("S", UseTrack_ExportProduction) _
                                                    .Add("T", UseTrack_ExportDependency) _
                                                    .Add("U", UseTrack_Notes) _
                                                    .Add("V", UseTrack_Simulation) _
                                                    .Add("W", UseTrack_Notes_PickaLoss) _
                                                    .Add("X", UseTrack_Notes_ExporttoExcel) _
                                                    .Add("Y", UseTrack_TargetsMain) _
                .Add("Z", "oldserver")


                col1.Insert(NewInfoBson)
                server.Disconnect()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

End Class
