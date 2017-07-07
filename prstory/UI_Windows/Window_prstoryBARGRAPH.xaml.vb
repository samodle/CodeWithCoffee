#Region "References"
Imports System.Collections.ObjectModel
Imports System.ComponentModel 'not sure if i need thsi
Imports System.Threading
Imports System.Windows.Media.Effects
Imports System.Net
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Drawing

Imports MongoDB.Driver
Imports MongoDB.Bson
Imports System.Windows.Media.Animation
Imports System.Windows.Media
'Imports System.Windows.Forms
#End Region


Public Class bargraphreportwindow

    'Private tmpRawDataWindow = New RawDataWindow

    'VARS ADDED FOR CONSTRUCTOR
    Private prStoryReport As prStoryMainPageReport
    Dim stopsWatchThread As Thread
    Private selectedfailuremode As String 'lg code
    Public Tier1Clicked_Unplanned As String = ""
    Public Tier2Clicked_Unplanned As String = ""
    Public Tier1Clicked_planned As String = ""

#Region "DT View Flexibility"
    'Line/Time Period #1
    Friend Card_Unplanned_T1 As New List(Of DTevent)
    Friend Card_Unplanned_T2 As New List(Of DTevent)
    Friend Card_Unplanned_T3A As New List(Of DTevent)
    Friend Card_Unplanned_T3B As New List(Of DTevent)
    Friend Card_Unplanned_T3C As New List(Of DTevent)

    Friend Card_Planned_T1 As New List(Of DTevent)
    Friend Card_Planned_T2 As New List(Of DTevent)

    'line top stops
    Friend Card_TopStops As New List(Of DTevent)
    Friend Card_TopThreeStops As New List(Of DTevent)

    'refresh the data from the prstory report
    Private Sub updateCardList(cardNumber As Integer, ScrollOffset As Integer, Mapping As Integer, Optional Selection_T1 As String = "", Optional Selection_T2 As String = "")
        With prStoryReport
            Select Case cardNumber
                Case prStoryCard.Changeover
                    .updateCardList_Planned_Tier2(ScrollOffset, Selection_T1)
                Case prStoryCard.Equipment
                    .updateCardList_Unplanned_Tier2(ScrollOffset, Selection_T1)
                Case prStoryCard.Equipment_One
                    .updateCardList_Unplanned_Tier3A(ScrollOffset, Selection_T1, Selection_T2)
                Case prStoryCard.Equipment_Two
                    .updateCardList_Unplanned_Tier3B(ScrollOffset, Selection_T1, Selection_T2)
                Case prStoryCard.Equipment_Three
                    .updateCardList_Unplanned_Tier3C(ScrollOffset, Selection_T1, Selection_T2)
                Case prStoryCard.Unplanned
                    .updateCardList_Unplanned_Tier1(ScrollOffset)
                Case prStoryCard.Planned
                    .updateCardList_Planned_Tier1(ScrollOffset)
                Case prStoryCard.Stops
                    .updateCardList_Stops(ScrollOffset)
                Case Else
                    Throw New unknownprstoryCardException
            End Select
        End With
    End Sub

    Sub launchchangeover()
        showchangeoverlabels()
        showPDTlabels()
        showPlannedchangeoverTIMEchart()
        hideEquipmentlabels()
        unplannedDTequipmentchart.Visibility = Visibility.Hidden
        unplannedDTequip1chart.Visibility = Visibility.Hidden
        unplannedDTequip2chart.Visibility = Visibility.Hidden
        unplannedDTequip3chart.Visibility = Visibility.Hidden
        PDTlabel2.Background = LabelSelectedColor

    End Sub

    Private Sub SetPDTLabelDefaultColor()

        PDTLabel1.Background = LabelDefaultColor
        PDTlabel2.Background = LabelDefaultColor
        PDTlabel3.Background = LabelDefaultColor
        PDTlabel4.Background = LabelDefaultColor
        PDTlabel5.Background = LabelDefaultColor
        PDTlabel6.Background = LabelDefaultColor
        PDTLabel7.Background = LabelDefaultColor
        PDTLabel8.Background = LabelDefaultColor
        PDTLabel9.Background = LabelDefaultColor
    End Sub

    Private Sub Tier1LabelClicktoShowTier2_Planned(sender As Object, e As MouseButtonEventArgs)
        If IsPickMode = True Then
            PickaLoss_CollectInfofromLabel(sender, 2)
            Exit Sub
        End If

        Tier1Clicked_planned = sender.content.ToString
        'update the Tier 2 Unplanned Card
        updateCardList(prStoryCard.Changeover, 0, DowntimeField.Tier1, sender.content.ToString, "")

        'update the graphics
        updateCard_Planned_Tier2()

        'from the original
        showchangeoverlabels()
        showPDTlabels()
        showPlannedchangeoverTIMEchart()
        hideEquipmentlabels()
        unplannedDTequipmentchart.Visibility = Visibility.Hidden
        unplannedDTequip1chart.Visibility = Visibility.Hidden
        unplannedDTequip2chart.Visibility = Visibility.Hidden
        unplannedDTequip3chart.Visibility = Visibility.Hidden
        SetPDTLabelDefaultColor()
        sender.background = LabelSelectedColor
        Card41Header.Content = sender.content.ToString 'cardnameLabeltext(41)
        ' Tier1Clicked_planned = sender.content.ToString
        If Card41Header.Content = "CO" Or InStr(Card41Header.Content, "Change", vbTextCompare) > 0 Or InStr(Card41Header.Content, "C/O", vbTextCompare) > 0 Then
            Card42Header.Content = "Changeover Events# & Time"
        Else
            Card42Header.Content = "Events# & MTTR"
        End If


    End Sub

    Private Sub Tier1LabelClicktoShowTier2(sender As Object, e As MouseButtonEventArgs)
        If IsPickMode = True Then
            PickaLoss_CollectInfofromLabel(sender, 1)
            Exit Sub
        End If


        If sender.content.ToString = "Total" Then Exit Sub

        Dim Tier3AString As String, Tier3BString As String, Tier3CString As String

        ScrollBase_Card3 = 0
        ScrollBase_Card4 = 0
        ScrollBase_Card5 = 0
        ScrollBase_Card6 = 0
        NavigationLeft_card3.Visibility = Visibility.Hidden
        NavigationLeft_card4.Visibility = Visibility.Hidden
        NavigationLeft_card5.Visibility = Visibility.Hidden
        NavigationLeft_card6.Visibility = Visibility.Hidden
        Tier1Clicked_Unplanned = sender.content.ToString
        'update the Tier 2 Unplanned Card
        updateCardList(prStoryCard.Equipment, ScrollBase_Card3, DowntimeField.Tier1, sender.content.ToString, "")

        'figure out which names to put on the three 'top three' cards
        If Card_Unplanned_T2.Count > 0 Then
            Tier3AString = Card_Unplanned_T2(0).Name
            If Card_Unplanned_T2.Count > 1 Then
                Tier3BString = Card_Unplanned_T2(1).Name
                If Card_Unplanned_T2.Count > 2 Then
                    Tier3CString = Card_Unplanned_T2(2).Name
                Else
                    Tier3CString = ""
                End If
            Else
                Tier3BString = ""
                Tier3CString = ""
            End If
        Else
            Tier3AString = ""
            Tier3BString = ""
            Tier3CString = ""
        End If

        'update tier 3 labels
        cardnameLabeltext(prStoryCard.Equipment_Three) = Tier3CString
        cardnameLabeltext(prStoryCard.Equipment_Two) = Tier3BString
        cardnameLabeltext(prStoryCard.Equipment_One) = Tier3AString

        'update the three Tier 3 Unplanned Cards
        updateCardList(prStoryCard.Equipment_One, 0, 0, sender.content.ToString, Tier3AString)
        updateCardList(prStoryCard.Equipment_Two, 0, 0, sender.content.ToString, Tier3BString)
        updateCardList(prStoryCard.Equipment_Three, 0, 0, sender.content.ToString, Tier3CString)

        'update the graphics
        updateCard_Unplanned_Tier2()
        updateCard_Unplanned_Tier3_All()


        showDTpercentframe()
        Card3Header.Content = sender.content.ToString
        sender.background = LabelSelectedColor ' New SolidColorBrush(Windows.Media.Color.FromRgb(0, 100, 255))
        showEquipmentlabels()
        ' EquipmentLabel1.Background = LabelSelectedColor
        PDTlabel2.Background = LabelDefaultColor
        RefreshHeaderColors()
        SetPDTLabelDefaultColor()
        If prStoryReport.getCardEventNumber(3) <= 6 Then NavigationRight_Card3.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(1) <= 6 Then NavigationRight_Card1.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(4) <= 3 Then NavigationRight_Card4.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(5) <= 3 Then NavigationRight_Card5.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(6) <= 3 Then NavigationRight_Card6.Visibility = Visibility.Hidden
        If UPDTlabel6.Content.ToString = OTHERS_STRING Then NavigationRight_Card1.Visibility = Visibility.Hidden


    End Sub
    Private Sub Tier2LabelClicktoShowTier3(sender As Object, e As MouseButtonEventArgs)

        If IsPickMode = True Then
            PickaLoss_CollectInfofromLabel(sender, 3)
            Exit Sub
        End If

        Dim IndexA As Integer
        'clear tier 3 labels
        cardnameLabeltext(prStoryCard.Equipment_Three) = ""
        cardnameLabeltext(prStoryCard.Equipment_Two) = ""
        cardnameLabeltext(prStoryCard.Equipment_One) = ""

        ScrollBase_Card4 = 0
        ScrollBase_Card5 = 0
        ScrollBase_Card6 = 0
        NavigationLeft_card4.Visibility = Visibility.Hidden
        NavigationLeft_card5.Visibility = Visibility.Hidden
        NavigationLeft_card6.Visibility = Visibility.Hidden

        'ScrollBase_Card3 = 0
        IndexA = Card_Unplanned_T2.IndexOf(New DTevent(sender.content.ToString, 0))

        If IndexA > -1 Then
            updateCardList(4, ScrollBase_Card3, DowntimeField.Tier1, Card3Header.Content.ToString, Card_Unplanned_T2(IndexA).Name)
            Card4Header.Content = Card_Unplanned_T2(IndexA).Name
            cardnameLabeltext(prStoryCard.Equipment_One) = Card_Unplanned_T2(IndexA).Name
            If IndexA + 1 < Card_Unplanned_T2.Count Then
                Card5Header.Content = Card_Unplanned_T2(IndexA + 1).Name
                cardnameLabeltext(prStoryCard.Equipment_Two) = Card_Unplanned_T2(IndexA + 1).Name
                updateCardList(5, ScrollBase_Card3, DowntimeField.Tier1, Card3Header.Content.ToString, Card_Unplanned_T2(IndexA + 1).Name)
                If IndexA + 2 < Card_Unplanned_T2.Count Then
                    updateCardList(6, ScrollBase_Card3, DowntimeField.Tier1, Card3Header.Content.ToString, Card_Unplanned_T2(IndexA + 2).Name)
                    Card6Header.Content = Card_Unplanned_T2(IndexA + 2).Name
                    cardnameLabeltext(prStoryCard.Equipment_Three) = Card_Unplanned_T2(IndexA + 2).Name
                Else
                    cardnameLabeltext(prStoryCard.Equipment_Three) = ""
                    updateCardList(6, ScrollBase_Card3, DowntimeField.Tier1, "x", "x")
                End If
            Else
                cardnameLabeltext(prStoryCard.Equipment_Two) = ""
                updateCardList(5, ScrollBase_Card3, DowntimeField.Tier1, "x", "x")
            End If
        End If
        updateCard_Unplanned_Tier3_All()
        showEquipmentlabels()
        RefreshHeaderColors()
        Card4Header.Background = LabelSelectedColor
        sender.background = LabelSelectedColor
        Tier2Clicked_Unplanned = sender.content.ToString
        'DecidetoShowNavigationButtons()
        If prStoryReport.getCardEventNumber(1) <= 6 Then NavigationRight_Card1.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(3) <= 6 Then NavigationRight_Card3.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(4) <= 3 Then NavigationRight_Card4.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(5) <= 3 Then NavigationRight_Card5.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(6) <= 3 Then NavigationRight_Card6.Visibility = Visibility.Hidden


    End Sub

    Private Sub Card1ScrollClick(sender As Object, e As MouseButtonEventArgs)


        If ScrollBase_Card1 = 0 Then NavigationLeft_card1.Visibility = Visibility.Hidden

        'If prStoryReport.getCardEventNumber(1) - 1 > 6 Then
        If sender Is NavigationRight_Card1 Then
            ScrollBase_Card1 = ScrollBase_Card1 + 1
        ElseIf sender Is NavigationLeft_card1 And ScrollBase_Card1 <> 0 Then
            ScrollBase_Card1 = ScrollBase_Card1 - 1
        End If




        If ScrollBase_Card1 = 0 Then NavigationLeft_card1.Visibility = Visibility.Hidden
        updateCardList(1, ScrollBase_Card1, DowntimeField.Tier1, "", "")
        '    createPDTchart(1, prStoryReport, False)
        updateCard_Unplanned_Tier1()
        showUPDTlabels()

        If prStoryReport.getCardEventNumber(1) + 2 - 6 = ScrollBase_Card1 Then

            NavigationRight_Card1.Visibility = Visibility.Hidden
        End If
        If prStoryReport.getCardEventNumber(1) <= 6 Then NavigationRight_Card1.Visibility = Visibility.Hidden
        If UPDTlabel6.Content.ToString = OTHERS_STRING Then NavigationRight_Card1.Visibility = Visibility.Hidden

        ScrollBase_Card4 = 0
        ScrollBase_Card5 = 0
        ScrollBase_Card6 = 0
    End Sub
    Private Sub Card3ScrollClick(sender As Object, e As MouseButtonEventArgs)

        If ScrollBase_Card3 = 0 Then NavigationLeft_card3.Visibility = Visibility.Hidden


        'If prStoryReport.getCardEventNumber(1) - 1 > 6 Then
        If sender Is NavigationRight_Card3 Then
            ScrollBase_Card3 = ScrollBase_Card3 + 1
        ElseIf sender Is NavigationLeft_card3 And ScrollBase_Card3 <> 0 Then
            ScrollBase_Card3 = ScrollBase_Card3 - 1
        End If




        If ScrollBase_Card3 = 0 Then NavigationLeft_card3.Visibility = Visibility.Hidden
        updateCardList(3, ScrollBase_Card3, DowntimeField.Tier2, Card3Header.Content, "")
        '     createPDTchart(3, prStoryReport, False)
        updateCard_Unplanned_Tier2()
        showEquipmentlabels()

        If prStoryReport.getCardEventNumber(3) - 6 = ScrollBase_Card3 Then
            NavigationRight_Card3.Visibility = Visibility.Hidden

        End If

        'If .Content.ToString = OTHERS_STRING Then NavigationRight_Card1.Visibility = Windows.Visibility.Hidden

        ScrollBase_Card4 = 0
        ScrollBase_Card5 = 0
        ScrollBase_Card6 = 0


    End Sub

    Private Sub Card4ScrollClick(sender As Object, e As MouseButtonEventArgs)

        If ScrollBase_Card4 = 0 Then NavigationLeft_card4.Visibility = Visibility.Hidden


        'If prStoryReport.getCardEventNumber(1) - 1 > 6 Then
        If sender Is NavigationRight_Card4 Then
            ScrollBase_Card4 = ScrollBase_Card4 + 1
        ElseIf sender Is NavigationLeft_card4 And ScrollBase_Card4 <> 0 Then
            ScrollBase_Card4 = ScrollBase_Card4 - 1
        End If


        If ScrollBase_Card4 = 0 Then NavigationLeft_card4.Visibility = Visibility.Hidden
        updateCardList(4, ScrollBase_Card4, DowntimeField.Tier3, Card3Header.Content, Card4Header.Content)
        '     createPDTchart(4, prStoryReport, False)
        updateCard_Unplanned_Tier3_All()
        showEquipmentlabels()

        If prStoryReport.getCardEventNumber(4) - 3 = ScrollBase_Card4 Then
            NavigationRight_Card4.Visibility = Visibility.Hidden

        End If

        'If .Content.ToString = OTHERS_STRING Then NavigationRight_Card1.Visibility = Windows.Visibility.Hidden


    End Sub
    Private Sub Card5ScrollClick(sender As Object, e As MouseButtonEventArgs)

        If ScrollBase_Card5 = 0 Then NavigationLeft_card5.Visibility = Visibility.Hidden


        'If prStoryReport.getCardEventNumber(1) - 1 > 6 Then
        If sender Is NavigationRight_Card5 Then
            ScrollBase_Card5 = ScrollBase_Card5 + 1
        ElseIf sender Is NavigationLeft_card5 And ScrollBase_Card5 <> 0 Then
            ScrollBase_Card5 = ScrollBase_Card5 - 1
        End If


        If ScrollBase_Card5 = 0 Then NavigationLeft_card5.Visibility = Visibility.Hidden
        updateCardList(5, ScrollBase_Card5, DowntimeField.Tier3, Card3Header.Content, Card5Header.Content)
        '     createPDTchart(5, prStoryReport, False)
        updateCard_Unplanned_Tier3_All()
        showEquipmentlabels()

        If prStoryReport.getCardEventNumber(5) - 3 = ScrollBase_Card5 Then
            NavigationRight_Card5.Visibility = Visibility.Hidden

        End If

        'If .Content.ToString = OTHERS_STRING Then NavigationRight_Card1.Visibility = Windows.Visibility.Hidden


    End Sub
    Private Sub Card6ScrollClick(sender As Object, e As MouseButtonEventArgs)

        If ScrollBase_Card6 = 0 Then NavigationLeft_card6.Visibility = Visibility.Hidden


        'If prStoryReport.getCardEventNumber(1) - 1 > 6 Then
        If sender Is NavigationRight_Card6 Then
            ScrollBase_Card6 = ScrollBase_Card6 + 1
        ElseIf sender Is NavigationLeft_card6 And ScrollBase_Card6 <> 0 Then
            ScrollBase_Card6 = ScrollBase_Card6 - 1
        End If


        If ScrollBase_Card6 = 0 Then NavigationLeft_card6.Visibility = Visibility.Hidden
        updateCardList(6, ScrollBase_Card6, DowntimeField.Tier3, Card3Header.Content, Card6Header.Content)
        '     createPDTchart(6, prStoryReport, False)
        updateCard_Unplanned_Tier3_All()
        showEquipmentlabels()

        If prStoryReport.getCardEventNumber(6) - 3 = ScrollBase_Card6 Then
            NavigationRight_Card6.Visibility = Visibility.Hidden

        End If

        'If .Content.ToString = OTHERS_STRING Then NavigationRight_Card1.Visibility = Windows.Visibility.Hidden


    End Sub
#End Region

#Region "GraphingVariables"

    Private Const stopbubbleMAXsize = 100
    Private Const stopbubbleMAXsize_expanded = 75
    Private Const barchartBarMaxSize = 160
    Private Const barchartBarMaxSize_stops = 350
    Private Const changeoverbubbleMaxSize = 50
    Private Const topstopbubbledefault_Height = 15

    Private SelectedFailuremonumber_inTopStopsforTrends = 0

    Private topstopsbar1 As Double
    Private topstopsbar2 As Double
    Private topstopsbar3 As Double
    Private topstopsbar4 As Double
    Private topstopsbar5 As Double
    Private topstopsbar6 As Double
    Private topstopsbar7 As Double
    Private topstopsbar8 As Double
    Private topstopsbar9 As Double
    Private topstopsbar10 As Double
    Private topstopsbar11 As Double
    Private topstopsbar12 As Double
    Private topstopsbar13 As Double
    Private topstopsbar14 As Double
    Private topstopsbar15 As Double

    Private topstopsbar1_PRloss As Double
    Private topstopsbar2_PRloss As Double
    Private topstopsbar3_PRloss As Double
    Private topstopsbar4_PRloss As Double
    Private topstopsbar5_PRloss As Double
    Private topstopsbar6_PRloss As Double
    Private topstopsbar7_PRloss As Double
    Private topstopsbar8_PRloss As Double
    Private topstopsbar9_PRloss As Double
    Private topstopsbar10_PRloss As Double
    Private topstopsbar11_PRloss As Double
    Private topstopsbar12_PRloss As Double
    Private topstopsbar13_PRloss As Double
    Private topstopsbar14_PRloss As Double
    Private topstopsbar15_PRloss As Double

    Private topstopsbar1_DTmin As Double
    Private topstopsbar2_DTmin As Double
    Private topstopsbar3_DTmin As Double
    Private topstopsbar4_DTmin As Double
    Private topstopsbar5_DTmin As Double
    Private topstopsbar6_DTmin As Double
    Private topstopsbar7_DTmin As Double
    Private topstopsbar8_DTmin As Double
    Private topstopsbar9_DTmin As Double
    Private topstopsbar10_DTmin As Double
    Private topstopsbar11_DTmin As Double
    Private topstopsbar12_DTmin As Double
    Private topstopsbar13_DTmin As Double
    Private topstopsbar14_DTmin As Double
    Private topstopsbar15_DTmin As Double


    Private topstop_mttr(0 To 14) As Double
    Private topstop_mtbf(0 To 14) As Double


    Private topstopsbar1_SPD As Double
    Private topstopsbar2_SPD As Double
    Private topstopsbar3_SPD As Double
    Private topstopsbar4_SPD As Double
    Private topstopsbar5_SPD As Double
    Private topstopsbar6_SPD As Double
    Private topstopsbar7_SPD As Double
    Private topstopsbar8_SPD As Double
    Private topstopsbar9_SPD As Double
    Private topstopsbar10_SPD As Double
    Private topstopsbar11_SPD As Double
    Private topstopsbar12_SPD As Double
    Private topstopsbar13_SPD As Double
    Private topstopsbar14_SPD As Double
    Private topstopsbar15_SPD As Double

    Private topstopsbar1_Stops As Double
    Private topstopsbar2_Stops As Double
    Private topstopsbar3_Stops As Double
    Private topstopsbar4_Stops As Double
    Private topstopsbar5_Stops As Double
    Private topstopsbar6_Stops As Double
    Private topstopsbar7_Stops As Double
    Private topstopsbar8_Stops As Double
    Private topstopsbar9_Stops As Double
    Private topstopsbar10_Stops As Double
    Private topstopsbar11_Stops As Double
    Private topstopsbar12_Stops As Double
    Private topstopsbar13_Stops As Double
    Private topstopsbar14_Stops As Double
    Private topstopsbar15_Stops As Double

    Private updtlabelstring(0 To 5) As String  ' new array for replacing independent variables
    Private updtlabel1string As String
    Private updtlabel2string As String
    Private updtlabel3string As String
    Private updtlabel4string As String
    Private updtlabel5string As String
    Private updtlabel6string As String
    Private UPDTbar1 As Double
    Private UPDTbar2 As Double
    Private UPDTbar3 As Double
    Private UPDTbar4 As Double
    Private UPDTbar5 As Double
    Private UPDTbar6 As Double
    Private UPDTbar1_PRloss As Double
    Private UPDTbar2_PRloss As Double
    Private UPDTbar3_PRloss As Double
    Private UPDTbar4_PRloss As Double
    Private UPDTbar5_PRloss As Double
    Private UPDTbar6_PRloss As Double

    Private UPDTbar1sim As Double
    Private UPDTbar2sim As Double
    Private UPDTbar3sim As Double
    Private UPDTbar4sim As Double
    Private UPDTbar5sim As Double
    Private UPDTbar6sim As Double
    Private UPDTbar1_PRlosssim As Double
    Private UPDTbar2_PRlosssim As Double
    Private UPDTbar3_PRlosssim As Double
    Private UPDTbar4_PRlosssim As Double
    Private UPDTbar5_PRlosssim As Double
    Private UPDTbar6_PRlosssim As Double





    Private Target_Prloss(0 To 9) As Double
    Private Target_PRloss_Tooltip(0 To 9) As String




    Private pdtlabelstring(0 To 8) As String  ' new array for replacing independent variables
    Private pdtlabel1string As String
    Private pdtlabel2string As String
    Private pdtlabel3string As String
    Private pdtlabel4string As String
    Private pdtlabel5string As String
    Private pdtlabel6string As String
    Private pdtlabel7string As String
    Private pdtlabel8string As String
    Private pdtlabel9string As String
    Private PDTbar1 As Double
    Private PDTbar2 As Double
    Private PDTbar3 As Double
    Private PDTbar4 As Double
    Private PDTbar5 As Double
    Private PDTbar6 As Double
    Private PDTbar7 As Double
    Private PDTbar8 As Double
    Private PDTbar9 As Double
    Private PDTbar1_PRloss As Double
    Private PDTbar2_PRloss As Double
    Private PDTbar3_PRloss As Double
    Private PDTbar4_PRloss As Double
    Private PDTbar5_PRloss As Double
    Private PDTbar6_PRloss As Double
    Private PDTbar7_PRloss As Double
    Private PDTbar8_PRloss As Double
    Private PDTbar9_PRloss As Double

    Private PDTbar1sim As Double
    Private PDTbar2sim As Double
    Private PDTbar3sim As Double
    Private PDTbar4sim As Double
    Private PDTbar5sim As Double
    Private PDTbar6sim As Double
    Private PDTbar7sim As Double
    Private PDTbar8sim As Double
    Private PDTbar9sim As Double
    Private PDTbar1_PRlosssim As Double
    Private PDTbar2_PRlosssim As Double
    Private PDTbar3_PRlosssim As Double
    Private PDTbar4_PRlosssim As Double
    Private PDTbar5_PRlosssim As Double
    Private PDTbar6_PRlosssim As Double
    Private PDTbar7_PRlosssim As Double
    Private PDTbar8_PRlosssim As Double
    Private PDTbar9_PRlosssim As Double


    Private changeoverlabelstring(0 To 6) As String  ' new array for replacing independent variables
    Private changeoverlabel1string As String
    Private changeoverlabel2string As String
    Private changeoverlabel3string As String
    Private changeoverlabel4string As String
    Private changeoverlabel5string As String
    Private changeoverlabel6string As String
    Private changeoverlabel7string As String
    Private Changeoverbar1 As Double
    Private Changeoverbar2 As Double
    Private Changeoverbar3 As Double
    Private Changeoverbar4 As Double
    Private Changeoverbar5 As Double
    Private Changeoverbar6 As Double
    Private Changeoverbar7 As Double
    Private Changeover1_PRloss As Double
    Private Changeover2_PRloss As Double
    Private Changeover3_PRloss As Double
    Private Changeover4_PRloss As Double
    Private Changeover5_PRloss As Double
    Private Changeover6_PRloss As Double
    Private Changeover7_PRloss As Double


    Private Changeoverbar1sim As Double
    Private Changeoverbar2sim As Double
    Private Changeoverbar3sim As Double
    Private Changeoverbar4sim As Double
    Private Changeoverbar5sim As Double
    Private Changeoverbar6sim As Double
    Private Changeoverbar7sim As Double
    Private Changeover1_PRlosssim As Double
    Private Changeover2_PRlosssim As Double
    Private Changeover3_PRlosssim As Double
    Private Changeover4_PRlosssim As Double
    Private Changeover5_PRlosssim As Double
    Private Changeover6_PRlosssim As Double
    Private Changeover7_PRlosssim As Double

    Private Changeovertime_TotalAvg As Double
    Private Changeovertime1 As Double
    Private Changeovertime2 As Double
    Private Changeovertime3 As Double
    Private Changeovertime4 As Double
    Private Changeovertime5 As Double
    Private Changeovertime6 As Double
    Private Changeovertime7 As Double
    Private changeoverNo_of_Events1 As Integer
    Private changeoverNo_of_Events2 As Integer
    Private changeoverNo_of_Events3 As Integer
    Private changeoverNo_of_Events4 As Integer
    Private changeoverNo_of_Events5 As Integer
    Private changeoverNo_of_Events6 As Integer
    Private changeoverNo_of_Events7 As Integer


    Private EquipMainlabelstring(0 To 5) As String  ' new array for replacing independent variables
    Private EquipmentLabel1string As String
    Private EquipmentLabel2string As String
    Private EquipmentLabel3string As String
    Private EquipmentLabel4string As String
    Private EquipmentLabel5string As String
    Private EquipmentLabel6string As String
    Private EquipMainbar1 As Double
    Private EquipMainbar2 As Double
    Private EquipMainbar3 As Double
    Private EquipMainbar4 As Double
    Private EquipMainbar5 As Double
    Private EquipMainbar6 As Double
    Private EquipMain1_PRloss As Double
    Private EquipMain2_PRloss As Double
    Private EquipMain3_PRloss As Double
    Private EquipMain4_PRloss As Double
    Private EquipMain5_PRloss As Double
    Private EquipMain6_PRloss As Double

    Private EquipMainbar1sim As Double
    Private EquipMainbar2sim As Double
    Private EquipMainbar3sim As Double
    Private EquipMainbar4sim As Double
    Private EquipMainbar5sim As Double
    Private EquipMainbar6sim As Double
    Private EquipMain1_PRlosssim As Double
    Private EquipMain2_PRlosssim As Double
    Private EquipMain3_PRlosssim As Double
    Private EquipMain4_PRlosssim As Double
    Private EquipMain5_PRlosssim As Double
    Private EquipMain6_PRlosssim As Double

    Private Equip1labelstring(0 To 2) As String  ' new array for replacing independent variables
    Private Equip1Label1string As String
    Private Equip1Label2string As String
    Private Equip1Label3string As String
    Private Equip1bar1 As Double
    Private Equip1bar2 As Double
    Private Equip1bar3 As Double
    Private Equip1_1_PRloss As Double
    Private Equip1_2_PRloss As Double
    Private Equip1_3_PRloss As Double
    Private Equip1bar1sim As Double
    Private Equip1bar2sim As Double
    Private Equip1bar3sim As Double
    Private Equip1_1_PRlosssim As Double
    Private Equip1_2_PRlosssim As Double
    Private Equip1_3_PRlosssim As Double


    Private Equip2labelstring(0 To 2) As String  ' new array for replacing independent variables
    Private Equip2Label1string As String
    Private Equip2Label2string As String
    Private Equip2Label3string As String
    Private Equip2bar1 As Double
    Private Equip2bar2 As Double
    Private Equip2bar3 As Double
    Private Equip2_1_PRloss As Double
    Private Equip2_2_PRloss As Double
    Private Equip2_3_PRloss As Double
    Private Equip2bar1sim As Double
    Private Equip2bar2sim As Double
    Private Equip2bar3sim As Double
    Private Equip2_1_PRlosssim As Double
    Private Equip2_2_PRlosssim As Double
    Private Equip2_3_PRlosssim As Double

    Private Equip3labelstring(0 To 2) As String  ' new array for replacing independent variables
    Private Equip3Label1string As String
    Private Equip3Label2string As String
    Private Equip3Label3string As String
    Private Equip3bar1 As Double
    Private Equip3bar2 As Double
    Private Equip3bar3 As Double
    Private Equip3_1_PRloss As Double
    Private Equip3_2_PRloss As Double
    Private Equip3_3_PRloss As Double
    Private Equip3bar1sim As Double
    Private Equip3bar2sim As Double
    Private Equip3bar3sim As Double
    Private Equip3_1_PRlosssim As Double
    Private Equip3_2_PRlosssim As Double
    Private Equip3_3_PRlosssim As Double


    Private Equip21labelstring(0 To 2) As String  ' new array for replacing independent variables
    Private Equip21label1string As String
    Private Equip21label2string As String
    Private Equip21bar1 As Double
    Private Equip21bar2 As Double
    Private Equip21_1_PRloss As Double
    Private Equip21_2_PRloss As Double

    Private Equip22labelstring(0 To 2) As String ' new array for replacing independent variables
    Private Equip22label1string As String
    Private Equip22label2string As String
    Private Equip22bar1 As Double
    Private Equip22bar2 As Double
    Private Equip22_1_PRloss As Double
    Private Equip22_2_PRloss As Double


    Private MaxPRindataset As Double ' = prStoryReport.UPDT
    Private MaxPRindataset_planned As Double '= prStoryReport.PDT
    Private MaxStopsindataset As Double
    Private MaxPRlossindataset_stops As Double
    Private MaxChangeoverBubble_COtime As Double

    Private selectedDTGroup As String
    Private selectedRL4bar As String
    Private selectedRLcolumn As Integer

    Private ScrollBase_Card1 As Integer = 0
    Private ScrollBase_Card3 As Integer = 0
    Private ScrollBase_Card4 As Integer = 0
    Private ScrollBase_Card5 As Integer = 0
    Private ScrollBase_Card6 As Integer = 0

#End Region

#Region "RawDataWindow"
    'Comment Window
    Dim _CommentCollection As New ObservableCollection(Of DTeventFields)()
    Public tempbubblesender As Object
    Private _lastHeaderClicked_Comment As GridViewColumnHeader = Nothing
    Private _lastDirection_Comment As ListSortDirection = ListSortDirection.Ascending
    Public ReadOnly Property CommentCollection() As ObservableCollection(Of DTeventFields)
        Get
            Return _CommentCollection
        End Get
    End Property
    Private Sub transferToCommentColelction(tmpEvent As DTevent)
        Dim commentIncrementer As Integer
        _CommentCollection.Clear()
        With tmpEvent
            For commentIncrementer = 0 To .RawRows.Count - 1 '.RawInfo.Count - 1
                _CommentCollection.Add(New DTeventFields(prStoryReport.MainLEDSReport.DT_Report.rawDTdata.UnplannedData(.RawRows(commentIncrementer))))
            Next
        End With
    End Sub
    Private Sub GridViewColumnHeaderClickedHandler_Comment(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim headerClicked As GridViewColumnHeader = TryCast(e.OriginalSource, GridViewColumnHeader)
        Dim direction As ListSortDirection

        If headerClicked IsNot Nothing Then

            'BusyIndicator.Visibility = Visibility.Visible
            'BusyIndicator.IsBusy = True

            If headerClicked.Role <> GridViewColumnHeaderRole.Padding Then
                If headerClicked IsNot _lastHeaderClicked_Comment Then
                    direction = ListSortDirection.Ascending
                Else
                    If _lastDirection_Comment = ListSortDirection.Ascending Then
                        direction = ListSortDirection.Descending
                    Else
                        direction = ListSortDirection.Ascending
                    End If
                End If

                Dim header As String = TryCast(headerClicked.Column.Header, String)
                Sort_Comments(header, direction)

                If direction = ListSortDirection.Ascending Then
                    headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowUp"), DataTemplate)
                Else
                    headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowDown"), DataTemplate)
                End If

                ' Remove arrow from previously sorted header
                If _lastHeaderClicked_Comment IsNot Nothing AndAlso _lastHeaderClicked_Comment IsNot headerClicked Then
                    _lastHeaderClicked_Comment.Column.HeaderTemplate = Nothing
                End If


                _lastHeaderClicked_Comment = headerClicked
                _lastDirection_Comment = direction
            End If
        End If
    End Sub
    Private Sub ReinitializeEquipCardsXyx(Optional card4 As Boolean = True, Optional card5 As Boolean = True, Optional card6 As Boolean = True)

        If card4 = True Then
            Equip1labelstring(0) = ""
            Equip1labelstring(1) = ""
            Equip1labelstring(2) = ""
            Equip1Label1string = ""
            Equip1Label2string = ""
            Equip1Label3string = ""
            Equip1bar1 = 0
            Equip1bar2 = 0
            Equip1bar3 = 0
            Equip1_1_PRloss = 0
            Equip1_2_PRloss = 0
            Equip1_3_PRloss = 0
        End If

        If card5 = True Then
            Equip2labelstring(0) = ""
            Equip2labelstring(1) = ""
            Equip2labelstring(2) = ""
            Equip2Label1string = ""
            Equip2Label2string = ""
            Equip2Label3string = ""
            Equip2bar1 = 0
            Equip2bar2 = 0
            Equip2bar3 = 0
            Equip2_1_PRloss = 0
            Equip2_2_PRloss = 0
            Equip2_3_PRloss = 0
        End If


        If card6 = True Then
            Equip3labelstring(0) = ""
            Equip3labelstring(1) = ""
            Equip3labelstring(2) = ""
            Equip3Label1string = ""
            Equip3Label2string = ""
            Equip3Label3string = ""
            Equip3bar1 = 0
            Equip3bar2 = 0
            Equip3bar3 = 0
            Equip3_1_PRloss = 0
            Equip3_2_PRloss = 0
            Equip3_3_PRloss = 0

        End If
    End Sub

    Private Sub Sort_Comments(ByVal sortBy As String, ByVal direction As ListSortDirection)
        Dim dataView As ICollectionView = CollectionViewSource.GetDefaultView(CommentList.ItemsSource)

        dataView.SortDescriptions.Clear()
        Dim sd As New SortDescription(sortBy, direction)
        dataView.SortDescriptions.Add(sd)
        dataView.Refresh()
    End Sub
#End Region

#Region "Menu 2.0"
    Public Menuitemclicked_number As Integer = -1

    Public Shared mybrushverylightgray_forcardbackground As New SolidColorBrush(Windows.Media.Color.FromRgb(248, 248, 248))

    Public multilineGroups As List(Of ProductionLineGroup)
    Private Sub LaunchMultiLineWindow()


        Dim multilinewindow As New Window_MultiLine
        multilinewindow.InitializeMultilineGroups(multilineGroups)
        multilinewindow.Owner = Me
        Me.Visibility = Visibility.Hidden
        multilinewindow.Show()
    End Sub

    Public Sub LaunchMenu(sender As Object, e As MouseButtonEventArgs)
        MenuCanvas.Visibility = Visibility.Visible
        AnimateMenuOpening()
        MenuSplashRectangle.Visibility = Visibility.Visible
    End Sub

    Public Sub AnimateMenuOpening()
        AnimateZoomUIElement_Margin(New Thickness(-280, 0, 1218, -4), New Thickness(0, 0, 938, -4), 0.15, MarginProperty, MenuCanvas)
    End Sub

    Public Sub CloseMenu(sender As Object, e As MouseButtonEventArgs)
        CloseMenu()
    End Sub

    Private Sub CloseMenu()
        MenuSplashRectangle.Visibility = Visibility.Hidden
        System.Windows.Forms.Application.DoEvents()
        MenuCanvas.Visibility = Visibility.Hidden
    End Sub
    Public Sub Menuitemmousemove(sender As Object, e As MouseEventArgs)
        Dim menuitem As Integer = -1
        If sender.[GetType]().ToString().IndexOf("Image") > -1 Then
            Dim tempsender As System.Windows.Controls.Image = DirectCast(sender, System.Windows.Controls.Image)
            menuitem = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name))
            If menuitem <> -1 Then

                getMenuItem_Label_fromitemindex(getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, -1, "", "Menu" & menuitem), -1, "", "Menu" & menuitem & "Label").Foreground = Media.Brushes.Orange
                getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, -1, "", "Menu" & menuitem).Background = mybrushverylightgray_forcardbackground
            End If
        ElseIf sender.[GetType]().ToString().IndexOf("Label") > -1 Then
            Dim tempsender As Label = DirectCast(sender, Label)
            menuitem = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name))
            tempsender.Foreground = Media.Brushes.Orange
            getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, -1, "", "Menu" & menuitem).Background = mybrushverylightgray_forcardbackground
        End If
    End Sub
    Public Sub Menuitemmouseleave(sender As Object, e As MouseEventArgs)
        Dim menuitem As Integer = -1
        If sender.[GetType]().ToString().IndexOf("Image") > -1 Then
            Dim tempsender As System.Windows.Controls.Image = DirectCast(sender, System.Windows.Controls.Image)
            menuitem = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name))
            If menuitem <> Menuitemclicked_number Then
                If menuitem <> -1 Then

                    getMenuItem_Label_fromitemindex(getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, -1, "", "Menu" & menuitem), -1, "", "Menu" & menuitem & "Label").Foreground = BrushColors.mybrushfontgray
                    getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, -1, "", "Menu" & menuitem).Background = Media.Brushes.White
                End If

            End If
        ElseIf sender.[GetType]().ToString().IndexOf("Label") > -1 Then
            Dim tempsender As Label = DirectCast(sender, Label)
            menuitem = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name))
            If menuitem <> Menuitemclicked_number Then
                tempsender.Foreground = BrushColors.mybrushfontgray
                '  getMenuItem_Image_fromitemindex(Menu_InternalInfiniteCanvas, menuitem);
                getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, -1, "", "Menu" & menuitem).Background = Media.Brushes.White
            End If
        End If
    End Sub
    Public Sub Menuitemclicked(sender As Object, e As MouseButtonEventArgs)
        restore_allmenuitems_color()


        Dim menuitem As Integer = -1
        If sender.[GetType]().ToString().IndexOf("Image") > -1 Then
            Dim tempsender As System.Windows.Controls.Image = DirectCast(sender, System.Windows.Controls.Image)
            menuitem = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name))
            If menuitem <> -1 Then
                Dim tempcanvas As Canvas
                tempcanvas = getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, menuitem)
                tempcanvas.Background = mybrushverylightgray_forcardbackground
                getMenuItem_Label_fromitemindex(tempcanvas, menuitem).Foreground = Media.Brushes.Orange
            End If
        ElseIf sender.[GetType]().ToString().IndexOf("Label") > -1 Then
            Dim tempsender As Label = DirectCast(sender, Label)
            menuitem = Convert.ToInt32(GlobalFcns.onlyDigits(tempsender.Name))
            '  getMenuItem_Image_fromitemindex(Menu_InternalInfiniteCanvas, menuitem);
            tempsender.Foreground = Media.Brushes.Orange

            getMenuItem_Canvas_fromitemindex(Menu_InternalInfiniteCanvas, menuitem).Background = mybrushverylightgray_forcardbackground
        End If
        Menuitemclicked_number = menuitem
        CloseMenu()
    End Sub

    Public Sub restore_allmenuitems_color()
        Menu1.Background = Media.Brushes.White
        Menu1Label.Foreground = BrushColors.mybrushfontgray

        Menu2.Background = Media.Brushes.White
        Menu2Label.Foreground = BrushColors.mybrushfontgray

        Menu3.Background = Media.Brushes.White
        Menu3Label.Foreground = BrushColors.mybrushfontgray

        Menu4.Background = Media.Brushes.White
        Menu4Label.Foreground = BrushColors.mybrushfontgray

        Menu5.Background = Media.Brushes.White
        Menu5Label.Foreground = BrushColors.mybrushfontgray

        'Menu6.Background = Media.Brushes.White
        'Menu6Label.Foreground = BrushColors.mybrushfontgray

        'Menu7.Background = Media.Brushes.White
        'Menu7Label.Foreground = BrushColors.mybrushfontgray

        '     Menu8.Background = Media.Brushes.White
        '    Menu8Label.Foreground = BrushColors.mybrushfontgray


        '   Menu9.Background = Media.Brushes.White
        '  Menu9Label.Foreground = BrushColors.mybrushfontgray

    End Sub

    Public Sub AnimateZoomUIElement_Margin(from As Thickness, [to] As Thickness, durn As Double, depprop As DependencyProperty, AnimatedObject As UIElement)
        ' Standard animation function
        Dim da = New ThicknessAnimation()
        ' da will contain the characteristics of the animation
        da.From = from
        ' position, where it starts 
        da.[To] = [to]
        ' position, where it ends
        da.Duration = New Duration(TimeSpan.FromSeconds(durn))
        ' how long animation lasts
        AnimatedObject.BeginAnimation(depprop, da)
        ' Animate object is the subject we are playing with. And Depprop determines what type of UI element it is (rectangle, label, control.. etc)
    End Sub

    Public Function getMenuItem_Label_fromitemindex(dep As DependencyObject, Optional menuitemindex As Integer = -1, Optional stringitemindex As String = "", Optional exactstring As String = "") As Label
        Dim sender As Label = Nothing
        Dim lbl As Label
        Dim j As Integer
        For j = 0 To VisualTreeHelper.GetChildrenCount(dep) - 1
            If VisualTreeHelper.GetChild(dep, j).[GetType]().ToString().IndexOf("Label") > -1 Then
                lbl = DirectCast(VisualTreeHelper.GetChild(dep, j), Label)

                If stringitemindex = "" AndAlso menuitemindex <> -1 Then
                    If lbl.Name.IndexOf(menuitemindex.ToString()) > -1 Then
                        sender = lbl
                    End If
                Else
                    If exactstring = "" Then
                        If lbl.Name.IndexOf(stringitemindex) > -1 Then
                            sender = lbl
                        End If
                    Else
                        If lbl.Name.ToString() = exactstring Then
                            sender = lbl

                        End If

                    End If
                End If

            End If
        Next

        Return sender
    End Function
    Public Function getMenuItem_Canvas_fromitemindex(dep As DependencyObject, Optional menuitemindex As Integer = -1, Optional stringitemindex As String = "", Optional exactstring As String = "") As Canvas
        Dim sender As Canvas = Nothing
        Dim cnv As Canvas
        Dim j As Integer
        For j = 0 To VisualTreeHelper.GetChildrenCount(dep) - 1
            If VisualTreeHelper.GetChild(dep, j).[GetType]().ToString().IndexOf("Canvas") > -1 Then
                cnv = DirectCast(VisualTreeHelper.GetChild(dep, j), Canvas)

                If stringitemindex = "" AndAlso menuitemindex <> -1 Then
                    If cnv.Name.IndexOf(menuitemindex.ToString()) > -1 Then
                        sender = cnv
                    End If
                Else
                    If exactstring = "" Then
                        If cnv.Name.IndexOf(stringitemindex) > -1 Then
                            sender = cnv
                        End If
                    Else
                        If cnv.Name.ToString() = exactstring Then
                            sender = cnv

                        End If

                    End If
                End If

            End If
        Next

        Return sender
    End Function


#End Region



    Sub LaunchExpandedIncontrol()
        If MasterDataSet.maxDTpct = 0 Then
            MsgBox("No failre modes selected to generate control chart.")
            Exit Sub
        End If

        Dim expandedincontrol As New Window_IncontrolExpanded
        HTML.createSPCchart(bubblenumberpublic)
        UseTrack_IncontrolControlChart = True
        expandedincontrol.ShowDialog()

    End Sub

    Public Sub showRateLossError()
        MsgBox("Blocked/Starved Data Error.")
    End Sub

    '\\\\\\\\\\\\\\\\\\\CONSTUCTOR\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Sub New(storyReport As prStoryMainPageReport)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add the event handler for handling UI thread exceptions to the event. 
        ';AddHandler CLS_ErrorHandling.dateRangeException, AddressOf WellThisSucks

        ' Add the event handler for handling non-UI thread exceptions to the event.  
        'TAKEN OUT FOR TESTING GLEDS ->     AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CatchAllUnhandledErrorsHere

        ' AddHandler Application.

        ' Add any initialization after the InitializeComponent() call.
        prStoryReport = storyReport
        Title = AllProdLines(selectedindexofLine_temp).Name & " Line - prstory"

        'set the max fields for the bar charts
        MaxPRindataset = prStoryReport.UPDT
        MaxPRindataset_planned = prStoryReport.PDT

        prStoryReport.setBargraphReportWindow(Me)
        '  AllPRSTORYMainPageReports.Add(storyReport)
        prStoryReport.InitializeBargraphWindowConnection_createAllLists()
        MaxStopsindataset = Card_TopStops(0).SPD

        'update the graphics
        updateCard_Unplanned_Tier1()
        updateCard_Planned_Tier1()
        updateCard_TopStops()

        If My.Settings.EnableTimeSpanExclusion Then
            TimeRangeIcon.Visibility = Visibility.Visible
        End If

    End Sub
    Private Sub IconMouseMove(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Hand

    End Sub
    Private Sub IconMouseLeave(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Arrow

    End Sub
    Private Sub CatchAllUnhandledErrorsHere()
        showErrorSplashPage()
    End Sub

    Private Sub CheckScreenResolution()
        Dim screenWidth As Integer = My.Computer.Screen.Bounds.Width
        Dim screenHeight As Integer = My.Computer.Screen.Bounds.Height

        If screenWidth < 1200 Or screenHeight < 700 Then Me.WindowState = Windows.WindowState.Maximized

    End Sub

#Region "Update Cards From Event Lists"
    Private Sub updateCard_Unplanned_Tier1()
        Dim tmpDtEvent As DTevent
        'clear tier 3 labels
        cardnameLabeltext(prStoryCard.Unplanned) = prStoryReport.getCardName(prStoryCard.Unplanned)
        If MaxPRindataset = 0 Then MaxPRindataset = 1

        For i As Integer = 0 To prStoryCardFields.Unplanned - 1
            tmpDtEvent = Card_Unplanned_T1(i)
            Select Case i
                Case 0
                    updtlabel1string = tmpDtEvent.Name
                    UPDTbar1 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar1_PRloss = tmpDtEvent.DTpct
                    UPDTbar1sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar1_PRlosssim = tmpDtEvent.DTpctSim
                Case 1
                    updtlabel2string = tmpDtEvent.Name
                    UPDTbar2 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar2_PRloss = tmpDtEvent.DTpct
                    UPDTbar2sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar2_PRlosssim = tmpDtEvent.DTpctSim
                Case 2
                    updtlabel3string = tmpDtEvent.Name
                    UPDTbar3 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar3_PRloss = tmpDtEvent.DTpct
                    UPDTbar3sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar3_PRlosssim = tmpDtEvent.DTpctSim
                Case 3
                    updtlabel4string = tmpDtEvent.Name
                    UPDTbar4 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar4_PRloss = tmpDtEvent.DTpct
                    UPDTbar4sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar4_PRlosssim = tmpDtEvent.DTpctSim
                Case 4
                    updtlabel5string = tmpDtEvent.Name
                    UPDTbar5 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar5_PRloss = tmpDtEvent.DTpct
                    UPDTbar5sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar5_PRlosssim = tmpDtEvent.DTpctSim
                Case 5
                    updtlabel6string = tmpDtEvent.Name
                    UPDTbar6 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar6_PRloss = tmpDtEvent.DTpct
                    UPDTbar6sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    UPDTbar6_PRlosssim = tmpDtEvent.DTpctSim
            End Select
        Next
    End Sub
    Private Sub updateCard_Unplanned_Tier2()
        Dim tmpDtEvent As DTevent
        For i As Integer = 0 To prStoryCardFields.Equipment - 1
            tmpDtEvent = Card_Unplanned_T2(i)
            Select Case i
                Case 0
                    EquipmentLabel1string = tmpDtEvent.Name
                    EquipMainbar1 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain1_PRloss = tmpDtEvent.DTpct
                    EquipMainbar1sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain1_PRlosssim = tmpDtEvent.DTpctSim
                Case 1
                    EquipmentLabel2string = tmpDtEvent.Name
                    EquipMainbar2 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain2_PRloss = tmpDtEvent.DTpct
                    EquipMainbar2sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain2_PRlosssim = tmpDtEvent.DTpctSim
                Case 2
                    EquipmentLabel3string = tmpDtEvent.Name
                    EquipMainbar3 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain3_PRloss = tmpDtEvent.DTpct
                    EquipMainbar3sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain3_PRlosssim = tmpDtEvent.DTpctSim
                Case 3
                    EquipmentLabel4string = tmpDtEvent.Name
                    EquipMainbar4 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain4_PRloss = tmpDtEvent.DTpct
                    EquipMainbar4sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain4_PRlosssim = tmpDtEvent.DTpctSim
                Case 4
                    EquipmentLabel5string = tmpDtEvent.Name
                    EquipMainbar5 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain5_PRloss = tmpDtEvent.DTpct
                    EquipMainbar5sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain5_PRlosssim = tmpDtEvent.DTpctSim
                Case 5
                    EquipmentLabel6string = tmpDtEvent.Name
                    EquipMainbar6 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain6_PRloss = tmpDtEvent.DTpct
                    EquipMainbar6sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    EquipMain6_PRlosssim = tmpDtEvent.DTpctSim
            End Select
        Next
    End Sub

    Private Sub updateCard_Unplanned_Tier3_All()
        updateCard_Unplanned_Tier3A()
        updateCard_Unplanned_Tier3B()
        updateCard_Unplanned_Tier3C()
    End Sub
    Private Sub updateCard_Unplanned_Tier3A()
        Dim tmpDtEvent As DTevent
        If Card_Unplanned_T3A.Count < 1 Then Exit Sub
        For i As Integer = 0 To Card_Unplanned_T3A.Count - 1
            tmpDtEvent = Card_Unplanned_T3A(i)
            Select Case i
                Case 0
                    Equip1Label1string = tmpDtEvent.Name
                    Equip1bar1 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip1_1_PRloss = tmpDtEvent.DTpct
                    Equip1bar1sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip1_1_PRlosssim = tmpDtEvent.DTpctSim
                Case 1
                    Equip1Label2string = tmpDtEvent.Name
                    Equip1bar2 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip1_2_PRloss = tmpDtEvent.DTpct
                    Equip1bar2sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip1_2_PRlosssim = tmpDtEvent.DTpctSim
                Case 2
                    Equip1Label3string = tmpDtEvent.Name
                    Equip1bar3 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip1_3_PRloss = tmpDtEvent.DTpct
                    Equip1bar3sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip1_3_PRlosssim = tmpDtEvent.DTpctSim
            End Select
        Next
    End Sub
    Private Sub updateCard_Unplanned_Tier3B()
        Dim tmpDtEvent As DTevent
        If Card_Unplanned_T3B.Count < 1 Then Exit Sub
        For i As Integer = 0 To Card_Unplanned_T3B.Count - 1
            tmpDtEvent = Card_Unplanned_T3B(i)
            Select Case i
                Case 0
                    Equip2Label1string = tmpDtEvent.Name
                    Equip2bar1 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip2_1_PRloss = tmpDtEvent.DTpct
                    Equip2bar1sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip2_1_PRlosssim = tmpDtEvent.DTpctSim
                Case 1
                    Equip2Label2string = tmpDtEvent.Name
                    Equip2bar2 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip2_2_PRloss = tmpDtEvent.DTpct
                    Equip2bar2sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip2_2_PRlosssim = tmpDtEvent.DTpctSim
                Case 2
                    Equip2Label3string = tmpDtEvent.Name
                    Equip2bar3 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip2_3_PRloss = tmpDtEvent.DTpct
                    Equip2bar3sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip2_3_PRlosssim = tmpDtEvent.DTpctSim
            End Select
        Next
    End Sub
    Private Sub updateCard_Unplanned_Tier3C()
        Dim tmpDtEvent As DTevent
        If Card_Unplanned_T3C.Count < 1 Then Exit Sub
        For i As Integer = 0 To Card_Unplanned_T3C.Count - 1
            tmpDtEvent = Card_Unplanned_T3C(i)
            Select Case i
                Case 0
                    Equip3Label1string = tmpDtEvent.Name
                    Equip3bar1 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip3_1_PRloss = tmpDtEvent.DTpct
                    Equip3bar1sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip3_1_PRlosssim = tmpDtEvent.DTpctSim
                Case 1
                    Equip3Label2string = tmpDtEvent.Name
                    Equip3bar2 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip3_2_PRloss = tmpDtEvent.DTpct
                    Equip3bar2sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip3_2_PRlosssim = tmpDtEvent.DTpctSim
                Case 2
                    Equip3Label3string = tmpDtEvent.Name
                    Equip3bar3 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset)
                    Equip3_3_PRloss = tmpDtEvent.DTpct
                    Equip3bar3sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset)
                    Equip3_3_PRlosssim = tmpDtEvent.DTpctSim
            End Select
        Next
    End Sub

    Private Sub updateCard_Planned_Tier1()
        Dim tmpDtEvent As DTevent

        cardnameLabeltext(prStoryCard.Planned) = prStoryReport.getCardName(prStoryCard.Planned)
        If MaxPRindataset_planned = 0 Then MaxPRindataset_planned = 1
        For i As Integer = 0 To prStoryCardFields.Planned - 1
            tmpDtEvent = Card_Planned_T1(i)
            Select Case i
                Case 0
                    pdtlabel1string = tmpDtEvent.Name
                    PDTbar1 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar1_PRloss = tmpDtEvent.DTpct
                    PDTbar1sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar1_PRlosssim = tmpDtEvent.DTpctSim
                Case 1
                    pdtlabel2string = tmpDtEvent.Name
                    PDTbar2 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar2_PRloss = tmpDtEvent.DTpct
                    PDTbar2sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar2_PRlosssim = tmpDtEvent.DTpctSim
                Case 2
                    pdtlabel3string = tmpDtEvent.Name
                    PDTbar3 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar3_PRloss = tmpDtEvent.DTpct
                    PDTbar3sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar3_PRlosssim = tmpDtEvent.DTpctSim
                Case 3
                    pdtlabel4string = tmpDtEvent.Name
                    PDTbar4 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar4_PRloss = tmpDtEvent.DTpct
                    PDTbar4sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar4_PRlosssim = tmpDtEvent.DTpctSim
                Case 4
                    pdtlabel5string = tmpDtEvent.Name
                    PDTbar5 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar5_PRloss = tmpDtEvent.DTpct
                    PDTbar5sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar5_PRlosssim = tmpDtEvent.DTpctSim
                Case 5
                    pdtlabel6string = tmpDtEvent.Name
                    PDTbar6 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar6_PRloss = tmpDtEvent.DTpct
                    PDTbar6sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar6_PRlosssim = tmpDtEvent.DTpctSim
                Case 6
                    pdtlabel7string = tmpDtEvent.Name
                    PDTbar7 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar7_PRloss = tmpDtEvent.DTpct
                    PDTbar7sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar7_PRlosssim = tmpDtEvent.DTpctSim
                Case 7
                    pdtlabel8string = tmpDtEvent.Name
                    PDTbar8 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar8_PRloss = tmpDtEvent.DTpct
                    PDTbar8sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar8_PRlosssim = tmpDtEvent.DTpctSim
                Case 8
                    pdtlabel9string = tmpDtEvent.Name
                    PDTbar9 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar9_PRloss = tmpDtEvent.DTpct
                    PDTbar9sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    PDTbar9_PRlosssim = tmpDtEvent.DTpctSim
            End Select
        Next
    End Sub
    Private Sub updateCard_Planned_Tier2()
        Dim tmpDtEvent As DTevent
        For i As Integer = 0 To prStoryCardFields.Changeover - 1
            tmpDtEvent = Card_Planned_T2(i)
            Select Case i
                Case 0
                    changeoverlabel1string = tmpDtEvent.Name
                    Changeoverbar1 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover1_PRloss = tmpDtEvent.DTpct
                    Changeovertime1 = Math.Round(tmpDtEvent.MTTR, 0)
                    MaxChangeoverBubble_COtime = Changeovertime1
                    changeoverNo_of_Events1 = tmpDtEvent.Stops
                    Changeoverbar1sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover1_PRlosssim = tmpDtEvent.DTpctSim
                Case 1
                    changeoverlabel2string = tmpDtEvent.Name
                    Changeoverbar2 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover2_PRloss = tmpDtEvent.DTpct
                    Changeovertime2 = Math.Round(tmpDtEvent.MTTR, 0)
                    changeoverNo_of_Events2 = tmpDtEvent.Stops
                    Changeoverbar2sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover2_PRlosssim = tmpDtEvent.DTpctSim
                    If Changeovertime2 > MaxChangeoverBubble_COtime Then MaxChangeoverBubble_COtime = Changeovertime2
                Case 2
                    changeoverlabel3string = tmpDtEvent.Name
                    Changeoverbar3 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover3_PRloss = tmpDtEvent.DTpct
                    Changeovertime3 = Math.Round(tmpDtEvent.MTTR, 0)
                    changeoverNo_of_Events3 = tmpDtEvent.Stops
                    Changeoverbar3sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover3_PRlosssim = tmpDtEvent.DTpctSim
                    If Changeovertime3 > MaxChangeoverBubble_COtime Then MaxChangeoverBubble_COtime = Changeovertime3
                Case 3
                    changeoverlabel4string = tmpDtEvent.Name
                    Changeoverbar4 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover4_PRloss = tmpDtEvent.DTpct
                    Changeovertime4 = Math.Round(tmpDtEvent.MTTR, 0)
                    changeoverNo_of_Events4 = tmpDtEvent.Stops
                    Changeoverbar4sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover4_PRlosssim = tmpDtEvent.DTpctSim
                    If Changeovertime4 > MaxChangeoverBubble_COtime Then MaxChangeoverBubble_COtime = Changeovertime4
                Case 4
                    changeoverlabel5string = tmpDtEvent.Name
                    Changeoverbar5 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover5_PRloss = tmpDtEvent.DTpct
                    Changeovertime5 = Math.Round(tmpDtEvent.MTTR, 0)
                    changeoverNo_of_Events5 = tmpDtEvent.Stops
                    Changeoverbar5sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover5_PRlosssim = tmpDtEvent.DTpctSim
                    If Changeovertime5 > MaxChangeoverBubble_COtime Then MaxChangeoverBubble_COtime = Changeovertime5
                Case 5
                    changeoverlabel6string = tmpDtEvent.Name
                    Changeoverbar6 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover6_PRloss = Math.Round(tmpDtEvent.DTpct, 0)
                    Changeovertime6 = Math.Round(tmpDtEvent.MTTR, 0)
                    changeoverNo_of_Events6 = tmpDtEvent.Stops
                    Changeoverbar6sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover6_PRlosssim = tmpDtEvent.DTpctSim
                    If Changeovertime6 > MaxChangeoverBubble_COtime Then MaxChangeoverBubble_COtime = Changeovertime6
                Case 6
                    changeoverlabel7string = tmpDtEvent.Name
                    Changeoverbar7 = tmpDtEvent.DTpct * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover7_PRloss = Math.Round(tmpDtEvent.DTpct, 0)
                    Changeovertime7 = Math.Round(tmpDtEvent.MTTR, 0)
                    changeoverNo_of_Events7 = tmpDtEvent.Stops
                    Changeoverbar7sim = tmpDtEvent.DTpctSim * (barchartBarMaxSize / MaxPRindataset_planned)
                    Changeover7_PRlosssim = tmpDtEvent.DTpctSim
                    If Changeovertime7 > MaxChangeoverBubble_COtime Then MaxChangeoverBubble_COtime = Changeovertime7
            End Select
        Next
    End Sub

    Public Sub updateCard_TopStops()
        Dim tmpDtEvent As DTevent
        If MaxStopsindataset = 0 Then MaxStopsindataset = 1
        For i As Integer = 0 To prStoryCardFields.Stops - 1
            tmpDtEvent = Card_TopStops(i)
            topstopname(i) = tmpDtEvent.Name
            Select Case i
                Case 0
                    topstopsbar1 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar1_SPD = tmpDtEvent.SPD
                    topstopsbar1_PRloss = Math.Round(tmpDtEvent.DTpct, 3) ' LG Code decimal points 
                    topstopsbar1_Stops = tmpDtEvent.Stops
                    topstopsbar1_DTmin = Math.Round(tmpDtEvent.DT, 1)

                Case 1
                    topstopsbar2 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar2_SPD = tmpDtEvent.SPD
                    topstopsbar2_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar2_Stops = tmpDtEvent.Stops
                    topstopsbar2_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 2
                    topstopsbar3 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar3_SPD = tmpDtEvent.SPD
                    topstopsbar3_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar3_Stops = tmpDtEvent.Stops
                    topstopsbar3_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 3
                    topstopsbar4 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar4_SPD = tmpDtEvent.SPD
                    topstopsbar4_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar4_Stops = tmpDtEvent.Stops
                    topstopsbar4_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 4
                    topstopsbar5 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar5_SPD = tmpDtEvent.SPD
                    topstopsbar5_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar5_Stops = tmpDtEvent.Stops
                    topstopsbar5_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 5
                    topstopsbar6 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar6_SPD = tmpDtEvent.SPD
                    topstopsbar6_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar6_Stops = tmpDtEvent.Stops
                    topstopsbar6_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 6
                    topstopsbar7 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar7_SPD = tmpDtEvent.SPD
                    topstopsbar7_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar7_Stops = tmpDtEvent.Stops
                    topstopsbar7_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 7
                    topstopsbar8 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar8_SPD = tmpDtEvent.SPD
                    topstopsbar8_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar8_Stops = tmpDtEvent.Stops
                    topstopsbar8_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 8
                    topstopsbar9 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar9_SPD = tmpDtEvent.SPD
                    topstopsbar9_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar9_Stops = tmpDtEvent.Stops
                    topstopsbar9_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 9
                    topstopsbar10 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar10_SPD = tmpDtEvent.SPD
                    topstopsbar10_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar10_Stops = tmpDtEvent.Stops
                    topstopsbar10_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 10
                    topstopsbar11 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar11_SPD = tmpDtEvent.SPD
                    topstopsbar11_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar11_Stops = tmpDtEvent.Stops
                    topstopsbar11_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 11
                    topstopsbar12 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar12_SPD = tmpDtEvent.SPD
                    topstopsbar12_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar12_Stops = tmpDtEvent.Stops
                    topstopsbar12_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 12
                    topstopsbar13 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar13_SPD = tmpDtEvent.SPD
                    topstopsbar13_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar13_Stops = tmpDtEvent.Stops
                    topstopsbar13_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 13
                    topstopsbar14 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar14_SPD = tmpDtEvent.SPD
                    topstopsbar14_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar14_Stops = tmpDtEvent.Stops
                    topstopsbar14_DTmin = Math.Round(tmpDtEvent.DT, 1)
                Case 14
                    topstopsbar15 = tmpDtEvent.SPD * (barchartBarMaxSize_stops / MaxStopsindataset)
                    topstopsbar15_SPD = tmpDtEvent.SPD
                    topstopsbar15_PRloss = Math.Round(tmpDtEvent.DTpct, 3)
                    topstopsbar15_Stops = tmpDtEvent.Stops
                    topstopsbar15_DTmin = Math.Round(tmpDtEvent.DT, 1)
            End Select

            topstop_mttr(i) = tmpDtEvent.MTTR
            topstop_mtbf(i) = tmpDtEvent.MTBF
        Next
    End Sub
#End Region
    Public Function GetMtDPR() As Double
        Dim prMotionReport As Motion_LinePRReport = New Motion_LinePRReport(AllProdLines(selectedindexofLine_temp), 1)

        If prMotionReport.findMtDStartDay = -1 Or prMotionReport.findAnyDayinDailyReports(endtimeselected) = -1 Then
            If endtimeselected.Day > 1 Then
                Dim x = endtimeselected.AddDays(-1)
                If prMotionReport.findMtDStartDay = -1 Or prMotionReport.findAnyDayinDailyReports(x) = -1 Then
                    Return -3.14159
                Else
                    Return Math.Round(prMotionReport.getHTMLdataString_AMCharts_PR_Weekly(prMotionReport.findMtDStartDay(), prMotionReport.findAnyDayinDailyReports(x)), 1)
                End If
            Else
                Return -3.14159
            End If
        Else
            Return Math.Round(prMotionReport.getHTMLdataString_AMCharts_PR_Weekly(prMotionReport.findMtDStartDay(), prMotionReport.findAnyDayinDailyReports(endtimeselected)), 1)
        End If

    End Function
    Sub bargraphreportload()
        Initialize_UseTrack_Variables()
        CheckScreenResolution()
        PrepareforAvailabilityMode()
        linename_label.Content = AllProdLines(selectedindexofLine_temp).ToString
        MainDateLabel.Content = datalabelcontent
        DTviewbox.Visibility = Visibility.Visible
        unplannedDTequipmentchart.Visibility = Visibility.Hidden
        unplannedDTequip1chart.Visibility = Visibility.Hidden
        unplannedDTequip2chart.Visibility = Visibility.Hidden
        unplannedDTequip3chart.Visibility = Visibility.Hidden
        misc_updt_chart1.Visibility = Visibility.Hidden
        misc_updt_chart2.Visibility = Visibility.Hidden
        topstopschart.Visibility = Visibility.Hidden
        dtgreenbox.Visibility = Visibility.Hidden
        stopsgreenbox.Visibility = Visibility.Hidden
        incontrolgreenbox.Visibility = Visibility.Hidden

        Dim showRateLossWithPR As Boolean = (AllProdLines(prStoryReport.ParentLineInt).SiteName = "GBO" And AllProdLines(prStoryReport.ParentLineInt).Sector <> SECTOR_BEAUTY)

        ' check for availability mode
        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            pr_label.Content = FormatPercent(prStoryReport.PR, 1) ' & " PR"
            PR_Label_Header.Content = "PR"
            If showRateLossWithPR Then
                pr_label2.Content = FormatPercent(prStoryReport.PR, 1) ' & " PR  " & FormatPercent(prStoryReport.rateLoss, 1) & "rate loss"
                pr_label2.Visibility = Visibility.Visible
                pr_label.Visibility = Visibility.Hidden
            Else
                pr_label2.Visibility = Visibility.Hidden
                pr_label.Visibility = Visibility.Visible
            End If
        Else
            pr_label.Content = FormatPercent(prStoryReport.PR, 1) ' & " Av."
            PR_Label_Header.Content = "Availability"
            cases_label.Visibility = Visibility.Hidden
            Cases_Label_Header.Visibility = Visibility.Hidden
            Rateloss_Label_Header.Visibility = Visibility.Hidden
            rateloss_label.Visibility = Visibility.Hidden
            If showRateLossWithPR Then
                pr_label2.Content = FormatPercent(prStoryReport.PR, 1) ' & "AV " & FormatPercent(prStoryReport.rateLoss, 1) & "rate loss"
                pr_label2.Visibility = Visibility.Visible
                pr_label.Visibility = Visibility.Hidden
            Else
                pr_label2.Visibility = Visibility.Hidden
                pr_label.Visibility = Visibility.Visible
            End If
        End If


        If prStoryReport.CasesAdjusted = 0 Then
            cases_label.Content = prStoryReport.CasesActual ' & " cases"
        Else
            cases_label.Content = prStoryReport.CasesAdjusted ' & " cases"
        End If


        rateloss_label.Content = FormatPercent(prStoryReport.rateLoss, 1) ' & " rate and quality loss"
        Stops_Label.Content = Math.Round(prStoryReport.StopsPerDay, 0) ' & " stops/day"
        Stops_Label.ToolTip = "Stops/Day = " & Math.Round(prStoryReport.ActualStops, 0) & " stops * 1440 (minutes/day) / " & Math.Round(prStoryReport.schedTime) & " min sched time. Stops normalized over 24hr period."
        MTBF_Label.Content = Math.Round(prStoryReport.MTBF, 0) & " min"
        If Not AllProdLines(prStoryReport.ParentLineInt).SQLdowntimeProcedure = DefaultProficyProductionProcedure.Maple And Not AllProdLines(prStoryReport.ParentLineInt).SQLdowntimeProcedure = DefaultProficyProductionProcedure.Maple_New Then
            MTBF_Label.ToolTip = "Total system unplanned MTBF.  Uptime = " & Math.Round(prStoryReport.MainLEDSReport.UT_DT, 2) & " Stops = " & prStoryReport.MainLEDSReport.Stops
        End If

        incontrolAnalysisSelectedEnddate = endtimeselected ' LG Code
        incontrolAnalysisSelectedStartdate = endtimeselected.AddDays(-1) ' LG Code

        'MaxPRlossindataset_stops = MasterDataSet.maxDTpct

        MaxPRlossindataset_stops = prStoryReport.EventMaxDTpct '0.15  ' sam to put max DT percent 
        HideMenu()
        hidestoplabels()
        hideRL4()
        showUPDTlabels()
        showPDTlabels()
        hideEquipmentlabels()
        hidechangeoverlabels()
        hidePlannedchangeoverTIMEchart()
        hideincontrolareaALLelements()
        HideIncontrolDatePicker()
        CloseNotesSplash()
        checkmappingavailability()
        CloseFloatingSimulator()
        HideSimRectangles_UPDT()
        HideSimRectangles_Changeover()
        HideSimRectangles_Equip()
        HideSimRectangles_PDT()
        hideErrorSplashPage()

        If AllProdLines(selectedindexofLine_temp).isFilterByBrandcode = False Then


            hidefiltericon()
            FilterIcon_inactive.Visibility = Visibility.Visible
            FilterOnOfflabel.Content = "Filter: OFF"

        Else
            FilterIcon_inactive.Visibility = Visibility.Hidden
            FilterIcon_active.Visibility = Visibility.Visible
            FilterOnOfflabel.Content = "Filter: ON"
        End If
        StopsQuestion.Visibility = Visibility.Hidden
        If AllProdLines(selectedindexofLine_temp)._isDualConstraint = True Then
            StopsQuestion.Visibility = Visibility.Visible
        End If

        DecidetoShowNavigationButtons()

        dtgreenbox.Visibility = Visibility.Visible
        stopsgreenbox.Visibility = Visibility.Hidden
        incontrolgreenbox.Visibility = Visibility.Hidden



        'If System.IO.File.Exists(SERVER_FOLDER_PATH & "pie.js") Then PR_Pie.Visibility = Windows.Visibility.Visible


        'MsgBox(DateDiff(DateInterval.Minute, starttimeselected, endtimeselected))

        DecidetoEnableNotes()

        If Not My.Settings.AdvancedSettings_UseSimulation Then
            StatusBar.Visibility = Visibility.Hidden
            CurrentStatusLabel.Visibility = Visibility.Hidden
        End If

        If IsExcludedEventsIncluded = True Then
            ExcludedEventsIcon.Visibility = Visibility.Visible
        Else
            ExcludedEventsIcon.Visibility = Visibility.Hidden
        End If

    End Sub
    Private Sub PrepareforAvailabilityMode()
        If My.Settings.AdvancedSettings_isAvailabilityMode = True Then
            pr_label.ToolTip = "Availibility of the line in the selected time period. Calculated using uptime and downtime data."
            TrendStarIcon0.ToolTip = "Daily, monthly and weekly avalibility trend of the line in last 90 days."
            TopStopsLegend_bubbletext.Content = "Av. Loss %"
            incontrolPRlabel.Content = "Av. Loss %"
        Else
            pr_label.ToolTip = "PR of the line in the selected time period. Calculated using production data, scheduled time and target rate."
            TrendStarIcon0.ToolTip = "Daily, monthly and weekly PR trend of the line in last 90 days."
            TopStopsLegend_bubbletext.Content = "PR Loss %"
            incontrolPRlabel.Content = "PR Loss %"
        End If
    End Sub
    Private Sub DecidetoEnableNotes()

        'If InStr(linename_label.Content, "Skin", vbTextCompare) Then
        If My.Settings.AdvancedSettings_UseNotes = True Or InStr(linename_label.Content, "Skin", vbTextCompare) > 0 Then
            My.Settings.AdvancedSettings_UseNotes = True
            NOtesLabel.Visibility = Visibility.Visible
            NotesPlusIcon.Visibility = Visibility.Visible
        Else

        End If


    End Sub

    Private Sub DATADR_launch()

    End Sub

    Private Sub DecidetoShowNavigationButtons()
        ScrollBase_Card1 = 0
        ScrollBase_Card3 = 0
        ScrollBase_Card4 = 0
        ScrollBase_Card5 = 0
        ScrollBase_Card6 = 0
        If prStoryReport.getCardEventNumber(1) <= 6 Then NavigationRight_Card1.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(3) <= 6 Then NavigationRight_Card3.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(4) <= 3 Then NavigationRight_Card4.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(5) <= 3 Then NavigationRight_Card5.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(6) <= 3 Then NavigationRight_Card6.Visibility = Visibility.Hidden
    End Sub
    Private Sub checkmappingavailability()

        If AllProdLines(selectedindexofLine_temp).prStoryMapping = prStoryMapping.NoMappingAvailable Then

            DTpercentframe.IsEnabled = False
            MsgBox("No mapping available!")


            '''''''''''copied from frameclick for stops frame click

            dtgreenbox.Visibility = Visibility.Hidden
            stopsgreenbox.Visibility = Visibility.Visible
            incontrolgreenbox.Visibility = Visibility.Hidden

            unplannedDTchart.Visibility = Visibility.Hidden
            plannedDTchart.Visibility = Visibility.Hidden
            unplannedDTequipmentchart.Visibility = Visibility.Hidden
            unplannedDTequip1chart.Visibility = Visibility.Hidden
            unplannedDTequip2chart.Visibility = Visibility.Hidden
            unplannedDTequip3chart.Visibility = Visibility.Hidden

            topstopschart.Visibility = Visibility.Visible
            ' RL4topstopschart.Visibility = Windows.Visibility.Visible
            misc_updt_chart1.Visibility = Visibility.Hidden
            misc_updt_chart2.Visibility = Visibility.Hidden
            'd3chart.Visibility = Windows.Visibility.Visible
            showstoplabels()
            Assignlabelnames(prStoryReport)
            hideUPDTlabels()
            hidePDTlabels()
            hideEquipmentlabels()
            hidechangeoverlabels()
            hidePlannedchangeoverTIMEchart()
            hideincontrolareaALLelements()
            HideIncontrolDatePicker()

        End If

    End Sub

#Region "Mouse Move / Click Handlers"
    Private Sub BubbleMouseMove(sender As Object, e As MouseEventArgs)
        Dim bubblenumber As Integer

        sender.opacity = 0.8
        If sender Is incontrol_Regenerate Or sender Is incontrol_date Then Exit Sub
        bubblenumber = onlyDigits(sender.name)

        ConsiderPickMode_for_Cursor(sender)


    End Sub
    Private Sub BubbleMouseLeave(sender As Object, e As MouseEventArgs)
        sender.opacity = 1.0
        ConsiderPickMode_for_Cursor(sender)
        If IsPickMode = True Then Cursor = Cursors.Pen
    End Sub

    Private Sub BarDoubleClick(sender As Object, e As MouseButtonEventArgs)
        If IsPickMode = True Then
            Exit Sub
        End If

        Dim failureName As String = topstopname(onlyDigits(sender.name) - 1)
        Dim rawRows As New List(Of Integer)

        With prStoryReport.MainLEDSReport.DT_Report.rawDTdata

            For i = 0 To .UnplannedData.Count - 1
                If My.Settings.defaultDownTimeField_Secondary = -1 Then
                    If .UnplannedData(i).getFieldFromInteger(My.Settings.defaultDownTimeField) = failureName Then
                        rawRows.Add(i)
                    End If
                Else
                    If .UnplannedData(i).getFieldFromInteger(My.Settings.defaultDownTimeField) & "-" & .UnplannedData(i).getFieldFromInteger(My.Settings.defaultDownTimeField_Secondary) = failureName Then
                        rawRows.Add(i)
                    End If
                End If
            Next

        End With

        Dim rawdatawindow2 As New RawDataWindow(-1, rawRows, prStoryReport, failureName) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)
        rawdatawindow2.updateValues(-1, rawRows, prStoryReport, failureName) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)

        '  System.Windows.Forms.Application.DoEvents()
        rawdatawindow2.Owner = Me
        rawdatawindow2.setBargraphReportWindow_forraw(Me)

        rawdatawindow2.Show()

    End Sub


    Private Sub BubbleDoubleClick(sender As Object, e As MouseButtonEventArgs)
        If IsPickMode = True Then
            Exit Sub
        End If

        Dim bubblenumber As Integer = onlyDigits(sender.Name)
        Dim failureName As String = stopbubblenames(bubblenumber - 1)
        Dim rawRows As New List(Of Integer)

        With prStoryReport.MainLEDSReport.DT_Report.rawDTdata

            For i = 0 To .UnplannedData.Count - 1
                If My.Settings.defaultDownTimeField_Secondary = -1 Then
                    If .UnplannedData(i).getFieldFromInteger(My.Settings.defaultDownTimeField) = failureName Then
                        rawRows.Add(i)
                    End If
                Else
                    If .UnplannedData(i).getFieldFromInteger(My.Settings.defaultDownTimeField) & "-" & .UnplannedData(i).getFieldFromInteger(My.Settings.defaultDownTimeField_Secondary) = failureName Then
                        rawRows.Add(i)
                    End If
                End If
            Next

        End With

        Dim rawdatawindow2 As New RawDataWindow(-1, rawRows, prStoryReport, failureName) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)
        rawdatawindow2.updateValues(-1, rawRows, prStoryReport, failureName) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)

        '  System.Windows.Forms.Application.DoEvents()
        rawdatawindow2.Owner = Me
        rawdatawindow2.setBargraphReportWindow_forraw(Me)

        rawdatawindow2.Show()

    End Sub

    Private Sub Bubbleclick(sender As Object, e As MouseButtonEventArgs)

        If IsPickMode = True Then
            PickaLoss_CollectInfofromLabel(sender, 51)
            Exit Sub
        End If



        Dim bubblenumber As Integer

        bubblenumber = onlyDigits(sender.Name)
        ' Incontrol_FailureMode.Content = "Failure mode - " & bubblenumber 'get failuremode name from sam's code
        If Not IsNothing(tempbubblesender) Then
            tempbubblesender.StrokeThickness = 0
        End If
        sender.Strokethickness = 1
        tempbubblesender = sender

        Incontrol_FailureMode.Content = stopbubblenames(bubblenumber - 1)
        bubblenumberpublic = bubblenumber - 1

        incontrolstopcount_circle.Visibility = Visibility.Visible
        incontrolstopcountLabel.Visibility = Visibility.Visible
        incontrolstopslabel.Visibility = Visibility.Visible
        incontrolPR_circle.Visibility = Visibility.Visible
        incontrolPRperLabel.Visibility = Visibility.Visible
        incontrolPRlabel.Visibility = Visibility.Visible
        'incontrolMTBF_circle.Visibility = Windows.Visibility.Visible
        'incontrolMTBFminLabel.Visibility = Windows.Visibility.Visible
        'incontrolMTBFlabel.Visibility = Windows.Visibility.Visible
        incontrolstopcountLabel.Content = stopbubblestops(bubblenumber - 1)
        incontrolPRperLabel.Content = FormatPercent(stopbubblePR(bubblenumber - 1), 1)
        incontrolMTBFminLabel.Content = Math.Round(stopbubbleMTBF(bubblenumber - 1))

    End Sub

    Private Sub FrameMouseMove(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Hand

        sender.Opacity = 0.8


        If sender Is DTpercentframe Then
            DTpercentframe.Opacity = 0.8
        End If

        If sender Is stopsframe Then
            stopsframe.Opacity = 0.8
        End If

        If sender Is incontrolframe Then
            incontrolframe.Opacity = 0.8
        End If

        If sender Is prstory_Main_icon Then
            prstory_Main_icon.Opacity = 0.8
        End If

        If sender Is RL4bar1 Then
            RL4bar1.Opacity = 0.8
        End If


        If sender Is RL4bar2 Then
            RL4bar2.Opacity = 0.8
        End If


        If sender Is RL4bar3 Then
            RL4bar3.Opacity = 0.8
        End If
        ConsiderPickMode_for_Cursor(sender)

    End Sub

    Private Sub executeRateLossCheck()
        ' If Me.Owner.Paren Then

        'End If
    End Sub

    Private Sub FrameMouseLeave(sender As Object, e As MouseEventArgs)

        executeRateLossCheck()

        Cursor = Cursors.Arrow
        sender.Opacity = 1.0
        If sender Is DTpercentframe Then
            DTpercentframe.Opacity = 1.0
        End If

        If sender Is stopsframe Then
            stopsframe.Opacity = 1.0
        End If

        If sender Is incontrolframe Then
            incontrolframe.Opacity = 1.0
        End If

        If sender Is prstory_Main_icon Then
            prstory_Main_icon.Opacity = 1.0
        End If

        If sender Is RL4bar1 Then
            RL4bar1.Opacity = 1.0
        End If


        If sender Is RL4bar2 Then
            RL4bar2.Opacity = 1.0
        End If


        If sender Is RL4bar3 Then
            RL4bar3.Opacity = 1.0
        End If


        ConsiderPickMode_for_Cursor(sender)
        If IsPickMode = True Then Cursor = Cursors.Pen
    End Sub

    Private Sub Frameclick(sender As Object, e As MouseButtonEventArgs)
        NotesBaseCanvas.Visibility = Visibility.Hidden
        NotesMenuMainCanvas.Visibility = Visibility.Hidden
        NOtesLabel.Visibility = Visibility.Visible

        If sender Is stopsframe Then

            UseTrack_TopStopsMain = True

            dtgreenbox.Visibility = Visibility.Hidden
            stopsgreenbox.Visibility = Visibility.Visible
            incontrolgreenbox.Visibility = Visibility.Hidden

            unplannedDTchart.Visibility = Visibility.Hidden
            plannedDTchart.Visibility = Visibility.Hidden
            unplannedDTequipmentchart.Visibility = Visibility.Hidden
            unplannedDTequip1chart.Visibility = Visibility.Hidden
            unplannedDTequip2chart.Visibility = Visibility.Hidden
            unplannedDTequip3chart.Visibility = Visibility.Hidden

            topstopschart.Visibility = Visibility.Visible
            ' RL4topstopschart.Visibility = Windows.Visibility.Visible
            misc_updt_chart1.Visibility = Visibility.Hidden
            misc_updt_chart2.Visibility = Visibility.Hidden
            'd3chart.Visibility = Windows.Visibility.Visible
            showstoplabels()
            Assignlabelnames(prStoryReport)
            hideUPDTlabels()
            hidePDTlabels()
            hideEquipmentlabels()
            hidechangeoverlabels()
            hidePlannedchangeoverTIMEchart()
            hideincontrolareaALLelements()
            HideIncontrolDatePicker()
            If DateDiff(DateInterval.Minute, starttimeselected, endtimeselected) < 1440 Then
                TopStopsLegend_bartext.Content = "Stops/Day"
                ToggleStops()
            End If

        End If

        If sender Is DTpercentframe Then
            showDTpercentframe()
            UseTrack_UPDTview = True
        End If


        If sender Is incontrolframe Then
            If DeactivateIncontrol = True Then
                MsgBox("Incontrol tool is not available, as there is not enough Scheduled Events for processing Chronic/Sporadic separation algorithms.")
                Exit Sub
            End If

            UseTrack_IncontrolMain = True

            If IsNothing(MasterDataSet) Then
                MsgBox("Give us a second, inControl is coming up.")
                Exit Sub
            End If

            incontrolExtraCanvas.Visibility = Visibility.Visible

            hidestoplabels()
            hideRL4()
            dtgreenbox.Visibility = Visibility.Hidden
            stopsgreenbox.Visibility = Visibility.Hidden
            incontrolgreenbox.Visibility = Visibility.Visible

            unplannedDTchart.Visibility = Visibility.Hidden
            plannedDTchart.Visibility = Visibility.Hidden
            unplannedDTequipmentchart.Visibility = Visibility.Hidden
            unplannedDTequip1chart.Visibility = Visibility.Hidden
            unplannedDTequip2chart.Visibility = Visibility.Hidden
            unplannedDTequip3chart.Visibility = Visibility.Hidden

            topstopschart.Visibility = Visibility.Hidden
            misc_updt_chart1.Visibility = Visibility.Hidden
            misc_updt_chart2.Visibility = Visibility.Hidden
            hideRL4()
            hideRL4rawstopslist()
            hidestoplabels()
            hideUPDTlabels()
            hidePDTlabels()
            hideEquipmentlabels()
            hidechangeoverlabels()
            hidePlannedchangeoverTIMEchart()
            HideIncontrolDatePicker()

            showincontrolareaALLelements()
            GenerateIncontrolCharts(incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate, True)
        Else
            incontrolExtraCanvas.Visibility = Visibility.Hidden

        End If

        If sender Is prstory_Main_icon Then


            dtgreenbox.Visibility = Visibility.Hidden
            stopsgreenbox.Visibility = Visibility.Hidden
            incontrolgreenbox.Visibility = Visibility.Hidden

            unplannedDTequipmentchart.Visibility = Visibility.Hidden
            unplannedDTequip1chart.Visibility = Visibility.Hidden
            unplannedDTequip2chart.Visibility = Visibility.Hidden
            unplannedDTequip3chart.Visibility = Visibility.Hidden
            topstopschart.Visibility = Visibility.Hidden
            misc_updt_chart1.Visibility = Visibility.Hidden
            misc_updt_chart2.Visibility = Visibility.Hidden



            unplannedDTchart.Visibility = Visibility.Visible
            plannedDTchart.Visibility = Visibility.Visible
            hideRL4()
            hidestoplabels()
            hideRL4rawstopslist()
            showPDTlabels()
            showUPDTlabels()
            hideEquipmentlabels()
            hidechangeoverlabels()
            hidePlannedchangeoverTIMEchart()
            hideincontrolareaALLelements()
            HideIncontrolDatePicker()
        End If

    End Sub
#End Region

    Private Sub showDTpercentframe()
        hidestoplabels()
        hideRL4()
        dtgreenbox.Visibility = Visibility.Visible
        stopsgreenbox.Visibility = Visibility.Hidden
        incontrolgreenbox.Visibility = Visibility.Hidden


        plannedDTchart.Visibility = Visibility.Hidden
        unplannedDTchart.Visibility = Visibility.Visible
        unplannedDTequipmentchart.Visibility = Visibility.Hidden
        unplannedDTequip1chart.Visibility = Visibility.Hidden
        unplannedDTequip2chart.Visibility = Visibility.Hidden
        unplannedDTequip3chart.Visibility = Visibility.Hidden
        topstopschart.Visibility = Visibility.Hidden
        'RL4topstopschart.Visibility = Windows.Visibility.Hidden
        plannedDTchart.Visibility = Visibility.Hidden
        misc_updt_chart1.Visibility = Visibility.Hidden
        misc_updt_chart2.Visibility = Visibility.Hidden
        'd3chart.Visibility = Windows.Visibility.Hidden
        showUPDTlabels()
        showPDTlabels()
        hideEquipmentlabels()
        hidechangeoverlabels()
        hidePlannedchangeoverTIMEchart()
        hideincontrolareaALLelements()
        HideIncontrolDatePicker()
        PDTlabel2.Background = LabelDefaultColor
        SetPDTLabelDefaultColor()

        If prStoryReport.getCardEventNumber(1) <= 6 Then NavigationRight_Card1.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(3) <= 6 Then NavigationRight_Card3.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(4) <= 3 Then NavigationRight_Card4.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(5) <= 3 Then NavigationRight_Card5.Visibility = Visibility.Hidden
        If prStoryReport.getCardEventNumber(6) <= 3 Then NavigationRight_Card6.Visibility = Visibility.Hidden

    End Sub

#Region "Show/Hide inControl"
    Sub showincontrolareaALLelements()
        IncontrolActiveArea.Visibility = Visibility.Visible
        IncontrolMAINarea.Visibility = Visibility.Visible
        incontrol_splitter.Visibility = Visibility.Visible
        Incontrol_FailureMode.Visibility = Visibility.Visible
        incontrollogo.Visibility = Visibility.Visible
        chronicarrow.Visibility = Visibility.Visible
        sporadicarrow.Visibility = Visibility.Visible
        incontroldummygapfiller.Visibility = Visibility.Visible
        IncontrolActiveArea_dummy.Visibility = Visibility.Visible

        '  legend_stable.Visibility = Windows.Visibility.Visible
        '  legend_unstable.Visibility = Windows.Visibility.Visible
        '   legend_green.Visibility = Windows.Visibility.Visible
        '   legend_Orange.Visibility = Windows.Visibility.Visible
        '   legend_red.Visibility = Windows.Visibility.Visible
        '   legend_yellow.Visibility = Windows.Visibility.Visible
        incontrol_expandedview_Launcher.Visibility = Visibility.Visible
        incontrol_date.Visibility = Visibility.Visible
        incontrol_Day_selection.Visibility = Visibility.Visible
        incontrol_Shift_selection.Visibility = Visibility.Visible
        stopbubble1.Visibility = Visibility.Visible
        stopbubble2.Visibility = Visibility.Visible
        stopbubble3.Visibility = Visibility.Visible
        stopbubble4.Visibility = Visibility.Visible
        stopbubble5.Visibility = Visibility.Visible
        stopbubble6.Visibility = Visibility.Visible
        stopbubble7.Visibility = Visibility.Visible
        stopbubble8.Visibility = Visibility.Visible
        stopbubble9.Visibility = Visibility.Visible
        stopbubble10.Visibility = Visibility.Visible
        stopbubble11.Visibility = Visibility.Visible
        stopbubble12.Visibility = Visibility.Visible
        stopbubble13.Visibility = Visibility.Visible
        stopbubble14.Visibility = Visibility.Visible
        stopbubble15.Visibility = Visibility.Visible


        incontrol_Shift_selection.Content = "Last " & Int(24 / AllProdLines(selectedindexofLine_temp).NumberOfShifts) & " hours"
    End Sub

    Sub hideincontrolareaALLelements()

        IncontrolActiveArea.Visibility = Visibility.Hidden
        IncontrolMAINarea.Visibility = Visibility.Hidden
        incontrol_splitter.Visibility = Visibility.Hidden
        Incontrol_FailureMode.Visibility = Visibility.Hidden
        incontrollogo.Visibility = Visibility.Hidden
        chronicarrow.Visibility = Visibility.Hidden
        sporadicarrow.Visibility = Visibility.Hidden
        incontroldummygapfiller.Visibility = Visibility.Hidden
        IncontrolActiveArea_dummy.Visibility = Visibility.Hidden
        incontrol_date.Visibility = Visibility.Hidden
        incontrol_Day_selection.Visibility = Visibility.Hidden
        incontrol_Shift_selection.Visibility = Visibility.Hidden
        legend_stable.Visibility = Visibility.Hidden
        legend_unstable.Visibility = Visibility.Hidden
        legend_green.Visibility = Visibility.Hidden
        legend_Orange.Visibility = Visibility.Hidden
        legend_red.Visibility = Visibility.Hidden
        legend_yellow.Visibility = Visibility.Hidden
        incontrol_expandedview_Launcher.Visibility = Visibility.Hidden
        incontrolstopcount_circle.Visibility = Visibility.Hidden
        incontrolstopcountLabel.Visibility = Visibility.Hidden
        incontrolstopslabel.Visibility = Visibility.Hidden
        incontrolPR_circle.Visibility = Visibility.Hidden
        incontrolPRperLabel.Visibility = Visibility.Hidden
        incontrolPRlabel.Visibility = Visibility.Hidden
        incontrolMTBF_circle.Visibility = Visibility.Hidden
        incontrolMTBFminLabel.Visibility = Visibility.Hidden
        incontrolMTBFlabel.Visibility = Visibility.Hidden

        hideStopbubbles()
    End Sub
    Sub hideStopbubbles()

        stopbubble1.Visibility = Visibility.Hidden
        stopbubble2.Visibility = Visibility.Hidden
        stopbubble3.Visibility = Visibility.Hidden
        stopbubble4.Visibility = Visibility.Hidden
        stopbubble5.Visibility = Visibility.Hidden
        stopbubble6.Visibility = Visibility.Hidden
        stopbubble7.Visibility = Visibility.Hidden
        stopbubble8.Visibility = Visibility.Hidden
        stopbubble9.Visibility = Visibility.Hidden
        stopbubble10.Visibility = Visibility.Hidden
        stopbubble11.Visibility = Visibility.Hidden
        stopbubble12.Visibility = Visibility.Hidden
        stopbubble13.Visibility = Visibility.Hidden
        stopbubble14.Visibility = Visibility.Hidden
        stopbubble15.Visibility = Visibility.Hidden
    End Sub
#End Region

#Region "RL4 Chart"
    Private Sub launchRL4(sender As Object, e As MouseButtonEventArgs)

        If IsPickMode = True And InStr(sender.name, "stoplabel") > 0 Then
            PickaLoss_CollectInfofromLabel(sender, 31)
            Exit Sub
        End If






        If InStr(sender.Name, "stoplabel") > 0 Then
            selectedDTGroup = sender.text
        ElseIf InStr(sender.Name, "TopStops_") > 0 Then
            selectedDTGroup = topstopname(onlyDigits(sender.name) - 1)
        End If

        SelectedFailuremonumber_inTopStopsforTrends = onlyDigits(sender.name) - 1

        selectedfailuremode = selectedDTGroup
        RL4circle.Opacity = 0.3
        RL3circle.Opacity = 0.3
        RL2circle.Opacity = 0.3
        RL1circle.Opacity = 0.3
        faultcodecircle.Opacity = 0.3
        RL4circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        RL3circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        RL2circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        RL1circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        faultcodecircle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))


        ' GenerateRL4chart(selectedDTGroup, DownTimeColumn.Reason4)
        Generate_after_Deciding_WhichRLCharttoGenerate()
        RL4circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(101, 222, 200))
        RL4circle.Opacity = 1.0

        If InStr(linename_label.Content, "Fam", vbTextCompare) > 0 Then

            RL2circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(101, 222, 200))
            RL2circle.Opacity = 1.0
            circleselect(RL2circle, e)

            RL3circle.Visibility = Visibility.Hidden
            RL3label.Visibility = Visibility.Hidden
            RL4circle.Visibility = Visibility.Hidden
            RL4label.Visibility = Visibility.Hidden

        End If


        MTBFlabel.Content = "MTBF: " & Math.Round(topstop_mtbf(SelectedFailuremonumber_inTopStopsforTrends), 1)
        MTTRLabel.Content = "MTTR: " & Math.Round(topstop_mttr(SelectedFailuremonumber_inTopStopsforTrends), 1)

    End Sub
    Sub Generate_after_Deciding_WhichRLCharttoGenerate()

        If InStr(linename_label.Content, "Fam", vbTextCompare) > 0 Then
            GenerateRL4chart(selectedDTGroup, DowntimeField.Reason2)
            Exit Sub
        End If

        Select Case My.Settings.defaultMappingLevel
            Case DowntimeField.DTGroup
                GenerateRL4chart(selectedDTGroup, DowntimeField.Reason4)
            Case DowntimeField.Reason4
                GenerateRL4chart(selectedDTGroup, DowntimeField.Reason4)
            Case DowntimeField.Reason3
                GenerateRL4chart(selectedDTGroup, DowntimeField.Reason4)
            Case DowntimeField.Reason2
                GenerateRL4chart(selectedDTGroup, DowntimeField.Reason4)
            Case DowntimeField.Reason1
                GenerateRL4chart(selectedDTGroup, DowntimeField.Reason4)
            Case DowntimeField.Fault
                GenerateRL4chart(selectedDTGroup, DowntimeField.Fault)
            Case DowntimeField.Location
                GenerateRL4chart(selectedDTGroup, DowntimeField.Reason4)


        End Select
    End Sub

    Public Sub GenerateRL4chart(selectedDTGroup1 As String, selectedcolumn As Integer)
        Const RL4barwidth = 230
        Dim fault1stops As Integer
        Dim fault2stops As Integer
        Dim fault3stops As Integer
        Dim fault1stopsandprlabel As String
        Dim fault2stopsandprlabel As String
        Dim fault3stopsandprlabel As String
        Dim fault1name, fault2name, fault3name As String
        Dim tmpDTevent_RL4 As DTevent
        Dim datalabel1Thickness As Thickness
        Dim datalabel2Thickness As Thickness
        Dim datalabel3Thickness As Thickness
        Dim Rl4barthickness As Thickness


        prStoryReport.setTopNReasonsBforReasonA(selectedDTGroup1, selectedcolumn)

        tmpDTevent_RL4 = Card_TopThreeStops(0) 'prStoryReport.getCardEventInfo(32, 0)
        fault1name = tmpDTevent_RL4.Name
        fault1stops = tmpDTevent_RL4.Stops
        fault1stopsandprlabel = tmpDTevent_RL4.Stops & " stops, " & Math.Round(tmpDTevent_RL4.DTpct, 2) & " %"

        tmpDTevent_RL4 = Card_TopThreeStops(1) 'prStoryReport.getCardEventInfo(32, 1)
        fault2name = tmpDTevent_RL4.Name
        fault2stops = tmpDTevent_RL4.Stops
        fault2stopsandprlabel = tmpDTevent_RL4.Stops & " stops, " & Math.Round(tmpDTevent_RL4.DTpct, 2) & " %"

        tmpDTevent_RL4 = Card_TopThreeStops(2) 'prStoryReport.getCardEventInfo(32, 2)
        fault3name = tmpDTevent_RL4.Name
        fault3stops = tmpDTevent_RL4.Stops
        fault3stopsandprlabel = tmpDTevent_RL4.Stops & " stops, " & Math.Round(tmpDTevent_RL4.DTpct, 2) & " %"
        'tmpDTevent_RL4.CommentList
        showRL4()


        RL4bar1.Width = RL4barwidth

        If fault1stops = 0 Then RL4bar1.Width = 0
        RL4bar2.Width = fault2stops * (RL4barwidth / fault1stops)
        RL4bar3.Width = fault3stops * (RL4barwidth / fault1stops)

        fault1label.Content = fault1name
        fault2label.Content = fault2name
        fault3label.Content = fault3name

        fault1label.ToolTip = fault1name
        fault2label.ToolTip = fault2name
        fault3label.ToolTip = fault3name

        datalabel1.Content = fault1stopsandprlabel
        datalabel2.Content = fault2stopsandprlabel
        datalabel3.Content = fault3stopsandprlabel

        datalabel1Thickness = datalabel1.Margin
        datalabel2Thickness = datalabel2.Margin
        datalabel3Thickness = datalabel3.Margin
        Rl4barthickness = RL4bar1.Margin

        RL4_Header_Label.Content = "Top 3 stops for " & selectedDTGroup1

        If fault1stops = 0 Then datalabel1.Visibility = Visibility.Hidden
        If fault2stops = 0 Then datalabel2.Visibility = Visibility.Hidden
        If fault3stops = 0 Then datalabel3.Visibility = Visibility.Hidden



        selectedRLcolumn = selectedcolumn  ' transfering data from a temp variable to a global variable


    End Sub

    Private Sub resetcircle(sender As Object, e As MouseEventArgs)
        RL4circle.Opacity = 1.0
        RL3circle.Opacity = 1.0
        RL2circle.Opacity = 1.0
        RL1circle.Opacity = 1.0
        faultcodecircle.Opacity = 1.0

        Cursor = Cursors.Arrow
    End Sub
    Private Sub animatecircle(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Hand

        If sender Is RL4circle Then

            RL4circle.Opacity = 0.8
        End If

        If sender Is RL3circle Then

            RL3circle.Opacity = 0.8
        End If

        If sender Is RL2circle Then

            RL2circle.Opacity = 0.8
        End If

        If sender Is RL1circle Then

            RL1circle.Opacity = 0.8
        End If
        If sender Is faultcodecircle Then

            faultcodecircle.Opacity = 0.8
        End If



    End Sub

    Sub circleselect(sender As Object, e As MouseButtonEventArgs)
        RL4circle.Opacity = 0.3
        RL3circle.Opacity = 0.3
        RL2circle.Opacity = 0.3
        RL1circle.Opacity = 0.3
        faultcodecircle.Opacity = 0.3
        RL4circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        RL3circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        RL2circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        RL1circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))
        faultcodecircle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(255, 255, 255))



        If sender Is RL4circle Or sender Is RL4label Then
            RL4circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(101, 222, 200))
            RL4circle.Opacity = 1.0
            GenerateRL4chart(selectedDTGroup, DowntimeField.Reason4)

        End If

        If sender Is RL3circle Or sender Is RL3label Then
            RL3circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(101, 222, 200))
            RL3circle.Opacity = 1.0
            GenerateRL4chart(selectedDTGroup, DowntimeField.Reason3)
        End If
        If sender Is RL2circle Or sender Is RL2label Then
            RL2circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(101, 222, 200))
            RL2circle.Opacity = 1.0
            GenerateRL4chart(selectedDTGroup, DowntimeField.Reason2)
        End If
        If sender Is RL1circle Or sender Is RL1label Then
            RL1circle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(101, 222, 200))
            RL1circle.Opacity = 1.0
            GenerateRL4chart(selectedDTGroup, DowntimeField.Reason1)
        End If
        If sender Is faultcodecircle Or sender Is faultcodelabel Then
            faultcodecircle.Fill = New SolidColorBrush(Windows.Media.Color.FromRgb(101, 222, 200))
            faultcodecircle.Opacity = 1.0
            GenerateRL4chart(selectedDTGroup, DowntimeField.Fault)
        End If

        hideRL4rawstopslist()

    End Sub


    Sub hideRL4()

        RL4circle.Visibility = Visibility.Hidden
        RL3circle.Visibility = Visibility.Hidden
        RL2circle.Visibility = Visibility.Hidden
        RL1circle.Visibility = Visibility.Hidden
        faultcodecircle.Visibility = Visibility.Hidden

        RL4label.Visibility = Visibility.Hidden
        RL3label.Visibility = Visibility.Hidden
        RL2label.Visibility = Visibility.Hidden
        RL1label.Visibility = Visibility.Hidden
        faultcodelabel.Visibility = Visibility.Hidden

        RL4bar1.Visibility = Visibility.Hidden
        RL4bar2.Visibility = Visibility.Hidden
        RL4bar3.Visibility = Visibility.Hidden

        fault1label.Visibility = Visibility.Hidden
        fault2label.Visibility = Visibility.Hidden
        fault3label.Visibility = Visibility.Hidden

        RL4area.Visibility = Visibility.Hidden
        RL4_Header_Label.Visibility = Visibility.Hidden
        closeRL4.Visibility = Visibility.Hidden

        datalabel1.Visibility = Visibility.Hidden
        datalabel2.Visibility = Visibility.Hidden
        datalabel3.Visibility = Visibility.Hidden

        RL4rawstopslist.Visibility = Visibility.Hidden
        CommentList.Visibility = Visibility.Hidden
        CommentListRawData_HeaderLabel.Visibility = Visibility.Hidden
        closeRL4_RL4rawstopslist.Visibility = Visibility.Hidden
        stopswatchicon.Visibility = Visibility.Hidden
        TrendStarIcon31.Visibility = Visibility.Hidden
        TrendLabel_Card31_RL4.Visibility = Visibility.Hidden
        StopswatchLabel_Card31_RL4.Visibility = Visibility.Hidden


        MTBFlabel.Visibility = Visibility.Hidden
        MTTRLabel.Visibility = Visibility.Hidden
        RL4interiorborder.Visibility = Visibility.Hidden

    End Sub
    Sub DecideWhichRLcirclestoshow()
        Select Case My.Settings.defaultMappingLevel

            Case DowntimeField.Reason1
                RL4circle.Visibility = Visibility.Visible
                RL3circle.Visibility = Visibility.Visible
                RL2circle.Visibility = Visibility.Visible

                RL4label.Visibility = Visibility.Visible
                RL3label.Visibility = Visibility.Visible
                RL2label.Visibility = Visibility.Visible
            Case DowntimeField.Reason2
                RL4circle.Visibility = Visibility.Visible
                RL3circle.Visibility = Visibility.Visible

                RL4label.Visibility = Visibility.Visible
                RL3label.Visibility = Visibility.Visible
            Case DowntimeField.Reason3
                RL4circle.Visibility = Visibility.Visible

                RL4label.Visibility = Visibility.Visible
            Case DowntimeField.Reason4
                RL4circle.Visibility = Visibility.Visible

                RL4label.Visibility = Visibility.Visible
            Case DowntimeField.Fault
                faultcodecircle.Visibility = Visibility.Visible
                faultcodelabel.Visibility = Visibility.Visible
            Case DowntimeField.Location
                RL4circle.Visibility = Visibility.Visible
                RL3circle.Visibility = Visibility.Visible
                RL2circle.Visibility = Visibility.Visible
                RL1circle.Visibility = Visibility.Visible
                faultcodecircle.Visibility = Visibility.Visible

                RL4label.Visibility = Visibility.Visible
                RL3label.Visibility = Visibility.Visible
                RL2label.Visibility = Visibility.Visible
                RL1label.Visibility = Visibility.Visible
                faultcodelabel.Visibility = Visibility.Visible
            Case DowntimeField.DTGroup
                RL4circle.Visibility = Visibility.Visible
                RL3circle.Visibility = Visibility.Visible
                RL2circle.Visibility = Visibility.Visible
                RL1circle.Visibility = Visibility.Visible
                faultcodecircle.Visibility = Visibility.Visible

                RL4label.Visibility = Visibility.Visible
                RL3label.Visibility = Visibility.Visible
                RL2label.Visibility = Visibility.Visible
                RL1label.Visibility = Visibility.Visible
                faultcodelabel.Visibility = Visibility.Visible

        End Select

        If InStr(linename_label.Content, "Fam") > 0 Then

            RL3circle.Visibility = Visibility.Hidden
            RL3label.Visibility = Visibility.Hidden
            RL4circle.Visibility = Visibility.Hidden
            RL4label.Visibility = Visibility.Hidden
        End If


    End Sub

    Sub showRL4()
        hideRL4()
        DecideWhichRLcirclestoshow()

        RL4label.Content = AllProdLines(selectedindexofLine_temp).Reason4Name
        RL3label.Content = AllProdLines(selectedindexofLine_temp).Reason3Name
        RL2label.Content = AllProdLines(selectedindexofLine_temp).Reason2Name
        RL1label.Content = AllProdLines(selectedindexofLine_temp).Reason1Name
        faultcodelabel.Content = AllProdLines(selectedindexofLine_temp).FaultCodeName


        RL4bar1.Visibility = Visibility.Visible
        RL4bar2.Visibility = Visibility.Visible
        RL4bar3.Visibility = Visibility.Visible

        fault1label.Visibility = Visibility.Visible
        fault2label.Visibility = Visibility.Visible
        fault3label.Visibility = Visibility.Visible

        RL4area.Visibility = Visibility.Visible
        RL4_Header_Label.Visibility = Visibility.Visible
        closeRL4.Visibility = Visibility.Visible

        datalabel1.Visibility = Visibility.Visible
        datalabel2.Visibility = Visibility.Visible
        datalabel3.Visibility = Visibility.Visible

        stopswatchicon.Visibility = Visibility.Visible
        TrendStarIcon31.Visibility = Visibility.Visible
        TrendLabel_Card31_RL4.Visibility = Visibility.Visible
        StopswatchLabel_Card31_RL4.Visibility = Visibility.Visible

        MTBFlabel.Visibility = Visibility.Visible
        MTTRLabel.Visibility = Visibility.Visible

        RL4interiorborder.Visibility = Visibility.Visible
    End Sub
    Private Sub RL4barselect(sender As Object, e As MouseButtonEventArgs)

        If sender Is RL4bar1 Or sender Is datalabel1 Then selectedRL4bar = fault1label.Content
        If sender Is RL4bar2 Or sender Is datalabel2 Then selectedRL4bar = fault2label.Content
        If sender Is RL4bar3 Or sender Is datalabel3 Then selectedRL4bar = fault3label.Content


        generateSTOPSRAWDATAlist(selectedRL4bar, selectedRLcolumn)
        ' CommentList.Visibility = Windows.Visibility.Hidden
    End Sub
#End Region

    Sub hidePlannedchangeoverTIMEchart()
        plannedchangeovertimechart.Visibility = Visibility.Hidden
        CObubble1.Visibility = Visibility.Hidden
        CObubble2.Visibility = Visibility.Hidden
        CObubble3.Visibility = Visibility.Hidden
        CObubble4.Visibility = Visibility.Hidden
        CObubble5.Visibility = Visibility.Hidden
        CObubble6.Visibility = Visibility.Hidden
        CObubble7.Visibility = Visibility.Hidden

        Card42Header.Visibility = Visibility.Hidden

        CObubble1_Label.Visibility = Visibility.Hidden
        CObubble2_Label.Visibility = Visibility.Hidden
        CObubble3_Label.Visibility = Visibility.Hidden
        CObubble4_Label.Visibility = Visibility.Hidden
        CObubble5_Label.Visibility = Visibility.Hidden
        CObubble6_Label.Visibility = Visibility.Hidden
        CObubble7_Label.Visibility = Visibility.Hidden
        TrendStarIcon42.Visibility = Visibility.Hidden
    End Sub
    Sub showPlannedchangeoverTIMEchart()
        UseTrack_PDTview = True
        plannedchangeovertimechart.Visibility = Visibility.Visible
        CObubble1.Visibility = Visibility.Visible
        CObubble2.Visibility = Visibility.Visible
        CObubble3.Visibility = Visibility.Visible
        CObubble4.Visibility = Visibility.Visible
        CObubble5.Visibility = Visibility.Visible
        CObubble6.Visibility = Visibility.Visible
        CObubble7.Visibility = Visibility.Visible

        CObubble1.Height = Math.Sqrt(Changeovertime1 / MaxChangeoverBubble_COtime) * changeoverbubbleMaxSize
        CObubble1.Width = CObubble1.Height
        CObubble2.Height = Math.Sqrt(Changeovertime2 / MaxChangeoverBubble_COtime) * changeoverbubbleMaxSize
        CObubble2.Width = CObubble2.Height
        CObubble3.Height = Math.Sqrt(Changeovertime3 / MaxChangeoverBubble_COtime) * changeoverbubbleMaxSize
        CObubble3.Width = CObubble3.Height
        CObubble4.Height = Math.Sqrt(Changeovertime4 / MaxChangeoverBubble_COtime) * changeoverbubbleMaxSize
        CObubble4.Width = CObubble4.Height
        CObubble5.Height = Math.Sqrt(Changeovertime5 / MaxChangeoverBubble_COtime) * changeoverbubbleMaxSize
        CObubble5.Width = CObubble5.Height
        CObubble6.Height = Math.Sqrt(Changeovertime6 / MaxChangeoverBubble_COtime) * changeoverbubbleMaxSize
        CObubble6.Width = CObubble6.Height
        CObubble7.Height = Math.Sqrt(Changeovertime7 / MaxChangeoverBubble_COtime) * changeoverbubbleMaxSize
        CObubble7.Width = CObubble7.Height

        Card42Header.Visibility = Visibility.Visible

        If Changeovertime1 > 0 Then CObubble1_Label.Visibility = Visibility.Visible
        If Changeovertime2 > 0 Then CObubble2_Label.Visibility = Visibility.Visible
        If Changeovertime3 > 0 Then CObubble3_Label.Visibility = Visibility.Visible
        If Changeovertime4 > 0 Then CObubble4_Label.Visibility = Visibility.Visible
        If Changeovertime5 > 0 Then CObubble5_Label.Visibility = Visibility.Visible
        If Changeovertime6 > 0 Then CObubble6_Label.Visibility = Visibility.Visible
        If Changeovertime7 > 0 Then CObubble7_Label.Visibility = Visibility.Visible

        CObubble1_Label.Content = changeoverlabel1string & vbNewLine & "#" & changeoverNo_of_Events1 & " ; " & Changeovertime1 & " mins"
        CObubble2_Label.Content = changeoverlabel2string & vbNewLine & "#" & changeoverNo_of_Events2 & " ; " & Changeovertime2 & " mins"
        CObubble3_Label.Content = changeoverlabel3string & vbNewLine & "#" & changeoverNo_of_Events3 & " ; " & Changeovertime3 & " mins"
        CObubble4_Label.Content = changeoverlabel4string & vbNewLine & "#" & changeoverNo_of_Events4 & " ; " & Changeovertime4 & " mins"
        CObubble5_Label.Content = changeoverlabel5string & vbNewLine & "#" & changeoverNo_of_Events5 & " ; " & Changeovertime5 & " mins"
        CObubble6_Label.Content = changeoverlabel6string & vbNewLine & "#" & changeoverNo_of_Events6 & " ; " & Changeovertime6 & " mins"
        CObubble7_Label.Content = changeoverlabel7string & vbNewLine & "#" & changeoverNo_of_Events7 & " ; " & Changeovertime7 & " mins"

        'Changeovertime_TotalAvg = Changeovertime1 + 

        TrendStarIcon42.Visibility = Visibility.Hidden
    End Sub

    Sub hidestoplabels()

        stoplabel_1.Visibility = Visibility.Hidden
        stoplabel_2.Visibility = Visibility.Hidden
        stoplabel_3.Visibility = Visibility.Hidden
        stoplabel_4.Visibility = Visibility.Hidden
        stoplabel_5.Visibility = Visibility.Hidden
        stoplabel_6.Visibility = Visibility.Hidden
        stoplabel_7.Visibility = Visibility.Hidden
        stoplabel_8.Visibility = Visibility.Hidden
        stoplabel_9.Visibility = Visibility.Hidden
        stoplabel_10.Visibility = Visibility.Hidden
        stoplabel_11.Visibility = Visibility.Hidden
        stoplabel_12.Visibility = Visibility.Hidden
        stoplabel_13.Visibility = Visibility.Hidden
        stoplabel_14.Visibility = Visibility.Hidden
        stoplabel_15.Visibility = Visibility.Hidden
        TopStops_Rect1.Visibility = Visibility.Hidden
        TopStops_Rect2.Visibility = Visibility.Hidden
        TopStops_Rect3.Visibility = Visibility.Hidden
        TopStops_Rect4.Visibility = Visibility.Hidden
        TopStops_Rect5.Visibility = Visibility.Hidden
        TopStops_Rect6.Visibility = Visibility.Hidden
        TopStops_Rect7.Visibility = Visibility.Hidden
        TopStops_Rect8.Visibility = Visibility.Hidden
        TopStops_Rect9.Visibility = Visibility.Hidden
        TopStops_Rect10.Visibility = Visibility.Hidden
        TopStops_Rect11.Visibility = Visibility.Hidden
        TopStops_Rect12.Visibility = Visibility.Hidden
        TopStops_Rect13.Visibility = Visibility.Hidden
        TopStops_Rect14.Visibility = Visibility.Hidden
        TopStops_Rect15.Visibility = Visibility.Hidden

        topstopsBubble1.Visibility = Visibility.Hidden
        topstopsBubble2.Visibility = Visibility.Hidden
        topstopsBubble3.Visibility = Visibility.Hidden
        topstopsBubble4.Visibility = Visibility.Hidden
        topstopsBubble5.Visibility = Visibility.Hidden
        topstopsBubble6.Visibility = Visibility.Hidden
        topstopsBubble7.Visibility = Visibility.Hidden
        topstopsBubble8.Visibility = Visibility.Hidden
        topstopsBubble9.Visibility = Visibility.Hidden
        topstopsBubble10.Visibility = Visibility.Hidden
        topstopsBubble11.Visibility = Visibility.Hidden
        topstopsBubble12.Visibility = Visibility.Hidden
        topstopsBubble13.Visibility = Visibility.Hidden
        topstopsBubble14.Visibility = Visibility.Hidden
        topstopsBubble15.Visibility = Visibility.Hidden


        TopStopsDataLabel1.Visibility = Visibility.Hidden
        TopStopsDataLabel2.Visibility = Visibility.Hidden
        TopStopsDataLabel3.Visibility = Visibility.Hidden
        TopStopsDataLabel4.Visibility = Visibility.Hidden
        TopStopsDataLabel5.Visibility = Visibility.Hidden
        TopStopsDataLabel6.Visibility = Visibility.Hidden
        TopStopsDataLabel7.Visibility = Visibility.Hidden
        TopStopsDataLabel8.Visibility = Visibility.Hidden
        TopStopsDataLabel9.Visibility = Visibility.Hidden
        TopStopsDataLabel10.Visibility = Visibility.Hidden
        TopStopsDataLabel11.Visibility = Visibility.Hidden
        TopStopsDataLabel12.Visibility = Visibility.Hidden
        TopStopsDataLabel13.Visibility = Visibility.Hidden
        TopStopsDataLabel14.Visibility = Visibility.Hidden
        TopStopsDataLabel15.Visibility = Visibility.Hidden



        resettopstopbubbles()
        Card31Header.Visibility = Visibility.Hidden
        TopStopsLegend_bar.Visibility = Visibility.Hidden
        TopStopsLegend_bartext.Visibility = Visibility.Hidden
        TopStopsLegend_bubble.Visibility = Visibility.Hidden
        TopStopsLegend_bubbletext.Visibility = Visibility.Hidden
        TopStopsLegend_Mapping.Visibility = Visibility.Hidden
        TrendStarIcon31.Visibility = Visibility.Hidden
        TrendLabel_Card31_RL4.Visibility = Visibility.Hidden
        StopswatchLabel_Card31_RL4.Visibility = Visibility.Hidden
    End Sub

    Sub BackButtonClicked()

        Me.Close()
        My.Settings.Save()
        '  Dim mainprstorywindow As New WindowMain_prstory
        Me.Owner.Visibility = Visibility.Visible


        bargraphreportwindow_Open = False
    End Sub
    Sub GoHomeButtonClicked()
        SendExceptionDatatoServer(ExceptionComments.Text.ToString, ErrorFunctionName)
        Me.Close()
        My.Settings.Save()
        Dim mainprstorywindow As New WindowMain_prstory
        Me.Owner.Visibility = Visibility.Visible


        bargraphreportwindow_Open = False
    End Sub

    Sub showstoplabels()
        Dim datalabelposition1 As Thickness
        Dim datalabelposition2 As Thickness
        Dim datalabelposition3 As Thickness
        Dim datalabelposition4 As Thickness
        Dim datalabelposition5 As Thickness
        Dim datalabelposition6 As Thickness
        Dim datalabelposition7 As Thickness
        Dim datalabelposition8 As Thickness
        Dim datalabelposition9 As Thickness
        Dim datalabelposition10 As Thickness
        Dim datalabelposition11 As Thickness
        Dim datalabelposition12 As Thickness
        Dim datalabelposition13 As Thickness
        Dim datalabelposition14 As Thickness
        Dim datalabelposition15 As Thickness

        Dim barposition As Thickness


        stoplabel_1.Visibility = Visibility.Visible
        stoplabel_2.Visibility = Visibility.Visible
        stoplabel_3.Visibility = Visibility.Visible
        stoplabel_4.Visibility = Visibility.Visible
        stoplabel_5.Visibility = Visibility.Visible
        stoplabel_6.Visibility = Visibility.Visible
        stoplabel_7.Visibility = Visibility.Visible
        stoplabel_8.Visibility = Visibility.Visible
        stoplabel_9.Visibility = Visibility.Visible
        stoplabel_10.Visibility = Visibility.Visible
        stoplabel_11.Visibility = Visibility.Visible
        stoplabel_12.Visibility = Visibility.Visible
        stoplabel_13.Visibility = Visibility.Visible
        stoplabel_14.Visibility = Visibility.Visible
        stoplabel_15.Visibility = Visibility.Visible

        Card31Header.Visibility = Visibility.Visible
        TopStopsLegend_bar.Visibility = Visibility.Visible
        TopStopsLegend_bartext.Visibility = Visibility.Visible
        TopStopsLegend_bubble.Visibility = Visibility.Visible
        TopStopsLegend_bubbletext.Visibility = Visibility.Visible
        'TopStopsLegend_Mapping.Visibility = Windows.Visibility.Visible

        Card31Header.Content = "TOP STOPS"

        TopStops_Rect1.Visibility = Visibility.Visible
        TopStops_Rect2.Visibility = Visibility.Visible
        TopStops_Rect3.Visibility = Visibility.Visible
        TopStops_Rect4.Visibility = Visibility.Visible
        TopStops_Rect5.Visibility = Visibility.Visible
        TopStops_Rect6.Visibility = Visibility.Visible
        TopStops_Rect7.Visibility = Visibility.Visible
        TopStops_Rect8.Visibility = Visibility.Visible
        TopStops_Rect9.Visibility = Visibility.Visible
        TopStops_Rect10.Visibility = Visibility.Visible
        TopStops_Rect11.Visibility = Visibility.Visible
        TopStops_Rect12.Visibility = Visibility.Visible
        TopStops_Rect13.Visibility = Visibility.Visible
        TopStops_Rect14.Visibility = Visibility.Visible
        TopStops_Rect15.Visibility = Visibility.Visible

        topstopsBubble1.Visibility = Visibility.Visible
        topstopsBubble2.Visibility = Visibility.Visible
        topstopsBubble3.Visibility = Visibility.Visible
        topstopsBubble4.Visibility = Visibility.Visible
        topstopsBubble5.Visibility = Visibility.Visible
        topstopsBubble6.Visibility = Visibility.Visible
        topstopsBubble7.Visibility = Visibility.Visible
        topstopsBubble8.Visibility = Visibility.Visible
        topstopsBubble9.Visibility = Visibility.Visible
        topstopsBubble10.Visibility = Visibility.Visible
        topstopsBubble11.Visibility = Visibility.Visible
        topstopsBubble12.Visibility = Visibility.Visible
        topstopsBubble13.Visibility = Visibility.Visible
        topstopsBubble14.Visibility = Visibility.Visible
        topstopsBubble15.Visibility = Visibility.Visible


        TopStopsDataLabel1.Visibility = Visibility.Visible
        TopStopsDataLabel2.Visibility = Visibility.Visible
        TopStopsDataLabel3.Visibility = Visibility.Visible
        TopStopsDataLabel4.Visibility = Visibility.Visible
        TopStopsDataLabel5.Visibility = Visibility.Visible
        TopStopsDataLabel6.Visibility = Visibility.Visible
        TopStopsDataLabel7.Visibility = Visibility.Visible
        TopStopsDataLabel8.Visibility = Visibility.Visible
        TopStopsDataLabel9.Visibility = Visibility.Visible
        TopStopsDataLabel10.Visibility = Visibility.Visible
        TopStopsDataLabel11.Visibility = Visibility.Visible
        TopStopsDataLabel12.Visibility = Visibility.Visible
        TopStopsDataLabel13.Visibility = Visibility.Visible
        TopStopsDataLabel14.Visibility = Visibility.Visible
        TopStopsDataLabel15.Visibility = Visibility.Visible


        TopStops_Rect1.Height = topstopsbar1
        TopStops_Rect2.Height = topstopsbar2
        TopStops_Rect3.Height = topstopsbar3
        TopStops_Rect4.Height = topstopsbar4
        TopStops_Rect5.Height = topstopsbar5
        TopStops_Rect6.Height = topstopsbar6
        TopStops_Rect7.Height = topstopsbar7
        TopStops_Rect8.Height = topstopsbar8
        TopStops_Rect9.Height = topstopsbar9
        TopStops_Rect10.Height = topstopsbar10
        TopStops_Rect11.Height = topstopsbar11
        TopStops_Rect12.Height = topstopsbar12
        TopStops_Rect13.Height = topstopsbar13
        TopStops_Rect14.Height = topstopsbar14
        TopStops_Rect15.Height = topstopsbar15



        barposition = TopStops_Rect1.Margin



        datalabelposition1 = TopStopsDataLabel1.Margin
        datalabelposition2 = TopStopsDataLabel2.Margin
        datalabelposition3 = TopStopsDataLabel3.Margin
        datalabelposition4 = TopStopsDataLabel4.Margin
        datalabelposition5 = TopStopsDataLabel5.Margin
        datalabelposition6 = TopStopsDataLabel6.Margin
        datalabelposition7 = TopStopsDataLabel7.Margin
        datalabelposition8 = TopStopsDataLabel8.Margin
        datalabelposition9 = TopStopsDataLabel9.Margin
        datalabelposition10 = TopStopsDataLabel10.Margin
        datalabelposition11 = TopStopsDataLabel11.Margin
        datalabelposition12 = TopStopsDataLabel12.Margin
        datalabelposition13 = TopStopsDataLabel13.Margin
        datalabelposition14 = TopStopsDataLabel14.Margin
        datalabelposition15 = TopStopsDataLabel15.Margin

        TopStopsDataLabel1.Margin = New Thickness(datalabelposition1.Left, 442.0 - (1.0 * topstopsbar1), 0, 0)
        TopStopsDataLabel2.Margin = New Thickness(datalabelposition2.Left, 442.0 - (1.0 * topstopsbar2), 0, 0)
        TopStopsDataLabel3.Margin = New Thickness(datalabelposition3.Left, 442.0 - (1.0 * topstopsbar3), 0, 0)
        TopStopsDataLabel4.Margin = New Thickness(datalabelposition4.Left, 442.0 - (1.0 * topstopsbar4), 0, 0)
        TopStopsDataLabel5.Margin = New Thickness(datalabelposition5.Left, 442.0 - (1.0 * topstopsbar5), 0, 0)
        TopStopsDataLabel6.Margin = New Thickness(datalabelposition6.Left, 442.0 - (1.0 * topstopsbar6), 0, 0)
        TopStopsDataLabel7.Margin = New Thickness(datalabelposition7.Left, 442.0 - (1.0 * topstopsbar7), 0, 0)
        TopStopsDataLabel8.Margin = New Thickness(datalabelposition8.Left, 442.0 - (1.0 * topstopsbar8), 0, 0)
        TopStopsDataLabel9.Margin = New Thickness(datalabelposition9.Left, 442.0 - (1.0 * topstopsbar9), 0, 0)
        TopStopsDataLabel10.Margin = New Thickness(datalabelposition10.Left, 442.0 - (1.0 * topstopsbar10), 0, 0)
        TopStopsDataLabel11.Margin = New Thickness(datalabelposition11.Left, 442.0 - (1.0 * topstopsbar11), 0, 0)
        TopStopsDataLabel12.Margin = New Thickness(datalabelposition12.Left, 442.0 - (1.0 * topstopsbar12), 0, 0)
        TopStopsDataLabel13.Margin = New Thickness(datalabelposition13.Left, 442.0 - (1.0 * topstopsbar13), 0, 0)
        TopStopsDataLabel14.Margin = New Thickness(datalabelposition14.Left, 442.0 - (1.0 * topstopsbar14), 0, 0)
        TopStopsDataLabel15.Margin = New Thickness(datalabelposition15.Left, 442.0 - (1.0 * topstopsbar15), 0, 0)



        TopStopsDataLabel1.Content = Math.Round(topstopsbar1_SPD, 1)
        TopStopsDataLabel2.Content = Math.Round(topstopsbar2_SPD, 1)
        TopStopsDataLabel3.Content = Math.Round(topstopsbar3_SPD, 1)
        TopStopsDataLabel4.Content = Math.Round(topstopsbar4_SPD, 1)
        TopStopsDataLabel5.Content = Math.Round(topstopsbar5_SPD, 1)
        TopStopsDataLabel6.Content = Math.Round(topstopsbar6_SPD, 1)
        TopStopsDataLabel7.Content = Math.Round(topstopsbar7_SPD, 1)
        TopStopsDataLabel8.Content = Math.Round(topstopsbar8_SPD, 1)
        TopStopsDataLabel9.Content = Math.Round(topstopsbar9_SPD, 1)
        TopStopsDataLabel10.Content = Math.Round(topstopsbar10_SPD, 1)
        TopStopsDataLabel11.Content = Math.Round(topstopsbar11_SPD, 1)
        TopStopsDataLabel12.Content = Math.Round(topstopsbar12_SPD, 1)
        TopStopsDataLabel13.Content = Math.Round(topstopsbar13_SPD, 1)
        TopStopsDataLabel14.Content = Math.Round(topstopsbar14_SPD, 1)
        TopStopsDataLabel15.Content = Math.Round(topstopsbar15_SPD, 1)

        TopStops_Rect1.ToolTip = "Stops per day " & Math.Round(topstopsbar1_SPD, 1)
        TopStops_Rect2.ToolTip = "Stops per day " & Math.Round(topstopsbar2_SPD, 1)
        TopStops_Rect3.ToolTip = "Stops per day " & Math.Round(topstopsbar3_SPD, 1)
        TopStops_Rect4.ToolTip = "Stops per day " & Math.Round(topstopsbar4_SPD, 1)
        TopStops_Rect5.ToolTip = "Stops per day " & Math.Round(topstopsbar5_SPD, 1)
        TopStops_Rect6.ToolTip = "Stops per day " & Math.Round(topstopsbar6_SPD, 1)
        TopStops_Rect7.ToolTip = "Stops per day " & Math.Round(topstopsbar7_SPD, 1)
        TopStops_Rect8.ToolTip = "Stops per day " & Math.Round(topstopsbar8_SPD, 1)
        TopStops_Rect9.ToolTip = "Stops per day " & Math.Round(topstopsbar9_SPD, 1)
        TopStops_Rect10.ToolTip = "Stops per day " & Math.Round(topstopsbar10_SPD, 1)
        TopStops_Rect11.ToolTip = "Stops per day " & Math.Round(topstopsbar11_SPD, 1)
        TopStops_Rect12.ToolTip = "Stops per day " & Math.Round(topstopsbar12_SPD, 1)
        TopStops_Rect13.ToolTip = "Stops per day " & Math.Round(topstopsbar13_SPD, 1)
        TopStops_Rect14.ToolTip = "Stops per day " & Math.Round(topstopsbar14_SPD, 1)
        TopStops_Rect15.ToolTip = "Stops per day " & Math.Round(topstopsbar15_SPD, 1)

        ''''''''''''

        Dim bubbleposition As Thickness
        Dim i As Integer
        Const maxbubbleheight = 350
        If MaxPRlossindataset_stops = 0 Then MaxPRlossindataset_stops = 1
        For i = 1 To 15
            Select Case i

                Case 1
                    bubbleposition = topstopsBubble1.Margin
                    topstopsBubble1.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar1_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble1.ToolTip = "PR Loss " & FormatPercent(topstopsbar1_PRloss, 1)

                Case 2
                    bubbleposition = topstopsBubble2.Margin
                    topstopsBubble2.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar2_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble2.ToolTip = "PR Loss " & FormatPercent(topstopsbar2_PRloss, 1)
                Case 3
                    bubbleposition = topstopsBubble3.Margin
                    topstopsBubble3.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar3_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble3.ToolTip = "PR Loss " & FormatPercent(topstopsbar3_PRloss, 1)
                Case 4
                    bubbleposition = topstopsBubble4.Margin
                    topstopsBubble4.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar4_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble4.ToolTip = "PR Loss " & FormatPercent(topstopsbar4_PRloss, 1)
                Case 5
                    bubbleposition = topstopsBubble5.Margin
                    topstopsBubble5.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar5_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble5.ToolTip = "PR Loss " & FormatPercent(topstopsbar5_PRloss, 1)
                Case 6
                    bubbleposition = topstopsBubble6.Margin
                    topstopsBubble6.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar6_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble6.ToolTip = "PR Loss " & FormatPercent(topstopsbar6_PRloss, 1)
                Case 7
                    bubbleposition = topstopsBubble7.Margin
                    topstopsBubble7.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar7_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble7.ToolTip = "PR Loss " & FormatPercent(topstopsbar7_PRloss, 1)
                Case 8
                    bubbleposition = topstopsBubble8.Margin
                    topstopsBubble8.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar8_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble8.ToolTip = "PR Loss " & FormatPercent(topstopsbar8_PRloss, 1)
                Case 9
                    bubbleposition = topstopsBubble9.Margin
                    topstopsBubble9.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar9_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble9.ToolTip = "PR Loss " & FormatPercent(topstopsbar9_PRloss, 1)
                Case 10
                    bubbleposition = topstopsBubble10.Margin
                    topstopsBubble10.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar10_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble10.ToolTip = "PR Loss " & FormatPercent(topstopsbar10_PRloss, 1)
                Case 11
                    bubbleposition = topstopsBubble11.Margin
                    topstopsBubble11.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar11_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble11.ToolTip = "PR Loss " & FormatPercent(topstopsbar11_PRloss, 1)
                Case 12
                    bubbleposition = topstopsBubble12.Margin
                    topstopsBubble12.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar12_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble12.ToolTip = "PR Loss " & FormatPercent(topstopsbar12_PRloss, 1)
                Case 13
                    bubbleposition = topstopsBubble13.Margin
                    topstopsBubble13.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar13_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble13.ToolTip = "PR Loss " & FormatPercent(topstopsbar13_PRloss, 1)
                Case 14
                    bubbleposition = topstopsBubble14.Margin
                    topstopsBubble14.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar14_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble14.ToolTip = "PR Loss " & FormatPercent(topstopsbar14_PRloss, 1)
                Case 15

                    bubbleposition = topstopsBubble15.Margin
                    topstopsBubble15.Margin = New Thickness(bubbleposition.Left, 450.0 - (maxbubbleheight * topstopsbar15_PRloss / MaxPRlossindataset_stops), 0, 0)
                    topstopsBubble15.ToolTip = "PR Loss " & FormatPercent(topstopsbar15_PRloss, 1)


            End Select

        Next




        TrendStarIcon31.Visibility = Visibility.Hidden
        TrendLabel_Card31_RL4.Visibility = Visibility.Hidden
        StopswatchLabel_Card31_RL4.Visibility = Visibility.Hidden
        resettopstopbubbles()

    End Sub

    Sub Assignlabelnames(prStoryReport As prStoryMainPageReport)
        Dim tmpDtEvent1 As DTevent

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 0)
        stoplabel_1.Text = tmpDtEvent1.Name
        stoplabel_1.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 1)
        stoplabel_2.Text = tmpDtEvent1.Name
        stoplabel_2.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 2)
        stoplabel_3.Text = tmpDtEvent1.Name
        stoplabel_3.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 3)
        stoplabel_4.Text = tmpDtEvent1.Name
        stoplabel_4.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 4)
        stoplabel_5.Text = tmpDtEvent1.Name
        stoplabel_5.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 5)
        stoplabel_6.Text = tmpDtEvent1.Name
        stoplabel_6.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 6)
        stoplabel_7.Text = tmpDtEvent1.Name
        stoplabel_7.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 7)
        stoplabel_8.Text = tmpDtEvent1.Name
        stoplabel_8.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 8)
        stoplabel_9.Text = tmpDtEvent1.Name
        stoplabel_9.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 9)
        stoplabel_10.Text = tmpDtEvent1.Name
        stoplabel_10.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 10)
        stoplabel_11.Text = tmpDtEvent1.Name
        stoplabel_11.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 11)
        stoplabel_12.Text = tmpDtEvent1.Name
        stoplabel_12.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 12)
        stoplabel_13.Text = tmpDtEvent1.Name
        stoplabel_13.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 13)
        stoplabel_14.Text = tmpDtEvent1.Name
        stoplabel_14.ToolTip = tmpDtEvent1.Name

        tmpDtEvent1 = prStoryReport.getCardEventInfo(31, 14)
        stoplabel_15.Text = tmpDtEvent1.Name
        stoplabel_15.ToolTip = tmpDtEvent1.Name



    End Sub

#Region "Stopswatch"
    Sub stopswatchlaunch()
        stopsWatchThread = New Thread(AddressOf stopswatchlaunch_Thread)
        stopsWatchThread.SetApartmentState(ApartmentState.STA)
        stopsWatchThread.Start() 'prStoryReport)
    End Sub
    Private Sub stopswatchlaunch_Thread() 'ByVal prReportX As Object)
        Dim stopswatchwindow As New Window_StopsWatch_Daily(prStoryReport, selectedfailuremode) 'prreportxyz)
        stopswatchwindow.Show()
        System.Windows.Threading.Dispatcher.Run()
    End Sub

    Sub generateSTOPSRAWDATAlist(selectedfailuremode As String, selectedcolumn As Integer)

        RL4rawstopslist.Visibility = Visibility.Visible
        closeRL4_RL4rawstopslist.Visibility = Visibility.Visible
        CommentList.Visibility = Visibility.Visible
        CommentListRawData_HeaderLabel.Visibility = Visibility.Visible
        CommentListRawData_HeaderLabel.Content = " Raw data - " & selectedfailuremode
        With prStoryReport
            If Card_TopThreeStops(0).Name.Equals(selectedfailuremode) Then
                transferToCommentColelction(Card_TopThreeStops(0))
                ''         transferToCommentColelction(.TopThreeList(0).getMeAsDTevent)
            ElseIf Card_TopThreeStops(1).Name.Equals(selectedfailuremode) Then
                transferToCommentColelction(Card_TopThreeStops(1))
                ''         transferToCommentColelction(.TopThreeList(1).getMeAsDTevent)
            ElseIf Card_TopThreeStops(2).Name.Equals(selectedfailuremode) Then
                transferToCommentColelction(Card_TopThreeStops(2))
                ''         transferToCommentColelction(.TopThreeList(2).getMeAsDTevent)
            Else
                Debugger.Break()
                'oops...
            End If
        End With
    End Sub

    Sub hideRL4rawstopslist()

        closeRL4_RL4rawstopslist.Visibility = Visibility.Hidden
        RL4rawstopslist.Visibility = Visibility.Hidden
        CommentList.Visibility = Visibility.Hidden
        CommentListRawData_HeaderLabel.Visibility = Visibility.Hidden

    End Sub
#End Region

#Region "Show / Hide Labels"
    Sub hideUPDTlabels()

        UPDTlabel1.Visibility = Visibility.Hidden
        UPDTlabel2.Visibility = Visibility.Hidden
        UPDTlabel3.Visibility = Visibility.Hidden
        UPDTlabel4.Visibility = Visibility.Hidden
        UPDTlabel5.Visibility = Visibility.Hidden
        UPDTlabel6.Visibility = Visibility.Hidden

        UPDT_Rect_1.Visibility = Visibility.Hidden
        UPDT_Rect_2.Visibility = Visibility.Hidden
        UPDT_Rect_3.Visibility = Visibility.Hidden
        UPDT_Rect_4.Visibility = Visibility.Hidden
        UPDT_Rect_5.Visibility = Visibility.Hidden
        UPDT_Rect_6.Visibility = Visibility.Hidden


        UPDT_datalabel1.Visibility = Visibility.Hidden
        UPDT_datalabel2.Visibility = Visibility.Hidden
        UPDT_datalabel3.Visibility = Visibility.Hidden
        UPDT_datalabel4.Visibility = Visibility.Hidden
        UPDT_datalabel5.Visibility = Visibility.Hidden
        UPDT_datalabel6.Visibility = Visibility.Hidden

        Card1Header.Visibility = Visibility.Hidden
        TrendStarIcon1.Visibility = Visibility.Hidden

        UPDT_Target_Rect1.Visibility = Visibility.Hidden
        UPDT_Target_Rect2.Visibility = Visibility.Hidden
        UPDT_Target_Rect3.Visibility = Visibility.Hidden
        UPDT_Target_Rect4.Visibility = Visibility.Hidden
        UPDT_Target_Rect5.Visibility = Visibility.Hidden
        UPDT_Target_Rect6.Visibility = Visibility.Hidden

        HideSimRectangles_UPDT()

        NavigationLeft_card1.Visibility = Visibility.Hidden
        NavigationRight_Card1.Visibility = Visibility.Hidden

    End Sub


    Private Sub LoadTargetsUI(cardnumbercalled As Integer, tempcardlist As List(Of DTevent))
        Dim i As Integer
        Dim maxreferencepr As Double

        Select Case cardnumbercalled
            Case 2
                maxreferencepr = MaxPRindataset_planned
            Case 41
                maxreferencepr = MaxPRindataset_planned
            Case Else
                maxreferencepr = MaxPRindataset
        End Select


        For i = 1 To prStoryReport.getCardEventFields(cardnumbercalled)
            Target_Prloss(i) = 0

            If IsNothing(AllProdLines(selectedindexofLine_temp).DowntimePercentTargets.getTargetValue(tempcardlist(i - 1).Name, cardnumbercalled)) Then GoTo skiptargetfor
            'Target_Prloss(i) = AllProductionLines(selectedindexofLine_temp).DowntimePercentTargets.getTargetValue(tempcardlist(i - 1).Name, cardnumbercalled)



            If AllProdLines(selectedindexofLine_temp).DowntimePercentTargets.getMaxDTpct(cardnumbercalled) < maxreferencepr Then
                Target_Prloss(i) = AllProdLines(selectedindexofLine_temp).DowntimePercentTargets.getTargetValue(tempcardlist(i - 1).Name, cardnumbercalled) * (barchartBarMaxSize / maxreferencepr)

            Else
                'Target_Prloss(i) = AllProductionLines(selectedindexofLine_temp).DowntimePercentTargets.getTargetValue(tempcardlist(i - 1).Name, cardnumbercalled) * (barchartBarMaxSize / AllProductionLines(selectedindexofLine_temp).DowntimePercentTargets.getMaxDTpct(cardnumbercalled))
                Target_Prloss(i) = AllProdLines(selectedindexofLine_temp).DowntimePercentTargets.getTargetValue(tempcardlist(i - 1).Name, cardnumbercalled) * (barchartBarMaxSize / maxreferencepr)
            End If

            Target_PRloss_Tooltip(i) = "Loss Target: " & FormatPercent(AllProdLines(selectedindexofLine_temp).DowntimePercentTargets.getTargetValue(tempcardlist(i - 1).Name, cardnumbercalled), 1)

skiptargetfor:
        Next




    End Sub

    Private Sub ShowTargets_Unplanned()
        Dim targetposition1 As Thickness
        Dim targetposition2 As Thickness
        Dim targetposition3 As Thickness
        Dim targetposition4 As Thickness
        Dim targetposition5 As Thickness
        Dim targetposition6 As Thickness
        Dim tier1_target_rect_top As Double

        targetposition1 = UPDT_Target_Rect1.Margin
        targetposition2 = UPDT_Target_Rect2.Margin
        targetposition3 = UPDT_Target_Rect3.Margin
        targetposition4 = UPDT_Target_Rect4.Margin
        targetposition5 = UPDT_Target_Rect5.Margin
        targetposition6 = UPDT_Target_Rect6.Margin

        tier1_target_rect_top = 230.0

        UPDT_Target_Rect1.Margin = New Thickness(targetposition1.Left, tier1_target_rect_top - (1.0 * Target_Prloss(1)), 0, 0)
        UPDT_Target_Rect2.Margin = New Thickness(targetposition2.Left, tier1_target_rect_top - (1.0 * Target_Prloss(2)), 0, 0)
        UPDT_Target_Rect3.Margin = New Thickness(targetposition3.Left, tier1_target_rect_top - (1.0 * Target_Prloss(3)), 0, 0)
        UPDT_Target_Rect4.Margin = New Thickness(targetposition4.Left, tier1_target_rect_top - (1.0 * Target_Prloss(4)), 0, 0)
        UPDT_Target_Rect5.Margin = New Thickness(targetposition5.Left, tier1_target_rect_top - (1.0 * Target_Prloss(5)), 0, 0)
        UPDT_Target_Rect6.Margin = New Thickness(targetposition6.Left, tier1_target_rect_top - (1.0 * Target_Prloss(6)), 0, 0)

        UPDT_Target_Rect1.Visibility = Visibility.Visible
        UPDT_Target_Rect2.Visibility = Visibility.Visible
        UPDT_Target_Rect3.Visibility = Visibility.Visible
        UPDT_Target_Rect4.Visibility = Visibility.Visible
        UPDT_Target_Rect5.Visibility = Visibility.Visible
        UPDT_Target_Rect6.Visibility = Visibility.Visible

        UPDT_Target_Rect1.ToolTip = Target_PRloss_Tooltip(1)
        UPDT_Target_Rect2.ToolTip = Target_PRloss_Tooltip(2)
        UPDT_Target_Rect3.ToolTip = Target_PRloss_Tooltip(3)
        UPDT_Target_Rect4.ToolTip = Target_PRloss_Tooltip(4)
        UPDT_Target_Rect5.ToolTip = Target_PRloss_Tooltip(5)
        UPDT_Target_Rect6.ToolTip = Target_PRloss_Tooltip(6)

    End Sub
    Private Sub ShowTargets_planned()
        Dim targetposition1 As Thickness
        Dim targetposition2 As Thickness
        Dim targetposition3 As Thickness
        Dim targetposition4 As Thickness
        Dim targetposition5 As Thickness
        Dim targetposition6 As Thickness
        Dim targetposition7 As Thickness
        Dim targetposition8 As Thickness
        Dim targetposition9 As Thickness


        Dim tier1_target_rect_top As Double

        targetposition1 = PDT_Target_Rect1.Margin
        targetposition2 = PDT_Target_Rect2.Margin
        targetposition3 = PDT_Target_Rect3.Margin
        targetposition4 = PDT_Target_Rect4.Margin
        targetposition5 = PDT_Target_Rect5.Margin
        targetposition6 = PDT_Target_Rect6.Margin
        targetposition7 = PDT_Target_Rect7.Margin
        targetposition8 = PDT_Target_Rect8.Margin
        targetposition9 = PDT_Target_Rect9.Margin

        tier1_target_rect_top = 230.0

        PDT_Target_Rect1.Margin = New Thickness(targetposition1.Left, tier1_target_rect_top - (1.0 * Target_Prloss(1)), 0, 0)
        PDT_Target_Rect2.Margin = New Thickness(targetposition2.Left, tier1_target_rect_top - (1.0 * Target_Prloss(2)), 0, 0)
        PDT_Target_Rect3.Margin = New Thickness(targetposition3.Left, tier1_target_rect_top - (1.0 * Target_Prloss(3)), 0, 0)
        PDT_Target_Rect4.Margin = New Thickness(targetposition4.Left, tier1_target_rect_top - (1.0 * Target_Prloss(4)), 0, 0)
        PDT_Target_Rect5.Margin = New Thickness(targetposition5.Left, tier1_target_rect_top - (1.0 * Target_Prloss(5)), 0, 0)
        PDT_Target_Rect6.Margin = New Thickness(targetposition6.Left, tier1_target_rect_top - (1.0 * Target_Prloss(6)), 0, 0)
        PDT_Target_Rect7.Margin = New Thickness(targetposition7.Left, tier1_target_rect_top - (1.0 * Target_Prloss(7)), 0, 0)
        PDT_Target_Rect8.Margin = New Thickness(targetposition8.Left, tier1_target_rect_top - (1.0 * Target_Prloss(8)), 0, 0)
        PDT_Target_Rect9.Margin = New Thickness(targetposition9.Left, tier1_target_rect_top - (1.0 * Target_Prloss(9)), 0, 0)



        PDT_Target_Rect1.Visibility = Visibility.Visible
        PDT_Target_Rect2.Visibility = Visibility.Visible
        PDT_Target_Rect3.Visibility = Visibility.Visible
        PDT_Target_Rect4.Visibility = Visibility.Visible
        PDT_Target_Rect5.Visibility = Visibility.Visible
        PDT_Target_Rect6.Visibility = Visibility.Visible
        PDT_Target_Rect7.Visibility = Visibility.Visible
        PDT_Target_Rect8.Visibility = Visibility.Visible
        PDT_Target_Rect9.Visibility = Visibility.Visible

        PDT_Target_Rect1.ToolTip = Target_PRloss_Tooltip(1)
        PDT_Target_Rect2.ToolTip = Target_PRloss_Tooltip(2)
        PDT_Target_Rect3.ToolTip = Target_PRloss_Tooltip(3)
        PDT_Target_Rect4.ToolTip = Target_PRloss_Tooltip(4)
        PDT_Target_Rect5.ToolTip = Target_PRloss_Tooltip(5)
        PDT_Target_Rect6.ToolTip = Target_PRloss_Tooltip(6)
        PDT_Target_Rect7.ToolTip = Target_PRloss_Tooltip(7)
        PDT_Target_Rect8.ToolTip = Target_PRloss_Tooltip(8)
        PDT_Target_Rect9.ToolTip = Target_PRloss_Tooltip(9)

    End Sub
    Private Sub ShowTargets_plannedT2()
        Dim targetposition1 As Thickness
        Dim targetposition2 As Thickness
        Dim targetposition3 As Thickness
        Dim targetposition4 As Thickness
        Dim targetposition5 As Thickness
        Dim targetposition6 As Thickness
        Dim targetposition7 As Thickness



        Dim tier1_target_rect_top As Double

        targetposition1 = CO_Target_Rect1.Margin
        targetposition2 = CO_Target_Rect2.Margin
        targetposition3 = CO_Target_Rect3.Margin
        targetposition4 = CO_Target_Rect4.Margin
        targetposition5 = CO_Target_Rect5.Margin
        targetposition6 = CO_Target_Rect6.Margin
        targetposition7 = CO_Target_Rect7.Margin


        tier1_target_rect_top = 490.0

        CO_Target_Rect1.Margin = New Thickness(targetposition1.Left, tier1_target_rect_top - (1.0 * Target_Prloss(1)), 0, 0)
        CO_Target_Rect2.Margin = New Thickness(targetposition2.Left, tier1_target_rect_top - (1.0 * Target_Prloss(2)), 0, 0)
        CO_Target_Rect3.Margin = New Thickness(targetposition3.Left, tier1_target_rect_top - (1.0 * Target_Prloss(3)), 0, 0)
        CO_Target_Rect4.Margin = New Thickness(targetposition4.Left, tier1_target_rect_top - (1.0 * Target_Prloss(4)), 0, 0)
        CO_Target_Rect5.Margin = New Thickness(targetposition5.Left, tier1_target_rect_top - (1.0 * Target_Prloss(5)), 0, 0)
        CO_Target_Rect6.Margin = New Thickness(targetposition6.Left, tier1_target_rect_top - (1.0 * Target_Prloss(6)), 0, 0)
        CO_Target_Rect7.Margin = New Thickness(targetposition7.Left, tier1_target_rect_top - (1.0 * Target_Prloss(7)), 0, 0)




        CO_Target_Rect1.Visibility = Visibility.Visible
        CO_Target_Rect2.Visibility = Visibility.Visible
        CO_Target_Rect3.Visibility = Visibility.Visible
        CO_Target_Rect4.Visibility = Visibility.Visible
        CO_Target_Rect5.Visibility = Visibility.Visible
        CO_Target_Rect6.Visibility = Visibility.Visible
        CO_Target_Rect7.Visibility = Visibility.Visible


        CO_Target_Rect1.ToolTip = Target_PRloss_Tooltip(1)
        CO_Target_Rect2.ToolTip = Target_PRloss_Tooltip(2)
        CO_Target_Rect3.ToolTip = Target_PRloss_Tooltip(3)
        CO_Target_Rect4.ToolTip = Target_PRloss_Tooltip(4)
        CO_Target_Rect5.ToolTip = Target_PRloss_Tooltip(5)
        CO_Target_Rect6.ToolTip = Target_PRloss_Tooltip(6)
        CO_Target_Rect7.ToolTip = Target_PRloss_Tooltip(7)


    End Sub
    Private Sub ShowTargets_UnplannedT2()
        Dim targetposition1 As Thickness
        Dim targetposition2 As Thickness
        Dim targetposition3 As Thickness
        Dim targetposition4 As Thickness
        Dim targetposition5 As Thickness
        Dim targetposition6 As Thickness



        Dim tier1_target_rect_top As Double

        targetposition1 = EquipMain_Target_Rect1.Margin
        targetposition2 = EquipMain_Target_Rect2.Margin
        targetposition3 = EquipMain_Target_Rect3.Margin
        targetposition4 = EquipMain_Target_Rect4.Margin
        targetposition5 = EquipMain_Target_Rect5.Margin
        targetposition6 = EquipMain_Target_Rect6.Margin


        tier1_target_rect_top = 490.0

        EquipMain_Target_Rect1.Margin = New Thickness(targetposition1.Left, tier1_target_rect_top - (1.0 * Target_Prloss(1)), 0, 0)
        EquipMain_Target_Rect2.Margin = New Thickness(targetposition2.Left, tier1_target_rect_top - (1.0 * Target_Prloss(2)), 0, 0)
        EquipMain_Target_Rect3.Margin = New Thickness(targetposition3.Left, tier1_target_rect_top - (1.0 * Target_Prloss(3)), 0, 0)
        EquipMain_Target_Rect4.Margin = New Thickness(targetposition4.Left, tier1_target_rect_top - (1.0 * Target_Prloss(4)), 0, 0)
        EquipMain_Target_Rect5.Margin = New Thickness(targetposition5.Left, tier1_target_rect_top - (1.0 * Target_Prloss(5)), 0, 0)
        EquipMain_Target_Rect6.Margin = New Thickness(targetposition6.Left, tier1_target_rect_top - (1.0 * Target_Prloss(6)), 0, 0)




        EquipMain_Target_Rect1.Visibility = Visibility.Visible
        EquipMain_Target_Rect2.Visibility = Visibility.Visible
        EquipMain_Target_Rect3.Visibility = Visibility.Visible
        EquipMain_Target_Rect4.Visibility = Visibility.Visible
        EquipMain_Target_Rect5.Visibility = Visibility.Visible
        EquipMain_Target_Rect6.Visibility = Visibility.Visible


        EquipMain_Target_Rect1.ToolTip = Target_PRloss_Tooltip(1)
        EquipMain_Target_Rect2.ToolTip = Target_PRloss_Tooltip(2)
        EquipMain_Target_Rect3.ToolTip = Target_PRloss_Tooltip(3)
        EquipMain_Target_Rect4.ToolTip = Target_PRloss_Tooltip(4)
        EquipMain_Target_Rect5.ToolTip = Target_PRloss_Tooltip(5)
        EquipMain_Target_Rect6.ToolTip = Target_PRloss_Tooltip(6)


    End Sub

    Private Sub ShowTargets_UnplannedT3_1()
        Dim targetposition1 As Thickness
        Dim targetposition2 As Thickness
        Dim targetposition3 As Thickness




        Dim tier1_target_rect_top As Double

        targetposition1 = Equip1_Target_Rect1.Margin
        targetposition2 = Equip1_Target_Rect2.Margin
        targetposition3 = Equip1_Target_Rect3.Margin



        tier1_target_rect_top = 490.0

        Equip1_Target_Rect1.Margin = New Thickness(targetposition1.Left, tier1_target_rect_top - (1.0 * Target_Prloss(1)), 0, 0)
        Equip1_Target_Rect2.Margin = New Thickness(targetposition2.Left, tier1_target_rect_top - (1.0 * Target_Prloss(2)), 0, 0)
        Equip1_Target_Rect3.Margin = New Thickness(targetposition3.Left, tier1_target_rect_top - (1.0 * Target_Prloss(3)), 0, 0)





        Equip1_Target_Rect1.Visibility = Visibility.Visible
        Equip1_Target_Rect2.Visibility = Visibility.Visible
        Equip1_Target_Rect3.Visibility = Visibility.Visible



        Equip1_Target_Rect1.ToolTip = Target_PRloss_Tooltip(1)
        Equip1_Target_Rect2.ToolTip = Target_PRloss_Tooltip(2)
        Equip1_Target_Rect3.ToolTip = Target_PRloss_Tooltip(3)



    End Sub
    Private Sub ShowTargets_UnplannedT3_2()
        Dim targetposition1 As Thickness
        Dim targetposition2 As Thickness
        Dim targetposition3 As Thickness




        Dim tier1_target_rect_top As Double

        targetposition1 = Equip2_Target_Rect1.Margin
        targetposition2 = Equip2_Target_Rect2.Margin
        targetposition3 = Equip2_Target_Rect3.Margin



        tier1_target_rect_top = 490.0

        Equip2_Target_Rect1.Margin = New Thickness(targetposition1.Left, tier1_target_rect_top - (1.0 * Target_Prloss(1)), 0, 0)
        Equip2_Target_Rect2.Margin = New Thickness(targetposition2.Left, tier1_target_rect_top - (1.0 * Target_Prloss(2)), 0, 0)
        Equip2_Target_Rect3.Margin = New Thickness(targetposition3.Left, tier1_target_rect_top - (1.0 * Target_Prloss(3)), 0, 0)





        Equip2_Target_Rect1.Visibility = Visibility.Visible
        Equip2_Target_Rect2.Visibility = Visibility.Visible
        Equip2_Target_Rect3.Visibility = Visibility.Visible



        Equip2_Target_Rect1.ToolTip = Target_PRloss_Tooltip(1)
        Equip2_Target_Rect2.ToolTip = Target_PRloss_Tooltip(2)
        Equip2_Target_Rect3.ToolTip = Target_PRloss_Tooltip(3)



    End Sub
    Private Sub ShowTargets_UnplannedT3_3()
        Dim targetposition1 As Thickness
        Dim targetposition2 As Thickness
        Dim targetposition3 As Thickness




        Dim tier1_target_rect_top As Double

        targetposition1 = Equip3_Target_Rect1.Margin
        targetposition2 = Equip3_Target_Rect2.Margin
        targetposition3 = Equip3_Target_Rect3.Margin



        tier1_target_rect_top = 490.0

        Equip3_Target_Rect1.Margin = New Thickness(targetposition1.Left, tier1_target_rect_top - (1.0 * Target_Prloss(1)), 0, 0)
        Equip3_Target_Rect2.Margin = New Thickness(targetposition2.Left, tier1_target_rect_top - (1.0 * Target_Prloss(2)), 0, 0)
        Equip3_Target_Rect3.Margin = New Thickness(targetposition3.Left, tier1_target_rect_top - (1.0 * Target_Prloss(3)), 0, 0)





        Equip3_Target_Rect1.Visibility = Visibility.Visible
        Equip3_Target_Rect2.Visibility = Visibility.Visible
        Equip3_Target_Rect3.Visibility = Visibility.Visible



        Equip3_Target_Rect1.ToolTip = Target_PRloss_Tooltip(1)
        Equip3_Target_Rect2.ToolTip = Target_PRloss_Tooltip(2)
        Equip3_Target_Rect3.ToolTip = Target_PRloss_Tooltip(3)



    End Sub
    Private Sub showUPDTlabels()
        Dim datalabelposition1 As Thickness
        Dim datalabelposition2 As Thickness
        Dim datalabelposition3 As Thickness
        Dim datalabelposition4 As Thickness
        Dim datalabelposition5 As Thickness
        Dim datalabelposition6 As Thickness

        Dim barposition As Thickness

        UPDTlabel1.Visibility = Visibility.Visible
        UPDTlabel2.Visibility = Visibility.Visible
        UPDTlabel3.Visibility = Visibility.Visible
        UPDTlabel4.Visibility = Visibility.Visible
        UPDTlabel5.Visibility = Visibility.Visible
        UPDTlabel6.Visibility = Visibility.Visible
        UPDT_Rect_1.Visibility = Visibility.Visible
        UPDT_Rect_2.Visibility = Visibility.Visible
        UPDT_Rect_3.Visibility = Visibility.Visible
        UPDT_Rect_4.Visibility = Visibility.Visible
        UPDT_Rect_5.Visibility = Visibility.Visible
        UPDT_Rect_6.Visibility = Visibility.Visible
        UPDT_datalabel1.Visibility = Visibility.Visible
        UPDT_datalabel2.Visibility = Visibility.Visible
        UPDT_datalabel3.Visibility = Visibility.Visible
        UPDT_datalabel4.Visibility = Visibility.Visible
        UPDT_datalabel5.Visibility = Visibility.Visible
        UPDT_datalabel6.Visibility = Visibility.Visible
        Card1Header.Visibility = Visibility.Visible

        Card1Header.Content = cardnameLabeltext(1)
        unplannedDTchart.Visibility = Visibility.Visible

        UPDTlabel1.Content = updtlabel1string
        UPDTlabel2.Content = updtlabel2string
        UPDTlabel3.Content = updtlabel3string
        UPDTlabel4.Content = updtlabel4string
        UPDTlabel5.Content = updtlabel5string
        UPDTlabel6.Content = updtlabel6string



        UPDTlabel1.Background = LabelDefaultColor
        UPDTlabel2.Background = LabelDefaultColor
        UPDTlabel3.Background = LabelDefaultColor
        UPDTlabel4.Background = LabelDefaultColor
        UPDTlabel5.Background = LabelDefaultColor
        UPDTlabel6.Background = LabelDefaultColor



        UPDTlabel1.ToolTip = updtlabel1string
        UPDTlabel2.ToolTip = updtlabel2string
        UPDTlabel3.ToolTip = updtlabel3string
        UPDTlabel4.ToolTip = updtlabel4string
        UPDTlabel5.ToolTip = updtlabel5string
        UPDTlabel6.ToolTip = updtlabel6string



        UPDT_Rect_1.Height = UPDTbar1
        UPDT_Rect_2.Height = UPDTbar2
        UPDT_Rect_3.Height = UPDTbar3
        UPDT_Rect_4.Height = UPDTbar4
        UPDT_Rect_5.Height = UPDTbar5
        UPDT_Rect_6.Height = UPDTbar6

        'SIM
        UPDT_Rect_1sim.Height = UPDTbar1sim
        UPDT_Rect_2sim.Height = UPDTbar2sim
        UPDT_Rect_3sim.Height = UPDTbar3sim
        UPDT_Rect_4sim.Height = UPDTbar4sim
        UPDT_Rect_5sim.Height = UPDTbar5sim
        UPDT_Rect_6sim.Height = UPDTbar6sim

        datalabelposition1 = UPDT_datalabel1.Margin
        datalabelposition2 = UPDT_datalabel2.Margin
        datalabelposition3 = UPDT_datalabel3.Margin
        datalabelposition4 = UPDT_datalabel4.Margin
        datalabelposition5 = UPDT_datalabel5.Margin
        datalabelposition6 = UPDT_datalabel6.Margin

        barposition = UPDT_Rect_1.Margin


        UPDT_datalabel1.Margin = New Thickness(datalabelposition1.Left, 210.0 - (1.0 * UPDTbar1), 0, 0)
        UPDT_datalabel2.Margin = New Thickness(datalabelposition2.Left, 210.0 - (1.0 * UPDTbar2), 0, 0)
        UPDT_datalabel3.Margin = New Thickness(datalabelposition3.Left, 210.0 - (1.0 * UPDTbar3), 0, 0)
        UPDT_datalabel4.Margin = New Thickness(datalabelposition4.Left, 210.0 - (1.0 * UPDTbar4), 0, 0)
        UPDT_datalabel5.Margin = New Thickness(datalabelposition5.Left, 210.0 - (1.0 * UPDTbar5), 0, 0)
        UPDT_datalabel6.Margin = New Thickness(datalabelposition6.Left, 210.0 - (1.0 * UPDTbar6), 0, 0)


        'SIM
        datalabelposition1 = UPDT_datalabel1sim.Margin
        datalabelposition2 = UPDT_datalabel2sim.Margin
        datalabelposition3 = UPDT_datalabel3sim.Margin
        datalabelposition4 = UPDT_datalabel4sim.Margin
        datalabelposition5 = UPDT_datalabel5sim.Margin
        datalabelposition6 = UPDT_datalabel6sim.Margin

        barposition = UPDT_Rect_1sim.Margin


        UPDT_datalabel1sim.Margin = New Thickness(datalabelposition1.Left, 210.0 - (1.0 * UPDTbar1sim), 0, 0)
        UPDT_datalabel2sim.Margin = New Thickness(datalabelposition2.Left, 210.0 - (1.0 * UPDTbar2sim), 0, 0)
        UPDT_datalabel3sim.Margin = New Thickness(datalabelposition3.Left, 210.0 - (1.0 * UPDTbar3sim), 0, 0)
        UPDT_datalabel4sim.Margin = New Thickness(datalabelposition4.Left, 210.0 - (1.0 * UPDTbar4sim), 0, 0)
        UPDT_datalabel5sim.Margin = New Thickness(datalabelposition5.Left, 210.0 - (1.0 * UPDTbar5sim), 0, 0)
        UPDT_datalabel6sim.Margin = New Thickness(datalabelposition6.Left, 210.0 - (1.0 * UPDTbar6sim), 0, 0)

        '




        UPDT_datalabel1.Content = FormatPercent(UPDTbar1_PRloss, 1)
        UPDT_datalabel2.Content = FormatPercent(UPDTbar2_PRloss, 1)
        UPDT_datalabel3.Content = FormatPercent(UPDTbar3_PRloss, 1)
        UPDT_datalabel4.Content = FormatPercent(UPDTbar4_PRloss, 1)
        UPDT_datalabel5.Content = FormatPercent(UPDTbar5_PRloss, 1)
        UPDT_datalabel6.Content = FormatPercent(UPDTbar6_PRloss, 1)

        'SIM
        UPDT_datalabel1sim.Content = FormatPercent(UPDTbar1_PRlosssim, 1)
        UPDT_datalabel2sim.Content = FormatPercent(UPDTbar2_PRlosssim, 1)
        UPDT_datalabel3sim.Content = FormatPercent(UPDTbar3_PRlosssim, 1)
        UPDT_datalabel4sim.Content = FormatPercent(UPDTbar4_PRlosssim, 1)
        UPDT_datalabel5sim.Content = FormatPercent(UPDTbar5_PRlosssim, 1)
        UPDT_datalabel6sim.Content = FormatPercent(UPDTbar6_PRlosssim, 1)
        '

        UPDT_Rect_1.ToolTip = "PR Loss " & FormatPercent(UPDTbar1_PRloss, 1)
        UPDT_Rect_2.ToolTip = "PR Loss " & FormatPercent(UPDTbar2_PRloss, 1)
        UPDT_Rect_3.ToolTip = "PR Loss " & FormatPercent(UPDTbar3_PRloss, 1)
        UPDT_Rect_4.ToolTip = "PR Loss " & FormatPercent(UPDTbar4_PRloss, 1)
        UPDT_Rect_5.ToolTip = "PR Loss " & FormatPercent(UPDTbar5_PRloss, 1)
        UPDT_Rect_6.ToolTip = "PR Loss " & FormatPercent(UPDTbar6_PRloss, 1)

        'SIM
        UPDT_Rect_1sim.ToolTip = "Simulated PR Loss " & FormatPercent(UPDTbar1_PRlosssim, 1)
        UPDT_Rect_2sim.ToolTip = "Simulated PR Loss " & FormatPercent(UPDTbar2_PRlosssim, 1)
        UPDT_Rect_3sim.ToolTip = "Simulated PR Loss " & FormatPercent(UPDTbar3_PRlosssim, 1)
        UPDT_Rect_4sim.ToolTip = "Simulated PR Loss " & FormatPercent(UPDTbar4_PRlosssim, 1)
        UPDT_Rect_5sim.ToolTip = "Simulated PR Loss " & FormatPercent(UPDTbar5_PRlosssim, 1)
        UPDT_Rect_6sim.ToolTip = "Simulated PR Loss " & FormatPercent(UPDTbar6_PRlosssim, 1)


        UPDT_Target_Rect1.Visibility = Visibility.Hidden 'Visible
        UPDT_Target_Rect2.Visibility = Visibility.Hidden 'Visible
        UPDT_Target_Rect3.Visibility = Visibility.Hidden 'Visible
        UPDT_Target_Rect4.Visibility = Visibility.Hidden 'Visible
        UPDT_Target_Rect5.Visibility = Visibility.Hidden 'Visible
        UPDT_Target_Rect6.Visibility = Visibility.Hidden 'Visible

        If IsSimulationMode Then ShowSimRectangles_UPDT()

        If AllProdLines(selectedindexofLine_temp).doIhaveTargets = True Then
            LoadTargetsUI(1, Card_Unplanned_T1)
            ShowTargets_Unplanned()
        End If



        If ScrollBase_Card1 <> 0 Then
            NavigationLeft_card1.Visibility = Visibility.Visible
        Else
            NavigationLeft_card1.Visibility = Visibility.Hidden
        End If

        NavigationRight_Card1.Visibility = Visibility.Visible

        TrendStarIcon1.Visibility = Visibility.Hidden

    End Sub
    Sub hidePDTlabels()

        PDTLabel1.Visibility = Visibility.Hidden
        PDTlabel2.Visibility = Visibility.Hidden
        PDTlabel3.Visibility = Visibility.Hidden
        PDTlabel4.Visibility = Visibility.Hidden
        PDTlabel5.Visibility = Visibility.Hidden
        PDTlabel6.Visibility = Visibility.Hidden
        PDTLabel7.Visibility = Visibility.Hidden
        PDTLabel8.Visibility = Visibility.Hidden
        PDTLabel9.Visibility = Visibility.Hidden

        PDT_Rect_1.Visibility = Visibility.Hidden
        PDT_Rect_2.Visibility = Visibility.Hidden
        PDT_Rect_3.Visibility = Visibility.Hidden
        PDT_Rect_4.Visibility = Visibility.Hidden
        PDT_Rect_5.Visibility = Visibility.Hidden
        PDT_Rect_6.Visibility = Visibility.Hidden
        PDT_Rect_7.Visibility = Visibility.Hidden
        PDT_Rect_8.Visibility = Visibility.Hidden
        PDT_Rect_9.Visibility = Visibility.Hidden
        PDT_datalabel1.Visibility = Visibility.Hidden
        PDT_datalabel2.Visibility = Visibility.Hidden
        PDT_datalabel3.Visibility = Visibility.Hidden
        PDT_datalabel4.Visibility = Visibility.Hidden
        PDT_datalabel5.Visibility = Visibility.Hidden
        PDT_datalabel6.Visibility = Visibility.Hidden
        PDT_datalabel7.Visibility = Visibility.Hidden
        PDT_datalabel8.Visibility = Visibility.Hidden
        PDT_datalabel9.Visibility = Visibility.Hidden

        Card2Header.Visibility = Visibility.Hidden
        TrendStarIcon2.Visibility = Visibility.Hidden


        PDT_Target_Rect1.Visibility = Visibility.Hidden
        PDT_Target_Rect2.Visibility = Visibility.Hidden
        PDT_Target_Rect3.Visibility = Visibility.Hidden
        PDT_Target_Rect4.Visibility = Visibility.Hidden
        PDT_Target_Rect5.Visibility = Visibility.Hidden
        PDT_Target_Rect6.Visibility = Visibility.Hidden
        PDT_Target_Rect7.Visibility = Visibility.Hidden
        PDT_Target_Rect8.Visibility = Visibility.Hidden
        PDT_Target_Rect9.Visibility = Visibility.Hidden

        HideSimRectangles_PDT()
    End Sub
    Sub showPDTlabels()

        Dim datalabelposition1 As Thickness
        Dim datalabelposition2 As Thickness
        Dim datalabelposition3 As Thickness
        Dim datalabelposition4 As Thickness
        Dim datalabelposition5 As Thickness
        Dim datalabelposition6 As Thickness
        Dim datalabelposition7 As Thickness
        Dim datalabelposition8 As Thickness
        Dim datalabelposition9 As Thickness

        Dim barposition As Thickness
        PDTLabel1.Visibility = Visibility.Visible
        PDTlabel2.Visibility = Visibility.Visible
        PDTlabel3.Visibility = Visibility.Visible
        PDTlabel4.Visibility = Visibility.Visible
        PDTlabel5.Visibility = Visibility.Visible
        PDTlabel6.Visibility = Visibility.Visible
        PDTLabel7.Visibility = Visibility.Visible
        PDTLabel8.Visibility = Visibility.Visible
        PDTLabel9.Visibility = Visibility.Visible

        Card2Header.Visibility = Visibility.Visible
        plannedDTchart.Visibility = Visibility.Visible
        Card2Header.Content = cardnameLabeltext(2)


        PDTLabel1.Content = pdtlabel1string
        PDTlabel2.Content = pdtlabel2string
        PDTlabel3.Content = pdtlabel3string
        PDTlabel4.Content = pdtlabel4string
        PDTlabel5.Content = pdtlabel5string
        PDTlabel6.Content = pdtlabel6string
        PDTLabel7.Content = pdtlabel7string
        PDTLabel8.Content = pdtlabel8string
        PDTLabel9.Content = pdtlabel9string

        PDTLabel1.ToolTip = pdtlabel1string
        PDTlabel2.ToolTip = pdtlabel2string
        PDTlabel3.ToolTip = pdtlabel3string
        PDTlabel4.ToolTip = pdtlabel4string
        PDTlabel5.ToolTip = pdtlabel5string
        PDTlabel6.ToolTip = pdtlabel6string
        PDTLabel7.ToolTip = pdtlabel7string
        PDTLabel8.ToolTip = pdtlabel8string
        PDTLabel9.ToolTip = pdtlabel9string

        PDT_Rect_1.Visibility = Visibility.Visible
        PDT_Rect_2.Visibility = Visibility.Visible
        PDT_Rect_3.Visibility = Visibility.Visible
        PDT_Rect_4.Visibility = Visibility.Visible
        PDT_Rect_5.Visibility = Visibility.Visible
        PDT_Rect_6.Visibility = Visibility.Visible
        PDT_Rect_7.Visibility = Visibility.Visible
        PDT_Rect_8.Visibility = Visibility.Visible
        PDT_Rect_9.Visibility = Visibility.Visible
        PDT_datalabel1.Visibility = Visibility.Visible
        PDT_datalabel2.Visibility = Visibility.Visible
        PDT_datalabel3.Visibility = Visibility.Visible
        PDT_datalabel4.Visibility = Visibility.Visible
        PDT_datalabel5.Visibility = Visibility.Visible
        PDT_datalabel6.Visibility = Visibility.Visible
        PDT_datalabel7.Visibility = Visibility.Visible
        PDT_datalabel8.Visibility = Visibility.Visible
        PDT_datalabel9.Visibility = Visibility.Visible



        PDT_Rect_1.Height = PDTbar1
        PDT_Rect_2.Height = PDTbar2
        PDT_Rect_3.Height = PDTbar3
        PDT_Rect_4.Height = PDTbar4
        PDT_Rect_5.Height = PDTbar5
        PDT_Rect_6.Height = PDTbar6
        PDT_Rect_7.Height = PDTbar7
        PDT_Rect_8.Height = PDTbar8
        PDT_Rect_9.Height = PDTbar9


        'SIM
        Try 'TEMPORARY TO MAKE THE UPDT PART RUN -sam 7/20
            PDT_Rect_1sim.Height = PDTbar1sim
            PDT_Rect_2sim.Height = PDTbar2sim
            PDT_Rect_3sim.Height = PDTbar3sim
            PDT_Rect_4sim.Height = PDTbar4sim
            PDT_Rect_5sim.Height = PDTbar5sim
            PDT_Rect_6sim.Height = PDTbar6sim
            PDT_Rect_7sim.Height = PDTbar7sim
            PDT_Rect_8sim.Height = PDTbar8sim
            PDT_Rect_9sim.Height = PDTbar9sim
        Catch ex As Exception
            PDT_Rect_1sim.Height = 0 ' PDTbar1sim
            PDT_Rect_2sim.Height = 0 ' PDTbar2sim
            PDT_Rect_3sim.Height = 0 ' PDTbar3sim
            PDT_Rect_4sim.Height = 0 ' PDTbar4sim
            PDT_Rect_5sim.Height = 0 ' PDTbar5sim
            PDT_Rect_6sim.Height = 0 ' PDTbar6sim
            PDT_Rect_7sim.Height = 0 ' PDTbar7sim
            PDT_Rect_8sim.Height = 0 ' PDTbar8sim
            PDT_Rect_9sim.Height = 0 ' PDTbar9sim

            PDTbar1sim = 0
            PDTbar2sim = 0
            PDTbar3sim = 0
            PDTbar4sim = 0
            PDTbar5sim = 0
            PDTbar6sim = 0
            PDTbar7sim = 0
            PDTbar8sim = 0
            PDTbar9sim = 0
        End Try


        datalabelposition1 = PDT_datalabel1.Margin
        datalabelposition2 = PDT_datalabel2.Margin
        datalabelposition3 = PDT_datalabel3.Margin
        datalabelposition4 = PDT_datalabel4.Margin
        datalabelposition5 = PDT_datalabel5.Margin
        datalabelposition6 = PDT_datalabel6.Margin
        datalabelposition7 = PDT_datalabel7.Margin
        datalabelposition8 = PDT_datalabel8.Margin
        datalabelposition9 = PDT_datalabel9.Margin

        barposition = PDT_Rect_1.Margin

        PDT_datalabel1.Margin = New Thickness(datalabelposition1.Left, 210.0 - (1.0 * PDTbar1), 0, 0)
        PDT_datalabel2.Margin = New Thickness(datalabelposition2.Left, 210.0 - (1.0 * PDTbar2), 0, 0)
        PDT_datalabel3.Margin = New Thickness(datalabelposition3.Left, 210.0 - (1.0 * PDTbar3), 0, 0)
        PDT_datalabel4.Margin = New Thickness(datalabelposition4.Left, 210.0 - (1.0 * PDTbar4), 0, 0)
        PDT_datalabel5.Margin = New Thickness(datalabelposition5.Left, 210.0 - (1.0 * PDTbar5), 0, 0)
        PDT_datalabel6.Margin = New Thickness(datalabelposition6.Left, 210.0 - (1.0 * PDTbar6), 0, 0)
        PDT_datalabel7.Margin = New Thickness(datalabelposition7.Left, 210.0 - (1.0 * PDTbar7), 0, 0)
        PDT_datalabel8.Margin = New Thickness(datalabelposition8.Left, 210.0 - (1.0 * PDTbar8), 0, 0)
        PDT_datalabel9.Margin = New Thickness(datalabelposition9.Left, 210.0 - (1.0 * PDTbar9), 0, 0)

        'SIM
        datalabelposition1 = PDT_datalabel1sim.Margin
        datalabelposition2 = PDT_datalabel2sim.Margin
        datalabelposition3 = PDT_datalabel3sim.Margin
        datalabelposition4 = PDT_datalabel4sim.Margin
        datalabelposition5 = PDT_datalabel5sim.Margin
        datalabelposition6 = PDT_datalabel6sim.Margin
        datalabelposition7 = PDT_datalabel7sim.Margin
        datalabelposition8 = PDT_datalabel8sim.Margin
        datalabelposition9 = PDT_datalabel9sim.Margin

        barposition = PDT_Rect_1sim.Margin

        PDT_datalabel1sim.Margin = New Thickness(datalabelposition1.Left, 210.0 - (1.0 * PDTbar1sim), 0, 0)
        PDT_datalabel2sim.Margin = New Thickness(datalabelposition2.Left, 210.0 - (1.0 * PDTbar2sim), 0, 0)
        PDT_datalabel3sim.Margin = New Thickness(datalabelposition3.Left, 210.0 - (1.0 * PDTbar3sim), 0, 0)
        PDT_datalabel4sim.Margin = New Thickness(datalabelposition4.Left, 210.0 - (1.0 * PDTbar4sim), 0, 0)
        PDT_datalabel5sim.Margin = New Thickness(datalabelposition5.Left, 210.0 - (1.0 * PDTbar5sim), 0, 0)
        PDT_datalabel6sim.Margin = New Thickness(datalabelposition6.Left, 210.0 - (1.0 * PDTbar6sim), 0, 0)
        PDT_datalabel7sim.Margin = New Thickness(datalabelposition7.Left, 210.0 - (1.0 * PDTbar7sim), 0, 0)
        PDT_datalabel8sim.Margin = New Thickness(datalabelposition8.Left, 210.0 - (1.0 * PDTbar8sim), 0, 0)
        PDT_datalabel9sim.Margin = New Thickness(datalabelposition9.Left, 210.0 - (1.0 * PDTbar9sim), 0, 0)
        '

        PDT_datalabel1.Content = FormatPercent(PDTbar1_PRloss, 1)
        PDT_datalabel2.Content = FormatPercent(PDTbar2_PRloss, 1)
        PDT_datalabel3.Content = FormatPercent(PDTbar3_PRloss, 1)
        PDT_datalabel4.Content = FormatPercent(PDTbar4_PRloss, 1)
        PDT_datalabel5.Content = FormatPercent(PDTbar5_PRloss, 1)
        PDT_datalabel6.Content = FormatPercent(PDTbar6_PRloss, 1)
        PDT_datalabel7.Content = FormatPercent(PDTbar7_PRloss, 1)
        PDT_datalabel8.Content = FormatPercent(PDTbar8_PRloss, 1)
        PDT_datalabel9.Content = FormatPercent(PDTbar9_PRloss, 1)

        'SIM
        PDT_datalabel1sim.Content = FormatPercent(PDTbar1_PRlosssim, 1)
        PDT_datalabel2sim.Content = FormatPercent(PDTbar2_PRlosssim, 1)
        PDT_datalabel3sim.Content = FormatPercent(PDTbar3_PRlosssim, 1)
        PDT_datalabel4sim.Content = FormatPercent(PDTbar4_PRlosssim, 1)
        PDT_datalabel5sim.Content = FormatPercent(PDTbar5_PRlosssim, 1)
        PDT_datalabel6sim.Content = FormatPercent(PDTbar6_PRlosssim, 1)
        PDT_datalabel7sim.Content = FormatPercent(PDTbar7_PRlosssim, 1)
        PDT_datalabel8sim.Content = FormatPercent(PDTbar8_PRlosssim, 1)
        PDT_datalabel9sim.Content = FormatPercent(PDTbar9_PRlosssim, 1)
        '



        PDT_Rect_1.ToolTip = "PR Loss " & FormatPercent(PDTbar1_PRloss, 1)
        PDT_Rect_2.ToolTip = "PR Loss " & FormatPercent(PDTbar2_PRloss, 1)
        PDT_Rect_3.ToolTip = "PR Loss " & FormatPercent(PDTbar3_PRloss, 1)
        PDT_Rect_4.ToolTip = "PR Loss " & FormatPercent(PDTbar4_PRloss, 1)
        PDT_Rect_5.ToolTip = "PR Loss " & FormatPercent(PDTbar5_PRloss, 1)
        PDT_Rect_6.ToolTip = "PR Loss " & FormatPercent(PDTbar6_PRloss, 1)
        PDT_Rect_7.ToolTip = "PR Loss " & FormatPercent(PDTbar7_PRloss, 1)
        PDT_Rect_8.ToolTip = "PR Loss " & FormatPercent(PDTbar8_PRloss, 1)
        PDT_Rect_9.ToolTip = "PR Loss " & FormatPercent(PDTbar9_PRloss, 1)
        'SIM
        PDT_Rect_1sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar1_PRlosssim, 1)
        PDT_Rect_2sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar2_PRlosssim, 1)
        PDT_Rect_3sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar3_PRlosssim, 1)
        PDT_Rect_4sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar4_PRlosssim, 1)
        PDT_Rect_5sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar5_PRlosssim, 1)
        PDT_Rect_6sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar6_PRlosssim, 1)
        PDT_Rect_7sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar7_PRlosssim, 1)
        PDT_Rect_8sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar8_PRlosssim, 1)
        PDT_Rect_9sim.ToolTip = "Simulated PR Loss " & FormatPercent(PDTbar9_PRlosssim, 1)


        TrendStarIcon2.Visibility = Visibility.Hidden

        PDT_Target_Rect1.Visibility = Visibility.Hidden 'Visible
        PDT_Target_Rect2.Visibility = Visibility.Hidden 'Visible
        PDT_Target_Rect3.Visibility = Visibility.Hidden ' Visible
        PDT_Target_Rect4.Visibility = Visibility.Hidden ' Visible
        PDT_Target_Rect5.Visibility = Visibility.Hidden ' Visible
        PDT_Target_Rect6.Visibility = Visibility.Hidden ' Visible
        PDT_Target_Rect7.Visibility = Visibility.Hidden ' Visible
        PDT_Target_Rect8.Visibility = Visibility.Hidden ' Visible
        PDT_Target_Rect9.Visibility = Visibility.Hidden ' Visible

        If IsSimulationMode Then ShowSimRectangles_PDT()

        If AllProdLines(selectedindexofLine_temp).doIhaveTargets = True Then
            LoadTargetsUI(2, Card_Planned_T1)
            ShowTargets_planned()
        End If
    End Sub


    Sub hidechangeoverlabels()

        changeoverchart.Visibility = Visibility.Hidden
        ChangeoverLabel1.Visibility = Visibility.Hidden
        ChangeoverLabel2.Visibility = Visibility.Hidden
        ChangeoverLabel3.Visibility = Visibility.Hidden
        ChangeoverLabel4.Visibility = Visibility.Hidden
        ChangeoverLabel5.Visibility = Visibility.Hidden
        ChangeoverLabel6.Visibility = Visibility.Hidden
        ChangeoverLabel7.Visibility = Visibility.Hidden
        changeover_Rect1.Visibility = Visibility.Hidden
        changeover_Rect2.Visibility = Visibility.Hidden
        changeover_Rect3.Visibility = Visibility.Hidden
        changeover_Rect4.Visibility = Visibility.Hidden
        changeover_Rect5.Visibility = Visibility.Hidden
        changeover_Rect6.Visibility = Visibility.Hidden
        changeover_Rect7.Visibility = Visibility.Hidden
        changeover_datalabel1.Visibility = Visibility.Hidden
        changeover_datalabel2.Visibility = Visibility.Hidden
        changeover_datalabel3.Visibility = Visibility.Hidden
        changeover_datalabel4.Visibility = Visibility.Hidden
        changeover_datalabel5.Visibility = Visibility.Hidden
        changeover_datalabel6.Visibility = Visibility.Hidden
        changeover_datalabel7.Visibility = Visibility.Hidden


        Card41Header.Visibility = Visibility.Hidden
        TrendStarIcon41.Visibility = Visibility.Hidden
        CO_Target_Rect1.Visibility = Visibility.Hidden
        CO_Target_Rect2.Visibility = Visibility.Hidden
        CO_Target_Rect3.Visibility = Visibility.Hidden
        CO_Target_Rect4.Visibility = Visibility.Hidden
        CO_Target_Rect5.Visibility = Visibility.Hidden
        CO_Target_Rect6.Visibility = Visibility.Hidden
        CO_Target_Rect7.Visibility = Visibility.Hidden

        HideSimRectangles_Changeover()
    End Sub

    Sub showchangeoverlabels()

        Dim datalabelposition1 As Thickness
        Dim datalabelposition2 As Thickness
        Dim datalabelposition3 As Thickness
        Dim datalabelposition4 As Thickness
        Dim datalabelposition5 As Thickness
        Dim datalabelposition6 As Thickness
        Dim datalabelposition7 As Thickness

        Dim barposition As Thickness
        changeoverchart.Visibility = Visibility.Visible
        ChangeoverLabel1.Visibility = Visibility.Visible
        ChangeoverLabel2.Visibility = Visibility.Visible
        ChangeoverLabel3.Visibility = Visibility.Visible
        ChangeoverLabel4.Visibility = Visibility.Visible
        ChangeoverLabel5.Visibility = Visibility.Visible
        ChangeoverLabel6.Visibility = Visibility.Visible
        ChangeoverLabel7.Visibility = Visibility.Visible
        changeover_Rect1.Visibility = Visibility.Visible
        changeover_Rect2.Visibility = Visibility.Visible
        changeover_Rect3.Visibility = Visibility.Visible
        changeover_Rect4.Visibility = Visibility.Visible
        changeover_Rect5.Visibility = Visibility.Visible
        changeover_Rect6.Visibility = Visibility.Visible
        changeover_Rect7.Visibility = Visibility.Visible
        changeover_datalabel1.Visibility = Visibility.Visible
        changeover_datalabel2.Visibility = Visibility.Visible
        changeover_datalabel3.Visibility = Visibility.Visible
        changeover_datalabel4.Visibility = Visibility.Visible
        changeover_datalabel5.Visibility = Visibility.Visible
        changeover_datalabel6.Visibility = Visibility.Visible
        changeover_datalabel7.Visibility = Visibility.Visible


        Card41Header.Visibility = Visibility.Visible

        '        Card41Header.Content = "Changeovers" 'cardnameLabeltext(41)


        ChangeoverLabel1.Content = changeoverlabel1string
        ChangeoverLabel2.Content = changeoverlabel2string
        ChangeoverLabel3.Content = changeoverlabel3string
        ChangeoverLabel4.Content = changeoverlabel4string
        ChangeoverLabel5.Content = changeoverlabel5string
        ChangeoverLabel6.Content = changeoverlabel6string
        ChangeoverLabel7.Content = changeoverlabel7string

        ChangeoverLabel1.ToolTip = changeoverlabel1string
        ChangeoverLabel2.ToolTip = changeoverlabel2string
        ChangeoverLabel3.ToolTip = changeoverlabel3string
        ChangeoverLabel4.ToolTip = changeoverlabel4string
        ChangeoverLabel5.ToolTip = changeoverlabel5string
        ChangeoverLabel6.ToolTip = changeoverlabel6string
        ChangeoverLabel7.ToolTip = changeoverlabel7string



        changeover_Rect1.Height = Changeoverbar1
        changeover_Rect2.Height = Changeoverbar2
        changeover_Rect3.Height = Changeoverbar3
        changeover_Rect4.Height = Changeoverbar4
        changeover_Rect5.Height = Changeoverbar5
        changeover_Rect6.Height = Changeoverbar6
        changeover_Rect7.Height = Changeoverbar7


        'SIM
        changeover_Rect1sim.Height = Changeoverbar1sim
        changeover_Rect2sim.Height = Changeoverbar2sim
        changeover_Rect3sim.Height = Changeoverbar3sim
        changeover_Rect4sim.Height = Changeoverbar4sim
        changeover_Rect5sim.Height = Changeoverbar5sim
        changeover_Rect6sim.Height = Changeoverbar6sim
        changeover_Rect7sim.Height = Changeoverbar7sim

        datalabelposition1 = changeover_datalabel1.Margin
        datalabelposition2 = changeover_datalabel2.Margin
        datalabelposition3 = changeover_datalabel3.Margin
        datalabelposition4 = changeover_datalabel4.Margin
        datalabelposition5 = changeover_datalabel5.Margin
        datalabelposition6 = changeover_datalabel6.Margin
        datalabelposition7 = changeover_datalabel7.Margin

        barposition = changeover_Rect1.Margin


        changeover_datalabel1.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Changeoverbar1), 0, 0)
        changeover_datalabel2.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Changeoverbar2), 0, 0)
        changeover_datalabel3.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Changeoverbar3), 0, 0)
        changeover_datalabel4.Margin = New Thickness(datalabelposition4.Left, 470.0 - (1.0 * Changeoverbar4), 0, 0)
        changeover_datalabel5.Margin = New Thickness(datalabelposition5.Left, 470.0 - (1.0 * Changeoverbar5), 0, 0)
        changeover_datalabel6.Margin = New Thickness(datalabelposition6.Left, 470.0 - (1.0 * Changeoverbar6), 0, 0)
        changeover_datalabel7.Margin = New Thickness(datalabelposition7.Left, 470.0 - (1.0 * Changeoverbar7), 0, 0)



        'SIM
        datalabelposition1 = changeover_datalabel1sim.Margin
        datalabelposition2 = changeover_datalabel2sim.Margin
        datalabelposition3 = changeover_datalabel3sim.Margin
        datalabelposition4 = changeover_datalabel4sim.Margin
        datalabelposition5 = changeover_datalabel5sim.Margin
        datalabelposition6 = changeover_datalabel6sim.Margin
        datalabelposition7 = changeover_datalabel7sim.Margin

        barposition = changeover_Rect1sim.Margin


        changeover_datalabel1sim.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Changeoverbar1sim), 0, 0)
        changeover_datalabel2sim.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Changeoverbar2sim), 0, 0)
        changeover_datalabel3sim.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Changeoverbar3sim), 0, 0)
        changeover_datalabel4sim.Margin = New Thickness(datalabelposition4.Left, 470.0 - (1.0 * Changeoverbar4sim), 0, 0)
        changeover_datalabel5sim.Margin = New Thickness(datalabelposition5.Left, 470.0 - (1.0 * Changeoverbar5sim), 0, 0)
        changeover_datalabel6sim.Margin = New Thickness(datalabelposition6.Left, 470.0 - (1.0 * Changeoverbar6sim), 0, 0)
        changeover_datalabel7sim.Margin = New Thickness(datalabelposition7.Left, 470.0 - (1.0 * Changeoverbar7sim), 0, 0)


        '

        changeover_datalabel1.Content = FormatPercent(Changeover1_PRloss, 1)
        changeover_datalabel2.Content = FormatPercent(Changeover2_PRloss, 1)
        changeover_datalabel3.Content = FormatPercent(Changeover3_PRloss, 1)
        changeover_datalabel4.Content = FormatPercent(Changeover4_PRloss, 1)
        changeover_datalabel5.Content = FormatPercent(Changeover5_PRloss, 1)
        changeover_datalabel6.Content = FormatPercent(Changeover6_PRloss, 1)
        changeover_datalabel7.Content = FormatPercent(Changeover7_PRloss, 1)

        'SIM

        changeover_datalabel1sim.Content = FormatPercent(Changeover1_PRlosssim, 1)
        changeover_datalabel2sim.Content = FormatPercent(Changeover2_PRlosssim, 1)
        changeover_datalabel3sim.Content = FormatPercent(Changeover3_PRlosssim, 1)
        changeover_datalabel4sim.Content = FormatPercent(Changeover4_PRlosssim, 1)
        changeover_datalabel5sim.Content = FormatPercent(Changeover5_PRlosssim, 1)
        changeover_datalabel6sim.Content = FormatPercent(Changeover6_PRlosssim, 1)
        changeover_datalabel7sim.Content = FormatPercent(Changeover7_PRlosssim, 1)
        '


        changeover_Rect1.ToolTip = "PR Loss " & FormatPercent(Changeover1_PRloss, 1)
        changeover_Rect2.ToolTip = "PR Loss " & FormatPercent(Changeover2_PRloss, 1)
        changeover_Rect3.ToolTip = "PR Loss " & FormatPercent(Changeover3_PRloss, 1)
        changeover_Rect4.ToolTip = "PR Loss " & FormatPercent(Changeover4_PRloss, 1)
        changeover_Rect5.ToolTip = "PR Loss " & FormatPercent(Changeover5_PRloss, 1)
        changeover_Rect6.ToolTip = "PR Loss " & FormatPercent(Changeover6_PRloss, 1)
        changeover_Rect7.ToolTip = "PR Loss " & FormatPercent(Changeover7_PRloss, 1)

        'SIM
        changeover_Rect1sim.ToolTip = "Simulated PR Loss " & FormatPercent(Changeover1_PRlosssim, 1)
        changeover_Rect2sim.ToolTip = "Simulated PR Loss " & FormatPercent(Changeover2_PRlosssim, 1)
        changeover_Rect3sim.ToolTip = "Simulated PR Loss " & FormatPercent(Changeover3_PRlosssim, 1)
        changeover_Rect4sim.ToolTip = "Simulated PR Loss " & FormatPercent(Changeover4_PRlosssim, 1)
        changeover_Rect5sim.ToolTip = "Simulated PR Loss " & FormatPercent(Changeover5_PRlosssim, 1)
        changeover_Rect6sim.ToolTip = "Simulated PR Loss " & FormatPercent(Changeover6_PRlosssim, 1)
        changeover_Rect7sim.ToolTip = "Simulated PR Loss " & FormatPercent(Changeover7_PRlosssim, 1)

        TrendStarIcon41.Visibility = Visibility.Hidden

        CO_Target_Rect1.Visibility = Visibility.Hidden 'Visible
        CO_Target_Rect2.Visibility = Visibility.Hidden ' Visible
        CO_Target_Rect3.Visibility = Visibility.Hidden 'Visible
        CO_Target_Rect4.Visibility = Visibility.Hidden ' Visible
        CO_Target_Rect5.Visibility = Visibility.Hidden ' Visible
        CO_Target_Rect6.Visibility = Visibility.Hidden ' Visible
        CO_Target_Rect7.Visibility = Visibility.Hidden ' Visible

        If IsSimulationMode = True Then ShowSimRectangles_Changeover()

        If AllProdLines(selectedindexofLine_temp).doIhaveTargets = True Then
            LoadTargetsUI(41, Card_Planned_T2)
            ShowTargets_plannedT2()
        End If

    End Sub
    Private Sub hideEquipmentlabels()

        EquipmentLabel1.Visibility = Visibility.Hidden
        EquipmentLabel2.Visibility = Visibility.Hidden
        EquipmentLabel3.Visibility = Visibility.Hidden
        EquipmentLabel4.Visibility = Visibility.Hidden
        EquipmentLabel5.Visibility = Visibility.Hidden
        EquipmentLabel6.Visibility = Visibility.Hidden
        Equip1Label1.Visibility = Visibility.Hidden
        Equip1Label2.Visibility = Visibility.Hidden
        Equip1Label3.Visibility = Visibility.Hidden
        Equip2Label1.Visibility = Visibility.Hidden
        Equip2Label2.Visibility = Visibility.Hidden
        Equip2Label3.Visibility = Visibility.Hidden
        Equip3Label1.Visibility = Visibility.Hidden
        Equip3Label2.Visibility = Visibility.Hidden
        Equip3Label3.Visibility = Visibility.Hidden
        Equip21Label1.Visibility = Visibility.Hidden
        Equip21Label2.Visibility = Visibility.Hidden
        Equip22Label1.Visibility = Visibility.Hidden
        Equip22Label2.Visibility = Visibility.Hidden


        EquipMain_Rect1.Visibility = Visibility.Hidden
        EquipMain_Rect2.Visibility = Visibility.Hidden
        EquipMain_Rect3.Visibility = Visibility.Hidden
        EquipMain_Rect4.Visibility = Visibility.Hidden
        EquipMain_Rect5.Visibility = Visibility.Hidden
        EquipMain_Rect6.Visibility = Visibility.Hidden
        EquipMain_datalabel1.Visibility = Visibility.Hidden
        EquipMain_datalabel2.Visibility = Visibility.Hidden
        EquipMain_datalabel3.Visibility = Visibility.Hidden
        EquipMain_datalabel4.Visibility = Visibility.Hidden
        EquipMain_datalabel5.Visibility = Visibility.Hidden
        EquipMain_datalabel6.Visibility = Visibility.Hidden

        Equip1_Rect1.Visibility = Visibility.Hidden
        Equip1_Rect2.Visibility = Visibility.Hidden
        Equip1_Rect3.Visibility = Visibility.Hidden
        Equip2_Rect1.Visibility = Visibility.Hidden
        Equip2_Rect2.Visibility = Visibility.Hidden
        Equip2_Rect3.Visibility = Visibility.Hidden
        Equip3_Rect1.Visibility = Visibility.Hidden
        Equip3_Rect2.Visibility = Visibility.Hidden
        Equip3_Rect3.Visibility = Visibility.Hidden
        Equip1_datalabel1.Visibility = Visibility.Hidden
        Equip1_datalabel2.Visibility = Visibility.Hidden
        Equip1_datalabel3.Visibility = Visibility.Hidden
        Equip2_datalabel1.Visibility = Visibility.Hidden
        Equip2_datalabel2.Visibility = Visibility.Hidden
        Equip2_datalabel3.Visibility = Visibility.Hidden
        Equip3_datalabel1.Visibility = Visibility.Hidden
        Equip3_datalabel2.Visibility = Visibility.Hidden
        Equip3_datalabel3.Visibility = Visibility.Hidden


        Equip21_Rect_1.Visibility = Visibility.Hidden
        Equip21_Rect_2.Visibility = Visibility.Hidden
        Equip22_Rect_1.Visibility = Visibility.Hidden
        Equip22_Rect_2.Visibility = Visibility.Hidden

        Equip21_datalabel1.Visibility = Visibility.Hidden
        Equip21_datalabel2.Visibility = Visibility.Hidden
        Equip22_datalabel1.Visibility = Visibility.Hidden
        Equip22_datalabel2.Visibility = Visibility.Hidden

        Card3Header.Visibility = Visibility.Hidden
        Card4Header.Visibility = Visibility.Hidden
        Card5Header.Visibility = Visibility.Hidden
        Card6Header.Visibility = Visibility.Hidden
        Card21Header.Visibility = Visibility.Hidden
        Card22Header.Visibility = Visibility.Hidden


        TrendStarIcon3.Visibility = Visibility.Hidden
        TrendStarIcon4.Visibility = Visibility.Hidden
        TrendStarIcon5.Visibility = Visibility.Hidden
        TrendStarIcon6.Visibility = Visibility.Hidden
        TrendStarIcon21.Visibility = Visibility.Hidden
        TrendStarIcon22.Visibility = Visibility.Hidden

        EquipMain_Target_Rect1.Visibility = Visibility.Hidden
        EquipMain_Target_Rect2.Visibility = Visibility.Hidden
        EquipMain_Target_Rect3.Visibility = Visibility.Hidden
        EquipMain_Target_Rect4.Visibility = Visibility.Hidden
        EquipMain_Target_Rect5.Visibility = Visibility.Hidden
        EquipMain_Target_Rect6.Visibility = Visibility.Hidden

        Equip1_Target_Rect1.Visibility = Visibility.Hidden
        Equip1_Target_Rect2.Visibility = Visibility.Hidden
        Equip1_Target_Rect3.Visibility = Visibility.Hidden

        Equip2_Target_Rect1.Visibility = Visibility.Hidden
        Equip2_Target_Rect2.Visibility = Visibility.Hidden
        Equip2_Target_Rect3.Visibility = Visibility.Hidden

        Equip3_Target_Rect1.Visibility = Visibility.Hidden
        Equip3_Target_Rect2.Visibility = Visibility.Hidden
        Equip3_Target_Rect3.Visibility = Visibility.Hidden

        HideSimRectangles_Equip()
        NavigationLeft_card3.Visibility = Visibility.Hidden
        NavigationRight_Card3.Visibility = Visibility.Hidden
        NavigationLeft_card4.Visibility = Visibility.Hidden
        NavigationRight_Card4.Visibility = Visibility.Hidden
        NavigationLeft_card5.Visibility = Visibility.Hidden
        NavigationRight_Card5.Visibility = Visibility.Hidden
        NavigationLeft_card6.Visibility = Visibility.Hidden
        NavigationRight_Card6.Visibility = Visibility.Hidden

    End Sub

    Private Sub showEquipmentlabels(Optional IsSourceScrollButton As Boolean = False)
        Dim datalabelposition1 As Thickness
        Dim datalabelposition2 As Thickness
        Dim datalabelposition3 As Thickness
        Dim datalabelposition4 As Thickness
        Dim datalabelposition5 As Thickness
        Dim datalabelposition6 As Thickness


        Dim barposition As Thickness
        unplannedDTequipmentchart.Visibility = Visibility.Visible
        unplannedDTequip1chart.Visibility = Visibility.Visible
        unplannedDTequip2chart.Visibility = Visibility.Visible
        unplannedDTequip3chart.Visibility = Visibility.Visible

        EquipmentLabel1.Visibility = Visibility.Visible
        EquipmentLabel2.Visibility = Visibility.Visible
        EquipmentLabel3.Visibility = Visibility.Visible
        EquipmentLabel4.Visibility = Visibility.Visible
        EquipmentLabel5.Visibility = Visibility.Visible
        EquipmentLabel6.Visibility = Visibility.Visible
        Equip1Label1.Visibility = Visibility.Visible
        Equip1Label2.Visibility = Visibility.Visible
        Equip1Label3.Visibility = Visibility.Visible
        Equip2Label1.Visibility = Visibility.Visible
        Equip2Label2.Visibility = Visibility.Visible
        Equip2Label3.Visibility = Visibility.Visible
        Equip3Label1.Visibility = Visibility.Visible
        Equip3Label2.Visibility = Visibility.Visible
        Equip3Label3.Visibility = Visibility.Visible


        Card3Header.Visibility = Visibility.Visible
        Card4Header.Visibility = Visibility.Visible
        Card5Header.Visibility = Visibility.Visible
        Card6Header.Visibility = Visibility.Visible

        'Card3Header.Content = cardnameLabeltext(3)  ' 
        Card4Header.Content = cardnameLabeltext(4)
        Card5Header.Content = cardnameLabeltext(5)
        Card6Header.Content = cardnameLabeltext(6)

        EquipmentLabel1.Content = EquipmentLabel1string
        EquipmentLabel2.Content = EquipmentLabel2string
        EquipmentLabel3.Content = EquipmentLabel3string
        EquipmentLabel4.Content = EquipmentLabel4string
        EquipmentLabel5.Content = EquipmentLabel5string
        EquipmentLabel6.Content = EquipmentLabel6string

        If IsSourceScrollButton = False Then
            EquipmentLabel1.Background = LabelDefaultColor
            EquipmentLabel2.Background = LabelDefaultColor
            EquipmentLabel3.Background = LabelDefaultColor
            EquipmentLabel4.Background = LabelDefaultColor
            EquipmentLabel5.Background = LabelDefaultColor
            EquipmentLabel6.Background = LabelDefaultColor
        End If

        EquipmentLabel1.ToolTip = EquipmentLabel1string
        EquipmentLabel2.ToolTip = EquipmentLabel2string
        EquipmentLabel3.ToolTip = EquipmentLabel3string
        EquipmentLabel4.ToolTip = EquipmentLabel4string
        EquipmentLabel5.ToolTip = EquipmentLabel5string
        EquipmentLabel6.ToolTip = EquipmentLabel6string


        Equip1Label1.Content = Equip1Label1string
        Equip1Label2.Content = Equip1Label2string
        Equip1Label3.Content = Equip1Label3string
        Equip2Label1.Content = Equip2Label1string
        Equip2Label2.Content = Equip2Label2string
        Equip2Label3.Content = Equip2Label3string
        Equip3Label1.Content = Equip3Label1string
        Equip3Label2.Content = Equip3Label2string
        Equip3Label3.Content = Equip3Label3string

        Equip1Label1.ToolTip = Equip1Label1string
        Equip1Label2.ToolTip = Equip1Label2string
        Equip1Label3.ToolTip = Equip1Label3string
        Equip2Label1.ToolTip = Equip2Label1string
        Equip2Label2.ToolTip = Equip2Label2string
        Equip2Label3.ToolTip = Equip2Label3string
        Equip3Label1.ToolTip = Equip3Label1string
        Equip3Label2.ToolTip = Equip3Label2string
        Equip3Label3.ToolTip = Equip3Label3string

        '''''''''''''


        EquipMain_Rect1.Visibility = Visibility.Visible
        EquipMain_Rect2.Visibility = Visibility.Visible
        EquipMain_Rect3.Visibility = Visibility.Visible
        EquipMain_Rect4.Visibility = Visibility.Visible
        EquipMain_Rect5.Visibility = Visibility.Visible
        EquipMain_Rect6.Visibility = Visibility.Visible
        EquipMain_datalabel1.Visibility = Visibility.Visible
        EquipMain_datalabel2.Visibility = Visibility.Visible
        EquipMain_datalabel3.Visibility = Visibility.Visible
        EquipMain_datalabel4.Visibility = Visibility.Visible
        EquipMain_datalabel5.Visibility = Visibility.Visible
        EquipMain_datalabel6.Visibility = Visibility.Visible

        EquipMain_Rect1.Height = EquipMainbar1
        EquipMain_Rect2.Height = EquipMainbar2
        EquipMain_Rect3.Height = EquipMainbar3
        EquipMain_Rect4.Height = EquipMainbar4
        EquipMain_Rect5.Height = EquipMainbar5
        EquipMain_Rect6.Height = EquipMainbar6

        'SIM
        EquipMain_Rect1sim.Height = EquipMainbar1sim
        EquipMain_Rect2sim.Height = EquipMainbar2sim
        EquipMain_Rect3sim.Height = EquipMainbar3sim
        EquipMain_Rect4sim.Height = EquipMainbar4sim
        EquipMain_Rect5sim.Height = EquipMainbar5sim
        EquipMain_Rect6sim.Height = EquipMainbar6sim
        '

        datalabelposition1 = EquipMain_datalabel1.Margin
        datalabelposition2 = EquipMain_datalabel2.Margin
        datalabelposition3 = EquipMain_datalabel3.Margin
        datalabelposition4 = EquipMain_datalabel4.Margin
        datalabelposition5 = EquipMain_datalabel5.Margin
        datalabelposition6 = EquipMain_datalabel6.Margin

        barposition = EquipMain_Rect1.Margin


        EquipMain_datalabel1.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * EquipMainbar1), 0, 0)
        EquipMain_datalabel2.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * EquipMainbar2), 0, 0)
        EquipMain_datalabel3.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * EquipMainbar3), 0, 0)
        EquipMain_datalabel4.Margin = New Thickness(datalabelposition4.Left, 470.0 - (1.0 * EquipMainbar4), 0, 0)
        EquipMain_datalabel5.Margin = New Thickness(datalabelposition5.Left, 470.0 - (1.0 * EquipMainbar5), 0, 0)
        EquipMain_datalabel6.Margin = New Thickness(datalabelposition6.Left, 470.0 - (1.0 * EquipMainbar6), 0, 0)

        'SIM
        datalabelposition1 = EquipMain_datalabel1sim.Margin
        datalabelposition2 = EquipMain_datalabel2sim.Margin
        datalabelposition3 = EquipMain_datalabel3sim.Margin
        datalabelposition4 = EquipMain_datalabel4sim.Margin
        datalabelposition5 = EquipMain_datalabel5sim.Margin
        datalabelposition6 = EquipMain_datalabel6sim.Margin

        barposition = EquipMain_Rect1sim.Margin


        EquipMain_datalabel1sim.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * EquipMainbar1sim), 0, 0)
        EquipMain_datalabel2sim.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * EquipMainbar2sim), 0, 0)
        EquipMain_datalabel3sim.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * EquipMainbar3sim), 0, 0)
        EquipMain_datalabel4sim.Margin = New Thickness(datalabelposition4.Left, 470.0 - (1.0 * EquipMainbar4sim), 0, 0)
        EquipMain_datalabel5sim.Margin = New Thickness(datalabelposition5.Left, 470.0 - (1.0 * EquipMainbar5sim), 0, 0)
        EquipMain_datalabel6sim.Margin = New Thickness(datalabelposition6.Left, 470.0 - (1.0 * EquipMainbar6sim), 0, 0)
        '

        '

        EquipMain_datalabel1.Content = FormatPercent(Math.Round(EquipMain1_PRloss, 3), 1)
        EquipMain_datalabel2.Content = FormatPercent(Math.Round(EquipMain2_PRloss, 3), 1)
        EquipMain_datalabel3.Content = FormatPercent(Math.Round(EquipMain3_PRloss, 3), 1)
        EquipMain_datalabel4.Content = FormatPercent(Math.Round(EquipMain4_PRloss, 3), 1)
        EquipMain_datalabel5.Content = FormatPercent(Math.Round(EquipMain5_PRloss, 3), 1)
        EquipMain_datalabel6.Content = FormatPercent(Math.Round(EquipMain6_PRloss, 3), 1)


        'SIM

        EquipMain_datalabel1sim.Content = FormatPercent(Math.Round(EquipMain1_PRlosssim, 3), 1)
        EquipMain_datalabel2sim.Content = FormatPercent(Math.Round(EquipMain2_PRlosssim, 3), 1)
        EquipMain_datalabel3sim.Content = FormatPercent(Math.Round(EquipMain3_PRlosssim, 3), 1)
        EquipMain_datalabel4sim.Content = FormatPercent(Math.Round(EquipMain4_PRlosssim, 3), 1)
        EquipMain_datalabel5sim.Content = FormatPercent(Math.Round(EquipMain5_PRlosssim, 3), 1)
        EquipMain_datalabel6sim.Content = FormatPercent(Math.Round(EquipMain6_PRlosssim, 3), 1)

        '

        EquipMain_Rect1.ToolTip = "PR Loss " & FormatPercent(EquipMain1_PRloss, 1)
        EquipMain_Rect2.ToolTip = "PR Loss " & FormatPercent(EquipMain2_PRloss, 1)
        EquipMain_Rect3.ToolTip = "PR Loss " & FormatPercent(EquipMain3_PRloss, 1)
        EquipMain_Rect4.ToolTip = "PR Loss " & FormatPercent(EquipMain4_PRloss, 1)
        EquipMain_Rect5.ToolTip = "PR Loss " & FormatPercent(EquipMain5_PRloss, 1)
        EquipMain_Rect6.ToolTip = "PR Loss " & FormatPercent(EquipMain6_PRloss, 1)


        'SIM
        EquipMain_Rect1sim.ToolTip = "Simulated PR Loss " & FormatPercent(EquipMain1_PRlosssim, 1)
        EquipMain_Rect2sim.ToolTip = "Simulated PR Loss " & FormatPercent(EquipMain2_PRlosssim, 1)
        EquipMain_Rect3sim.ToolTip = "Simulated PR Loss " & FormatPercent(EquipMain3_PRlosssim, 1)
        EquipMain_Rect4sim.ToolTip = "Simulated PR Loss " & FormatPercent(EquipMain4_PRlosssim, 1)
        EquipMain_Rect5sim.ToolTip = "Simulated PR Loss " & FormatPercent(EquipMain5_PRlosssim, 1)
        EquipMain_Rect6sim.ToolTip = "Simulated PR Loss " & FormatPercent(EquipMain6_PRlosssim, 1)

        '''''''''''''''''''
        Equip1_Rect1.Visibility = Visibility.Visible
        Equip1_Rect2.Visibility = Visibility.Visible
        Equip1_Rect3.Visibility = Visibility.Visible
        Equip2_Rect1.Visibility = Visibility.Visible
        Equip2_Rect2.Visibility = Visibility.Visible
        Equip2_Rect3.Visibility = Visibility.Visible
        Equip3_Rect1.Visibility = Visibility.Visible
        Equip3_Rect2.Visibility = Visibility.Visible
        Equip3_Rect3.Visibility = Visibility.Visible
        Equip1_datalabel1.Visibility = Visibility.Visible
        Equip1_datalabel2.Visibility = Visibility.Visible
        Equip1_datalabel3.Visibility = Visibility.Visible
        Equip2_datalabel1.Visibility = Visibility.Visible
        Equip2_datalabel2.Visibility = Visibility.Visible
        Equip2_datalabel3.Visibility = Visibility.Visible
        Equip3_datalabel1.Visibility = Visibility.Visible
        Equip3_datalabel2.Visibility = Visibility.Visible
        Equip3_datalabel3.Visibility = Visibility.Visible

        Equip1_Rect1.Height = Equip1bar1
        Equip1_Rect2.Height = Equip1bar2
        Equip1_Rect3.Height = Equip1bar3

        Equip2_Rect1.Height = Equip2bar1
        Equip2_Rect2.Height = Equip2bar2
        Equip2_Rect3.Height = Equip2bar3

        Equip3_Rect1.Height = Equip3bar1
        Equip3_Rect2.Height = Equip3bar2
        Equip3_Rect3.Height = Equip3bar3

        'Sim
        Equip1_Rect1sim.Height = Equip1bar1sim
        Equip1_Rect2sim.Height = Equip1bar2sim
        Equip1_Rect3sim.Height = Equip1bar3sim

        Equip2_Rect1sim.Height = Equip2bar1sim
        Equip2_Rect2sim.Height = Equip2bar2sim
        Equip2_Rect3sim.Height = Equip2bar3sim

        Equip3_Rect1sim.Height = Equip3bar1sim
        Equip3_Rect2sim.Height = Equip3bar2sim
        Equip3_Rect3sim.Height = Equip3bar3sim
        ''''''''''

        barposition = Equip1_Rect1.Margin


        'equip1
        datalabelposition1 = Equip1_datalabel1.Margin
        datalabelposition2 = Equip1_datalabel2.Margin
        datalabelposition3 = Equip1_datalabel3.Margin

        Equip1_datalabel1.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Equip1bar1), 0, 0)
        Equip1_datalabel2.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Equip1bar2), 0, 0)
        Equip1_datalabel3.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Equip1bar3), 0, 0)


        'SIM equip1
        datalabelposition1 = Equip1_datalabel1sim.Margin
        datalabelposition2 = Equip1_datalabel2sim.Margin
        datalabelposition3 = Equip1_datalabel3sim.Margin

        Equip1_datalabel1sim.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Equip1bar1sim), 0, 0)
        Equip1_datalabel2sim.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Equip1bar2sim), 0, 0)
        Equip1_datalabel3sim.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Equip1bar3sim), 0, 0)
        ''''''''




        Equip1_datalabel1.Content = FormatPercent(Math.Round(Equip1_1_PRloss, 3), 1)
        Equip1_datalabel2.Content = FormatPercent(Math.Round(Equip1_2_PRloss, 3), 1)
        Equip1_datalabel3.Content = FormatPercent(Math.Round(Equip1_3_PRloss, 3), 1)

        'SIM equip1
        Equip1_datalabel1sim.Content = FormatPercent(Math.Round(Equip1_1_PRlosssim, 3), 1)
        Equip1_datalabel2sim.Content = FormatPercent(Math.Round(Equip1_2_PRlosssim, 3), 1)
        Equip1_datalabel3sim.Content = FormatPercent(Math.Round(Equip1_3_PRlosssim, 3), 1)

        ''''''




        Equip1_Rect1.ToolTip = "PR Loss " & FormatPercent(Equip1_1_PRloss, 1)
        Equip1_Rect2.ToolTip = "PR Loss " & FormatPercent(Equip1_2_PRloss, 1)
        Equip1_Rect3.ToolTip = "PR Loss " & FormatPercent(Equip1_3_PRloss, 1)

        'SIM
        Equip1_Rect1sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip1_1_PRlosssim, 1)
        Equip1_Rect2sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip1_2_PRlosssim, 1)
        Equip1_Rect3sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip1_3_PRlosssim, 1)
        '''''

        'equip2
        datalabelposition1 = Equip2_datalabel1.Margin
        datalabelposition2 = Equip2_datalabel2.Margin
        datalabelposition3 = Equip2_datalabel3.Margin

        Equip2_datalabel1.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Equip2bar1), 0, 0)
        Equip2_datalabel2.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Equip2bar2), 0, 0)
        Equip2_datalabel3.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Equip2bar3), 0, 0)


        'SIM Equip2
        datalabelposition1 = Equip2_datalabel1sim.Margin
        datalabelposition2 = Equip2_datalabel2sim.Margin
        datalabelposition3 = Equip2_datalabel3sim.Margin

        Equip2_datalabel1sim.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Equip2bar1sim), 0, 0)
        Equip2_datalabel2sim.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Equip2bar2sim), 0, 0)
        Equip2_datalabel3sim.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Equip2bar3sim), 0, 0)
        '


        Equip2_datalabel1.Content = FormatPercent(Math.Round(Equip2_1_PRloss, 3), 1)
        Equip2_datalabel2.Content = FormatPercent(Math.Round(Equip2_2_PRloss, 3), 1)
        Equip2_datalabel3.Content = FormatPercent(Math.Round(Equip2_3_PRloss, 3), 1)

        'SIM Equip2
        Equip2_datalabel1sim.Content = FormatPercent(Math.Round(Equip2_1_PRlosssim, 3), 1)
        Equip2_datalabel2sim.Content = FormatPercent(Math.Round(Equip2_2_PRlosssim, 3), 1)
        Equip2_datalabel3sim.Content = FormatPercent(Math.Round(Equip2_3_PRlosssim, 3), 1)

        ''


        Equip2_Rect1.ToolTip = "PR Loss " & FormatPercent(Equip2_1_PRloss, 1)
        Equip2_Rect2.ToolTip = "PR Loss " & FormatPercent(Equip2_2_PRloss, 1)
        Equip2_Rect3.ToolTip = "PR Loss " & FormatPercent(Equip2_3_PRloss, 1)

        'SIM
        Equip2_Rect1sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip2_1_PRlosssim, 1)
        Equip2_Rect2sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip2_2_PRlosssim, 1)
        Equip2_Rect3sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip2_3_PRlosssim, 1)

        'equip3
        datalabelposition1 = Equip3_datalabel1.Margin
        datalabelposition2 = Equip3_datalabel2.Margin
        datalabelposition3 = Equip3_datalabel3.Margin

        Equip3_datalabel1.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Equip3bar1), 0, 0)
        Equip3_datalabel2.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Equip3bar2), 0, 0)
        Equip3_datalabel3.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Equip3bar3), 0, 0)


        'SIM Equip3
        datalabelposition1 = Equip3_datalabel1sim.Margin
        datalabelposition2 = Equip3_datalabel2sim.Margin
        datalabelposition3 = Equip3_datalabel3sim.Margin

        Equip3_datalabel1sim.Margin = New Thickness(datalabelposition1.Left, 470.0 - (1.0 * Equip3bar1sim), 0, 0)
        Equip3_datalabel2sim.Margin = New Thickness(datalabelposition2.Left, 470.0 - (1.0 * Equip3bar2sim), 0, 0)
        Equip3_datalabel3sim.Margin = New Thickness(datalabelposition3.Left, 470.0 - (1.0 * Equip3bar3sim), 0, 0)


        '


        Equip3_datalabel1.Content = FormatPercent(Math.Round(Equip3_1_PRloss, 3), 1)
        Equip3_datalabel2.Content = FormatPercent(Math.Round(Equip3_2_PRloss, 3), 1)
        Equip3_datalabel3.Content = FormatPercent(Math.Round(Equip3_3_PRloss, 3), 1)

        'SIM Equip3
        Equip3_datalabel1sim.Content = FormatPercent(Math.Round(Equip3_1_PRlosssim, 3), 1)
        Equip3_datalabel2sim.Content = FormatPercent(Math.Round(Equip3_2_PRlosssim, 3), 1)
        Equip3_datalabel3sim.Content = FormatPercent(Math.Round(Equip3_3_PRlosssim, 3), 1)
        ''


        Equip3_Rect1.ToolTip = "PR Loss " & FormatPercent(Equip3_1_PRloss, 1)
        Equip3_Rect2.ToolTip = "PR Loss " & FormatPercent(Equip3_2_PRloss, 1)
        Equip3_Rect3.ToolTip = "PR Loss " & FormatPercent(Equip3_3_PRloss, 1)

        'SIM
        Equip3_Rect1sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip3_1_PRlosssim, 1)
        Equip3_Rect2sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip3_2_PRlosssim, 1)
        Equip3_Rect3sim.ToolTip = "Simulated PR Loss " & FormatPercent(Equip3_3_PRlosssim, 1)
        ''''''''


        TrendStarIcon3.Visibility = Visibility.Hidden
        TrendStarIcon4.Visibility = Visibility.Hidden
        TrendStarIcon5.Visibility = Visibility.Hidden
        TrendStarIcon6.Visibility = Visibility.Hidden
        TrendStarIcon21.Visibility = Visibility.Hidden
        TrendStarIcon22.Visibility = Visibility.Hidden

        EquipMain_Target_Rect1.Visibility = Visibility.Hidden 'Visible
        EquipMain_Target_Rect2.Visibility = Visibility.Hidden ' Visible
        EquipMain_Target_Rect3.Visibility = Visibility.Hidden ' Visible
        EquipMain_Target_Rect4.Visibility = Visibility.Hidden ' Visible
        EquipMain_Target_Rect5.Visibility = Visibility.Hidden ' Visible
        EquipMain_Target_Rect6.Visibility = Visibility.Hidden ' Visible

        Equip1_Target_Rect1.Visibility = Visibility.Hidden ' Visible
        Equip1_Target_Rect2.Visibility = Visibility.Hidden ' Visible
        Equip1_Target_Rect3.Visibility = Visibility.Hidden ' Visible

        Equip2_Target_Rect1.Visibility = Visibility.Hidden 'Visible
        Equip2_Target_Rect2.Visibility = Visibility.Hidden ' Visible
        Equip2_Target_Rect3.Visibility = Visibility.Hidden 'Visible

        Equip3_Target_Rect1.Visibility = Visibility.Hidden ' Visible
        Equip3_Target_Rect2.Visibility = Visibility.Hidden ' Visible
        Equip3_Target_Rect3.Visibility = Visibility.Hidden ' Visible

        If IsSimulationMode = True Then ShowSimRectangles_Equip()

        If AllProdLines(selectedindexofLine_temp).doIhaveTargets = True Then
            LoadTargetsUI(3, Card_Unplanned_T2)
            ShowTargets_UnplannedT2()
            LoadTargetsUI(4, Card_Unplanned_T3A)
            ShowTargets_UnplannedT3_1()
            LoadTargetsUI(5, Card_Unplanned_T3B)
            ShowTargets_UnplannedT3_2()
            LoadTargetsUI(6, Card_Unplanned_T3C)
            ShowTargets_UnplannedT3_3()
        End If


        If ScrollBase_Card3 <> 0 Then
            NavigationLeft_card3.Visibility = Visibility.Visible
        Else
            NavigationLeft_card3.Visibility = Visibility.Hidden
        End If

        If ScrollBase_Card4 <> 0 Then
            NavigationLeft_card4.Visibility = Visibility.Visible
        Else
            NavigationLeft_card4.Visibility = Visibility.Hidden
        End If

        If ScrollBase_Card5 <> 0 Then
            NavigationLeft_card5.Visibility = Visibility.Visible
        Else
            NavigationLeft_card5.Visibility = Visibility.Hidden
        End If

        If ScrollBase_Card6 <> 0 Then
            NavigationLeft_card6.Visibility = Visibility.Visible
        Else
            NavigationLeft_card6.Visibility = Visibility.Hidden
        End If




        If prStoryReport.getCardEventNumber(3) - ScrollBase_Card3 > 6 Then
            NavigationRight_Card3.Visibility = Visibility.Visible
        End If
        If prStoryReport.getCardEventNumber(4) - ScrollBase_Card4 > 3 Then
            NavigationRight_Card4.Visibility = Visibility.Visible
        End If
        If prStoryReport.getCardEventNumber(5) - ScrollBase_Card5 > 3 Then
            NavigationRight_Card5.Visibility = Visibility.Visible
        End If
        If prStoryReport.getCardEventNumber(6) - ScrollBase_Card6 > 3 Then
            NavigationRight_Card6.Visibility = Visibility.Visible
        End If

    End Sub
#End Region

#Region "inControl"
    Private Sub ToggleDayANDShift_Incontrol(sender As Object, e As MouseButtonEventArgs)
        Dim xhoursagodate As Object
        If sender Is incontrol_Day_selection Then
            incontrol_Day_selection.Background = CardHeaderdefaultColor
            incontrol_Day_selection.Foreground = mybrushlanguagewhite
            incontrol_Shift_selection.Background = mybrushlightgray

            IsIncontrol_Shiftmode = False
            GenerateIncontrolCharts(incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
        ElseIf sender Is incontrol_Shift_selection Then
            incontrol_Shift_selection.Background = CardHeaderdefaultColor
            incontrol_Shift_selection.Foreground = mybrushlanguagewhite
            incontrol_Day_selection.Background = mybrushlightgray
            xhoursagodate = DateAdd(DateInterval.Hour, -1 * (Math.Round(24 / AllProdLines(selectedindexofLine_temp).NumberOfShifts)), incontrolAnalysisSelectedEnddate)
            IsIncontrol_Shiftmode = True
            UseTrack_IncontrolControlShift = True
            GenerateIncontrolCharts(xhoursagodate, incontrolAnalysisSelectedEnddate) 'GenerateIncontrolCharts(xhoursagodate, incontrolAnalysisSelectedEnddate)

        End If


    End Sub

    Private Sub GenerateIncontrolCharts(startdate As Object, enddate As Object, Optional IsInitialize As Boolean = False)
        Dim tmpEvent As inControlDTevent
        'Dim MasterDataSet As inControlReport
        Dim bubbleposition As New Thickness
        Dim incontrolareaposition As New Thickness
        Dim bubblecount As Integer
        Dim PRLossmax As Double
        Dim stopsmax As Integer
        Dim CSscore As Integer
        Dim leftorigin As Double
        Dim toporigin As Double
        Dim leftposition As Double
        Dim topposition As Double
        Dim areawidth As Double
        Dim areaheight As Double
        Dim stopscount As Integer
        Dim prlossactual As Double
        Dim tooltipstring As String
        Dim failuremodename As String
        Dim MTBFmin As Double
        Dim WEscore As Integer
        Dim xhoursagodate As Object

        If IsIncontrol_Shiftmode = False Then
            If Not IsInitialize Then MasterDataSet = New inControlReport(AllProdLines(selectedindexofLine_temp), incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate) ', selectedRLcolumn)
            incontrol_date.Content = "Most recent run day:" & vbNewLine & Format(MasterDataSet.AnalysisStartDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(MasterDataSet.AnalysisEndDate, "MMMM dd, yyyy HH:mm").ToString

        Else
            xhoursagodate = DateAdd(DateInterval.Hour, -1 * (Math.Round(24 / AllProdLines(selectedindexofLine_temp).NumberOfShifts)), incontrolAnalysisSelectedEnddate)
            If Not IsInitialize Then MasterDataSet = New inControlReport(AllProdLines(selectedindexofLine_temp), xhoursagodate, incontrolAnalysisSelectedEnddate) ', selectedRLcolumn)
            'MasterDataSet = New inControlReport(AllProductionLines(selectedindexofLine_temp), xhoursagodate, incontrolAnalysisSelectedEnddate) ', selectedRLcolumn)
            MasterDataSet.analyzeNewTimePeriod(xhoursagodate, incontrolAnalysisSelectedEnddate)
            incontrol_date.Content = "Most recent run shift:" & vbNewLine & Format(MasterDataSet.AnalysisStartDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(MasterDataSet.AnalysisEndDate, "MMMM dd, yyyy HH:mm").ToString

        End If

        'incontrol_date.Content = Format(incontrolAnalysisSelectedStartdate, "MMMM dd, yyyy HH:mm") & vbNewLine & Format(incontrolAnalysisSelectedEnddate, "MMMM dd, yyyy HH:mm")
        While IsNothing(MasterDataSet)
            System.Threading.Thread.Sleep(500)
        End While


        PRLossmax = MasterDataSet.maxDTpct '0.15
        stopsmax = MasterDataSet.maxStops
        incontrolareaposition = IncontrolActiveArea.Margin
        leftorigin = incontrolareaposition.Left + 10
        toporigin = incontrolareaposition.Top + 40.0
        areaheight = 300
        areawidth = 810
        hideStopbubbles()
        HideIncontrolDatePicker()
        Incontrol_FailureMode.Content = "Click on a loss bubble"

        For bubblecount = 0 To inCONTROL_EventsToShow
            If Not bubblecount > MasterDataSet.inControlEvents.Count - 1 Then
                tmpEvent = MasterDataSet.inControlEvents(bubblecount)

                WEscore = tmpEvent.WesternRulesScore
                CSscore = tmpEvent.ChronicSporadicRanking
                stopscount = tmpEvent.Stops
                prlossactual = tmpEvent.DTpct
                ' MTBFmin = tmpEvent.DailyMTBF(bubblecount)
                failuremodename = tmpEvent.Name
                tooltipstring = "Failure Mode: " & failuremodename & vbNewLine & "Stops: " & stopscount & vbNewLine & "PR Loss: " & FormatPercent(prlossactual)

                stopbubblenames(bubblecount) = failuremodename
                stopbubblestops(bubblecount) = stopscount
                stopbubblePR(bubblecount) = prlossactual
                stopbubbleMTBF(bubblecount) = MTBFmin
                stopbubble90daystopsperday(bubblecount) = Math.Round(tmpEvent.AdjMu_InvMTDF, 1)

                stopbubbleAnalysisstopsperday(bubblecount) = Math.Round(tmpEvent.SPD, 1)

                Select Case bubblecount
                    Case 0

                        stopbubble1.Fill = decidecolor(WEscore)
                        stopbubble1.Visibility = Visibility.Visible
                        stopbubble1.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble1.Width = stopbubble1.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble1.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble1.Height / 2)
                        stopbubble1.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble1.ToolTip = tooltipstring

                    Case 1

                        stopbubble2.Fill = decidecolor(WEscore)
                        stopbubble2.Visibility = Visibility.Visible
                        stopbubble2.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble2.Width = stopbubble2.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble2.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble2.Height / 2)
                        stopbubble2.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble2.ToolTip = tooltipstring
                    Case 2

                        stopbubble3.Fill = decidecolor(WEscore)
                        stopbubble3.Visibility = Visibility.Visible
                        stopbubble3.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble3.Width = stopbubble3.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble3.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble3.Height / 2)
                        stopbubble3.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble3.ToolTip = tooltipstring
                    Case 3

                        stopbubble4.Fill = decidecolor(WEscore)
                        stopbubble4.Visibility = Visibility.Visible
                        stopbubble4.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble4.Width = stopbubble4.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble4.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble4.Height / 2)
                        stopbubble4.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble4.ToolTip = tooltipstring
                    Case 4

                        stopbubble5.Fill = decidecolor(WEscore)
                        stopbubble5.Visibility = Visibility.Visible
                        stopbubble5.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble5.Width = stopbubble5.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble5.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble5.Height / 2)
                        stopbubble5.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble5.ToolTip = tooltipstring
                    Case 5
                        '        prlossactual = bubblecount / 100   'update this line from Sam's code
                        ''         stopscount = bubblecount 'update this line from Sam's code
                        '         CSscore = 8  'update this line from Sam's code
                        '         WEscore = 1 'update this line from Sam's code

                        stopbubble6.Fill = decidecolor(WEscore)
                        stopbubble6.Visibility = Visibility.Visible
                        stopbubble6.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble6.Width = stopbubble6.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble6.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble6.Height / 2)
                        stopbubble6.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble6.ToolTip = tooltipstring
                    Case 6
                        '            prlossactual = bubblecount / 100   'update this line from Sam's code
                        '            stopscount = bubblecount 'update this line from Sam's code
                        '            CSscore = 1  'update this line from Sam's code
                        '            WEscore = 1 'update this line from Sam's code

                        stopbubble7.Fill = decidecolor(WEscore)
                        stopbubble7.Visibility = Visibility.Visible
                        stopbubble7.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble7.Width = stopbubble7.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble7.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble7.Height / 2)
                        stopbubble7.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble7.ToolTip = tooltipstring
                    Case 7
                        '          prlossactual = bubblecount / 100   'update this line from Sam's code
                        '          stopscount = bubblecount 'update this line from Sam's code
                        '          CSscore = 9  'update this line from Sam's code
                        '          WEscore = 1 'update this line from Sam's code

                        stopbubble8.Fill = decidecolor(WEscore)
                        stopbubble8.Visibility = Visibility.Visible
                        stopbubble8.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble8.Width = stopbubble8.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble8.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble8.Height / 2)
                        stopbubble8.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble8.ToolTip = tooltipstring
                    Case 8
                        '          prlossactual = bubblecount / 100   'update this line from Sam's code
                        '          stopscount = bubblecount 'update this line from Sam's code
                        '          CSscore = 3  'update this line from Sam's code
                        '          WEscore = 1 'update this line from Sam's code

                        stopbubble9.Fill = decidecolor(WEscore)
                        stopbubble9.Visibility = Visibility.Visible
                        stopbubble9.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble9.Width = stopbubble9.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble9.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble9.Height / 2)
                        stopbubble9.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble9.ToolTip = tooltipstring
                    Case 9
                        '          prlossactual = bubblecount / 100   'update this line from Sam's code
                        '          stopscount = bubblecount 'update this line from Sam's code
                        '          CSscore = 4  'update this line from Sam's code
                        '          WEscore = 1 'update this line from Sam's code

                        stopbubble10.Fill = decidecolor(WEscore)
                        stopbubble10.Visibility = Visibility.Visible
                        stopbubble10.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble10.Width = stopbubble10.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble10.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble10.Height / 2)
                        stopbubble10.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble10.ToolTip = tooltipstring
                    Case 10
                        '          prlossactual = bubblecount / 100   'update this line from Sam's code
                        '          stopscount = bubblecount 'update this line from Sam's code
                        '          CSscore = 2  'update this line from Sam's code
                        '          WEscore = 1 'update this line from Sam's code

                        stopbubble11.Fill = decidecolor(WEscore)
                        stopbubble11.Visibility = Visibility.Visible
                        stopbubble11.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble11.Width = stopbubble11.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble11.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble11.Height / 2)
                        stopbubble11.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble11.ToolTip = tooltipstring

                    Case 11
                        '          prlossactual = bubblecount / 100   'update this line from Sam's code
                        '          stopscount = bubblecount 'update this line from Sam's code
                        '          CSscore = 6  'update this line from Sam's code
                        '          WEscore = 1 'update this line from Sam's code

                        stopbubble12.Fill = decidecolor(WEscore)
                        stopbubble12.Visibility = Visibility.Visible
                        stopbubble12.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble12.Width = stopbubble12.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble12.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble12.Height / 2)
                        stopbubble12.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble12.ToolTip = tooltipstring
                    Case 12
                        '        prlossactual = bubblecount / 100   'update this line from Sam's code
                        '        stopscount = bubblecount 'update this line from Sam's code
                        '        CSscore = 7  'update this line from Sam's code
                        '        WEscore = 1 'update this line from Sam's code

                        stopbubble13.Fill = decidecolor(WEscore)
                        stopbubble13.Visibility = Visibility.Visible
                        stopbubble13.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble13.Width = stopbubble13.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble13.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble13.Height / 2)
                        stopbubble13.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble13.ToolTip = tooltipstring
                    Case 13
                        '            prlossactual = bubblecount / 100   'update this line from Sam's code
                        '            stopscount = bubblecount 'update this line from Sam's code
                        '            CSscore = 8  'update this line from Sam's code
                        '            WEscore = 1 'update this line from Sam's code

                        stopbubble14.Fill = decidecolor(WEscore)
                        stopbubble14.Visibility = Visibility.Visible
                        stopbubble14.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble14.Width = stopbubble14.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble14.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble14.Height / 2)
                        stopbubble14.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble14.ToolTip = tooltipstring
                    Case 14
                        '                    prlossactual = bubblecount / 100   'update this line from Sam's code
                        '                   stopscount = bubblecount 'update this line from Sam's code
                        '                  CSscore = 10  'update this line from Sam's code
                        '                 WEscore = 1 'update this line from Sam's code

                        stopbubble15.Fill = decidecolor(WEscore)
                        stopbubble15.Visibility = Visibility.Visible
                        stopbubble15.Height = Math.Sqrt((prlossactual / PRLossmax)) * (stopbubbleMAXsize)
                        stopbubble15.Width = stopbubble15.Height
                        leftposition = leftorigin + (CSscore * (areawidth / 10)) - (stopbubble15.Height / 2)
                        topposition = toporigin + areaheight - ((areaheight * (stopscount / stopsmax))) - (stopbubble15.Height / 2)
                        stopbubble15.Margin = New Thickness(leftposition, topposition, 0.0, 0.0)
                        stopbubble15.ToolTip = tooltipstring
                End Select
            End If
        Next

        If MasterDataSet.maxDTpct = 0 Then MsgBox("The line seems to have been scheduled for a run, but no unplanned stops were recorded for the day.")

    End Sub

    Private Sub RegenerateIncontrolCharts()
        If IsNothing(incontrolStartDatePicker.SelectedDate) Then incontrolStartDatePicker.SelectedDate = incontrolAnalysisSelectedStartdate
        HideIncontrolDatePicker()
        incontrolAnalysisSelectedStartdate = incontrolStartDatePicker.SelectedDate
        incontrolAnalysisSelectedEnddate = endtimeselected '.AddDays(-1) ' LG Code
        incontrolEndDatePicker.Content = Format(incontrolAnalysisSelectedEnddate, "MMMM dd, yyyy HH:mm")
        'incontrol_date.Content = Format(incontrolAnalysisSelectedStartdate, "MMMM dd, yyyy HH:mm") & vbNewLine & Format(incontrolAnalysisSelectedEnddate, "MMMM dd, yyyy HH:mm")
        Cursor = Cursors.Wait
        GenerateIncontrolCharts(incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
        Cursor = Cursors.Arrow

    End Sub

    Private Sub ShowIncontrolDatePicker()
        incontrolStartDateLabel.Visibility = Visibility.Visible
        incontrolEndDateLabel.Visibility = Visibility.Visible
        incontrolStartDatePicker.Visibility = Visibility.Visible
        incontrolEndDatePicker.Visibility = Visibility.Visible
        incontrol_Regenerate.Visibility = Visibility.Visible
        incontrolEndDatePicker.Content = Format(endtimeselected.AddDays(-1), "MMMM dd, yyyy HH:mm")
    End Sub
    Private Sub HideIncontrolDatePicker()
        incontrolStartDateLabel.Visibility = Visibility.Hidden
        incontrolEndDateLabel.Visibility = Visibility.Hidden
        incontrolStartDatePicker.Visibility = Visibility.Hidden
        incontrolEndDatePicker.Visibility = Visibility.Hidden
        incontrol_Regenerate.Visibility = Visibility.Hidden
    End Sub

#End Region

    Private Sub ToggleStops()
        If TopStopsLegend_bartext.Content = "Stops/Day" Then
            TopStopsLegend_bartext.Content = "Stops"

            TopStopsDataLabel1.Content = Math.Round(topstopsbar1_Stops, 1)
            TopStopsDataLabel2.Content = Math.Round(topstopsbar2_Stops, 1)
            TopStopsDataLabel3.Content = Math.Round(topstopsbar3_Stops, 1)
            TopStopsDataLabel4.Content = Math.Round(topstopsbar4_Stops, 1)
            TopStopsDataLabel5.Content = Math.Round(topstopsbar5_Stops, 1)
            TopStopsDataLabel6.Content = Math.Round(topstopsbar6_Stops, 1)
            TopStopsDataLabel7.Content = Math.Round(topstopsbar7_Stops, 1)
            TopStopsDataLabel8.Content = Math.Round(topstopsbar8_Stops, 1)
            TopStopsDataLabel9.Content = Math.Round(topstopsbar9_Stops, 1)
            TopStopsDataLabel10.Content = Math.Round(topstopsbar10_Stops, 1)
            TopStopsDataLabel11.Content = Math.Round(topstopsbar11_Stops, 1)
            TopStopsDataLabel12.Content = Math.Round(topstopsbar12_Stops, 1)
            TopStopsDataLabel13.Content = Math.Round(topstopsbar13_Stops, 1)
            TopStopsDataLabel14.Content = Math.Round(topstopsbar14_Stops, 1)
            TopStopsDataLabel15.Content = Math.Round(topstopsbar15_Stops, 1)


        ElseIf TopStopsLegend_bartext.Content = "Stops" Then
            TopStopsLegend_bartext.Content = "Stops/Day"

            TopStopsDataLabel1.Content = Math.Round(topstopsbar1_SPD, 1)
            TopStopsDataLabel2.Content = Math.Round(topstopsbar2_SPD, 1)
            TopStopsDataLabel3.Content = Math.Round(topstopsbar3_SPD, 1)
            TopStopsDataLabel4.Content = Math.Round(topstopsbar4_SPD, 1)
            TopStopsDataLabel5.Content = Math.Round(topstopsbar5_SPD, 1)
            TopStopsDataLabel6.Content = Math.Round(topstopsbar6_SPD, 1)
            TopStopsDataLabel7.Content = Math.Round(topstopsbar7_SPD, 1)
            TopStopsDataLabel8.Content = Math.Round(topstopsbar8_SPD, 1)
            TopStopsDataLabel9.Content = Math.Round(topstopsbar9_SPD, 1)
            TopStopsDataLabel10.Content = Math.Round(topstopsbar10_SPD, 1)
            TopStopsDataLabel11.Content = Math.Round(topstopsbar11_SPD, 1)
            TopStopsDataLabel12.Content = Math.Round(topstopsbar12_SPD, 1)
            TopStopsDataLabel13.Content = Math.Round(topstopsbar13_SPD, 1)
            TopStopsDataLabel14.Content = Math.Round(topstopsbar14_SPD, 1)
            TopStopsDataLabel15.Content = Math.Round(topstopsbar15_SPD, 1)


        End If

    End Sub
    Private Sub TogglePR()
        Dim fixedPRAVlossstring As String

        If My.Settings.AdvancedSettings_isAvailabilityMode = True Then
            fixedPRAVlossstring = "Av. Loss"
        Else
            fixedPRAVlossstring = "PR Loss"
        End If

        If TopStopsLegend_bubbletext.Content = "PR Loss %" Or TopStopsLegend_bubbletext.Content = "Av. Loss %" Then

            TopStopsLegend_bubbletext.Content = "DT min"
            TopStopsPRDataLabel1.Content = topstopsbar1_DTmin
            TopStopsPRDataLabel2.Content = topstopsbar2_DTmin
            TopStopsPRDataLabel3.Content = topstopsbar3_DTmin
            TopStopsPRDataLabel4.Content = topstopsbar4_DTmin
            TopStopsPRDataLabel5.Content = topstopsbar5_DTmin
            TopStopsPRDataLabel6.Content = topstopsbar6_DTmin
            TopStopsPRDataLabel7.Content = topstopsbar7_DTmin
            TopStopsPRDataLabel8.Content = topstopsbar8_DTmin
            TopStopsPRDataLabel9.Content = topstopsbar9_DTmin
            TopStopsPRDataLabel10.Content = topstopsbar10_DTmin
            TopStopsPRDataLabel11.Content = topstopsbar11_DTmin
            TopStopsPRDataLabel12.Content = topstopsbar12_DTmin
            TopStopsPRDataLabel13.Content = topstopsbar13_DTmin
            TopStopsPRDataLabel14.Content = topstopsbar14_DTmin
            TopStopsPRDataLabel15.Content = topstopsbar15_DTmin


            topstopsBubble1.ToolTip = "DT min " & topstopsbar1_DTmin
            topstopsBubble2.ToolTip = "DT min " & topstopsbar2_DTmin
            topstopsBubble3.ToolTip = "DT min " & topstopsbar3_DTmin
            topstopsBubble4.ToolTip = "DT min " & topstopsbar4_DTmin
            topstopsBubble5.ToolTip = "DT min " & topstopsbar5_DTmin
            topstopsBubble6.ToolTip = "DT min " & topstopsbar6_DTmin
            topstopsBubble7.ToolTip = "DT min " & topstopsbar7_DTmin
            topstopsBubble8.ToolTip = "DT min " & topstopsbar8_DTmin
            topstopsBubble9.ToolTip = "DT min " & topstopsbar9_DTmin
            topstopsBubble10.ToolTip = "DT min " & topstopsbar10_DTmin
            topstopsBubble11.ToolTip = "DT min " & topstopsbar11_DTmin
            topstopsBubble12.ToolTip = "DT min " & topstopsbar12_DTmin
            topstopsBubble13.ToolTip = "DT min " & topstopsbar13_DTmin
            topstopsBubble14.ToolTip = "DT min " & topstopsbar14_DTmin
            topstopsBubble15.ToolTip = "DT min " & topstopsbar15_DTmin

            'topstopsBubble1.ToolTip = 

        ElseIf TopStopsLegend_bubbletext.Content = "DT min" Then
            TopStopsLegend_bubbletext.Content = "PR Loss %"
            If My.Settings.AdvancedSettings_isAvailabilityMode Then TopStopsLegend_bubbletext.Content = "Av. Loss %"

            TopStopsPRDataLabel1.Content = FormatPercent(topstopsbar1_PRloss, 1)
            TopStopsPRDataLabel2.Content = FormatPercent(topstopsbar2_PRloss, 1)
            TopStopsPRDataLabel3.Content = FormatPercent(topstopsbar3_PRloss, 1)
            TopStopsPRDataLabel4.Content = FormatPercent(topstopsbar4_PRloss, 1)
            TopStopsPRDataLabel5.Content = FormatPercent(topstopsbar5_PRloss, 1)
            TopStopsPRDataLabel6.Content = FormatPercent(topstopsbar6_PRloss, 1)
            TopStopsPRDataLabel7.Content = FormatPercent(topstopsbar7_PRloss, 1)
            TopStopsPRDataLabel8.Content = FormatPercent(topstopsbar8_PRloss, 1)
            TopStopsPRDataLabel9.Content = FormatPercent(topstopsbar9_PRloss, 1)
            TopStopsPRDataLabel10.Content = FormatPercent(topstopsbar10_PRloss, 1)
            TopStopsPRDataLabel11.Content = FormatPercent(topstopsbar11_PRloss, 1)
            TopStopsPRDataLabel12.Content = FormatPercent(topstopsbar12_PRloss, 1)
            TopStopsPRDataLabel13.Content = FormatPercent(topstopsbar13_PRloss, 1)
            TopStopsPRDataLabel14.Content = FormatPercent(topstopsbar14_PRloss, 1)
            TopStopsPRDataLabel15.Content = FormatPercent(topstopsbar15_PRloss, 1)

            topstopsBubble1.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar1_PRloss, 1)
            topstopsBubble2.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar2_PRloss, 1)
            topstopsBubble3.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar3_PRloss, 1)
            topstopsBubble4.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar4_PRloss, 1)
            topstopsBubble5.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar5_PRloss, 1)
            topstopsBubble6.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar6_PRloss, 1)
            topstopsBubble7.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar7_PRloss, 1)
            topstopsBubble8.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar8_PRloss, 1)
            topstopsBubble9.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar9_PRloss, 1)
            topstopsBubble10.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar10_PRloss, 1)
            topstopsBubble11.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar11_PRloss, 1)
            topstopsBubble12.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar12_PRloss, 1)
            topstopsBubble13.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar13_PRloss, 1)
            topstopsBubble14.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar14_PRloss, 1)
            topstopsBubble15.ToolTip = fixedPRAVlossstring & FormatPercent(topstopsbar15_PRloss, 1)


        End If


    End Sub


    Private Sub launchcalendar()
        If incontrolStartDateLabel.Visibility = Windows.Visibility.Hidden Then
            ShowIncontrolDatePicker()
        Else
            HideIncontrolDatePicker()
        End If
    End Sub


    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        sender.Opacity = 0.7


        'keep cursor as pen if pickmode is ON
        ConsiderPickMode_for_Cursor(sender)


    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        sender.Opacity = 1.0
        'keep cursor as pen if pickmode is ON
        ConsiderPickMode_for_Cursor(sender)
        If IsPickMode = True Then Cursor = Cursors.Pen

    End Sub

    Private Sub ConsiderPickMode_for_Cursor(sender As Object)
        If IsPickMode = True Then
            If InStr(sender.name, "label") > 0 Or InStr(sender.name, "stopbubble") > 0 Then
                If InStr(sender.name, "PDT") > 0 Or InStr(sender.name, "Equip") > 0 Or InStr(sender.name, "change") > 0 Or InStr(sender.name, "stop") > 0 Then
                    Cursor = Cursors.Pen
                End If
            End If
        End If
    End Sub

#Region "Raw Data"
    Private RawDataLock As Boolean = True
    Private Sub LaunchRawDataforSelectedDTfield(sender As Object, e As MouseButtonEventArgs)
        Dim Timer As Single = 0
        If RawDataLock Then
            RawDataLock = False
            BusyIndicator.IsBusy = True
            Dim sngWaitEnd As Single = 5
            'sngWaitEnd = Timer + 2
            Do
                Threading.Thread.Sleep(100)
                Forms.Application.DoEvents()
                Timer = Timer + 1
            Loop Until Timer >= sngWaitEnd

            LaunchRawDataforSelectedDTfieldx(sender, e)

            BusyIndicator.IsBusy = False
            RawDataLock = True
        End If
    End Sub
    Private Sub LaunchRawDataforSelectedDTfieldx(sender As Object, e As MouseButtonEventArgs)

        Dim sourcecardTemp As String
        sourcecardTemp = sender.name
        If IsPickMode = True Then
            If InStr(sourcecardTemp, "Equip1", vbTextCompare) > 0 Then
                PickaLoss_CollectInfofromLabel(sender, 4)
                Exit Sub
            ElseIf InStr(sourcecardTemp, "Equip2", vbTextCompare) > 0 Then
                PickaLoss_CollectInfofromLabel(sender, 5)
                Exit Sub
            ElseIf InStr(sourcecardTemp, "Equip3", vbTextCompare) > 0 Then
                PickaLoss_CollectInfofromLabel(sender, 6)
                Exit Sub
            ElseIf InStr(sourcecardTemp, "changeover", vbTextCompare) > 0 Then
                PickaLoss_CollectInfofromLabel(sender, 41)

                Exit Sub
            End If


        End If

        If IsSimulationMode Then
            If InStr(sender.name, "_Rect") > 0 Or InStr(sender.name, "_datalabel") > 0 Then
                If InStr(sender.name, "UPDT") > 0 Then
                    LocateLossSImulator(sender, 1)
                ElseIf InStr(sender.name, "EquipMain") > 0 Then
                    LocateLossSImulator(sender, 3)
                ElseIf InStr(sender.name, "changeover") > 0 Then
                    LocateLossSImulator(sender, 41)
                ElseIf InStr(sender.name, "Equip1", vbTextCompare) > 0 Then
                    LocateLossSImulator(sender, 4)
                ElseIf InStr(sender.name, "Equip2", vbTextCompare) > 0 Then
                    LocateLossSImulator(sender, 5)
                ElseIf InStr(sender.name, "Equip3", vbTextCompare) > 0 Then
                    LocateLossSImulator(sender, 6)
                Else
                    LocateLossSImulator(sender, 2)
                End If


                Exit Sub
            End If
        End If



        '  Dim RawWindowThread As Thread
        Dim sourcefield As String

        Dim sourcecard As Integer
        Dim fieldnumber As Integer
        '  Dim targetEvent As DTevent
        ' Dim paramObj(2) As Object 'i think the (#) is the actual 'count', not the zero based one

        sourcefield = ""

        If Not InStr(sender.name, "data", vbTextCompare) > 0 Then
            If InStr(sender.name, "Label", vbTextCompare) > 0 Then Exit Sub
        End If
        sourcecardTemp = sender.name

        If sender.name = "" Then Exit Sub

        sourcecard = 1
        fieldnumber = onlyDigits(sender.name)
        If InStr(sender.name, "Rect", vbTextCompare) > 0 Then fieldnumber = onlyDigits(Strings.Mid(sender.name, InStr(sender.name, "Rect", vbTextCompare)))
        If InStr(sender.name, "data", vbTextCompare) > 0 Then fieldnumber = onlyDigits(Strings.Mid(sender.name, InStr(sender.name, "data", vbTextCompare)))

        If InStr(sourcecardTemp, "UPDT", vbTextCompare) > 0 Then
            sourcecard = 1
        End If
        If InStr(sourcecardTemp, "EquipMain", vbTextCompare) > 0 Then
            sourcecard = 3
        ElseIf InStr(sourcecardTemp, "Equip1", vbTextCompare) > 0 Then
            sourcecard = 4
        ElseIf InStr(sourcecardTemp, "Equip2", vbTextCompare) > 0 Then
            sourcecard = 5
        ElseIf InStr(sourcecardTemp, "Equip3", vbTextCompare) > 0 Then
            sourcecard = 6
        ElseIf InStr(sourcecardTemp, "Equip21", vbTextCompare) > 0 Then
            sourcecard = 21
        ElseIf InStr(sourcecardTemp, "Equip22", vbTextCompare) > 0 Then
            sourcecard = 22
        ElseIf InStr(sourcecardTemp, "changeover", vbTextCompare) > 0 Then
            sourcecard = 41
            prStoryReport.PlannedVariance_GenerateHTML()
        End If

        If InStr(sourcecardTemp, "UPDT", vbTextCompare) = 0 And InStr(sourcecardTemp, "PDT", vbTextCompare) > 0 Then
            sourcecard = 2
            sourcefield = pdtlabelstring(fieldnumber - 1)
        End If

        '  paramObj(0) = sourcecard
        '  paramObj(1) = fieldnumber
        '  Dim RawWindowThread As Thread = New Thread(AddressOf LaunchARawDataWindow_Thread)
        '  RawWindowThread.SetApartmentState(ApartmentState.STA)
        ' RawWindowThread.Start(paramObj)
        ' LaunchARawDataWindow_Thread(paramObj)

        '  End Sub

        '   Private Sub LaunchARawDataWindow_Thread(ByVal paramObj As Object)
        ' Cursor = Cursors.Wait
        ' System.Windows.Forms.Application.DoEvents()
        UseTrack_RawDatawindow_Main = True

        '  Dim sourceCard As Integer
        Dim targetEvent As DTevent
        '  Dim fieldNumber As Integer
        '  sourcecard = paramObj(0)
        ' fieldnumber = paramObj(1)

        Select Case sourcecard
            Case prStoryCard.Unplanned
                targetEvent = Card_Unplanned_T1(fieldnumber - 1)
            Case prStoryCard.Planned
                targetEvent = Card_Planned_T1(fieldnumber - 1)
            Case prStoryCard.Equipment
                targetEvent = Card_Unplanned_T2(fieldnumber - 1)
            Case prStoryCard.Equipment_One
                targetEvent = Card_Unplanned_T3A(fieldnumber - 1)
            Case prStoryCard.Equipment_Two
                targetEvent = Card_Unplanned_T3B(fieldnumber - 1)
            Case prStoryCard.Equipment_Three
                targetEvent = Card_Unplanned_T3C(fieldnumber - 1)
            Case prStoryCard.Changeover
                targetEvent = Card_Planned_T2(fieldnumber - 1)
            Case Else
                Throw New unknownprstoryCardException
        End Select

        If targetEvent.RawRows.Count > 0 Then 'if there is some data

            '  Dim weStillWantToDoThis As Boolean = True
            Dim rawdatawindow2 As New RawDataWindow(sourcecard, targetEvent.RawRows, prStoryReport, targetEvent.Name) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)
            rawdatawindow2.updateValues(sourcecard, targetEvent.RawRows, prStoryReport, targetEvent.Name) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)



            '  System.Windows.Forms.Application.DoEvents()
            rawdatawindow2.Owner = Me
            rawdatawindow2.setBargraphReportWindow_forraw(Me)

            rawdatawindow2.Show()
            '  rawdatawindow2.FinishConstruction()
            ' System.Windows.Threading.Dispatcher.Run()
        Else
            MsgBox("No Stops For Selected Field!", vbExclamation, "Mode Selected With No Stops/DT")
        End If
        Cursor = Cursors.Arrow
    End Sub
#End Region


    Private MultiLine2Window As Window_Multiline2

    Private Sub LaunchMultiLineDemo(sender As Object, e As MouseButtonEventArgs)
        MultiLine2Window = New Window_Multiline2(prStoryReport.ParentLineInt, prStoryReport.StartDate, prStoryReport.EndDate) 'RawDataWindow(sourcecard, targetEvent.RawRows, prStoryReport, targetEvent.Name) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)

        MultiLine2Window.Show()
    End Sub


    Private rawRateLossDataWindow As Window_RateLossReport
    Private Sub LaunchRawRateWindow_Thread()
        '   If targetEvent.RawRows.Count > 0 Then 'if there is some data
        Dim includeList As New List(Of String)
        Dim filterField As DowntimeField = DowntimeField.Team
        Dim doweFilter As Boolean = False

        If FilterOnOfflabel.Content = "Filter: ON" Then 'we filtered something!
            doweFilter = True
            includeList = AllProdLines(selectedindexofLine_temp).BrandCodesWeWant
            Select Case ProductFilterComboBox.SelectedValue

                Case "SKUs"
                    filterField = DowntimeField.Product
                Case "Teams"

                Case Else
                    Exit Sub

            End Select
        End If


        rawRateLossDataWindow = New Window_RateLossReport(AllProdLines(prStoryReport.ParentLineInt).Name, prStoryReport.schedTime, prStoryReport.StartDate, prStoryReport.EndDate, AllProdLines(prStoryReport.ParentLineInt).RawRateLossDataArray, doweFilter, includeList, filterField, selectedindexofLine_temp) 'RawDataWindow(sourcecard, targetEvent.RawRows, prStoryReport, targetEvent.Name) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)
        rawRateLossDataWindow.Show()
    End Sub

    Private RawRateWindowThread As Thread
    Private Sub LaunchRawDataforRateLoss(sender As Object, e As MouseButtonEventArgs)

        LaunchRawRateWindow_Thread()
    End Sub

    Private Sub showErrorSplashPage()
        SplashRectangle.Visibility = Visibility.Visible
        ErrorLabel_A.Visibility = Visibility.Visible
        ErrorLabel_B.Visibility = Visibility.Visible
        ErrorLabel_C.Visibility = Visibility.Visible
        ErrorLabel_D.Visibility = Visibility.Visible
        ExceptionComments.Visibility = Visibility.Visible
        ExceptionComments.Text = ""
        ErrorButton_A.Visibility = Visibility.Visible
    End Sub
    Private Sub hideErrorSplashPage()
        SplashRectangle.Visibility = Visibility.Hidden
        ErrorLabel_A.Visibility = Visibility.Hidden
        ErrorLabel_B.Visibility = Visibility.Hidden
        ErrorLabel_C.Visibility = Visibility.Hidden
        ErrorLabel_D.Visibility = Visibility.Hidden
        ErrorButton_A.Visibility = Visibility.Hidden
        ExceptionComments.Visibility = Visibility.Hidden
        NOtesLabel.Visibility = Visibility.Hidden
    End Sub

    Private Sub HideMenu() 'optional LeaveSplashRectangle as boolean = false)
        '  if Not leavesplashrectangle then
        SplashRectangle.Visibility = Visibility.Hidden
        '  End If

        MappingTextLabel.Visibility = Visibility.Hidden
        'WeibullTextLabel.Visibility = Visibility.Hidden
        MRIdemoTextLabel.Visibility = Visibility.Hidden

        ExportTextLabel.Visibility = Visibility.Hidden
        MappingIconLabel.Visibility = Visibility.Hidden
        ChangeMapping_SubCanvas.Visibility = Visibility.Hidden

        exportdataMenuButton.Visibility = Visibility.Hidden

        'LaunchDataDRbutton.Visibility = Visibility.Hidden
        'WeibullLaunchButton.Visibility = Visibility.Hidden
        'LiveLineLaunchButton.Visibility = Visibility.Hidden
        DependencyLaunchButton.Visibility = Visibility.Hidden
        SKUlistbox.Visibility = Visibility.Hidden
        SKUfilterCancelbutton.Visibility = Visibility.Hidden
        SKUfilterDonebutton.Visibility = Visibility.Hidden
        SKUListBoxLabel.Visibility = Visibility.Hidden
        SKUListFilterByLabel.Visibility = Visibility.Hidden
        ProductFilterComboBox.Visibility = Visibility.Hidden

        LossTreeDownloadIcon.Visibility = Visibility.Hidden
        LossTreeDownloadIcon2.Visibility = Visibility.Hidden
        ProductionDownloadIcon.Visibility = Visibility.Hidden
        DowntimeDownloadIcon.Visibility = Visibility.Hidden
        DependencyDownloadIcon.Visibility = Visibility.Hidden

        CSV_DT_CheckBox.Visibility = Visibility.Hidden
        CSV_PROD_CheckBox.Visibility = Visibility.Hidden
        CSV_LossTree_CheckBox.Visibility = Visibility.Hidden
        CSV_LossTree_CheckBox2.Visibility = Visibility.Hidden
        CSV_RE.Visibility = Visibility.Hidden

        SplashRectangle.Visibility = Visibility.Hidden
        SKUListView.Visibility = Visibility.Hidden
        SKUListViewReset.Visibility = Visibility.Hidden
        SKUListViewOK.Visibility = Visibility.Hidden
        SKUListViewLabel.Visibility = Visibility.Hidden
    End Sub

    Private Sub LaunchMenu()
        SplashRectangle.Visibility = Visibility.Visible

        MappingTextLabel.Visibility = Visibility.Visible
        'WeibullTextLabel.Visibility = Visibility.Visible
        ExportTextLabel.Visibility = Visibility.Visible

        MappingIconLabel.Visibility = Visibility.Visible
        exportdataMenuButton.Visibility = Visibility.Visible
        'WeibullLaunchButton.Visibility = Visibility.Visible
        LiveLineLaunchButton.Visibility = Visibility.Visible
    End Sub

#Region "Dependency Analysis"
    Private Sub LaunchREDEP()
        '   Dim reDepWin As New Window_REdependency(RE_getDependencyAnalysisForLine(prStoryReport.ParentLineInt))
        ' reDepWin.Show()
    End Sub
#End Region

#Region "Weibull"
    Private Sub LaunchWeibull()
        '  Dim WeibullThread As Thread
        HideMenu()
        UseTrack_WeibullMain = True
        Weibull_Thread()
        '  Dim weibullwindow As New Window_Weibull(prStoryReport)
        ' WeibullThread = New Thread(AddressOf Weibull_Thread)
        ' WeibullThread.SetApartmentState(ApartmentState.STA)
        ' WeibullThread.Start()

    End Sub
    Private Sub Weibull_Thread()
        Dim weibullwindowX As New Window_Weibull(prStoryReport)

        weibullwindowX.Show()
        ' System.Windows.Threading.Dispatcher.Run()
    End Sub
#End Region

#Region "Mapping & Remapping"
    Private formatIndex As Integer
    Private shapeIndex As Integer
    Private productGroupIndex As Integer
    Private Sub PopulateMappingList()
        Dim tmpIndex As Integer
        Dim blankMappingString As String = " "
        'PRIMARY MAPPING LEVEL
        MappingSelectionCombo.Visibility = Visibility.Visible
        MappingSelectionCombo.Items.Clear()
        MappingSelectionCombo.Items.Add(AllProdLines(selectedindexofLine_temp).Reason1Name) ' 0
        MappingSelectionCombo.Items.Add(AllProdLines(selectedindexofLine_temp).Reason2Name) '1
        MappingSelectionCombo.Items.Add(AllProdLines(selectedindexofLine_temp).Reason3Name) '2
        MappingSelectionCombo.Items.Add(AllProdLines(selectedindexofLine_temp).Reason4Name) '3
        MappingSelectionCombo.Items.Add(AllProdLines(selectedindexofLine_temp).DTgroupName) '4
        MappingSelectionCombo.Items.Add(AllProdLines(selectedindexofLine_temp).FaultCodeName) '5
        MappingSelectionCombo.Items.Add("SKU") '6

        MappingSelectionCombo.Items.Add("Tier 1") 'LG Code  '7
        MappingSelectionCombo.Items.Add("Tier 2") 'LG Code  '8
        MappingSelectionCombo.Items.Add("Tier 3") 'LG Code '9

        tmpIndex = 9
        If AllProdLines(selectedindexofLine_temp).formatMapping <> MappingByFormat.NoMapping Then
            tmpIndex += 1
            MappingSelectionCombo.Items.Add("Format")
            formatIndex = tmpIndex
        End If
        If AllProdLines(selectedindexofLine_temp).shapeMapping <> MappingByShape.NoMapping Then
            tmpIndex += 1
            MappingSelectionCombo.Items.Add("Shape")
            shapeIndex = tmpIndex
        End If
        If AllProdLines(selectedindexofLine_temp).FieldCheck_ProductGroup Then
            tmpIndex += 1
            MappingSelectionCombo.Items.Add("Product Groups")
            productGroupIndex = tmpIndex
        End If


        Select Case My.Settings.defaultDownTimeField
            Case DowntimeField.Reason1
                MappingSelectionCombo.SelectedIndex = 0
            Case DowntimeField.Reason2
                MappingSelectionCombo.SelectedIndex = 1
            Case DowntimeField.Reason3
                MappingSelectionCombo.SelectedIndex = 2
            Case DowntimeField.Reason4
                MappingSelectionCombo.SelectedIndex = 3
            Case DowntimeField.DTGroup
                MappingSelectionCombo.SelectedIndex = 4
            Case DowntimeField.Fault
                MappingSelectionCombo.SelectedIndex = 5
            Case DowntimeField.ProductCode
                MappingSelectionCombo.SelectedIndex = 6
            Case DowntimeField.Format
                MappingSelectionCombo.SelectedIndex = formatIndex
            Case DowntimeField.Shape
                MappingSelectionCombo.SelectedIndex = shapeIndex
            Case DowntimeField.ProductGroup
                MappingSelectionCombo.SelectedIndex = productGroupIndex
            Case DowntimeField.Tier1
                MappingSelectionCombo.SelectedIndex = 7
            Case DowntimeField.Tier2
                MappingSelectionCombo.SelectedIndex = 8
            Case DowntimeField.Tier3
                MappingSelectionCombo.SelectedIndex = 9
        End Select
        ' MappingSelectionCombo.SelectedValue = tempreasonlevel


        'SECONDARY MAPPING LEVEL
        MappingSelectionCombo2.Visibility = Visibility.Visible
        MappingSelectionCombo2.Items.Clear()
        MappingSelectionCombo2.Items.Add(blankMappingString)
        MappingSelectionCombo2.Items.Add(AllProdLines(selectedindexofLine_temp).Reason1Name) ' 0
        MappingSelectionCombo2.Items.Add(AllProdLines(selectedindexofLine_temp).Reason2Name) '1
        MappingSelectionCombo2.Items.Add(AllProdLines(selectedindexofLine_temp).Reason3Name) '2
        MappingSelectionCombo2.Items.Add(AllProdLines(selectedindexofLine_temp).Reason4Name) '3
        MappingSelectionCombo2.Items.Add(AllProdLines(selectedindexofLine_temp).DTgroupName) '4
        MappingSelectionCombo2.Items.Add(AllProdLines(selectedindexofLine_temp).FaultCodeName) '5
        MappingSelectionCombo2.Items.Add("SKU")
        MappingSelectionCombo2.Items.Add("Tier 1") 'LG Code  '7
        MappingSelectionCombo2.Items.Add("Tier 2") 'LG Code  '8
        MappingSelectionCombo2.Items.Add("Tier 3") 'LG Code '9


        tmpIndex = 10
        If AllProdLines(selectedindexofLine_temp).formatMapping <> MappingByFormat.NoMapping Then
            ' tmpIndex += 1
            MappingSelectionCombo2.Items.Add("Format")
            ' formatIndex = tmpIndex
        End If
        If AllProdLines(selectedindexofLine_temp).shapeMapping <> MappingByShape.NoMapping Then
            ' tmpIndex += 1
            MappingSelectionCombo2.Items.Add("Shape")
            ' shapeIndex = tmpIndex
        End If
        If AllProdLines(selectedindexofLine_temp).FieldCheck_ProductGroup Then
            ' tmpIndex += 1
            MappingSelectionCombo2.Items.Add("Product Groups")
            '   productGroupIndex = tmpIndex
        End If


        Select Case My.Settings.defaultDownTimeField_Secondary
            Case -1
                MappingSelectionCombo2.SelectedValue = blankMappingString
            Case DowntimeField.Reason1
                MappingSelectionCombo2.SelectedIndex = 1
            Case DowntimeField.Reason2
                MappingSelectionCombo2.SelectedIndex = 2
            Case DowntimeField.Reason3
                MappingSelectionCombo2.SelectedIndex = 3
            Case DowntimeField.Reason4
                MappingSelectionCombo2.SelectedIndex = 4
            Case DowntimeField.DTGroup
                MappingSelectionCombo2.SelectedIndex = 5
            Case DowntimeField.Fault
                MappingSelectionCombo2.SelectedIndex = 6
            Case DowntimeField.ProductCode
                MappingSelectionCombo2.SelectedIndex = 7
            Case DowntimeField.Tier1
                MappingSelectionCombo2.SelectedIndex = 8
            Case DowntimeField.Tier2
                MappingSelectionCombo2.SelectedIndex = 9
            Case DowntimeField.Tier3
                MappingSelectionCombo2.SelectedIndex = 10
            Case DowntimeField.Format
                MappingSelectionCombo2.SelectedIndex = formatIndex + 1
            Case DowntimeField.Shape
                MappingSelectionCombo2.SelectedIndex = shapeIndex + 1
            Case DowntimeField.ProductGroup
                MappingSelectionCombo2.SelectedIndex = productGroupIndex + 1
        End Select


    End Sub


    Private Sub ShowMappingSelection()
        PopulateMappingList()
        ChangeMapping_SubCanvas.Visibility = Visibility.Visible
    End Sub


    Private Sub MappingSelectionChange()
        'primary mapping level
        Select Case MappingSelectionCombo.SelectedIndex
            Case 0
                My.Settings.defaultDownTimeField = DowntimeField.Reason1
            Case 1
                My.Settings.defaultDownTimeField = DowntimeField.Reason2
            Case 2
                My.Settings.defaultDownTimeField = DowntimeField.Reason3
            Case 3
                My.Settings.defaultDownTimeField = DowntimeField.Reason4
            Case 4
                My.Settings.defaultDownTimeField = DowntimeField.DTGroup
            Case 5
                My.Settings.defaultDownTimeField = DowntimeField.Fault
            Case 6
                My.Settings.defaultDownTimeField = DowntimeField.ProductCode
            Case formatIndex
                My.Settings.defaultDownTimeField = DowntimeField.Format
            Case shapeIndex
                My.Settings.defaultDownTimeField = DowntimeField.Shape
            Case productGroupIndex
                My.Settings.defaultDownTimeField = DowntimeField.ProductGroup
            Case 7
                My.Settings.defaultDownTimeField = DowntimeField.Tier1
            Case 8
                My.Settings.defaultDownTimeField = DowntimeField.Tier2
            Case 9
                My.Settings.defaultDownTimeField = DowntimeField.Tier3

        End Select
        tempreasonlevel = MappingSelectionCombo.SelectedValue

        'secondary mapping level
        Select Case MappingSelectionCombo2.SelectedIndex
            Case -1
                My.Settings.defaultDownTimeField_Secondary = -1
            Case 0
                My.Settings.defaultDownTimeField_Secondary = -1
            Case 1
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Reason1
            Case 2
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Reason2
            Case 3
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Reason3
            Case 4
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Reason4
            Case 5
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.DTGroup
            Case 6
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Fault
            Case 7
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.ProductCode
            Case formatIndex + 1
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Format
            Case shapeIndex + 1
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Shape
            Case productGroupIndex + 1
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.ProductGroup
            Case 8
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Tier1
            Case 9
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Tier2
            Case 10
                My.Settings.defaultDownTimeField_Secondary = DowntimeField.Tier3
        End Select

        'work hard hard work
        MainDateLabel.Content = Format(starttimeselected, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(endtimeselected, "MMMM dd, yyyy HH:mm").ToString & vbNewLine

        prStoryReport.reMapReport()
        AllProdLines(selectedindexofLine_temp).reMapRawData()
        MaxStopsindataset = Card_TopStops(0).SPD
        MaxPRlossindataset_stops = prStoryReport.EventMaxDTpct 'LG Code
        updateCard_TopStops()


        If stoplabel_1.IsVisible Then 'if we're looking at top stops, do this
            showstoplabels()
            Assignlabelnames(prStoryReport)
        End If

        If incontrolgreenbox.Visibility = Windows.Visibility.Visible Then
            GenerateIncontrolCharts(incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
        Else
            MasterDataSet = New inControlReport(AllProdLines(selectedindexofLine_temp), incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
        End If
        HideMenu()
        hideRL4()

        UseTrack_ChangeMapping = True
        IsRemappingDoneOnce = True
    End Sub
    Sub mappingiconclick()


        ChangeMapping_SubCanvas.Visibility = Visibility.Hidden
        MappingSelectionChange()


    End Sub
#End Region

    Public Sub showDefaultErrorMessage(header As String)
        MessageBox.Show("Something went wrong! Sorry for the inconvenience.  Please restart prstory and try again. " & vbCrLf & "If the problem persists, please report the issue to your SPOC.",
                       header,
                Forms.MessageBoxButtons.OK,
Forms.MessageBoxIcon.Warning,
Forms.MessageBoxDefaultButton.Button1)
    End Sub

#Region "Menu"
    Sub DownloadLossTree()
        UseTrack_ExportLossTree = True
        Try
            ExportLossAllocationtoExcel()
        Catch e As Exception
            showDefaultErrorMessage("Excel Export Error: " & e.Message)
        End Try
    End Sub
    Sub DownloadLossTree2()
        UseTrack_ExportLossTree = True
        Try
            ExportLossAllocationtoExcel2()
        Catch e As Exception
            showDefaultErrorMessage("CSV Export Error: " & e.Message)
        End Try
    End Sub
    Sub DownloadDowntime()
        UseTrack_ExportDowntime = True
        Try
            CSV_exportRawLEDsData(AllProdLines(selectedindexofLine_temp), prStoryReport, True, False, False, False)
        Catch e As Exception
            showDefaultErrorMessage("Downtime Export Error: " & e.Message)
        End Try
    End Sub
    Sub DownloadProduction()
        UseTrack_ExportProduction = True
        Try
            CSV_exportRawLEDsData(AllProdLines(selectedindexofLine_temp), prStoryReport, False, True, False, False)
        Catch e As Exception
            showDefaultErrorMessage("Production Export Error: " & e.Message)
        End Try
    End Sub
    Sub DownloadDependency()
        UseTrack_ExportDependency = True
        Try
            CSV_exportRawLEDsData(AllProdLines(selectedindexofLine_temp), prStoryReport, False, False, True, False)
        Catch e As Exception
            showDefaultErrorMessage("Export Error")
        End Try
    End Sub



    Sub showCSVSelection()
        CSV_DT_CheckBox.Visibility = Visibility.Visible
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
            CSV_PROD_CheckBox.Visibility = Visibility.Visible
            ProductionDownloadIcon.Visibility = Visibility.Visible
        End If
        CSV_LossTree_CheckBox.Visibility = Visibility.Visible
        CSV_LossTree_CheckBox2.Visibility = Visibility.Visible
        CSV_RE.Visibility = Visibility.Visible

        LossTreeDownloadIcon.Visibility = Visibility.Visible
        LossTreeDownloadIcon2.Visibility = Visibility.Visible

        DowntimeDownloadIcon.Visibility = Visibility.Visible
        DependencyDownloadIcon.Visibility = Visibility.Visible

    End Sub
    Sub exportSelectedCSVData()
        ' CSV_exportRawLEDsData(AllProductionLines(selectedindexofLine_temp), prStoryReport, CSV_DT_CheckBox.IsChecked, CSV_PROD_CheckBox.IsChecked, CSV_RE.IsChecked, CSV_LossTree_CheckBox.IsChecked)
        CSV_DT_CheckBox.Visibility = Visibility.Hidden
        CSV_PROD_CheckBox.Visibility = Visibility.Hidden
        CSV_LossTree_CheckBox.Visibility = Visibility.Hidden
        CSV_LossTree_CheckBox2.Visibility = Visibility.Hidden
        CSV_RE.Visibility = Visibility.Hidden

        LossTreeDownloadIcon.Visibility = Visibility.Hidden
        LossTreeDownloadIcon2.Visibility = Visibility.Hidden
        ProductionDownloadIcon.Visibility = Visibility.Hidden
        DowntimeDownloadIcon.Visibility = Visibility.Hidden
        DependencyDownloadIcon.Visibility = Visibility.Hidden

    End Sub
    Sub LaunchSKUlistbox()
        SKUlistbox.Visibility = Visibility.Visible
        SKUfilterCancelbutton.Visibility = Visibility.Visible
        SKUfilterDonebutton.Visibility = Visibility.Visible
        SKUListBoxLabel.Visibility = Visibility.Visible
        SKUListFilterByLabel.Visibility = Visibility.Visible
        ProductFilterComboBox.Visibility = Visibility.Visible
        SKUlistbox.ItemsSource = Nothing
        SKUlistbox.Items.Clear()
        SplashRectangle.Visibility = Visibility.Visible
        Populate_ProductFilterComboBox()
        SKUListBoxLabel.Content = "Choose from the dropdown what you want to filter by. SKUs, Teams..?"
    End Sub
    Sub ProductFilterComboBoxSelectionChange()

        Select Case ProductFilterComboBox.SelectedValue

            Case "SKUs"
                prStoryReport.updateProductList(DowntimeField.Product)
            Case "Teams"
                prStoryReport.updateProductList(DowntimeField.Team)
            Case "Formats"
                prStoryReport.updateProductList(DowntimeField.Format)
            Case "Shapes"
                prStoryReport.updateProductList(DowntimeField.Shape)
            Case "Product Groups"
                prStoryReport.updateProductList(DowntimeField.ProductGroup)
            Case Else
                Exit Sub

        End Select

        SKUlistbox.ItemsSource = Nothing
        SKUlistbox.Items.Clear()
        SKUlistbox.ItemsSource = prStoryReport.ProductList
        SKUListBoxLabel.Content = "Select one or more items from the following list of " & ProductFilterComboBox.SelectedValue
    End Sub
    Sub Populate_ProductFilterComboBox()
        ProductFilterComboBox.Items.Clear()
        ProductFilterComboBox.Items.Add("SKUs")
        ProductFilterComboBox.Items.Add("Teams")

        'check if line uses format or mapping
        If AllProdLines(selectedindexofLine_temp).formatMapping <> MappingByFormat.NoMapping Then
            ProductFilterComboBox.Items.Add("Formats")
        End If
        If AllProdLines(selectedindexofLine_temp).shapeMapping <> MappingByShape.NoMapping Then
            ProductFilterComboBox.Items.Add("Shapes")
        End If
        If AllProdLines(selectedindexofLine_temp).FieldCheck_ProductGroup Then
            ProductFilterComboBox.Items.Add("Product Groups")
        End If

    End Sub
    Sub CancelSKUlistbox()

        SKUlistbox.Visibility = Visibility.Hidden
        SKUfilterCancelbutton.Visibility = Visibility.Hidden
        SKUfilterDonebutton.Visibility = Visibility.Hidden
        SKUListBoxLabel.Visibility = Visibility.Hidden
        SKUListFilterByLabel.Visibility = Visibility.Hidden
        ProductFilterComboBox.Visibility = Visibility.Hidden
        HideMenu()
    End Sub

    Sub DoneSKUlistbox()
        UseTrack_Filter = True

        SKUlistbox.Visibility = Visibility.Hidden
        SKUfilterCancelbutton.Visibility = Visibility.Hidden
        SKUfilterDonebutton.Visibility = Visibility.Hidden
        SKUListBoxLabel.Visibility = Visibility.Hidden
        SKUListFilterByLabel.Visibility = Visibility.Hidden
        ProductFilterComboBox.Visibility = Visibility.Hidden
        AllProdLines(selectedindexofLine_temp).isFilterByBrandcode = True
        AllProdLines(selectedindexofLine_temp).BrandCodesWeWant.Clear()
        For i As Integer = 0 To SKUlistbox.SelectedItems.Count - 1
            AllProdLines(selectedindexofLine_temp).BrandCodesWeWant.Add(CStr(SKUlistbox.SelectedItems(i)))
        Next

        Select Case ProductFilterComboBox.SelectedValue

            Case "SKUs"
                prStoryReport.reFilterData_SKU(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
                AllProdLines(selectedindexofLine_temp).reFilterData_SKU(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
            Case "Teams"
                prStoryReport.reFilterData_Team(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
                AllProdLines(selectedindexofLine_temp).reFilterData_Team(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
            Case "Formats"
                prStoryReport.reFilterData_Format(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
                AllProdLines(selectedindexofLine_temp).reFilterData_Format(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
            Case "Shapes"
                prStoryReport.reFilterData_Shape(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
                AllProdLines(selectedindexofLine_temp).reFilterData_Shape(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
            Case "Product Groups"
                prStoryReport.reFilterData_ProductGroup(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
                AllProdLines(selectedindexofLine_temp).reFilterData_ProductGroup(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
            Case Else
                Exit Sub

        End Select


        MaxPRindataset = prStoryReport.UPDT
        MaxPRindataset_planned = prStoryReport.PDT

        MaxStopsindataset = Card_TopStops(0).SPD


        updateCard_TopStops()
        updateCard_Planned_Tier1()
        updateCard_Unplanned_Tier1()

        If stoplabel_1.IsVisible Then 'if we're looking at top stops, do this
            showstoplabels()
            Assignlabelnames(prStoryReport)
        ElseIf UPDTlabel1.IsVisible Then
            showDTpercentframe()
        End If

        If incontrolgreenbox.Visibility = Windows.Visibility.Visible Then
            GenerateIncontrolCharts(incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
        Else
            Try
                MasterDataSet = New inControlReport(AllProdLines(selectedindexofLine_temp), incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
            Catch
            End Try
        End If

        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            pr_label.Content = FormatPercent(prStoryReport.PR, 1) ' & " PR"
            PR_Label_Header.Content = "PR"
        Else
            pr_label.Content = FormatPercent(prStoryReport.PR, 1) ' & " Av."
            PR_Label_Header.Content = "Availability"
            cases_label.Visibility = Visibility.Hidden
            rateloss_label.Visibility = Visibility.Hidden
            Cases_Label_Header.Visibility = Visibility.Hidden
            Rateloss_Label_Header.Visibility = Visibility.Hidden
        End If

        If prStoryReport.CasesAdjusted = 0 Then
            cases_label.Content = prStoryReport.CasesActual ' & " cases"
        Else
            cases_label.Content = prStoryReport.CasesAdjusted '& " cases"
        End If

        rateloss_label.Content = FormatPercent(prStoryReport.rateLoss, 1) ' ' & " rate and quality loss"
        Stops_Label.Content = Math.Round(prStoryReport.StopsPerDay, 0) ' & " stops/day"
        MTBF_Label.Content = Math.Round(prStoryReport.MTBF, 0) & " min"
        HideMenu()
        FilterIcon_inactive.Visibility = Visibility.Hidden
        FilterIcon_active.Visibility = Visibility.Visible
        FilterOnOfflabel.Content = "Filter: ON"

        SplashRectangle.Visibility = Visibility.Hidden

        IsRemappingDoneOnce = True
    End Sub
    Private Sub RefreshHeaderColors()

        Card4Header.Background = CardHeaderdefaultColor
        Card5Header.Background = CardHeaderdefaultColor
        Card6Header.Background = CardHeaderdefaultColor

    End Sub
#End Region

#Region "Printing"
    Public Sub PrintScreen()
        Dim printdlg As New PrintDialog
        HideMenu()
        HideDropShadows(unplannedDTchart)
        HideDropShadows(plannedDTchart)
        HideDropShadows(unplannedDTequipmentchart)
        HideDropShadows(unplannedDTequip1chart)
        HideDropShadows(unplannedDTequip2chart)
        HideDropShadows(unplannedDTequip3chart)
        HideDropShadows(topstopschart)
        HideDropShadows(plannedchangeovertimechart)
        HideDropShadows(changeoverchart)
        HideDropShadows(IncontrolActiveArea)
        maincanvas.Background = mybrushcolorlesswhite


        'printdlg.PrintVisual(Me, "Choices")
        If printdlg.ShowDialog = True Then printdlg.PrintVisual(PrintGrid, "prStory")

        ShowDropShadows(unplannedDTchart)
        ShowDropShadows(plannedDTchart)
        ShowDropShadows(unplannedDTequipmentchart)
        ShowDropShadows(unplannedDTequip1chart)
        ShowDropShadows(unplannedDTequip2chart)
        ShowDropShadows(unplannedDTequip3chart)
        ShowDropShadows(topstopschart)
        ShowDropShadows(plannedchangeovertimechart)
        ShowDropShadows(changeoverchart)
        ShowDropShadows(IncontrolActiveArea)

        maincanvas.Background = mybrushgray

    End Sub
    Private Sub HideDropShadows(shape As Object)
        'Exit Sub
        'unplannedDTchart.Effect


        Dim NO_dropshaddowobj As New DropShadowEffect
        NO_dropshaddowobj.ShadowDepth = 0
        NO_dropshaddowobj.BlurRadius = 0

        shape.effect = NO_dropshaddowobj

    End Sub

    Private Sub ShowDropShadows(shape As Object)
        ' Exit Sub
        Dim dropshaddowobj As New DropShadowEffect
        dropshaddowobj.ShadowDepth = 2
        dropshaddowobj.BlurRadius = 10
        'dropshaddowobj.Color = Color.FromArgb(143, 143, 143)

        shape.effect = dropshaddowobj
    End Sub
#End Region

    Sub LaunchTrends(sender As Object, e As MouseButtonEventArgs)
        Dim RefreshExportsThread As New Thread(AddressOf refreshTrendCharts_Exports)
        Cursor = Cursors.Wait

        '   Dim motionchartwindow As New Window_MotionChart '(onlyDigits(sender.name))
        If AllProdLines(selectedindexofLine_temp).IsStartupMode Or AllProdLines(selectedindexofLine_temp).ToString.Contains("IP74") Or AllProdLines(selectedindexofLine_temp).ToString = "GBO APDO J" Or AllProdLines(selectedindexofLine_temp).ToString = "DIBH11" Then
            MsgBox("Not enough data to show trends.")
            Exit Sub
        End If

        If IsRemappingDoneOnce = True Then
            Cursor = Cursors.Wait
            Dim stopsMotionReport As MotionReport = New MotionReport(AllProdLines(selectedindexofLine_temp), starttimeselected, endtimeselected, prStoryReport.MainLEDSReport.DT_Report.UnplannedEventDirectory, 1)
            Dim k As Integer

            RefreshExportsThread.Start()
            GoTo skipRefreshNonThreadedExports
            For k = 0 To Math.Min(14, stopsMotionReport.motionEvents_All15.Count - 1)

                exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, True, k)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, False, k)

                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, True, k)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, False, k)

                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, True, k)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, False, k)

                exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF(stopsMotionReport, True, k)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Weekly(stopsMotionReport, True, k)
                exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Monthly(stopsMotionReport, True, k)

            Next k
skipRefreshNonThreadedExports:
        End If

        motionchartsource = onlyDigits(sender.name)
        If motionchartsource = 31 Then
            Dim motionchartwindow = New Window_MotionChart(True, SelectedFailuremonumber_inTopStopsforTrends, topstopname(SelectedFailuremonumber_inTopStopsforTrends))

            motionchartwindow.EventDirectory = prStoryReport.MainLEDSReport.DT_Report.UnplannedEventDirectory  'added for s shape
            motionchartwindow.ledsReport = prStoryReport.MainLEDSReport
            motionchartwindow.downtimeData = AllProdLines(selectedindexofLine_temp).rawDowntimeData

            motionchartwindow.Owner = Me
            motionchartwindow.ShowDialog()
            motionchartwindow.Topmost = True
        Else
            Dim motionchartwindow = New Window_MotionChart(False)
            motionchartwindow.Owner = Me
            motionchartwindow.ShowDialog()
            motionchartwindow.Topmost = True
        End If
        Cursor = Cursors.Arrow
    End Sub
    Public Sub refreshTrendCharts_Exports()
        Dim k As Integer
        Dim stopsMotionReport As MotionReport = New MotionReport(AllProdLines(selectedindexofLine_temp), starttimeselected, endtimeselected, prStoryReport.MainLEDSReport.DT_Report.UnplannedEventDirectory, 1)

        For k = 0 To Math.Min(14, stopsMotionReport.motionEvents_All15.Count - 1)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, False, k)

            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, False, k)

            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, False, k)

            exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Weekly(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Monthly(stopsMotionReport, True, k)
        Next k
    End Sub
    Sub topstopbubbleclick(sender As Object, e As MouseButtonEventArgs)
        topstopbubble_Animate(sender)
    End Sub

    Sub topstopbubble_Animate(bubble As Object)

        Dim topstopbubbleposition As Thickness
        resettopstopbubbles()


        topstopbubbleposition = bubble.margin
        bubble.height = 25
        bubble.width = 25



        Select Case onlyDigits(bubble.name)

            Case 1
                TopStopsPRDataLabel1.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel1.Content = FormatPercent(topstopsbar1_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel1.Content = topstopsbar1_DTmin
                TopStopsPRDataLabel1.Visibility = Visibility.Visible
            Case 2
                TopStopsPRDataLabel2.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel2.Content = FormatPercent(topstopsbar2_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel2.Content = topstopsbar2_DTmin
                TopStopsPRDataLabel2.Visibility = Visibility.Visible
            Case 3
                TopStopsPRDataLabel3.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel3.Content = FormatPercent(topstopsbar3_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel3.Content = topstopsbar3_DTmin
                TopStopsPRDataLabel3.Visibility = Visibility.Visible
            Case 4
                TopStopsPRDataLabel4.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel4.Content = FormatPercent(topstopsbar4_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel4.Content = topstopsbar4_DTmin
                TopStopsPRDataLabel4.Visibility = Visibility.Visible
            Case 5
                TopStopsPRDataLabel5.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel5.Content = FormatPercent(topstopsbar5_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel5.Content = topstopsbar5_DTmin
                TopStopsPRDataLabel5.Visibility = Visibility.Visible
            Case 6
                TopStopsPRDataLabel6.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel6.Content = FormatPercent(topstopsbar6_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel6.Content = topstopsbar6_DTmin
                TopStopsPRDataLabel6.Visibility = Visibility.Visible
            Case 7
                TopStopsPRDataLabel7.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel7.Content = FormatPercent(topstopsbar7_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel7.Content = topstopsbar7_DTmin
                TopStopsPRDataLabel7.Visibility = Visibility.Visible
            Case 8
                TopStopsPRDataLabel8.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel8.Content = FormatPercent(topstopsbar8_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel8.Content = topstopsbar8_DTmin
                TopStopsPRDataLabel8.Visibility = Visibility.Visible
            Case 9
                TopStopsPRDataLabel9.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel9.Content = FormatPercent(topstopsbar9_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel9.Content = topstopsbar9_DTmin
                TopStopsPRDataLabel9.Visibility = Visibility.Visible
            Case 10
                TopStopsPRDataLabel10.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel10.Content = FormatPercent(topstopsbar10_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel10.Content = topstopsbar10_DTmin
                TopStopsPRDataLabel10.Visibility = Visibility.Visible
            Case 11
                TopStopsPRDataLabel11.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel11.Content = FormatPercent(topstopsbar11_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel11.Content = topstopsbar11_DTmin
                TopStopsPRDataLabel11.Visibility = Visibility.Visible
            Case 12
                TopStopsPRDataLabel12.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel12.Content = FormatPercent(topstopsbar12_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel12.Content = topstopsbar12_DTmin
                TopStopsPRDataLabel12.Visibility = Visibility.Visible
            Case 13
                TopStopsPRDataLabel13.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel13.Content = FormatPercent(topstopsbar13_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel13.Content = topstopsbar13_DTmin
                TopStopsPRDataLabel13.Visibility = Visibility.Visible
            Case 14
                TopStopsPRDataLabel14.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel14.Content = FormatPercent(topstopsbar14_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel14.Content = topstopsbar14_DTmin
                TopStopsPRDataLabel14.Visibility = Visibility.Visible
            Case 15
                TopStopsPRDataLabel15.Margin = New Thickness(topstopbubbleposition.Left - 3, topstopbubbleposition.Top - 20, 0, 0)
                TopStopsPRDataLabel15.Content = FormatPercent(topstopsbar15_PRloss, 1)
                If TopStopsLegend_bubbletext.Content = "DT min" Then TopStopsPRDataLabel15.Content = topstopsbar15_DTmin
                TopStopsPRDataLabel15.Visibility = Visibility.Visible





        End Select

    End Sub

    Sub resettopstopbubbles()

        topstopsBubble1.Height = topstopbubbledefault_Height
        topstopsBubble1.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel1.Visibility = Visibility.Hidden
        TopStopsPRDataLabel1.Content = FormatPercent(topstopsbar1_PRloss, 1)

        topstopsBubble2.Height = topstopbubbledefault_Height
        topstopsBubble2.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel2.Visibility = Visibility.Hidden
        TopStopsPRDataLabel2.Content = FormatPercent(topstopsbar2_PRloss, 1)

        topstopsBubble3.Height = topstopbubbledefault_Height
        topstopsBubble3.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel3.Visibility = Visibility.Hidden
        TopStopsPRDataLabel3.Content = FormatPercent(topstopsbar3_PRloss, 1)

        topstopsBubble4.Height = topstopbubbledefault_Height
        topstopsBubble4.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel4.Visibility = Visibility.Hidden
        TopStopsPRDataLabel4.Content = FormatPercent(topstopsbar4_PRloss, 1)

        topstopsBubble5.Height = topstopbubbledefault_Height
        topstopsBubble5.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel5.Visibility = Visibility.Hidden
        TopStopsPRDataLabel5.Content = FormatPercent(topstopsbar5_PRloss, 1)

        topstopsBubble6.Height = topstopbubbledefault_Height
        topstopsBubble6.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel6.Visibility = Visibility.Hidden
        TopStopsPRDataLabel6.Content = FormatPercent(topstopsbar6_PRloss, 1)

        topstopsBubble7.Height = topstopbubbledefault_Height
        topstopsBubble7.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel7.Visibility = Visibility.Hidden
        TopStopsPRDataLabel7.Content = FormatPercent(topstopsbar7_PRloss, 1)


        topstopsBubble8.Height = topstopbubbledefault_Height
        topstopsBubble8.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel8.Visibility = Visibility.Hidden
        TopStopsPRDataLabel8.Content = FormatPercent(topstopsbar8_PRloss, 1)

        topstopsBubble9.Height = topstopbubbledefault_Height
        topstopsBubble9.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel9.Visibility = Visibility.Hidden
        TopStopsPRDataLabel9.Content = FormatPercent(topstopsbar9_PRloss, 1)

        topstopsBubble10.Height = topstopbubbledefault_Height
        topstopsBubble10.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel10.Visibility = Visibility.Hidden
        TopStopsPRDataLabel10.Content = FormatPercent(topstopsbar10_PRloss, 1)

        topstopsBubble11.Height = topstopbubbledefault_Height
        topstopsBubble11.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel11.Visibility = Visibility.Hidden
        TopStopsPRDataLabel11.Content = FormatPercent(topstopsbar11_PRloss, 1)

        topstopsBubble12.Height = topstopbubbledefault_Height
        topstopsBubble12.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel12.Visibility = Visibility.Hidden
        TopStopsPRDataLabel12.Content = FormatPercent(topstopsbar12_PRloss, 1)


        topstopsBubble13.Height = topstopbubbledefault_Height
        topstopsBubble13.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel13.Visibility = Visibility.Hidden
        TopStopsPRDataLabel13.Content = FormatPercent(topstopsbar13_PRloss, 1)

        topstopsBubble14.Height = topstopbubbledefault_Height
        topstopsBubble14.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel14.Visibility = Visibility.Hidden
        TopStopsPRDataLabel14.Content = FormatPercent(topstopsbar14_PRloss, 1)

        topstopsBubble15.Height = topstopbubbledefault_Height
        topstopsBubble15.Width = topstopbubbledefault_Height
        TopStopsPRDataLabel15.Visibility = Visibility.Hidden
        TopStopsPRDataLabel15.Content = FormatPercent(topstopsbar15_PRloss, 1)

    End Sub

    Sub hidefiltericon()
        FilterIcon_active.Visibility = Visibility.Hidden
    End Sub
    Sub launchfilteredSKUlist()
        SplashRectangle.Visibility = Visibility.Visible
        SKUListView.Visibility = Visibility.Visible
        SKUListView.ItemsSource = Nothing
        SKUListView.Items.Clear()
        SKUListView.ItemsSource = AllProdLines(selectedindexofLine_temp).BrandCodesWeWant
        SKUListViewReset.Visibility = Visibility.Visible
        SKUListViewOK.Visibility = Visibility.Visible
        SKUListViewLabel.Visibility = Visibility.Visible
    End Sub

    Sub okskulistbox()
        HideMenu()
    End Sub

    Sub resetskulistbox()
        If IsSimulationMode Then Exit Sub
        HideMenu()
        FilterIcon_active.Visibility = Visibility.Hidden
        FilterIcon_inactive.Visibility = Visibility.Visible
        FilterOnOfflabel.Content = "Filter: OFF"
        AllProdLines(selectedindexofLine_temp).isFilterByBrandcode = False
        AllProdLines(selectedindexofLine_temp).BrandCodesWeWant.Clear()

        prStoryReport.reFilterData_ClearAllFilters()
        AllProdLines(selectedindexofLine_temp).rawDowntimeData.reFilterData_ClearAllFilters()

        MaxPRindataset = prStoryReport.UPDT
        MaxPRindataset_planned = prStoryReport.PDT

        MaxStopsindataset = Card_TopStops(0).SPD
        updateCard_TopStops()
        updateCard_Planned_Tier1()
        updateCard_Unplanned_Tier1()

        If stoplabel_1.IsVisible Then 'if we're looking at top stops, do this
            showstoplabels()
            Assignlabelnames(prStoryReport)
        ElseIf UPDTlabel1.IsVisible Then
            showDTpercentframe()
        End If

        If incontrolgreenbox.Visibility = Windows.Visibility.Visible Then
            GenerateIncontrolCharts(incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
        Else
            MasterDataSet = New inControlReport(AllProdLines(selectedindexofLine_temp), incontrolAnalysisSelectedStartdate, incontrolAnalysisSelectedEnddate)
        End If

        If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
            pr_label.Content = FormatPercent(prStoryReport.PR, 1) & " PR"
            PR_Label_Header.Content = "PR"
        Else
            pr_label.Content = FormatPercent(prStoryReport.PR, 1) & " Av."
            PR_Label_Header.Content = "Availability"
            cases_label.Visibility = Visibility.Hidden
            rateloss_label.Visibility = Visibility.Hidden
            Cases_Label_Header.Visibility = Visibility.Hidden
            Rateloss_Label_Header.Visibility = Visibility.Hidden
        End If

        If prStoryReport.CasesAdjusted = 0 Then
            cases_label.Content = prStoryReport.CasesActual ' & " cases"
        Else
            cases_label.Content = prStoryReport.CasesAdjusted ' & " cases"
        End If


        rateloss_label.Content = FormatPercent(prStoryReport.rateLoss, 1) ' & " rate and quality loss"
        Stops_Label.Content = Math.Round(prStoryReport.StopsPerDay, 0) ' & " stops/day"
        MTBF_Label.Content = Math.Round(prStoryReport.MTBF, 0) & " min"
    End Sub

    Private Sub BargraphreportwindowClose(ByVal sender As Object, ByVal e As CancelEventArgs)

        If bargraphreportwindow_Open = True Then

            If InStr(sender.ToString, "bargraphreportwindow", vbTextCompare) > 0 Then
                AllProdLines(selectedindexofLine_temp).rawDowntimeData.reFilterData_ClearAllFilters()
                Dim mainprstorywindow As New WindowMain_prstory
                Me.Owner.Visibility = Visibility.Visible
                SendUserAnalyticsDatatoServer()

                RateLossReport_Icon.Visibility = Visibility.Hidden
                RateLossReport_Label.Visibility = Visibility.Hidden
            End If

        End If

        bargraphreportwindow_Open = False

    End Sub
#Region "User Analytics"
    Private Sub SendUserAnalyticsDatatoServer()
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
                                                    .Add("Line", String.Format(linename_label.Content)) _
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
                                                    .Add("Z", "oldserver") _
                                                    .Add("Z1", PRSTORY_VERSION_NUMBER)


                col1.Insert(NewInfoBson)
                server.Disconnect()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub


    Private Sub SendExceptionDatatoServer(comments As String, functionname As String)
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
                col1 = db("ExLog")

                Dim currentloginname As String
                currentloginname = Environment.UserName
                Dim NewInfoBson As BsonDocument = New BsonDocument() _
                                                    .Add("who", String.Format(currentloginname)) _
                                                    .Add("when", String.Format(Now(), "MM dd yyyy hh:mm")) _
                                                    .Add("Line", String.Format(linename_label.Content)) _
                                                    .Add("comm", comments) _
                                                    .Add("fcn", functionname) _
                                                    .Add("starttime", String.Format(starttimeselected)) _
                                                    .Add("endtime", String.Format(endtimeselected)) _
                                                    .Add("AvMode", My.Settings.AdvancedSettings_isAvailabilityMode) _
                                                    .Add("SimMode", IsSimulationMode)


                col1.Insert(NewInfoBson)
                server.Disconnect()
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
#End Region


    Private Sub LaunchStopDetails()
        Select Case My.Settings.AdvancedSettings_MultiConstraintAnalysisMode

            Case MultiConstraintAnalysis.SingleConstraint
                MsgBox("The stops/day shown here are for constraint unit-op only, since the app is running on 'Single Constraint' mode. The PR hurt generated from rate loss stios are NOT being accounted for." & vbNewLine & vbNewLine & "If you want to change the mode, go back to Main screen > Settings > Advanced Settings > MultiConstraint Settings.")
            Case MultiConstraintAnalysis.NoRateLossStops
                MsgBox("The stops/day shown here are for constraint unit-op only, since the app is running on 'No Rate Loss Stops' mode. However, the PR hurt generated from rate loss stops are being accounted for in the respective unit ops." & vbNewLine & vbNewLine & "If you want to change the mode, go back to Main screen > Settings > Advanced Settings > MultiConstraint Settings.")
            Case MultiConstraintAnalysis.RateLossAsStops
                MsgBox("The stops/day shown here are for constraint unit-op and rate loss unit-ops (like dual Fillers), since the app is running on 'Rate Loss as Stops' mode. The PR hurt generated from rate loss stops are also being accounted for in the respective unit-ops." & vbNewLine & vbNewLine & "If you want to change the mode, go back to Main screen > Settings > Advanced Settings > MultiConstraint Settings.")
        End Select
    End Sub

    Private Sub Showdashboard1more()
        If MoreLabel_Area.Visibility = Windows.Visibility.Visible Then
            MoreLabel_Area.Visibility = Visibility.Hidden
            Moreiconopen.Visibility = Visibility.Visible
            Moreiconclose.Visibility = Visibility.Hidden
            Exit Sub
        End If


        MoreLabel_Area.Visibility = Visibility.Visible

        Dim mtdPRstring As String

        Try
            mtdPRstring = "MtD PR " & GetMtDPR() & "%"
        Catch
            mtdPRstring = ""
        End Try

        Dim rateLossString As String = ""
        If RateLossReport_Icon.Visibility = Visibility.Visible Then

            Dim includeList As New List(Of String)
            Dim filterField As DowntimeField = DowntimeField.Team
            Dim doweFilter As Boolean = False

            If FilterOnOfflabel.Content = "Filter: ON" Then 'we filtered something!
                doweFilter = True
                includeList = AllProdLines(selectedindexofLine_temp).BrandCodesWeWant
                Select Case ProductFilterComboBox.SelectedValue

                    Case "SKUs"
                        filterField = DowntimeField.Product
                    Case "Teams"

                    Case Else
                        Exit Sub
                End Select
            End If

            rawRateLossDataWindow = New Window_RateLossReport(AllProdLines(prStoryReport.ParentLineInt).Name, prStoryReport.schedTime, prStoryReport.StartDate, prStoryReport.EndDate, AllProdLines(prStoryReport.ParentLineInt).RawRateLossDataArray, doweFilter, includeList, filterField, selectedindexofLine_temp) 'RawDataWindow(sourcecard, targetEvent.RawRows, prStoryReport, targetEvent.Name) ', sourcecard, sourcefield, weStillWantToDoThis, prStoryReport)

            While Not rawRateLossDataWindow.isDone
                System.Threading.Thread.Sleep(100)
            End While

            If AllProdLines(selectedindexofLine_temp).parentSite.Name = SITE_CRUX Then
                rateLossString = ""
            Else
                rateLossString = Math.Round((rawRateLossDataWindow.netRateLossTime / prStoryReport.schedTime) * 100, 1) & " % rate loss"
            End If
        End If

        If My.Settings.AdvancedSettings_isAvailabilityMode = True Then
            MoreLabel_Area.Content = Math.Round(prStoryReport.schedTime, 0) & " mins sched. time" & vbNewLine & vbNewLine & prStoryReport.ActualStops & " actual stops" & vbNewLine & vbNewLine & Math.Round(prStoryReport.MainLEDSReport.DT_Report.MTTR, 2) & " min MTTR" & vbNewLine & vbNewLine & rateLossString '& vbNewLine & vbNewLine & cases_label.Content & vbNewLine & vbNewLine & rateloss_label.Content & vbNewLine & vbNewLine & prStoryReport.MSU & " MSU"
        Else
            MoreLabel_Area.Content = Math.Round(prStoryReport.schedTime, 0) & " mins sched. time" & vbNewLine & vbNewLine & cases_label.Content & " cases" & vbNewLine & vbNewLine & rateloss_label.Content & " rate/quality loss" & vbNewLine & vbNewLine & prStoryReport.MSU & " SU" & vbNewLine & vbNewLine & prStoryReport.ActualStops & " actual stops" & vbNewLine & vbNewLine & mtdPRstring & vbNewLine & vbNewLine & Math.Round(prStoryReport.MainLEDSReport.DT_Report.MTTR, 2) & " min MTTR" & vbNewLine & vbNewLine & rateLossString
        End If


        Moreiconopen.Visibility = Visibility.Hidden
        Moreiconclose.Visibility = Visibility.Visible
    End Sub


    Public Sub mapping_instaLaunch()
        LaunchMenu()
        ShowMappingSelection()
    End Sub

    Private Sub Hidedashboard1more()
        Showdashboard1more()

    End Sub

    Private Sub ExpandBarGraphReportWindow()

        WindowExpandIcon.Visibility = Visibility.Hidden
        WindowCollapseIcon.Visibility = Visibility.Visible
        Dim hx As Integer
        For hx = 0 To 100 Step 25
            Me.Height = Me.Height + 25
        Next

    End Sub

    Private Sub CollapseBarGraphReportWindow()
        WindowCollapseIcon.Visibility = Visibility.Hidden
        WindowExpandIcon.Visibility = Visibility.Visible
        Me.Height = 706
    End Sub
#Region "Notes"
    Private Sub ShowNotes()
        UseTrack_Notes = True
        NOtesLabel.Background = mybrushNOTESblue
        NotesBaseCanvas.Visibility = Visibility.Visible
        NotesSplash.Visibility = Visibility.Visible
        NOtesLabel_Header.Visibility = Visibility.Visible
        NotesClose.Visibility = Visibility.Visible
        NOtesLabel.Visibility = Visibility.Hidden
        NotesPlusIcon.Visibility = Visibility.Hidden


        If currentNotesfilename <> "" Then
            OpenNoteEditor(currentNotesfilename, False)
        Else
            NotesMenuMainCanvas.Visibility = Visibility.Visible
            NotesClose.Visibility = Visibility.Visible
            NotesMinimize.Visibility = Visibility.Visible
        End If


    End Sub
    Private Sub OpenNoteEditor(filenameselected As String, createmode As Boolean)

        IsNotesCreateMode = createmode
        If IsPickMode <> True Then ClearNotesforNew()

        NotesCanvas.Visibility = Visibility.Visible

        NOtesLabel.Visibility = Visibility.Hidden
        NotesPlusIcon.Visibility = Visibility.Hidden
        NotesSplash.Visibility = Visibility.Visible
        NOtesLabel_Header.Visibility = Visibility.Visible
        NotesUpload.Visibility = Visibility.Visible
        NotesClose.Visibility = Visibility.Visible
        NotesMinimize.Visibility = Visibility.Visible
        NotesScrollArea.Visibility = Visibility.Visible
        NotesIntroduction.Visibility = Visibility.Visible
        NotesExport.Visibility = Visibility.Visible
        SyncImage.Visibility = Visibility.Visible
        ExportExcelLabel.Visibility = Visibility.Visible
        Notesfilenamelabel.Visibility = Visibility.Visible
        NotesStartDate.Visibility = Visibility.Visible
        NotesEndDate.Visibility = Visibility.Visible
        NotesCreatorName.Visibility = Visibility.Visible
        NotesUploadAlert.Visibility = Visibility.Hidden

        Textbox1Header.Visibility = Visibility.Visible
        Textbox2Header.Visibility = Visibility.Visible
        Textbox3Header.Visibility = Visibility.Visible
        Textbox4Header.Visibility = Visibility.Visible
        Textbox5Header.Visibility = Visibility.Visible
        Textbox6Header.Visibility = Visibility.Visible
        Textbox7Header.Visibility = Visibility.Visible
        Textbox8Header.Visibility = Visibility.Visible

        Notesfilenamelabel.Content = "File name: " & filenameselected


        If createmode Then

            NotesStartDate.Content = "Analysis date range: " & starttimeselected & " to "
            NotesEndDate.Content = endtimeselected
            NotesCreatorName.Content = "Last modified by " & Environment.UserName


        End If






        If filenameselected = "" Or createmode = True Then Exit Sub
        SyncNoteswithCloud(filenameselected)

        If IsPickMode = False Then

            ReadCSVfilesforNotes(filenameselected)
        End If
    End Sub


    Private Sub ClearNotesforNew()

        NotesStartDate.Content = ""
        NotesEndDate.Content = ""
        NotesCreatorName.Content = ""





        Row1textbox7.Text = ""
        Row1textbox8.Text = ""
        Row1textbox1.Text = ""
        Row1textbox2.Text = ""
        Row1textbox3.Text = ""
        Row1textbox4.Text = ""
        Row1textbox5.Text = ""
        Row1textbox6.Text = ""

        Row2textbox7.Text = ""
        Row2textbox8.Text = ""
        Row2textbox1.Text = ""
        Row2textbox2.Text = ""
        Row2textbox3.Text = ""
        Row2textbox4.Text = ""
        Row2textbox5.Text = ""
        Row2textbox6.Text = ""

        Row3textbox7.Text = ""
        Row3textbox8.Text = ""
        Row3textbox1.Text = ""
        Row3textbox2.Text = ""
        Row3textbox3.Text = ""
        Row3textbox4.Text = ""
        Row3textbox5.Text = ""
        Row3textbox6.Text = ""

        Row4textbox7.Text = ""
        Row4textbox8.Text = ""
        Row4textbox1.Text = ""
        Row4textbox2.Text = ""
        Row4textbox3.Text = ""
        Row4textbox4.Text = ""
        Row4textbox5.Text = ""
        Row4textbox6.Text = ""

        Row5textbox7.Text = ""
        Row5textbox8.Text = ""
        Row5textbox1.Text = ""
        Row5textbox2.Text = ""
        Row5textbox3.Text = ""
        Row5textbox4.Text = ""
        Row5textbox5.Text = ""
        Row5textbox6.Text = ""

        Row6textbox7.Text = ""
        Row6textbox8.Text = ""
        Row6textbox1.Text = ""
        Row6textbox2.Text = ""
        Row6textbox3.Text = ""
        Row6textbox4.Text = ""
        Row6textbox5.Text = ""
        Row6textbox6.Text = ""

        Row7textbox7.Text = ""
        Row7textbox8.Text = ""
        Row7textbox1.Text = ""
        Row7textbox2.Text = ""
        Row7textbox3.Text = ""
        Row7textbox4.Text = ""
        Row7textbox5.Text = ""
        Row7textbox6.Text = ""

        Row8textbox7.Text = ""
        Row8textbox8.Text = ""
        Row8textbox1.Text = ""
        Row8textbox2.Text = ""
        Row8textbox3.Text = ""
        Row8textbox4.Text = ""
        Row8textbox5.Text = ""
        Row8textbox6.Text = ""

        Row9textbox7.Text = ""
        Row9textbox8.Text = ""
        Row9textbox1.Text = ""
        Row9textbox2.Text = ""
        Row9textbox3.Text = ""
        Row9textbox4.Text = ""
        Row9textbox5.Text = ""
        Row9textbox6.Text = ""

        Row10textbox7.Text = ""
        Row10textbox8.Text = ""
        Row10textbox1.Text = ""
        Row10textbox2.Text = ""
        Row10textbox3.Text = ""
        Row10textbox4.Text = ""
        Row10textbox5.Text = ""
        Row10textbox6.Text = ""
    End Sub


    Private Sub SyncNoteswithCloud(filenameselected As String)


        filedownload(filenameselected)
    End Sub
    Private Sub MinimizeNotesSplashbyUser()
        CloseNotesSplash()
        NotesUploadtoCloud()
        NOtesLabel.Background = mybrushNOTESORANGE
    End Sub

    Private Sub HideNotesSplashbyUser()
        CloseNotesSplash()
        NotesUploadtoCloud()
        currentNotesfilename = ""
        NOtesLabel.Background = mybrushNOTESblue
    End Sub

    Private Sub CloseNotesSplash()
        'NotesUploadtoCloud()

        If My.Settings.AdvancedSettings_UseNotes = True Then

            NOtesLabel.Visibility = Visibility.Visible
            NotesPlusIcon.Visibility = Visibility.Visible
        End If



        NotesSplash.Visibility = Visibility.Hidden
        NOtesLabel_Header.Visibility = Visibility.Hidden
        NotesUpload.Visibility = Visibility.Hidden
        NotesClose.Visibility = Visibility.Hidden
        NotesMinimize.Visibility = Visibility.Hidden
        NotesScrollArea.Visibility = Visibility.Hidden
        NotesIntroduction.Visibility = Visibility.Hidden
        NotesExport.Visibility = Visibility.Hidden
        SyncImage.Visibility = Visibility.Hidden
        ExportExcelLabel.Visibility = Visibility.Hidden
        NotesUploadAlert.Visibility = Visibility.Hidden
        NotesMenuMainCanvas.Visibility = Visibility.Hidden
        Notesfilenamelabel.Visibility = Visibility.Hidden
        NotesStartDate.Visibility = Visibility.Hidden
        NotesEndDate.Visibility = Visibility.Hidden
        NotesCreatorName.Visibility = Visibility.Hidden
        NotesMenuAuxCanvas_OPEN_ExistingNotes.Visibility = Visibility.Hidden

        Textbox1Header.Visibility = Visibility.Hidden
        Textbox2Header.Visibility = Visibility.Hidden
        Textbox3Header.Visibility = Visibility.Hidden
        Textbox4Header.Visibility = Visibility.Hidden
        Textbox5Header.Visibility = Visibility.Hidden
        Textbox6Header.Visibility = Visibility.Hidden
        Textbox7Header.Visibility = Visibility.Hidden
        Textbox8Header.Visibility = Visibility.Hidden

        NotesMenuAuxCanvas_OPEN_ExistingNotes.Visibility = Visibility.Hidden
        NotesMenuMainCanvas.Visibility = Visibility.Hidden
    End Sub

    Private Sub NotesUploadtoCloud()

        Dim notesfilename As String
        Dim starttimestring As String

        starttimestring = Month(Now()) & "_" & Day(Now()) & "_" & Year(Now()) & "_" & Hour(Now()) & "_" & Minute(Now())

        If IsNotesCreateMode Then
            notesfilename = linename_label.Content & "_" & starttimestring & ".csv"
        Else
            notesfilename = currentNotesfilename

        End If
        If notesfilename <> "" Then
            Try

                CreateCSVfileforNotes(notesfilename)

                If CheckIfFtpFileExists("ftp://prstory.pg.com/log.txt", "normalusers", "pgdigitalfactory412") Then
                    fileupload(notesfilename)
                    currentNotesfilename = notesfilename
                    Notesfilenamelabel.Content = "File name: " & currentNotesfilename
                Else

                    If System.IO.File.Exists(PATH_PRSTORY_TARGETS & notesfilename) Then
                        System.IO.File.Copy(PATH_PRSTORY_TARGETS & notesfilename, "C:\Users\" & Environment.UserName & "\Desktop\" & notesfilename, True)
                    End If
                    MsgBox("We tried to upload, but the cloud is nowhere to be found.  Please try again later." & vbNewLine & "Also, we just saved a local copy of the file to your desktop." & vbNewLine & "File name: " & notesfilename & vbNewLine & "You can email the file without changing the file name or contents to  das.l@pg.com and odle.sr@pg.com, and the file will be uploaded to the cloud.")
                End If
            Catch
                If System.IO.File.Exists(PATH_PRSTORY_TARGETS & notesfilename) Then
                    System.IO.File.Copy(PATH_PRSTORY_TARGETS & notesfilename, "C:\Users\" & Environment.UserName & "\Desktop\" & notesfilename, True)

                End If
                MsgBox("We tried to upload, but the cloud is nowhere to be found.  Please try again later." & vbNewLine & "Also, we just saved a local copy of the file to your desktop." & vbNewLine & "File name: " & notesfilename & vbNewLine & "You can email the file without changing the file name or contents to  das.l@pg.com and odle.sr@pg.com, and the file will be uploaded to the cloud.")
            End Try
        End If

    End Sub

    Private Sub CreateNewNote()
        NotesMenuMainCanvas.Visibility = Visibility.Hidden
        OpenNoteEditor("", True)
    End Sub
    Private Sub OpenExistingNote()
        NotesMenuMainCanvas.Visibility = Visibility.Hidden
        NotesMenuAuxCanvas_OPEN_ExistingNotes.Visibility = Visibility.Visible

    End Sub
    Private Sub SearchForExistingNoteson_SelectedDate()
        Dim selectedday As Integer = Day(ExistingNotesDatePicker.SelectedDate)
        Dim selectedmonth As Integer = Month(ExistingNotesDatePicker.SelectedDate)
        Dim selectedyear As Integer = Year(ExistingNotesDatePicker.SelectedDate)
        Dim uripath As String = "ftp://prstory.pg.com/prstory_notes/"
        ExistingNotesList.ItemsSource = Nothing
        Dim ftpRequest As FtpWebRequest = Nothing
        Dim ftpResponse As FtpWebResponse = Nothing
        Dim sline As String = ""
        Dim strReader As StreamReader = Nothing
        Dim filelist As New List(Of String)
        Dim selectedfilelist As New List(Of String)

        Try
            ftpRequest = CType(WebRequest.Create(uripath), FtpWebRequest)

            With ftpRequest
                .Credentials = New NetworkCredential("normalusers", "pgdigitalfactory412")
                .Method = WebRequestMethods.Ftp.ListDirectory
            End With

            ftpResponse = CType(ftpRequest.GetResponse, FtpWebResponse)

            strReader = New StreamReader(ftpResponse.GetResponseStream)

            If strReader IsNot Nothing Then sline = strReader.ReadLine

            While sline IsNot Nothing
                filelist.Add(sline)
                sline = strReader.ReadLine
            End While
        Catch
        Finally
            If ftpResponse IsNot Nothing Then
                ftpResponse.Close()
                ftpResponse = Nothing
            End If

            If strReader IsNot Nothing Then
                strReader.Close()
                strReader = Nothing
            End If



        End Try
        If Not IsNothing(filelist) Then
            For Each item In filelist
                If InStr(item, linename_label.Content) > 0 Then
                    If InStr(item, selectedmonth & "_" & selectedday & "_" & selectedyear) > 0 Then
                        selectedfilelist.Add(item)
                    End If
                End If
            Next



            ExistingNotesList.ItemsSource = selectedfilelist
            ExistingNotesList.SelectedIndex = 0
        End If






    End Sub

    Private Sub SelectExistingNote()
        If ExistingNotesList.SelectedIndex > -1 Then
            NotesMenuAuxCanvas_OPEN_ExistingNotes.Visibility = Visibility.Hidden
            currentNotesfilename = ExistingNotesList.SelectedValue
            OpenNoteEditor(ExistingNotesList.SelectedValue, False)

        End If
    End Sub
    Private Sub CancelExistingNote()

        NotesMenuAuxCanvas_OPEN_ExistingNotes.Visibility = Visibility.Hidden
        NotesMenuMainCanvas.Visibility = Visibility.Visible

        NotesClose.Visibility = Visibility.Visible

    End Sub
    Private Sub filedownload(notesfilename As String)
        Try
            If CheckIfFtpFileExists("ftp://prstory.pg.com/prstory_notes/" & notesfilename, "normalusers", "pgdigitalfactory412") Then
                Dim sourcewebaddress As Uri = New Uri("http://prstory.pg.com/prstory_notes/" & notesfilename)

                Dim destinationfolderaddress As String = "C:/Users/Public/prstory/targets/" & notesfilename

                Dim myWebClient As New WebClient()
                My.Computer.Network.DownloadFile(sourcewebaddress, destinationfolderaddress, "", "", False, 10000, True)
            Else

            End If
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub
    Private Sub fileupload(notesfilename As String)
        Dim myWebClient As New WebClient()
        Dim sourcewebaddress As Uri = New Uri("http://prstory.pg.com/prstory_notes/" & notesfilename)
        Dim destinationfolderaddress As String = "C:/Users/Public/prstory/targets/" & notesfilename
        Try
            Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create("ftp://prstory.pg.com/prstory_notes/" & notesfilename), System.Net.FtpWebRequest)
            request.Credentials = New System.Net.NetworkCredential("normalusers", "pgdigitalfactory412")
            request.Method = System.Net.WebRequestMethods.Ftp.UploadFile

            Dim files() As Byte = System.IO.File.ReadAllBytes(destinationfolderaddress)

            Dim strz As System.IO.Stream = request.GetRequestStream()
            strz.Write(files, 0, files.Length)
            strz.Close()
            strz.Dispose()
        Catch ex As Exception
            Exit Sub
        End Try
        'CreateObject("WScript.Shell").Popup("Synced with cloud", 1, "prstory NOTES")
    End Sub



    Private Sub ReadCSVfilesforNotes(notesfilename As String)
        Dim TextLine As String
        Dim rowcount As Integer = 0
        Dim linecontent(0 To 7) As String

        If System.IO.File.Exists(PATH_PRSTORY_TARGETS & notesfilename) Then

            Dim objReader As New System.IO.StreamReader(PATH_PRSTORY_TARGETS & notesfilename)
            Do While objReader.Peek() <> -1
                TextLine = objReader.ReadLine()
                If InStr(TextLine, "~-~", vbTextCompare) > 0 Then
                    linecontent = Split(TextLine, "~-~")
                Else
                    GoTo skiptonextline
                End If


                Select Case rowcount

                    Case 0
                        Textbox7Header.Text = linecontent(0)
                        Textbox8Header.Text = linecontent(1)
                        Textbox1Header.Text = linecontent(2)
                        Textbox2Header.Text = linecontent(3)
                        Textbox3Header.Text = linecontent(4)
                        Textbox4Header.Text = linecontent(5)
                        Textbox5Header.Text = linecontent(6)
                        Textbox6Header.Text = linecontent(7)


                    Case 1
                        Row1textbox7.Text = linecontent(0)
                        Row1textbox8.Text = linecontent(1)
                        Row1textbox1.Text = linecontent(2)
                        Row1textbox2.Text = linecontent(3)
                        Row1textbox3.Text = linecontent(4)
                        Row1textbox4.Text = linecontent(5)
                        Row1textbox5.Text = linecontent(6)
                        Row1textbox6.Text = linecontent(7)
                    Case 2
                        Row2textbox7.Text = linecontent(0)
                        Row2textbox8.Text = linecontent(1)
                        Row2textbox1.Text = linecontent(2)
                        Row2textbox2.Text = linecontent(3)
                        Row2textbox3.Text = linecontent(4)
                        Row2textbox4.Text = linecontent(5)
                        Row2textbox5.Text = linecontent(6)
                        Row2textbox6.Text = linecontent(7)
                    Case 3
                        Row3textbox7.Text = linecontent(0)
                        Row3textbox8.Text = linecontent(1)
                        Row3textbox1.Text = linecontent(2)
                        Row3textbox2.Text = linecontent(3)
                        Row3textbox3.Text = linecontent(4)
                        Row3textbox4.Text = linecontent(5)
                        Row3textbox5.Text = linecontent(6)
                        Row3textbox6.Text = linecontent(7)
                    Case 4
                        Row4textbox7.Text = linecontent(0)
                        Row4textbox8.Text = linecontent(1)
                        Row4textbox1.Text = linecontent(2)
                        Row4textbox2.Text = linecontent(3)
                        Row4textbox3.Text = linecontent(4)
                        Row4textbox4.Text = linecontent(5)
                        Row4textbox5.Text = linecontent(6)
                        Row4textbox6.Text = linecontent(7)
                    Case 5
                        Row5textbox7.Text = linecontent(0)
                        Row5textbox8.Text = linecontent(1)
                        Row5textbox1.Text = linecontent(2)
                        Row5textbox2.Text = linecontent(3)
                        Row5textbox3.Text = linecontent(4)
                        Row5textbox4.Text = linecontent(5)
                        Row5textbox5.Text = linecontent(6)
                        Row5textbox6.Text = linecontent(7)
                    Case 6
                        Row6textbox7.Text = linecontent(0)
                        Row6textbox8.Text = linecontent(1)
                        Row6textbox1.Text = linecontent(2)
                        Row6textbox2.Text = linecontent(3)
                        Row6textbox3.Text = linecontent(4)
                        Row6textbox4.Text = linecontent(5)
                        Row6textbox5.Text = linecontent(6)
                        Row6textbox6.Text = linecontent(7)
                    Case 7
                        Row7textbox7.Text = linecontent(0)
                        Row7textbox8.Text = linecontent(1)
                        Row7textbox1.Text = linecontent(2)
                        Row7textbox2.Text = linecontent(3)
                        Row7textbox3.Text = linecontent(4)
                        Row7textbox4.Text = linecontent(5)
                        Row7textbox5.Text = linecontent(6)
                        Row7textbox6.Text = linecontent(7)
                    Case 8
                        Row8textbox7.Text = linecontent(0)
                        Row8textbox8.Text = linecontent(1)
                        Row8textbox1.Text = linecontent(2)
                        Row8textbox2.Text = linecontent(3)
                        Row8textbox3.Text = linecontent(4)
                        Row8textbox4.Text = linecontent(5)
                        Row8textbox5.Text = linecontent(6)
                        Row8textbox6.Text = linecontent(7)
                    Case 9
                        Row9textbox7.Text = linecontent(0)
                        Row9textbox8.Text = linecontent(1)
                        Row9textbox1.Text = linecontent(2)
                        Row9textbox2.Text = linecontent(3)
                        Row9textbox3.Text = linecontent(4)
                        Row9textbox4.Text = linecontent(5)
                        Row9textbox5.Text = linecontent(6)
                        Row9textbox6.Text = linecontent(7)
                    Case 10
                        Row10textbox7.Text = linecontent(0)
                        Row10textbox8.Text = linecontent(1)
                        Row10textbox1.Text = linecontent(2)
                        Row10textbox2.Text = linecontent(3)
                        Row10textbox3.Text = linecontent(4)
                        Row10textbox4.Text = linecontent(5)
                        Row10textbox5.Text = linecontent(6)
                        Row10textbox6.Text = linecontent(7)


                    Case 12
                        NotesCreatorName.Content = "Last modified by " & linecontent(0)
                    Case 13
                        NotesStartDate.Content = linecontent(0)
                    Case 14
                        NotesEndDate.Content = linecontent(0)
                End Select



skiptonextline:
                rowcount += 1
            Loop
        End If
    End Sub

    Private Sub CreateCSVfileforNotes(notesfilename As String)


        Dim fsT As Object
        Dim tmpWriteString As String


        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object



        tmpWriteString = Textbox7Header.Text & "~-~" & Textbox8Header.Text & "~-~" & Textbox1Header.Text & "~-~" & Textbox2Header.Text & "~-~" & Textbox3Header.Text & "~-~" & Textbox4Header.Text & "~-~" & Textbox5Header.Text & "~-~" & Textbox6Header.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row1textbox7.Text & "~-~" & Row1textbox8.Text & "~-~" & Row1textbox1.Text & "~-~" & Row1textbox2.Text & "~-~" & Row1textbox3.Text & "~-~" & Row1textbox4.Text & "~-~" & Row1textbox5.Text & "~-~" & Row1textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row2textbox7.Text & "~-~" & Row2textbox8.Text & "~-~" & Row2textbox1.Text & "~-~" & Row2textbox2.Text & "~-~" & Row2textbox3.Text & "~-~" & Row2textbox4.Text & "~-~" & Row2textbox5.Text & "~-~" & Row2textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row3textbox7.Text & "~-~" & Row3textbox8.Text & "~-~" & Row3textbox1.Text & "~-~" & Row3textbox2.Text & "~-~" & Row3textbox3.Text & "~-~" & Row3textbox4.Text & "~-~" & Row3textbox5.Text & "~-~" & Row3textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row4textbox7.Text & "~-~" & Row4textbox8.Text & "~-~" & Row4textbox1.Text & "~-~" & Row4textbox2.Text & "~-~" & Row4textbox3.Text & "~-~" & Row4textbox4.Text & "~-~" & Row4textbox5.Text & "~-~" & Row4textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row5textbox7.Text & "~-~" & Row5textbox8.Text & "~-~" & Row5textbox1.Text & "~-~" & Row5textbox2.Text & "~-~" & Row5textbox3.Text & "~-~" & Row5textbox4.Text & "~-~" & Row5textbox5.Text & "~-~" & Row5textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row6textbox7.Text & "~-~" & Row6textbox8.Text & "~-~" & Row6textbox1.Text & "~-~" & Row6textbox2.Text & "~-~" & Row6textbox3.Text & "~-~" & Row6textbox4.Text & "~-~" & Row6textbox5.Text & "~-~" & Row6textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row7textbox7.Text & "~-~" & Row7textbox8.Text & "~-~" & Row7textbox1.Text & "~-~" & Row7textbox2.Text & "~-~" & Row7textbox3.Text & "~-~" & Row7textbox4.Text & "~-~" & Row7textbox5.Text & "~-~" & Row7textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row8textbox7.Text & "~-~" & Row8textbox8.Text & "~-~" & Row8textbox1.Text & "~-~" & Row8textbox2.Text & "~-~" & Row8textbox3.Text & "~-~" & Row8textbox4.Text & "~-~" & Row8textbox5.Text & "~-~" & Row8textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row9textbox7.Text & "~-~" & Row9textbox8.Text & "~-~" & Row9textbox1.Text & "~-~" & Row9textbox2.Text & "~-~" & Row9textbox3.Text & "~-~" & Row9textbox4.Text & "~-~" & Row9textbox5.Text & "~-~" & Row9textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Row10textbox7.Text & "~-~" & Row10textbox8.Text & "~-~" & Row10textbox1.Text & "~-~" & Row10textbox2.Text & "~-~" & Row10textbox3.Text & "~-~" & Row10textbox4.Text & "~-~" & Row10textbox5.Text & "~-~" & Row10textbox6.Text
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = Now.ToString
        fsT.writetext(tmpWriteString & vbCrLf)


        Try
            tmpWriteString = Environment.UserName & "~-~"
            fsT.writetext(tmpWriteString & vbCrLf)
        Catch ex As Exception

        End Try


        tmpWriteString = NotesStartDate.Content & "~-~"
        fsT.writetext(tmpWriteString & vbCrLf)

        tmpWriteString = NotesEndDate.Content & "~-~"
        fsT.writetext(tmpWriteString & vbCrLf)


        'fin
        Try
            fsT.SaveToFile(PATH_PRSTORY_TARGETS & notesfilename, 2) 'Save binary data To disk
        Catch ex1 As Exception
            System.Threading.Thread.Sleep(1000)
            Try
                fsT.SaveToFile(PATH_PRSTORY_TARGETS & notesfilename, 2) 'Save binary data To disk
            Catch ex2 As Exception
                System.Threading.Thread.Sleep(1000)
                Try
                    fsT.SaveToFile(PATH_PRSTORY_TARGETS & notesfilename, 2) 'Save binary data To disk
                Catch ex3 As Exception
                    'MsgBox("Error Saving Notes. Please try again.", vbExclamation)
                    ' SendExceptionDatatoServer("", "CreateCSVfileforNotes")
                End Try

            End Try
        End Try

        fsT = Nothing


    End Sub
    Private Sub ExportNotestoExcel()
        UseTrack_Notes_ExporttoExcel = True
        'HideNotesSplashbyUser()
        'HideNotesSplashbyUser()
        'TakeScreenShotofMyself()
        Dim appXL As Excel.Application
        Dim wbXl As Excel.Workbook
        Dim shXL As Excel.Worksheet
        Dim raXL As Excel.Range
        ' Start Excel and get Application object.
        appXL = CreateObject("Excel.Application")
        appXL.Visible = True
        ' Add a new workbook.
        wbXl = appXL.Workbooks.Add
        shXL = wbXl.ActiveSheet
        ' Add table headers going cell by cell.
        Dim HeaderLine(0 To 8) As String
        Dim Textline(10, 8) As String

        HeaderLine(2) = Textbox1Header.Text
        HeaderLine(3) = Textbox2Header.Text
        HeaderLine(4) = Textbox3Header.Text
        HeaderLine(5) = Textbox4Header.Text
        HeaderLine(6) = Textbox5Header.Text
        HeaderLine(7) = Textbox6Header.Text
        HeaderLine(0) = Textbox7Header.Text
        HeaderLine(1) = Textbox8Header.Text


        Textline(0, 2) = Row1textbox1.Text
        Textline(0, 3) = Row1textbox2.Text
        Textline(0, 4) = Row1textbox3.Text
        Textline(0, 5) = Row1textbox4.Text
        Textline(0, 6) = Row1textbox5.Text
        Textline(0, 7) = Row1textbox6.Text
        Textline(0, 0) = Row1textbox7.Text
        Textline(0, 1) = Row1textbox8.Text

        Textline(1, 2) = Row2textbox1.Text
        Textline(1, 3) = Row2textbox2.Text
        Textline(1, 4) = Row2textbox3.Text
        Textline(1, 5) = Row2textbox4.Text
        Textline(1, 6) = Row2textbox5.Text
        Textline(1, 7) = Row2textbox6.Text
        Textline(1, 0) = Row2textbox7.Text
        Textline(1, 1) = Row2textbox8.Text

        Textline(2, 2) = Row3textbox1.Text
        Textline(2, 3) = Row3textbox2.Text
        Textline(2, 4) = Row3textbox3.Text
        Textline(2, 5) = Row3textbox4.Text
        Textline(2, 6) = Row3textbox5.Text
        Textline(2, 7) = Row3textbox6.Text
        Textline(2, 0) = Row3textbox7.Text
        Textline(2, 1) = Row3textbox8.Text


        Textline(3, 2) = Row4textbox1.Text
        Textline(3, 3) = Row4textbox2.Text
        Textline(3, 4) = Row4textbox3.Text
        Textline(3, 5) = Row4textbox4.Text
        Textline(3, 6) = Row4textbox5.Text
        Textline(3, 7) = Row4textbox6.Text
        Textline(3, 0) = Row4textbox7.Text
        Textline(3, 1) = Row4textbox8.Text


        Textline(4, 2) = Row5textbox1.Text
        Textline(4, 3) = Row5textbox2.Text
        Textline(4, 4) = Row5textbox3.Text
        Textline(4, 5) = Row5textbox4.Text
        Textline(4, 6) = Row5textbox5.Text
        Textline(4, 7) = Row5textbox6.Text
        Textline(4, 0) = Row5textbox7.Text
        Textline(4, 1) = Row5textbox8.Text


        Textline(5, 2) = Row6textbox1.Text
        Textline(5, 3) = Row6textbox2.Text
        Textline(5, 4) = Row6textbox3.Text
        Textline(5, 5) = Row6textbox4.Text
        Textline(5, 6) = Row6textbox5.Text
        Textline(5, 7) = Row6textbox6.Text
        Textline(5, 0) = Row6textbox7.Text
        Textline(5, 1) = Row6textbox8.Text


        Textline(6, 2) = Row7textbox1.Text
        Textline(6, 3) = Row7textbox2.Text
        Textline(6, 4) = Row7textbox3.Text
        Textline(6, 5) = Row7textbox4.Text
        Textline(6, 6) = Row7textbox5.Text
        Textline(6, 7) = Row7textbox6.Text
        Textline(6, 0) = Row7textbox7.Text
        Textline(6, 1) = Row7textbox8.Text



        Textline(7, 2) = Row8textbox1.Text
        Textline(7, 3) = Row8textbox2.Text
        Textline(7, 4) = Row8textbox3.Text
        Textline(7, 5) = Row8textbox4.Text
        Textline(7, 6) = Row8textbox5.Text
        Textline(7, 7) = Row8textbox6.Text
        Textline(7, 0) = Row8textbox7.Text
        Textline(7, 1) = Row8textbox8.Text


        Textline(8, 2) = Row9textbox1.Text
        Textline(8, 3) = Row9textbox2.Text
        Textline(8, 4) = Row9textbox3.Text
        Textline(8, 5) = Row9textbox4.Text
        Textline(8, 6) = Row9textbox5.Text
        Textline(8, 7) = Row9textbox6.Text
        Textline(8, 0) = Row9textbox7.Text
        Textline(8, 1) = Row9textbox8.Text


        Textline(9, 2) = Row10textbox1.Text
        Textline(9, 3) = Row10textbox2.Text
        Textline(9, 4) = Row10textbox3.Text
        Textline(9, 5) = Row10textbox4.Text
        Textline(9, 6) = Row10textbox5.Text
        Textline(9, 7) = Row10textbox6.Text
        Textline(9, 0) = Row10textbox7.Text
        Textline(9, 1) = Row10textbox8.Text

        shXL.Range("A1").Value = "prstory NOTES"
        shXL.Range("A2").Value = "Safety"
        shXL.Range("A3").Value = "Quality"
        shXL.Range("A4").Value = "CIL"
        shXL.Range("A6").Value = "Deviations"
        shXL.Range("A8").Value = "PR"
        shXL.Range("A9").Value = "UPDT"
        shXL.Range("A10").Value = "PDT"
        shXL.Range("A11").Value = "Stops"
        shXL.Range("A12").Value = "Stops per day"
        shXL.Range("A13").Value = "MtD PR"
        shXL.Range("A14").Value = "Defects Found"
        shXL.Range("A15").Value = "Defects Fixed"

        shXL.Range("B8").Value = FormatPercent(prStoryReport.PR, 1)
        shXL.Range("B11").Value = Math.Round(prStoryReport.ActualStops)
        shXL.Range("B9").Value = UPDT_datalabel1.Content
        shXL.Range("B10").Value = PDT_datalabel1.Content
        shXL.Range("B13").Value = GetMtDPR() & "%"
        shXL.Range("B12").Value = Math.Round(prStoryReport.StopsPerDay, 1)



        shXL.Range("B1").Value = linename_label.Content

        Dim starttimestring As String
        Dim endtimestring As String

        starttimestring = Month(starttimeselected) & "_" & Day(starttimeselected) & "_" & Year(starttimeselected) & "_" & Hour(starttimeselected) & "_" & Minute(starttimeselected)
        endtimestring = Month(endtimeselected) & "_" & Day(endtimeselected) & "_" & Year(endtimeselected) & "_" & Hour(endtimeselected) & "_" & Minute(endtimeselected)
        shXL.Range("D1").Value = starttimeselected & " to " & endtimeselected

        shXL.Range("A16", "H16").Value = HeaderLine
        shXL.Range("A17", "H26").Value = Textline

        With shXL.Range("A2", "A15")
            .Font.Bold = True
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            .Font.Size = 13
            .Font.Color = System.Drawing.Color.DarkSlateGray
            .Font.FontStyle = "Sans Serif"
        End With

        'With shXL.Range("A1", "D5")
        ' .Font.Bold = True
        ' .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        ' .Font.Size = 10
        ' .Font.Color = System.Drawing.Color.DarkSlateGray
        ' .Font.FontStyle = "Sans Serif"
        ' End With

        With shXL.Range("A16", "H16")
            .Font.Bold = True
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            .Font.Size = 15
            .Font.Color = System.Drawing.Color.Blue
            .Font.FontStyle = "Sans Serif"
        End With

        With shXL.Range("A17", "H26")
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            .Font.Size = 13
            .Font.Color = System.Drawing.Color.DarkSlateGray
            .Font.FontStyle = "Sans Serif"
        End With
        raXL = shXL.Range("A13", "B13")
        raXL.ColumnWidth = 20
        raXL = shXL.Range("C13", "H13")
        raXL.ColumnWidth = 25
        raXL.WrapText = True
        raXL = shXL.Range("A16", "H26")
        raXL.WrapText = True
        raXL.EntireRow.AutoFit()
        'raXL = DirectCast(shXL.Range("A13"), Excel.Range)


        ' shXL.Shapes.AddPicture(PATH_PRSTORY_TARGETS & "tmp.bmp", Microsoft.Office.Core.MsoTriState.msoFalse, _
        '       Microsoft.Office.Core.MsoTriState.msoCTrue, shXL.Range("A13").Left, shXL.Range("A13").Top, 800, 451)




        appXL.Visible = True
        appXL.UserControl = True
        ' Release object references.
        raXL = Nothing
        shXL = Nothing
        wbXl = Nothing
        'appXL.Quit()
        appXL = Nothing

        My.Computer.Clipboard.Clear()
        Exit Sub
Err_Handler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    End Sub

    Private Sub TakeScreenShotofMyself()

        Dim bounds As Rectangle
        Dim screenshot As System.Drawing.Bitmap
        'Dim screenshot As New Bitmap(Me.Width, Me.Height)
        'Dim sc As New Bitmap(Me.Width, Me.Height)
        Dim graph As Graphics
        ' bounds = My.Computer.Screen.Bounds
        bounds.Size = New System.Drawing.Size(1220, 668)

        screenshot = New System.Drawing.Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb)
        graph = Graphics.FromImage(screenshot)
        graph.CopyFromScreen(Me.Left + 20, Me.Top + 33, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy)
        screenshot.Save(PATH_PRSTORY_TARGETS & "tmp.bmp", Imaging.ImageFormat.Bmp)

    End Sub
#End Region
#Region "PickaLoss"

    Private Sub PickaLossInitialize(sender As Object, e As MouseButtonEventArgs)
        Dim whichrow As Integer
        whichrow = onlyDigits(sender.name)

        If IsPickMode = False Then
            IsPickMode = True
            ManageStatusBar("PickModeOn")
            sender.opacity = 0.5
            WhichPickRow = whichrow
            Cursor = Cursors.Pen
            MinimizeNotesSplashbyUser()
        Else
            IsPickMode = False
            ManageStatusBar("Normal")
            Cursor = Cursors.Arrow
            sender.opacity = 1
        End If
    End Sub
    Private Sub PickaLoss_CollectInfofromLabel(clickedlabel As Object, cardnumber As Integer)
        Dim lossname As String, prloss As Double, stops As Integer, stopsperday As Double, DTmin As Double
        Dim tmpDTeventforPickaLoss As DTevent
        Dim tmpDTeventforPickaLoss_incontrol As inControlDTevent
        If cardnumber = 51 Then
        ElseIf cardnumber = 31 Then
            If clickedlabel.text.ToString = "" Then Exit Sub
        Else
            If clickedlabel.content.ToString = "" Then Exit Sub
        End If

        Select Case cardnumber

            Case 1  ' UPDT tier 1
                lossname = clickedlabel.content.ToString
                tmpDTeventforPickaLoss = Card_Unplanned_T1(onlyDigits(clickedlabel.name) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)
            Case 2 ' PDT tier 1
                lossname = clickedlabel.content.ToString
                tmpDTeventforPickaLoss = Card_Planned_T1(onlyDigits(clickedlabel.name) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)
            Case 3 ' UPDT tier 2
                lossname = clickedlabel.content.ToString
                tmpDTeventforPickaLoss = Card_Unplanned_T2(onlyDigits(clickedlabel.name) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)
            Case 4 ' UPDT tier 3A
                lossname = clickedlabel.content.ToString
                tmpDTeventforPickaLoss = Card_Unplanned_T3A(onlyDigits(Strings.Mid(clickedlabel.name, 7)) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)

            Case 5  ' UPDT tier 3 B
                lossname = clickedlabel.content.ToString
                tmpDTeventforPickaLoss = Card_Unplanned_T3B(onlyDigits(Strings.Mid(clickedlabel.name, 7)) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)

            Case 6  ' UPDT tier 3 C

                lossname = clickedlabel.content.ToString
                tmpDTeventforPickaLoss = Card_Unplanned_T3C(onlyDigits(Strings.Mid(clickedlabel.name, 7)) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)

            Case 41 ' PDT tier 2
                lossname = clickedlabel.content.ToString
                tmpDTeventforPickaLoss = Card_Planned_T2(onlyDigits(clickedlabel.name) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)

            Case 31 ' top stops 15
                lossname = clickedlabel.text.ToString  ' text block does not have content, it has text
                tmpDTeventforPickaLoss = Card_TopStops(onlyDigits(clickedlabel.name) - 1)
                prloss = tmpDTeventforPickaLoss.DTpct
                DTmin = Math.Round(tmpDTeventforPickaLoss.DT)
                stops = tmpDTeventforPickaLoss.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)

            Case 51 ' incontrol

                tmpDTeventforPickaLoss_incontrol = MasterDataSet.inControlEvents(onlyDigits(clickedlabel.name) - 1)
                lossname = tmpDTeventforPickaLoss_incontrol.Name    ' bubbles do not have content, hence we get their names from the array
                prloss = tmpDTeventforPickaLoss_incontrol.DTpct
                'DTmin = Math.Round(tmpDTeventforPickaLoss_incontrol.DT)
                stops = tmpDTeventforPickaLoss_incontrol.Stops
                stopsperday = Math.Round(tmpDTeventforPickaLoss_incontrol.SPD, 1)
                PickaLoss_TransferInfotoNotes(lossname, prloss, stops, stopsperday, DTmin)

        End Select

    End Sub
    Private Sub PickaLoss_TransferInfotoNotes(lossname As String, prloss As Double, stops As Integer, Optional stopsperday As Double = 0.0, Optional DTmin As Double = 0.0)
        Cursor = Cursors.Arrow

        Select Case WhichPickRow
            Case 1
                Row1textbox7.Text = stops
                Row1textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row1textbox1.Text = lossname
                PickRow1.Opacity = 1
            Case 2
                Row2textbox7.Text = stops
                Row2textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row2textbox1.Text = lossname
                PickRow2.Opacity = 1
            Case 3
                Row3textbox7.Text = stops
                Row3textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row3textbox1.Text = lossname
                PickRow3.Opacity = 1
            Case 4
                Row4textbox7.Text = stops
                Row4textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row4textbox1.Text = lossname
                PickRow4.Opacity = 1
            Case 5
                Row5textbox7.Text = stops
                Row5textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row5textbox1.Text = lossname
                PickRow5.Opacity = 1
            Case 6
                Row6textbox7.Text = stops
                Row6textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row6textbox1.Text = lossname
                PickRow6.Opacity = 1
            Case 7
                Row7textbox7.Text = stops
                Row7textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row7textbox1.Text = lossname
                PickRow7.Opacity = 1
            Case 8
                Row8textbox7.Text = stops
                Row8textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row8textbox1.Text = lossname
                PickRow8.Opacity = 1
            Case 9
                Row9textbox7.Text = stops
                Row9textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row9textbox1.Text = lossname
                PickRow9.Opacity = 1
            Case 10
                Row10textbox7.Text = stops
                Row10textbox8.Text = FormatPercent(prloss, 1) & ", " & DTmin & "min"
                Row10textbox1.Text = lossname
                PickRow10.Opacity = 1

        End Select
        UseTrack_Notes_PickaLoss = True
        NotesUploadtoCloud()
        ShowNotes()
        IsPickMode = False
        ManageStatusBar("Normal")

    End Sub

#End Region
#Region "Status_and_Mode"
    Private Sub ManageStatusBar(mode As String)
        Select Case mode

            Case "Normal"
                CurrentStatusLabel.Content = "Normal Mode"
                StatusBar.Fill = mybrushdarkgray
                IsSimulationMode = False
            Case "PickModeOn"
                CurrentStatusLabel.Content = "Normal Mode - Pick a Loss is ON - Click on any failure mode / unit-op's gray label to pick"
                StatusBar.Fill = mybrushbrightblue
            Case "Simulation"
                CurrentStatusLabel.Content = "Simulation Mode"
                StatusBar.Fill = mybrushbrightorange
            Case Else
                CurrentStatusLabel.Content = "Normal Mode"
                StatusBar.Fill = mybrushdarkgray
        End Select



    End Sub
    Private Sub ChangeMode()
        Select Case CurrentStatusLabel.Content

            Case "Normal Mode"
                ManageStatusBar("Simulation")
                InitializeSimulationMode()
                UseTrack_Simulation = True
            Case "Simulation Mode"
                ManageStatusBar("Normal")
                BringbacktoNormal()

            Case "Normal Mode - Pick a Loss is ON - Click on any failure mode / unit-op's gray label to pick"


            Case Else
                ManageStatusBar("Normal")



        End Select


    End Sub

#End Region

#Region "Simulation"
    Private Sub InitializeSimulationMode()
        'housekeeping
        ''disable a few buttons
        ''change color of all rectangles
        ''
        IsSimulationMode = True
        showDTpercentframe()
        ChangeRectangleColors_ALL()
        ChangeDataLabelColors_ALL()
        ChangeRectangleWidth_All()
        incontrolframe.Visibility = Visibility.Hidden
        stopsframe.Visibility = Visibility.Hidden
        MenuCanvasA.Visibility = Visibility.Hidden
        If AllProdLines(selectedindexofLine_temp).isFilterByBrandcode = True Then
            FilterIcon_active.Visibility = True
            FilterOnOfflabel.Visibility = True
        End If

        ExporttoExcel_labelLA.Visibility = Visibility.Visible
        ExporttoExcelicon_LA.Visibility = Visibility.Visible
        ExportReasonTreetoExcel.Visibility = Visibility.Hidden

        FilterIcon_inactive.Visibility = Visibility.Hidden
        FilterOnOfflabel.Visibility = Visibility.Hidden
        pr_labelsim.Visibility = Visibility.Visible
        prStoryReport.initializeSimulationMode()
    End Sub
    Private Sub BringbacktoNormal()

        ResetallRectangleColors()
        ResetDataLabelColors_ALL()
        ResetRectangleSizes_All()
        IsSimulationMode = False
        CloseFloatingSimulator()
        incontrolframe.Visibility = Visibility.Visible
        stopsframe.Visibility = Visibility.Visible
        MenuCanvasA.Visibility = Visibility.Visible
        HideSimRectangles_UPDT()
        HideSimRectangles_Changeover()
        HideSimRectangles_PDT()
        HideSimRectangles_Equip()
        pr_labelsim.Visibility = Visibility.Hidden
        ExporttoExcel_labelLA.Visibility = Visibility.Hidden
        ExporttoExcelicon_LA.Visibility = Visibility.Hidden

        If AllProdLines(selectedindexofLine_temp).isFilterByBrandcode = True Then
            FilterIcon_active.Visibility = Visibility.Visible
        Else
            FilterIcon_inactive.Visibility = Visibility.Visible
        End If
        FilterOnOfflabel.Visibility = Visibility.Visible
        '  ExportReasonTreetoExcel.Visibility = Windows.Visibility.Visible
    End Sub


    Private Sub ExportLossAllocationtoExcel2()
        Try

            Dim appPath As String, startTime As String, endTime As String
            ' exportPath = My.Settings.CSV_ExportPath
            Dim dialog As New System.Windows.Forms.FolderBrowserDialog()
            dialog.RootFolder = Environment.SpecialFolder.Desktop
            dialog.SelectedPath = "C:\"
            dialog.Description = "Select Path To Save prStory Raw Data Files"
            If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                appPath = dialog.SelectedPath
                Dim parentLine = AllProdLines(prStoryReport.ParentLineInt)
                startTime = parentLine.rawProfStartTime.Day & Month(parentLine.rawProfStartTime) & Year(parentLine.rawProfStartTime)
                endTime = parentLine.rawProfEndTime.Day & Month(parentLine.rawProfEndTime) & Year(parentLine.rawProfEndTime)
                exportLossTreeAsCSV(appPath & "\" & parentLine.Name & "_" & "LossTree_" & startTime & "_" & endTime & ".csv", parentLine)


                Process.Start(appPath)
            End If

        Catch ex As Exception
            MessageBox.Show(".CSV Export Failed. " & ex.Message)

        End Try


    End Sub

    Private Sub exportLossTreeAsCSV(targetPath As String, parentLine As ProdLine) 'myArray As Array)
        Dim fsT As Object
        Dim fileName As String
        Dim colArr(0 To 15) As Integer
        Dim lineName As String
        lineName = parentLine.Name
        fileName = targetPath 'error hider
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object



        Dim rowR As Integer = 4
        Dim colC As Integer = 2
        Dim netuptimeDT As Double
        Dim netschedtime As Double
        Dim Tier1Incrementer As Integer, Tier2Incrementer As Integer, Tier3Incrementer As Integer, tmpTier2List As List(Of DTevent), tmpTier3list As List(Of DTevent)


        netschedtime = prStoryReport.MainLEDSReport.schedTime
        netuptimeDT = prStoryReport.MainLEDSReport.UT_DT

        fsT.writetext("prstory LOSS TREE")
        fsT.writetext(vbCrLf)
        fsT.writetext(linename_label.Content)
        fsT.writetext(vbCrLf)

        Dim starttimestring As String
        Dim endtimestring As String

        starttimestring = Month(starttimeselected) & "_" & Day(starttimeselected) & "_" & Year(starttimeselected) & "_" & Hour(starttimeselected) & "_" & Minute(starttimeselected)
        endtimestring = Month(endtimeselected) & "_" & Day(endtimeselected) & "_" & Year(endtimeselected) & "_" & Hour(endtimeselected) & "_" & Minute(endtimeselected)
        fsT.writetext(starttimeselected & "," & endtimeselected)

        fsT.writetext(vbCrLf)
        fsT.writetext("Loss Areas" & "," & "Stops" & "," & "Stops/day" & "," & "DT(min)" & "," & "DT%" & "," & "MTTR" & "," & "MTBF")
        fsT.writetext(vbCrLf)


        With prStoryReport

            For Tier1Incrementer = 0 To .UnplannedList.Count - 1
                fsT.writetext(.UnplannedList(Tier1Incrementer).Name & "," & Math.Round(.UnplannedList(Tier1Incrementer).Stops) & "," & Math.Round(.UnplannedList(Tier1Incrementer).SPD, 1) & "," & Math.Round(.UnplannedList(Tier1Incrementer).DT) & "," & FormatPercent(.UnplannedList(Tier1Incrementer).DTpct, 1) & "," & Math.Round(.UnplannedList(Tier1Incrementer).MTTR, 1) & "," & Math.Round(netuptimeDT / .UnplannedList(Tier1Incrementer).Stops, 1))
                fsT.writetext(vbCrLf)


                tmpTier2List = .MainLEDSReport.DT_Report.getTier2Directory(.UnplannedList(Tier1Incrementer).Name)
                tmpTier2List.Sort()
                For Tier2Incrementer = 0 To tmpTier2List.Count - 1

                    tmpTier2List(Tier2Incrementer).DTpct = netschedtime
                    fsT.writetext(tmpTier2List(Tier2Incrementer).Name & "," & Math.Round(tmpTier2List(Tier2Incrementer).Stops) & "," & Math.Round(tmpTier2List(Tier2Incrementer).SPD, 1) & "," & Math.Round(tmpTier2List(Tier2Incrementer).DT) & "," & FormatPercent(tmpTier2List(Tier2Incrementer).DT / netschedtime, 1) & "," & Math.Round(tmpTier2List(Tier2Incrementer).MTTR, 1) & "," & Math.Round(netuptimeDT / tmpTier2List(Tier2Incrementer).Stops, 1))
                    fsT.writetext(vbCrLf)


                    tmpTier3list = .MainLEDSReport.DT_Report.getTier3Directory(.UnplannedList(Tier1Incrementer).Name, tmpTier2List(Tier2Incrementer).Name)
                    tmpTier3list.Sort()
                    For Tier3Incrementer = 0 To tmpTier3list.Count - 1
                        tmpTier3list(Tier3Incrementer).DTpct = netschedtime

                        fsT.writetext(tmpTier3list(Tier3Incrementer).Name & "," & Math.Round(tmpTier3list(Tier3Incrementer).Stops) & "," & Math.Round(tmpTier3list(Tier3Incrementer).SPD, 1) & "," & Math.Round(tmpTier3list(Tier3Incrementer).DT) & "," & FormatPercent(tmpTier3list(Tier3Incrementer).DT / netschedtime, 1) & "," & Math.Round(tmpTier3list(Tier3Incrementer).MTTR, 1) & "," & Math.Round(netuptimeDT / tmpTier3list(Tier3Incrementer).Stops, 1))
                        fsT.writetext(vbCrLf)
                    Next
                Next
            Next

            fsT.writetext(pr_label.Content)
            fsT.writetext(vbCrLf)

        End With

        'fin
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing

    End Sub



    Private Sub ExportLossAllocationtoExcel()
        Try
            Dim appXL As Excel.Application
            Dim wbXl As Excel.Workbook
            Dim shXL As Excel.Worksheet
            Dim raXL As Excel.Range
            Dim rowR As Integer = 4
            Dim colC As Integer = 2
            Dim netuptimeDT As Double
            Dim netschedtime As Double
            Dim Tier1Incrementer As Integer, Tier2Incrementer As Integer, Tier3Incrementer As Integer, tmpTier2List As List(Of DTevent), tmpTier3list As List(Of DTevent)
            ' Start Excel and get Application object.

            appXL = CreateObject("Excel.Application")
            appXL.Visible = True
            ' Add a new workbook.
            wbXl = appXL.Workbooks.Add
            shXL = wbXl.ActiveSheet
            ' Add table headers going cell by cell.


            netschedtime = prStoryReport.MainLEDSReport.schedTime
            netuptimeDT = prStoryReport.MainLEDSReport.UT_DT
            If My.Settings.AdvancedSettings_UseSimulation Then
                shXL.Range("A1").Value = "prstory LOSS ALLOCATION"
            Else
                shXL.Range("A1").Value = "prstory LOSS TREE"
            End If
            shXL.Range("B1").Value = linename_label.Content

            Dim starttimestring As String
            Dim endtimestring As String

            starttimestring = Month(starttimeselected) & "_" & Day(starttimeselected) & "_" & Year(starttimeselected) & "_" & Hour(starttimeselected) & "_" & Minute(starttimeselected)
            endtimestring = Month(endtimeselected) & "_" & Day(endtimeselected) & "_" & Year(endtimeselected) & "_" & Hour(endtimeselected) & "_" & Minute(endtimeselected)
            shXL.Range("C1").Value = starttimeselected
            shXL.Range("D1").Value = endtimeselected

            With shXL.Range("A1", "D1")
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .Font.Size = 10
                .Font.Color = System.Drawing.Color.DarkSlateGray
                .Font.FontStyle = "Sans Serif"
            End With

            With shXL.Range("B2", "Q2")
                .Font.Bold = True
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .Font.Size = 15
                .Font.Color = System.Drawing.Color.DarkSlateGray
                .Font.FontStyle = "Sans Serif"

            End With
            shXL.Range("B3").Value = "Loss Areas"
            shXL.Range("C3").Value = "Stops"
            shXL.Range("D3").Value = "Stops/day"
            shXL.Range("E3").Value = "DT(min)"
            shXL.Range("F3").Value = "DT%"
            shXL.Range("G3").Value = "MTTR"
            shXL.Range("H3").Value = "MTBF"
            If IsSimulationMode Then
                shXL.Range("K3").Value = "Loss Areas"
                shXL.Range("L3").Value = "Stops"
                shXL.Range("M3").Value = "Stops/day"
                shXL.Range("N3").Value = "DT(min)"
                shXL.Range("O3").Value = "DT%"
                shXL.Range("P3").Value = "MTTR"
                shXL.Range("Q3").Value = "MTBF"
                shXL.Range("A2").Value = ""
                shXL.Range("B2").Value = "Base System"
                shXL.Range("K2").Value = "New System"
            End If




            With prStoryReport

                For Tier1Incrementer = 0 To .UnplannedList.Count - 1
                    shXL.Cells(rowR, colC) = "  " & .UnplannedList(Tier1Incrementer).Name
                    shXL.Range("B" & rowR).Font.Size = 14
                    shXL.Range("B" & rowR).Font.Bold = True
                    shXL.Cells(rowR, colC + 1) = Math.Round(.UnplannedList(Tier1Incrementer).Stops)
                    shXL.Cells(rowR, colC + 2) = Math.Round(.UnplannedList(Tier1Incrementer).SPD, 1)
                    shXL.Cells(rowR, colC + 3) = Math.Round(.UnplannedList(Tier1Incrementer).DT)
                    shXL.Cells(rowR, colC + 4) = FormatPercent(.UnplannedList(Tier1Incrementer).DTpct, 1)
                    shXL.Cells(rowR, colC + 5) = Math.Round(.UnplannedList(Tier1Incrementer).MTTR, 1)
                    shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / .UnplannedList(Tier1Incrementer).Stops, 1)

                    If IsSimulationMode Then
                        colC = colC + 9
                        shXL.Cells(rowR, colC) = "  " & .UnplannedList(Tier1Incrementer).Name
                        shXL.Range("K" & rowR).Font.Size = 14
                        shXL.Range("K" & rowR).Font.Bold = True
                        shXL.Cells(rowR, colC + 1) = Math.Round(.UnplannedList(Tier1Incrementer).StopsSim)
                        shXL.Cells(rowR, colC + 2) = Math.Round(.UnplannedList(Tier1Incrementer).SPDsim, 1)
                        shXL.Cells(rowR, colC + 3) = Math.Round(.UnplannedList(Tier1Incrementer).DTsim)
                        shXL.Cells(rowR, colC + 4) = FormatPercent(.UnplannedList(Tier1Incrementer).DTpctSim, 1)
                        shXL.Cells(rowR, colC + 5) = Math.Round(.UnplannedList(Tier1Incrementer).MTTRsim, 1)
                        shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / .UnplannedList(Tier1Incrementer).StopsSim, 1)
                        colC = 2
                    End If
                    rowR += 1


                    tmpTier2List = .MainLEDSReport.DT_Report.getTier2Directory(.UnplannedList(Tier1Incrementer).Name)
                    tmpTier2List.Sort()
                    For Tier2Incrementer = 0 To tmpTier2List.Count - 1


                        shXL.Cells(rowR, colC) = "     " & tmpTier2List(Tier2Incrementer).Name
                        shXL.Range("B" & rowR).Font.Size = 13
                        shXL.Range("B" & rowR).Font.Bold = True
                        tmpTier2List(Tier2Incrementer).DTpct = netschedtime
                        shXL.Cells(rowR, colC + 1) = Math.Round(tmpTier2List(Tier2Incrementer).Stops)
                        shXL.Cells(rowR, colC + 2) = Math.Round(tmpTier2List(Tier2Incrementer).SPD, 1)
                        shXL.Cells(rowR, colC + 3) = Math.Round(tmpTier2List(Tier2Incrementer).DT)
                        shXL.Cells(rowR, colC + 4) = FormatPercent(tmpTier2List(Tier2Incrementer).DT / netschedtime, 1)
                        shXL.Cells(rowR, colC + 5) = Math.Round(tmpTier2List(Tier2Incrementer).MTTR, 1)
                        shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / tmpTier2List(Tier2Incrementer).Stops, 1)

                        If IsSimulationMode Then
                            colC += 9
                            shXL.Cells(rowR, colC) = "     " & tmpTier2List(Tier2Incrementer).Name
                            shXL.Range("K" & rowR).Font.Size = 13
                            shXL.Range("K" & rowR).Font.Bold = True

                            shXL.Cells(rowR, colC + 1) = Math.Round(tmpTier2List(Tier2Incrementer).StopsSim)
                            shXL.Cells(rowR, colC + 2) = Math.Round(tmpTier2List(Tier2Incrementer).SPDsim, 1)
                            shXL.Cells(rowR, colC + 3) = Math.Round(tmpTier2List(Tier2Incrementer).DTsim)
                            shXL.Cells(rowR, colC + 4) = FormatPercent(tmpTier2List(Tier2Incrementer).DTsim / netschedtime, 1)
                            shXL.Cells(rowR, colC + 5) = Math.Round(tmpTier2List(Tier2Incrementer).MTTRsim, 1)
                            shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / tmpTier2List(Tier2Incrementer).StopsSim, 1)
                            colC = 2
                        End If

                        rowR += 1

                        tmpTier3list = .MainLEDSReport.DT_Report.getTier3Directory(.UnplannedList(Tier1Incrementer).Name, tmpTier2List(Tier2Incrementer).Name)
                        tmpTier3list.Sort()
                        For Tier3Incrementer = 0 To tmpTier3list.Count - 1

                            shXL.Cells(rowR, colC) = "         " & tmpTier3list(Tier3Incrementer).Name
                            shXL.Range("B" & rowR).Font.Size = 11
                            shXL.Range("B" & rowR).Font.Color = System.Drawing.Color.DarkSlateGray
                            tmpTier3list(Tier3Incrementer).DTpct = netschedtime
                            shXL.Cells(rowR, colC + 1) = Math.Round(tmpTier3list(Tier3Incrementer).Stops)
                            shXL.Cells(rowR, colC + 2) = Math.Round(tmpTier3list(Tier3Incrementer).SPD, 1)
                            shXL.Cells(rowR, colC + 3) = Math.Round(tmpTier3list(Tier3Incrementer).DT)
                            shXL.Cells(rowR, colC + 4) = FormatPercent(tmpTier3list(Tier3Incrementer).DT / netschedtime, 1)
                            shXL.Cells(rowR, colC + 5) = Math.Round(tmpTier3list(Tier3Incrementer).MTTR, 1)
                            shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / tmpTier3list(Tier3Incrementer).Stops, 1)

                            If IsSimulationMode Then
                                colC += 9
                                shXL.Cells(rowR, colC) = "         " & tmpTier3list(Tier3Incrementer).Name
                                shXL.Range("K" & rowR).Font.Size = 11
                                shXL.Range("K" & rowR).Font.Color = System.Drawing.Color.DarkSlateGray
                                shXL.Cells(rowR, colC + 1) = Math.Round(tmpTier3list(Tier3Incrementer).StopsSim)
                                shXL.Cells(rowR, colC + 2) = Math.Round(tmpTier3list(Tier3Incrementer).SPDsim, 1)
                                shXL.Cells(rowR, colC + 3) = Math.Round(tmpTier3list(Tier3Incrementer).DTsim)
                                shXL.Cells(rowR, colC + 4) = FormatPercent(tmpTier3list(Tier3Incrementer).DTsim / netschedtime, 1)
                                shXL.Cells(rowR, colC + 5) = Math.Round(tmpTier3list(Tier3Incrementer).MTTRsim, 1)
                                shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / tmpTier3list(Tier3Incrementer).StopsSim, 1)
                                colC = 2
                            End If

                            rowR += 1
                        Next
                    Next
                Next

                rowR += 2

                For Tier1Incrementer = 0 To .PlannedList.Count - 1
                    shXL.Cells(rowR, colC) = "  " & .PlannedList(Tier1Incrementer).Name
                    If Tier1Incrementer = 0 Then
                        shXL.Range("B" & rowR).Value = "Total Planned"
                    End If
                    shXL.Range("B" & rowR).Font.Size = 14
                    shXL.Range("B" & rowR).Font.Bold = True

                    shXL.Cells(rowR, colC + 1) = Math.Round(.PlannedList(Tier1Incrementer).Stops)
                    shXL.Cells(rowR, colC + 2) = Math.Round(.PlannedList(Tier1Incrementer).SPD, 1)
                    shXL.Cells(rowR, colC + 3) = Math.Round(.PlannedList(Tier1Incrementer).DT)
                    shXL.Cells(rowR, colC + 4) = FormatPercent(.PlannedList(Tier1Incrementer).DTpct, 1)
                    shXL.Cells(rowR, colC + 5) = Math.Round(.PlannedList(Tier1Incrementer).MTTR, 1)
                    shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / .PlannedList(Tier1Incrementer).Stops, 1)

                    If IsSimulationMode Then
                        colC = colC + 9
                        shXL.Cells(rowR, colC) = "  " & .PlannedList(Tier1Incrementer).Name
                        If Tier1Incrementer = 0 Then
                            shXL.Range("K" & rowR).Value = "Total Planned"
                        End If
                        shXL.Range("K" & rowR).Font.Size = 14
                        shXL.Range("K" & rowR).Font.Bold = True
                        shXL.Cells(rowR, colC + 1) = Math.Round(.PlannedList(Tier1Incrementer).StopsSim)
                        shXL.Cells(rowR, colC + 2) = Math.Round(.PlannedList(Tier1Incrementer).SPDsim, 1)
                        shXL.Cells(rowR, colC + 3) = Math.Round(.PlannedList(Tier1Incrementer).DTsim)
                        shXL.Cells(rowR, colC + 4) = FormatPercent(.PlannedList(Tier1Incrementer).DTpctSim, 1)
                        shXL.Cells(rowR, colC + 5) = Math.Round(.PlannedList(Tier1Incrementer).MTTRsim, 1)
                        shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / .PlannedList(Tier1Incrementer).StopsSim, 1)
                        colC = 2
                    End If
                    rowR += 1

                    tmpTier2List = .MainLEDSReport.DT_Report.getPlannedTier2Directory(.PlannedList(Tier1Incrementer).Name)
                    tmpTier2List.Sort()
                    For Tier2Incrementer = 0 To tmpTier2List.Count - 1


                        shXL.Cells(rowR, colC) = "     " & tmpTier2List(Tier2Incrementer).Name
                        shXL.Range("B" & rowR).Font.Size = 13
                        shXL.Range("B" & rowR).Font.Bold = True
                        tmpTier2List(Tier2Incrementer).DTpct = netschedtime
                        shXL.Cells(rowR, colC + 1) = Math.Round(tmpTier2List(Tier2Incrementer).Stops)
                        shXL.Cells(rowR, colC + 2) = Math.Round(tmpTier2List(Tier2Incrementer).SPD, 1)
                        shXL.Cells(rowR, colC + 3) = Math.Round(tmpTier2List(Tier2Incrementer).DT)
                        shXL.Cells(rowR, colC + 4) = FormatPercent(tmpTier2List(Tier2Incrementer).DT / netschedtime, 1)
                        shXL.Cells(rowR, colC + 5) = Math.Round(tmpTier2List(Tier2Incrementer).MTTR, 1)
                        shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / tmpTier2List(Tier2Incrementer).Stops, 1)

                        If IsSimulationMode Then
                            colC += 9
                            shXL.Cells(rowR, colC) = "     " & tmpTier2List(Tier2Incrementer).Name
                            shXL.Range("K" & rowR).Font.Size = 13
                            shXL.Range("K" & rowR).Font.Bold = True

                            shXL.Cells(rowR, colC + 1) = Math.Round(tmpTier2List(Tier2Incrementer).StopsSim)
                            shXL.Cells(rowR, colC + 2) = Math.Round(tmpTier2List(Tier2Incrementer).SPDsim, 1)
                            shXL.Cells(rowR, colC + 3) = Math.Round(tmpTier2List(Tier2Incrementer).DTsim)
                            shXL.Cells(rowR, colC + 4) = FormatPercent(tmpTier2List(Tier2Incrementer).DTsim / netschedtime, 1)
                            shXL.Cells(rowR, colC + 5) = Math.Round(tmpTier2List(Tier2Incrementer).MTTRsim, 1)
                            shXL.Cells(rowR, colC + 6) = Math.Round(netuptimeDT / tmpTier2List(Tier2Incrementer).StopsSim, 1)
                            colC = 2
                        End If

                        rowR += 1

                    Next

                Next




                rowR += 2

                If My.Settings.AdvancedSettings_isAvailabilityMode = False Then
                    shXL.Range("B" & rowR).Value = rateloss_label.Content
                    shXL.Range("B" & rowR).Font.Size = 12
                    If IsSimulationMode Then
                        shXL.Range("K" & rowR).Value = rateloss_label.Content
                        shXL.Range("K" & rowR).Font.Size = 12
                    End If

                End If

                rowR += 2



                shXL.Range("B" & rowR).Value = pr_label.Content
                shXL.Range("B" & rowR).Font.Size = 15
                shXL.Range("B" & rowR).Font.Bold = True

                If IsSimulationMode Then
                    shXL.Range("K" & rowR).Value = pr_labelsim.Content
                    shXL.Range("K" & rowR).Font.Size = 15
                    shXL.Range("K" & rowR).Font.Bold = True
                End If

            End With

            shXL.Range("B4").Value = "Total Unplanned"
            If IsSimulationMode Then
                shXL.Range("K4").Value = "Total Unplanned"
            End If

            raXL = shXL.Range("A4:Z20")
            raXL.EntireColumn.AutoFit()








            appXL.Visible = True
            appXL.UserControl = True
            ' Release object references.
            raXL = Nothing
            shXL = Nothing
            wbXl = Nothing

            appXL = Nothing

        Catch ex As Exception
            MessageBox.Show("Excel Export Failred. " & ex.Message)
        End Try


    End Sub
    Private Sub Stopsim()
        OriginStopSim = True
        KickStartSim()
        OriginStopSim = False
    End Sub
    Private Sub LocateLossSImulator(sender As Object, cardnumber As Integer)

        Dim canvaslocator As New Thickness
        Dim rectanglelocator As New Thickness
        Dim tmpdtevent_forSimInputs As DTevent


        If InStr(sender.name, "_Rect") > 0 Or InStr(sender.name, "_datalabel") > 0 Then
            If cardnumber = 1 And onlyDigits(sender.name) = 1 Then Exit Sub
            If cardnumber = 2 And onlyDigits(sender.name) = 1 Then Exit Sub

            rectanglelocator = sender.margin
            FloatingSimulator.Margin = New Thickness(rectanglelocator.Left, rectanglelocator.Top - 40, rectanglelocator.Left + 135, rectanglelocator.Top + 46)
            FloatingSimulator.Visibility = Visibility.Visible

            Select Case cardnumber
                Case 1
                    tmpdtevent_forSimInputs = Card_Unplanned_T1(onlyDigits(sender.name) - 1)
                Case 2
                    tmpdtevent_forSimInputs = Card_Planned_T1(onlyDigits(sender.name) - 1)
                Case 3
                    tmpdtevent_forSimInputs = Card_Unplanned_T2(onlyDigits(sender.name) - 1)
                Case 4
                    tmpdtevent_forSimInputs = Card_Unplanned_T3A(onlyDigits(Strings.Mid(sender.name, 7)) - 1)
                Case 5
                    tmpdtevent_forSimInputs = Card_Unplanned_T3B(onlyDigits(Strings.Mid(sender.name, 7)) - 1)
                Case 6
                    tmpdtevent_forSimInputs = Card_Unplanned_T3C(onlyDigits(Strings.Mid(sender.name, 7)) - 1)
                Case 41
                    tmpdtevent_forSimInputs = Card_Planned_T2(onlyDigits(sender.name) - 1)

            End Select
            If tmpdtevent_forSimInputs.Name <> "" Then
                MTBFactualbox.Content = Math.Round(tmpdtevent_forSimInputs.MTBF, 1)
                MTTRactualbox.Content = Math.Round(tmpdtevent_forSimInputs.MTTR, 1)
                MTBFnewbox.Content = Math.Round(tmpdtevent_forSimInputs.MTBFsim, 1)
                MTTRnewbox.Content = Math.Round(tmpdtevent_forSimInputs.MTTRsim, 1)


                MTBFsim.Text = Math.Round(tmpdtevent_forSimInputs.MTBF_userScaleFactor, 1)
                MTTRsim.Text = Math.Round(tmpdtevent_forSimInputs.MTTR_userScaleFactor, 1)
                Dtactualboxsim.Content = FormatPercent(tmpdtevent_forSimInputs.DTpct, 1)
                DTsim.Text = FormatPercent(tmpdtevent_forSimInputs.DTpctSim, 1)
                LOssName_Sim.Content = tmpdtevent_forSimInputs.Name
                CardNumSim.Content = cardnumber
            Else

                CloseFloatingSimulator()
                Exit Sub

            End If


        End If

    End Sub
    Private Sub CloseFloatingSimulator()
        FloatingSimulator.Visibility = Visibility.Hidden

    End Sub
    Private Sub ChangeDataLabelColors_ALL()
        UPDT_datalabel1.Foreground = mybrushlightgray
        UPDT_datalabel2.Foreground = mybrushlightgray
        UPDT_datalabel3.Foreground = mybrushlightgray
        UPDT_datalabel4.Foreground = mybrushlightgray
        UPDT_datalabel5.Foreground = mybrushlightgray
        UPDT_datalabel6.Foreground = mybrushlightgray



        PDT_datalabel1.Foreground = mybrushlightgray
        PDT_datalabel2.Foreground = mybrushlightgray
        PDT_datalabel3.Foreground = mybrushlightgray
        PDT_datalabel4.Foreground = mybrushlightgray
        PDT_datalabel5.Foreground = mybrushlightgray
        PDT_datalabel6.Foreground = mybrushlightgray
        PDT_datalabel7.Foreground = mybrushlightgray
        PDT_datalabel8.Foreground = mybrushlightgray
        PDT_datalabel9.Foreground = mybrushlightgray

        EquipMain_datalabel1.Foreground = mybrushlightgray
        EquipMain_datalabel2.Foreground = mybrushlightgray
        EquipMain_datalabel3.Foreground = mybrushlightgray
        EquipMain_datalabel4.Foreground = mybrushlightgray
        EquipMain_datalabel5.Foreground = mybrushlightgray
        EquipMain_datalabel6.Foreground = mybrushlightgray

        Equip1_datalabel1.Foreground = mybrushlightgray
        Equip1_datalabel2.Foreground = mybrushlightgray
        Equip1_datalabel3.Foreground = mybrushlightgray
        Equip2_datalabel1.Foreground = mybrushlightgray
        Equip2_datalabel2.Foreground = mybrushlightgray
        Equip2_datalabel3.Foreground = mybrushlightgray
        Equip3_datalabel1.Foreground = mybrushlightgray
        Equip3_datalabel2.Foreground = mybrushlightgray
        Equip3_datalabel3.Foreground = mybrushlightgray

        changeover_datalabel1.Foreground = mybrushlightgray
        changeover_datalabel2.Foreground = mybrushlightgray
        changeover_datalabel3.Foreground = mybrushlightgray
        changeover_datalabel4.Foreground = mybrushlightgray
        changeover_datalabel5.Foreground = mybrushlightgray
        changeover_datalabel6.Foreground = mybrushlightgray
        changeover_datalabel7.Foreground = mybrushlightgray



    End Sub
    Private Sub ResetDataLabelColors_ALL()
        UPDT_datalabel1.Foreground = mybrushfontgray
        UPDT_datalabel2.Foreground = mybrushfontgray
        UPDT_datalabel3.Foreground = mybrushfontgray
        UPDT_datalabel4.Foreground = mybrushfontgray
        UPDT_datalabel5.Foreground = mybrushfontgray
        UPDT_datalabel6.Foreground = mybrushfontgray



        PDT_datalabel1.Foreground = mybrushfontgray
        PDT_datalabel2.Foreground = mybrushfontgray
        PDT_datalabel3.Foreground = mybrushfontgray
        PDT_datalabel4.Foreground = mybrushfontgray
        PDT_datalabel5.Foreground = mybrushfontgray
        PDT_datalabel6.Foreground = mybrushfontgray
        PDT_datalabel7.Foreground = mybrushfontgray
        PDT_datalabel8.Foreground = mybrushfontgray
        PDT_datalabel9.Foreground = mybrushfontgray

        EquipMain_datalabel1.Foreground = mybrushfontgray
        EquipMain_datalabel2.Foreground = mybrushfontgray
        EquipMain_datalabel3.Foreground = mybrushfontgray
        EquipMain_datalabel4.Foreground = mybrushfontgray
        EquipMain_datalabel5.Foreground = mybrushfontgray
        EquipMain_datalabel6.Foreground = mybrushfontgray

        Equip1_datalabel1.Foreground = mybrushfontgray
        Equip1_datalabel2.Foreground = mybrushfontgray
        Equip1_datalabel3.Foreground = mybrushfontgray
        Equip2_datalabel1.Foreground = mybrushfontgray
        Equip2_datalabel2.Foreground = mybrushfontgray
        Equip2_datalabel3.Foreground = mybrushfontgray
        Equip3_datalabel1.Foreground = mybrushfontgray
        Equip3_datalabel2.Foreground = mybrushfontgray
        Equip3_datalabel3.Foreground = mybrushfontgray

        changeover_datalabel1.Foreground = mybrushfontgray
        changeover_datalabel2.Foreground = mybrushfontgray
        changeover_datalabel3.Foreground = mybrushfontgray
        changeover_datalabel4.Foreground = mybrushfontgray
        changeover_datalabel5.Foreground = mybrushfontgray
        changeover_datalabel6.Foreground = mybrushfontgray
        changeover_datalabel7.Foreground = mybrushfontgray



    End Sub
    Private Sub ChangeRectangleColors_ALL()
        UPDT_Rect_1.Fill = mybrushlightgray
        UPDT_Rect_2.Fill = mybrushlightgray
        UPDT_Rect_3.Fill = mybrushlightgray
        UPDT_Rect_4.Fill = mybrushlightgray
        UPDT_Rect_5.Fill = mybrushlightgray
        UPDT_Rect_6.Fill = mybrushlightgray

        PDT_Rect_1.Fill = mybrushlightgray
        PDT_Rect_2.Fill = mybrushlightgray
        PDT_Rect_3.Fill = mybrushlightgray
        PDT_Rect_4.Fill = mybrushlightgray
        PDT_Rect_5.Fill = mybrushlightgray
        PDT_Rect_6.Fill = mybrushlightgray
        PDT_Rect_7.Fill = mybrushlightgray
        PDT_Rect_8.Fill = mybrushlightgray
        PDT_Rect_9.Fill = mybrushlightgray

        EquipMain_Rect1.Fill = mybrushlightgray
        EquipMain_Rect2.Fill = mybrushlightgray
        EquipMain_Rect3.Fill = mybrushlightgray
        EquipMain_Rect4.Fill = mybrushlightgray
        EquipMain_Rect5.Fill = mybrushlightgray
        EquipMain_Rect6.Fill = mybrushlightgray

        Equip1_Rect1.Fill = mybrushlightgray
        Equip1_Rect2.Fill = mybrushlightgray
        Equip1_Rect3.Fill = mybrushlightgray
        Equip2_Rect1.Fill = mybrushlightgray
        Equip2_Rect2.Fill = mybrushlightgray
        Equip2_Rect3.Fill = mybrushlightgray
        Equip3_Rect1.Fill = mybrushlightgray
        Equip3_Rect2.Fill = mybrushlightgray
        Equip3_Rect3.Fill = mybrushlightgray

        changeover_Rect1.Fill = mybrushlightgray
        changeover_Rect2.Fill = mybrushlightgray
        changeover_Rect3.Fill = mybrushlightgray
        changeover_Rect4.Fill = mybrushlightgray
        changeover_Rect5.Fill = mybrushlightgray
        changeover_Rect6.Fill = mybrushlightgray
        changeover_Rect7.Fill = mybrushlightgray



    End Sub
    Private Sub ToggleSim()

        If DTactuallabel.Visibility = Windows.Visibility.Visible Then
            ShowSIMMTBF()

            DTsim.Visibility = Visibility.Hidden
            DTactuallabel.Visibility = Visibility.Hidden
            Dtactualboxsim.Visibility = Visibility.Hidden

            MTBFHeaderLabel.Background = mybrushlightgreen
            DTHeaderLabel.Background = mybrushlightgray
        Else
            HideSIMMTBF()

            DTsim.Visibility = Visibility.Visible
            DTactuallabel.Visibility = Visibility.Visible
            Dtactualboxsim.Visibility = Visibility.Visible

            MTBFHeaderLabel.Background = mybrushlightgray
            DTHeaderLabel.Background = mybrushlightgreen
        End If

    End Sub
    Private Sub ShowSIMMTBF()

        MTBFsim.Visibility = Visibility.Visible
        MTTRsim.Visibility = Visibility.Visible
        MTBFactualbox.Visibility = Visibility.Visible
        MTTRactualbox.Visibility = Visibility.Visible
        MTBFactuallabel.Visibility = Visibility.Visible
        MTTRactuallabel.Visibility = Visibility.Visible
    End Sub
    Private Sub HideSIMMTBF()

        MTBFsim.Visibility = Visibility.Hidden
        MTTRsim.Visibility = Visibility.Hidden
        MTBFactualbox.Visibility = Visibility.Hidden
        MTTRactualbox.Visibility = Visibility.Hidden
        MTBFactuallabel.Visibility = Visibility.Hidden
        MTTRactuallabel.Visibility = Visibility.Hidden
    End Sub
    Private Sub ResetallRectangleColors()
        UPDT_Rect_1.Fill = mybrushbrightblue
        UPDT_Rect_2.Fill = mybrushbrightblue
        UPDT_Rect_3.Fill = mybrushbrightblue
        UPDT_Rect_4.Fill = mybrushbrightblue
        UPDT_Rect_5.Fill = mybrushbrightblue
        UPDT_Rect_6.Fill = mybrushbrightblue

        PDT_Rect_1.Fill = mybrushbrightblue
        PDT_Rect_2.Fill = mybrushbrightblue
        PDT_Rect_3.Fill = mybrushbrightblue
        PDT_Rect_4.Fill = mybrushbrightblue
        PDT_Rect_5.Fill = mybrushbrightblue
        PDT_Rect_6.Fill = mybrushbrightblue
        PDT_Rect_7.Fill = mybrushbrightblue
        PDT_Rect_8.Fill = mybrushbrightblue
        PDT_Rect_9.Fill = mybrushbrightblue

        EquipMain_Rect1.Fill = mybrushbrightblue
        EquipMain_Rect2.Fill = mybrushbrightblue
        EquipMain_Rect3.Fill = mybrushbrightblue
        EquipMain_Rect4.Fill = mybrushbrightblue
        EquipMain_Rect5.Fill = mybrushbrightblue
        EquipMain_Rect6.Fill = mybrushbrightblue

        Equip1_Rect1.Fill = mybrushbrightblue
        Equip1_Rect2.Fill = mybrushbrightblue
        Equip1_Rect1.Fill = mybrushbrightblue
        Equip2_Rect1.Fill = mybrushbrightblue
        Equip2_Rect2.Fill = mybrushbrightblue
        Equip2_Rect3.Fill = mybrushbrightblue
        Equip3_Rect1.Fill = mybrushbrightblue
        Equip3_Rect2.Fill = mybrushbrightblue
        Equip3_Rect3.Fill = mybrushbrightblue

        changeover_Rect1.Fill = mybrushbrightblue
        changeover_Rect2.Fill = mybrushbrightblue
        changeover_Rect3.Fill = mybrushbrightblue
        changeover_Rect4.Fill = mybrushbrightblue
        changeover_Rect5.Fill = mybrushbrightblue
        changeover_Rect6.Fill = mybrushbrightblue
        changeover_Rect7.Fill = mybrushbrightblue




    End Sub
    Private Sub ResetRectangleSizes_All()

        UPDT_Rect_1.Width = 30
        UPDT_Rect_2.Width = 30
        UPDT_Rect_3.Width = 30
        UPDT_Rect_4.Width = 30
        UPDT_Rect_5.Width = 30
        UPDT_Rect_6.Width = 30

        PDT_Rect_1.Width = 30
        PDT_Rect_2.Width = 30
        PDT_Rect_3.Width = 30
        PDT_Rect_4.Width = 30
        PDT_Rect_5.Width = 30
        PDT_Rect_6.Width = 30
        PDT_Rect_7.Width = 30
        PDT_Rect_8.Width = 30
        PDT_Rect_9.Width = 30

        EquipMain_Rect1.Width = 30
        EquipMain_Rect2.Width = 30
        EquipMain_Rect3.Width = 30
        EquipMain_Rect4.Width = 30
        EquipMain_Rect5.Width = 30
        EquipMain_Rect6.Width = 30

        Equip1_Rect1.Width = 30
        Equip1_Rect2.Width = 30
        Equip1_Rect1.Width = 30
        Equip2_Rect1.Width = 30
        Equip2_Rect2.Width = 30
        Equip2_Rect3.Width = 30
        Equip3_Rect1.Width = 30
        Equip3_Rect2.Width = 30
        Equip3_Rect3.Width = 30

        changeover_Rect1.Width = 30
        changeover_Rect2.Width = 30
        changeover_Rect3.Width = 30
        changeover_Rect4.Width = 30
        changeover_Rect5.Width = 30
        changeover_Rect6.Width = 30
        changeover_Rect7.Width = 30
    End Sub
    Private Sub ChangeRectangleWidth_All()

        UPDT_Rect_1.Width = 25
        UPDT_Rect_2.Width = 25
        UPDT_Rect_3.Width = 25
        UPDT_Rect_4.Width = 25
        UPDT_Rect_5.Width = 25
        UPDT_Rect_6.Width = 25

        PDT_Rect_1.Width = 25
        PDT_Rect_2.Width = 25
        PDT_Rect_3.Width = 25
        PDT_Rect_4.Width = 25
        PDT_Rect_5.Width = 25
        PDT_Rect_6.Width = 25
        PDT_Rect_7.Width = 25
        PDT_Rect_8.Width = 25
        PDT_Rect_9.Width = 25

        EquipMain_Rect1.Width = 25
        EquipMain_Rect2.Width = 25
        EquipMain_Rect3.Width = 25
        EquipMain_Rect4.Width = 25
        EquipMain_Rect5.Width = 25
        EquipMain_Rect6.Width = 25

        Equip1_Rect1.Width = 25
        Equip1_Rect2.Width = 25
        Equip1_Rect1.Width = 25
        Equip2_Rect1.Width = 25
        Equip2_Rect2.Width = 25
        Equip2_Rect3.Width = 25
        Equip3_Rect1.Width = 25
        Equip3_Rect2.Width = 25
        Equip3_Rect3.Width = 25

        changeover_Rect1.Width = 25
        changeover_Rect2.Width = 25
        changeover_Rect3.Width = 25
        changeover_Rect4.Width = 25
        changeover_Rect5.Width = 25
        changeover_Rect6.Width = 25
        changeover_Rect7.Width = 25

    End Sub
    Private Sub HideSimRectangles_UPDT()
        UPDT_Rect_1sim.Visibility = Visibility.Hidden
        UPDT_Rect_2sim.Visibility = Visibility.Hidden
        UPDT_Rect_3sim.Visibility = Visibility.Hidden
        UPDT_Rect_4sim.Visibility = Visibility.Hidden
        UPDT_Rect_5sim.Visibility = Visibility.Hidden
        UPDT_Rect_6sim.Visibility = Visibility.Hidden

        UPDT_datalabel1sim.Visibility = Visibility.Hidden
        UPDT_datalabel2sim.Visibility = Visibility.Hidden
        UPDT_datalabel3sim.Visibility = Visibility.Hidden
        UPDT_datalabel4sim.Visibility = Visibility.Hidden
        UPDT_datalabel5sim.Visibility = Visibility.Hidden
        UPDT_datalabel6sim.Visibility = Visibility.Hidden

    End Sub
    Private Sub HideSimRectangles_PDT()


        PDT_Rect_1sim.Visibility = Visibility.Hidden
        PDT_Rect_2sim.Visibility = Visibility.Hidden
        PDT_Rect_3sim.Visibility = Visibility.Hidden
        PDT_Rect_4sim.Visibility = Visibility.Hidden
        PDT_Rect_5sim.Visibility = Visibility.Hidden
        PDT_Rect_6sim.Visibility = Visibility.Hidden
        PDT_Rect_7sim.Visibility = Visibility.Hidden
        PDT_Rect_8sim.Visibility = Visibility.Hidden
        PDT_Rect_9sim.Visibility = Visibility.Hidden

        PDT_datalabel1sim.Visibility = Visibility.Hidden
        PDT_datalabel2sim.Visibility = Visibility.Hidden
        PDT_datalabel3sim.Visibility = Visibility.Hidden
        PDT_datalabel4sim.Visibility = Visibility.Hidden
        PDT_datalabel5sim.Visibility = Visibility.Hidden
        PDT_datalabel6sim.Visibility = Visibility.Hidden
        PDT_datalabel7sim.Visibility = Visibility.Hidden
        PDT_datalabel8sim.Visibility = Visibility.Hidden
        PDT_datalabel9sim.Visibility = Visibility.Hidden

    End Sub
    Private Sub HideSimRectangles_Equip()

        EquipMain_Rect1sim.Visibility = Visibility.Hidden
        EquipMain_Rect2sim.Visibility = Visibility.Hidden
        EquipMain_Rect3sim.Visibility = Visibility.Hidden
        EquipMain_Rect4sim.Visibility = Visibility.Hidden
        EquipMain_Rect5sim.Visibility = Visibility.Hidden
        EquipMain_Rect6sim.Visibility = Visibility.Hidden

        Equip1_Rect1sim.Visibility = Visibility.Hidden
        Equip1_Rect2sim.Visibility = Visibility.Hidden
        Equip1_Rect3sim.Visibility = Visibility.Hidden

        Equip2_Rect1sim.Visibility = Visibility.Hidden
        Equip2_Rect2sim.Visibility = Visibility.Hidden
        Equip2_Rect3sim.Visibility = Visibility.Hidden

        Equip3_Rect1sim.Visibility = Visibility.Hidden
        Equip3_Rect2sim.Visibility = Visibility.Hidden
        Equip3_Rect3sim.Visibility = Visibility.Hidden


        EquipMain_datalabel1sim.Visibility = Visibility.Hidden
        EquipMain_datalabel2sim.Visibility = Visibility.Hidden
        EquipMain_datalabel3sim.Visibility = Visibility.Hidden
        EquipMain_datalabel4sim.Visibility = Visibility.Hidden
        EquipMain_datalabel5sim.Visibility = Visibility.Hidden
        EquipMain_datalabel6sim.Visibility = Visibility.Hidden

        Equip1_datalabel1sim.Visibility = Visibility.Hidden
        Equip1_datalabel2sim.Visibility = Visibility.Hidden
        Equip1_datalabel3sim.Visibility = Visibility.Hidden

        Equip2_datalabel1sim.Visibility = Visibility.Hidden
        Equip2_datalabel2sim.Visibility = Visibility.Hidden
        Equip2_datalabel3sim.Visibility = Visibility.Hidden

        Equip3_datalabel1sim.Visibility = Visibility.Hidden
        Equip3_datalabel2sim.Visibility = Visibility.Hidden
        Equip3_datalabel3sim.Visibility = Visibility.Hidden

    End Sub
    Private Sub HideSimRectangles_Changeover()

        changeover_Rect1sim.Visibility = Visibility.Hidden
        changeover_Rect2sim.Visibility = Visibility.Hidden
        changeover_Rect3sim.Visibility = Visibility.Hidden
        changeover_Rect4sim.Visibility = Visibility.Hidden
        changeover_Rect5sim.Visibility = Visibility.Hidden
        changeover_Rect6sim.Visibility = Visibility.Hidden
        changeover_Rect7sim.Visibility = Visibility.Hidden

        changeover_datalabel1sim.Visibility = Visibility.Hidden
        changeover_datalabel2sim.Visibility = Visibility.Hidden
        changeover_datalabel3sim.Visibility = Visibility.Hidden
        changeover_datalabel4sim.Visibility = Visibility.Hidden
        changeover_datalabel5sim.Visibility = Visibility.Hidden
        changeover_datalabel6sim.Visibility = Visibility.Hidden
        changeover_datalabel7sim.Visibility = Visibility.Hidden
    End Sub
    Private Sub ShowSimRectangles_UPDT()


        UPDT_Rect_1sim.Visibility = Visibility.Visible
        UPDT_Rect_2sim.Visibility = Visibility.Visible
        UPDT_Rect_3sim.Visibility = Visibility.Visible
        UPDT_Rect_4sim.Visibility = Visibility.Visible
        UPDT_Rect_5sim.Visibility = Visibility.Visible
        UPDT_Rect_6sim.Visibility = Visibility.Visible

        If UPDT_datalabel1sim.Content <> "0.0%" Then
            UPDT_datalabel1sim.Visibility = Visibility.Visible
            UPDT_datalabel2sim.Visibility = Visibility.Visible
            UPDT_datalabel3sim.Visibility = Visibility.Visible
            UPDT_datalabel4sim.Visibility = Visibility.Visible
            UPDT_datalabel5sim.Visibility = Visibility.Visible
            UPDT_datalabel6sim.Visibility = Visibility.Visible
        End If

    End Sub
    Private Sub ShowSimRectangles_PDT()


        PDT_Rect_1sim.Visibility = Visibility.Visible
        PDT_Rect_2sim.Visibility = Visibility.Visible
        PDT_Rect_3sim.Visibility = Visibility.Visible
        PDT_Rect_4sim.Visibility = Visibility.Visible
        PDT_Rect_5sim.Visibility = Visibility.Visible
        PDT_Rect_6sim.Visibility = Visibility.Visible
        PDT_Rect_7sim.Visibility = Visibility.Visible
        PDT_Rect_8sim.Visibility = Visibility.Visible
        PDT_Rect_9sim.Visibility = Visibility.Visible

        If PDT_datalabel1sim.Content <> "0.0%" Then
            PDT_datalabel1sim.Visibility = Visibility.Visible
            PDT_datalabel2sim.Visibility = Visibility.Visible
            PDT_datalabel3sim.Visibility = Visibility.Visible
            PDT_datalabel4sim.Visibility = Visibility.Visible
            PDT_datalabel5sim.Visibility = Visibility.Visible
            PDT_datalabel6sim.Visibility = Visibility.Visible
            PDT_datalabel7sim.Visibility = Visibility.Visible
            PDT_datalabel8sim.Visibility = Visibility.Visible
            PDT_datalabel9sim.Visibility = Visibility.Visible
        End If


    End Sub
    Private Sub ShowSimRectangles_Equip()

        EquipMain_Rect1sim.Visibility = Visibility.Visible
        EquipMain_Rect2sim.Visibility = Visibility.Visible
        EquipMain_Rect3sim.Visibility = Visibility.Visible
        EquipMain_Rect4sim.Visibility = Visibility.Visible
        EquipMain_Rect5sim.Visibility = Visibility.Visible
        EquipMain_Rect6sim.Visibility = Visibility.Visible

        Equip1_Rect1sim.Visibility = Visibility.Visible
        Equip1_Rect2sim.Visibility = Visibility.Visible
        Equip1_Rect3sim.Visibility = Visibility.Visible

        Equip2_Rect1sim.Visibility = Visibility.Visible
        Equip2_Rect2sim.Visibility = Visibility.Visible
        Equip2_Rect3sim.Visibility = Visibility.Visible

        Equip3_Rect1sim.Visibility = Visibility.Visible
        Equip3_Rect2sim.Visibility = Visibility.Visible
        Equip3_Rect3sim.Visibility = Visibility.Visible
        If UPDT_datalabel1sim.Content <> "0.0%" And PDT_datalabel1sim.Content <> "0.0%" Then
            EquipMain_datalabel1sim.Visibility = Visibility.Visible
            EquipMain_datalabel2sim.Visibility = Visibility.Visible
            EquipMain_datalabel3sim.Visibility = Visibility.Visible
            EquipMain_datalabel4sim.Visibility = Visibility.Visible
            EquipMain_datalabel5sim.Visibility = Visibility.Visible
            EquipMain_datalabel6sim.Visibility = Visibility.Visible

            Equip1_datalabel1sim.Visibility = Visibility.Visible
            Equip1_datalabel2sim.Visibility = Visibility.Visible
            Equip1_datalabel3sim.Visibility = Visibility.Visible

            Equip2_datalabel1sim.Visibility = Visibility.Visible
            Equip2_datalabel2sim.Visibility = Visibility.Visible
            Equip2_datalabel3sim.Visibility = Visibility.Visible

            Equip3_datalabel1sim.Visibility = Visibility.Visible
            Equip3_datalabel2sim.Visibility = Visibility.Visible
            Equip3_datalabel3sim.Visibility = Visibility.Visible
        End If
    End Sub
    Private Sub ShowSimRectangles_Changeover()

        changeover_Rect1sim.Visibility = Visibility.Visible
        changeover_Rect2sim.Visibility = Visibility.Visible
        changeover_Rect3sim.Visibility = Visibility.Visible
        changeover_Rect4sim.Visibility = Visibility.Visible
        changeover_Rect5sim.Visibility = Visibility.Visible
        changeover_Rect6sim.Visibility = Visibility.Visible
        changeover_Rect7sim.Visibility = Visibility.Visible
        If UPDT_datalabel1sim.Content <> "0.0%" And PDT_datalabel1sim.Content <> "0.0%" Then
            changeover_datalabel1sim.Visibility = Visibility.Visible
            changeover_datalabel2sim.Visibility = Visibility.Visible
            changeover_datalabel3sim.Visibility = Visibility.Visible
            changeover_datalabel4sim.Visibility = Visibility.Visible
            changeover_datalabel5sim.Visibility = Visibility.Visible
            changeover_datalabel6sim.Visibility = Visibility.Visible
            changeover_datalabel7sim.Visibility = Visibility.Visible
        End If

    End Sub
    Private Sub ColorModifiedBars(cardnumber As Integer, labelname As String, mtbfscale As Double, mttrscale As Double)
        Dim mybrushdynamicmodified As SolidColorBrush

        If mtbfscale <> 1 Or mttrscale <> 1 Then
            mybrushdynamicmodified = mybrushmodifiedGreen
        Else
            mybrushdynamicmodified = mybrushbrightorange
        End If


        Select Case cardnumber

            Case 1
                Select Case Card_Unplanned_T1.IndexOf(New DTevent(labelname, 0)) + 1
                    Case 1
                        UPDT_Rect_1sim.Fill = mybrushdynamicmodified
                    Case 2
                        UPDT_Rect_2sim.Fill = mybrushdynamicmodified
                    Case 3
                        UPDT_Rect_3sim.Fill = mybrushdynamicmodified
                    Case 4
                        UPDT_Rect_4sim.Fill = mybrushdynamicmodified
                    Case 5
                        UPDT_Rect_5sim.Fill = mybrushdynamicmodified
                    Case 6
                        UPDT_Rect_6sim.Fill = mybrushdynamicmodified

                End Select
            Case 2

                Select Case Card_Planned_T1.IndexOf(New DTevent(labelname, 0)) + 1
                    Case 1
                        PDT_Rect_1sim.Fill = mybrushdynamicmodified
                    Case 2
                        PDT_Rect_2sim.Fill = mybrushdynamicmodified
                    Case 3
                        PDT_Rect_3sim.Fill = mybrushdynamicmodified
                    Case 4
                        PDT_Rect_4sim.Fill = mybrushdynamicmodified
                    Case 5
                        PDT_Rect_5sim.Fill = mybrushdynamicmodified
                    Case 6
                        PDT_Rect_6sim.Fill = mybrushdynamicmodified
                    Case 7
                        PDT_Rect_7sim.Fill = mybrushdynamicmodified
                    Case 8
                        PDT_Rect_8sim.Fill = mybrushdynamicmodified
                    Case 9
                        PDT_Rect_9sim.Fill = mybrushdynamicmodified
                End Select

            Case 3
                Select Case Card_Unplanned_T2.IndexOf(New DTevent(labelname, 0)) + 1
                    Case 1
                        EquipMain_Rect1sim.Fill = mybrushdynamicmodified
                    Case 2
                        EquipMain_Rect2sim.Fill = mybrushdynamicmodified
                    Case 3
                        EquipMain_Rect3sim.Fill = mybrushdynamicmodified
                    Case 4
                        EquipMain_Rect4sim.Fill = mybrushdynamicmodified
                    Case 5
                        EquipMain_Rect5sim.Fill = mybrushdynamicmodified
                    Case 6
                        EquipMain_Rect6sim.Fill = mybrushdynamicmodified

                End Select

            Case 4
                Select Case Card_Unplanned_T3A.IndexOf(New DTevent(labelname, 0)) + 1
                    Case 1
                        Equip1_Rect1sim.Fill = mybrushdynamicmodified
                    Case 2
                        Equip1_Rect2sim.Fill = mybrushdynamicmodified
                    Case 3
                        Equip1_Rect3sim.Fill = mybrushdynamicmodified

                End Select

            Case 5

                Select Case Card_Unplanned_T3B.IndexOf(New DTevent(labelname, 0)) + 1
                    Case 1
                        Equip2_Rect1sim.Fill = mybrushdynamicmodified
                    Case 2
                        Equip2_Rect2sim.Fill = mybrushdynamicmodified
                    Case 3
                        Equip2_Rect3sim.Fill = mybrushdynamicmodified

                End Select

            Case 6

                Select Case Card_Unplanned_T3C.IndexOf(New DTevent(labelname, 0)) + 1
                    Case 1
                        Equip3_Rect1sim.Fill = mybrushdynamicmodified
                    Case 2
                        Equip3_Rect2sim.Fill = mybrushdynamicmodified
                    Case 3
                        Equip3_Rect3sim.Fill = mybrushdynamicmodified

                End Select

            Case 41
                Select Case Card_Planned_T2.IndexOf(New DTevent(labelname, 0)) + 1
                    Case 1
                        changeover_Rect1sim.Fill = mybrushdynamicmodified
                    Case 2
                        changeover_Rect2sim.Fill = mybrushdynamicmodified
                    Case 3
                        changeover_Rect3sim.Fill = mybrushdynamicmodified
                    Case 4
                        changeover_Rect4sim.Fill = mybrushdynamicmodified
                    Case 5
                        changeover_Rect5sim.Fill = mybrushdynamicmodified
                    Case 6
                        changeover_Rect6sim.Fill = mybrushdynamicmodified
                    Case 7
                        changeover_Rect7sim.Fill = mybrushdynamicmodified
                End Select


        End Select

    End Sub
    Private Sub KickStartSim()

        If MTTRsim.Text = "" Then MTTRsim.Text = "1"
        If MTBFsim.Text = "" Then MTBFsim.Text = "1"

        If OriginStopSim = True Then
            MTTRsim.Text = "1"
            MTBFsim.Text = "1"

        End If


        '  MsgBox(CardNumSim.Content)
        If CDbl(CardNumSim.Content) <> 0.0 Then
            ColorModifiedBars(CDbl(CardNumSim.Content), LOssName_Sim.Content, CDbl(MTBFsim.Text), CDbl(MTTRsim.Text))
            Select Case onlyDigits(CardNumSim.Content)
                Case 1
                    prStoryReport.GenerateNewLossAllocation(onlyDigits(CardNumSim.Content), LOssName_Sim.Content, CDbl(MTTRsim.Text), CDbl(MTBFsim.Text))
                Case 3
                    prStoryReport.GenerateNewLossAllocation(onlyDigits(CardNumSim.Content), LOssName_Sim.Content, CDbl(MTTRsim.Text), CDbl(MTBFsim.Text), Tier1Clicked_Unplanned)
                Case 2
                    prStoryReport.GenerateNewLossAllocation(onlyDigits(CardNumSim.Content), LOssName_Sim.Content, CDbl(MTTRsim.Text), CDbl(MTBFsim.Text))
                Case 41
                    prStoryReport.GenerateNewLossAllocation(onlyDigits(CardNumSim.Content), LOssName_Sim.Content, CDbl(MTTRsim.Text), CDbl(MTBFsim.Text), Tier1Clicked_planned)
                Case Else
                    prStoryReport.GenerateNewLossAllocation(onlyDigits(CardNumSim.Content), LOssName_Sim.Content, CDbl(MTTRsim.Text), CDbl(MTBFsim.Text), Tier1Clicked_Unplanned, Tier2Clicked_Unplanned)
            End Select

        End If

        pr_labelsim.Visibility = Visibility.Visible
        '      If My.Settings.AdvancedSettings_isAvailabilityMode Then
        'pr_labelsim.Content = FormatPercent(prStoryReport.AvSys_Sim, 1) & " Av."
        '      Else
        '      If prStoryReport.AvSys_Sim = 0 Then
        '      pr_labelsim.Content = pr_label.Content
        '      Else
        '      pr_labelsim.Content = FormatPercent(prStoryReport.AvSys_Sim - prStoryReport.rateLoss, 1) & " PR"
        '      End If
        '   End If
        '



        'If Not My.Settings.AdvancedSettings_isAvailabilityMode Then pr_labelsim.Content = FormatPercent(prStoryReport.AvSys_Sim - prStoryReport.rateLoss, 1) & " PR"


        If prStoryReport.EventMaxDTpctUnplannedSim >= MaxPRindataset Then
            MaxPRindataset = prStoryReport.EventMaxDTpctUnplannedSim '+ prStoryReport.rateLoss
        Else
            MaxPRindataset = prStoryReport.UPDT
        End If
        If prStoryReport.EventMaxDTpctplannedSim >= MaxPRindataset_planned Then
            MaxPRindataset_planned = prStoryReport.EventMaxDTpctplannedSim '+ prStoryReport.rateLoss
        Else
            MaxPRindataset_planned = prStoryReport.PDT
        End If


        updateCardList(1, ScrollBase_Card1, DowntimeField.Tier1, "", "")
        updateCard_Unplanned_Tier1()


        updateCardList(2, 0, DowntimeField.Tier1, "", "")
        updateCard_Planned_Tier1()
        showUPDTlabels()
        showPDTlabels()

        If Card41Header.Visibility = Windows.Visibility.Visible Then
            updateCardList(41, 0, DowntimeField.Tier1, Tier1Clicked_planned, "")
            updateCard_Planned_Tier2()
            showchangeoverlabels()
        End If

        If Card3Header.Visibility = Windows.Visibility.Visible Then

            updateCardList(3, ScrollBase_Card3, DowntimeField.Tier1, Card3Header.Content, "")
            updateCard_Unplanned_Tier2()




            '''''''copied code from Tier2Clicked togenerateTier3
            Dim indexA As Integer
            indexA = Card_Unplanned_T2.IndexOf(New DTevent(Card4Header.Content.ToString, 0))

            If indexA > -1 Then
                updateCardList(4, ScrollBase_Card3, DowntimeField.Tier1, Card3Header.Content.ToString, Card_Unplanned_T2(indexA).Name)
                Card4Header.Content = Card_Unplanned_T2(indexA).Name
                cardnameLabeltext(prStoryCard.Equipment_One) = Card_Unplanned_T2(indexA).Name
                If indexA + 1 < Card_Unplanned_T2.Count Then
                    Card5Header.Content = Card_Unplanned_T2(indexA + 1).Name
                    cardnameLabeltext(prStoryCard.Equipment_Two) = Card_Unplanned_T2(indexA + 1).Name
                    updateCardList(5, ScrollBase_Card3, DowntimeField.Tier1, Card3Header.Content.ToString, Card_Unplanned_T2(indexA + 1).Name)
                    If indexA + 2 < Card_Unplanned_T2.Count Then
                        updateCardList(6, ScrollBase_Card3, DowntimeField.Tier1, Card3Header.Content.ToString, Card_Unplanned_T2(indexA + 2).Name)
                        Card6Header.Content = Card_Unplanned_T2(indexA + 2).Name
                        cardnameLabeltext(prStoryCard.Equipment_Three) = Card_Unplanned_T2(indexA + 2).Name
                    Else
                        cardnameLabeltext(prStoryCard.Equipment_Three) = ""
                        updateCardList(6, ScrollBase_Card3, DowntimeField.Tier1, "x", "x")
                    End If
                Else
                    cardnameLabeltext(prStoryCard.Equipment_Two) = ""
                    updateCardList(5, ScrollBase_Card3, DowntimeField.Tier1, "x", "x")
                End If
            End If

            ''''''''''''''''''''''
            updateCard_Unplanned_Tier3_All()

            showEquipmentlabels()
            If prStoryReport.getCardEventNumber(1) <= 6 Then NavigationRight_Card1.Visibility = Visibility.Hidden
            If prStoryReport.getCardEventNumber(3) <= 6 Then NavigationRight_Card3.Visibility = Visibility.Hidden
        End If







    End Sub
#End Region

#Region "DataDoctorLaunch"

    Public Sub LaunchDataDoctorWindow()

        '  Dim datadocwindow As New DataDoctorwindow
        '  datadocwindow.ShowDialog()



    End Sub


#End Region

    Sub launchDateExclusionSettings()
        MessageBox.Show("This icon appears when certain time spans have been EXCLUDED from each day that has been analyzed by prstory. To disable this feature, return to the main menu and click Settings -> Pick Date Range. Then select the icon in the bottom right corner of the menu, uncheck the checkbox and confirm your selection.", "DAILY TIME SPAN EXCLUSION ACTIVE", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Exclamation)
    End Sub

#Region "Team Results"

    Private Sub LaunchTeamResults(sender As Object, e As MouseButtonEventArgs)

        Dim i As Integer
        Dim listofteamreports_raw As New List(Of prStoryMainPageReport)
        Dim multireportwindow As Window_Multiline
        Dim paramObj(10) As Object
        Dim listofteamnames As New List(Of String)

        prStoryReport.updateProductList(DowntimeField.Team)
        prStoryReport.ProductList.Sort()

        paramObj(0) = linename_label.Content
        paramObj(1) = pr_label.Content & " PR"
        paramObj(4) = Stops_Label.Content & " Stops/Day"
        paramObj(2) = UPDT_datalabel1.Content 'UPDT %
        paramObj(3) = PDT_datalabel1.Content 'PDT %
        paramObj(5) = MTBF_Label.Content & " MTBF"
        paramObj(6) = Math.Round(prStoryReport.schedTime, 0)
        paramObj(7) = Math.Round(prStoryReport.MSU / 1000, 2)
        paramObj(9) = cases_label.Content & " cases"

        Try
            resetskulistbox()

            For i = 0 To prStoryReport.ProductList.Count - 1
                If Not prStoryReport.ProductList(i) = Nothing Then
                    Dim tempprstoryreportforteams As New prStoryMainPageReport(selectedindexofLine_temp, starttimeselected, endtimeselected)
                    'tempprstoryreportforteams = prStoryReport
                    AllProdLines(selectedindexofLine_temp).BrandCodesWeWant.Add(CStr(prStoryReport.ProductList(i)))
                    listofteamnames.Add(CStr(prStoryReport.ProductList(i)))
                    tempprstoryreportforteams.reFilterData_Team(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant, True)
                    AllProdLines(selectedindexofLine_temp).reFilterData_Team(AllProdLines(selectedindexofLine_temp).BrandCodesWeWant)
                    listofteamreports_raw.Add(tempprstoryreportforteams)

                    AllProdLines(selectedindexofLine_temp).BrandCodesWeWant.Clear()
                    ' tempprstoryreportforteams.reFilterData_ClearAllFilters()
                    'prStoryReport.reFilterData_ClearAllFilters()
                    AllProdLines(selectedindexofLine_temp).rawDowntimeData.reFilterData_ClearAllFilters()

                End If
            Next
            paramObj(8) = listofteamnames
            multireportwindow = New Window_Multiline(True, selectedindexofLine_temp, listofteamreports_raw, MainDateLabel.Content.ToString, paramObj)
            multireportwindow.Show()
        Catch
            MessageBox.Show("Team Analysis is unavailable due to insufficient number of Scheduled Events in the raw data.")
        End Try
    End Sub

#End Region

    Public Sub reinitializePublicVariables()
        topstopsbar1 = 0
        topstopsbar2 = 0
        topstopsbar3 = 0
        topstopsbar4 = 0
        topstopsbar5 = 0
        topstopsbar6 = 0
        topstopsbar7 = 0
        topstopsbar8 = 0
        topstopsbar9 = 0
        topstopsbar10 = 0
        topstopsbar11 = 0
        topstopsbar12 = 0
        topstopsbar13 = 0
        topstopsbar14 = 0
        topstopsbar15 = 0

        topstopsbar1_PRloss = 0
        topstopsbar2_PRloss = 0
        topstopsbar3_PRloss = 0
        topstopsbar4_PRloss = 0
        topstopsbar5_PRloss = 0
        topstopsbar6_PRloss = 0
        topstopsbar7_PRloss = 0
        topstopsbar8_PRloss = 0
        topstopsbar9_PRloss = 0
        topstopsbar10_PRloss = 0
        topstopsbar11_PRloss = 0
        topstopsbar12_PRloss = 0
        topstopsbar13_PRloss = 0
        topstopsbar14_PRloss = 0
        topstopsbar15_PRloss = 0

        topstopsbar1_SPD = 0
        topstopsbar2_SPD = 0
        topstopsbar3_SPD = 0
        topstopsbar4_SPD = 0
        topstopsbar5_SPD = 0
        topstopsbar6_SPD = 0
        topstopsbar7_SPD = 0
        topstopsbar8_SPD = 0
        topstopsbar9_SPD = 0
        topstopsbar10_SPD = 0
        topstopsbar11_SPD = 0
        topstopsbar12_SPD = 0
        topstopsbar13_SPD = 0
        topstopsbar14_SPD = 0
        topstopsbar15_SPD = 0

        updtlabel1string = ""
        updtlabel2string = ""
        updtlabel3string = ""
        updtlabel4string = ""
        updtlabel5string = ""
        updtlabel6string = ""
        UPDTbar1 = 0
        UPDTbar2 = 0
        UPDTbar3 = 0
        UPDTbar4 = 0
        UPDTbar5 = 0
        UPDTbar6 = 0
        UPDTbar1_PRloss = 0
        UPDTbar2_PRloss = 0
        UPDTbar3_PRloss = 0
        UPDTbar4_PRloss = 0
        UPDTbar5_PRloss = 0
        UPDTbar6_PRloss = 0

        pdtlabel1string = ""
        pdtlabel2string = ""
        pdtlabel3string = ""
        pdtlabel4string = ""
        pdtlabel5string = ""
        pdtlabel6string = ""
        pdtlabel7string = ""
        pdtlabel8string = ""
        pdtlabel9string = ""
        PDTbar1 = 0
        PDTbar2 = 0
        PDTbar3 = 0
        PDTbar4 = 0
        PDTbar5 = 0
        PDTbar6 = 0
        PDTbar7 = 0
        PDTbar8 = 0
        PDTbar9 = 0
        PDTbar1_PRloss = 0
        PDTbar2_PRloss = 0
        PDTbar3_PRloss = 0
        PDTbar4_PRloss = 0
        PDTbar5_PRloss = 0
        PDTbar6_PRloss = 0
        PDTbar7_PRloss = 0
        PDTbar8_PRloss = 0
        PDTbar9_PRloss = 0



        changeoverlabel1string = ""
        changeoverlabel2string = ""
        changeoverlabel3string = ""
        changeoverlabel4string = ""
        changeoverlabel5string = ""
        changeoverlabel6string = ""
        changeoverlabel7string = ""
        Changeoverbar1 = 0
        Changeoverbar2 = 0
        Changeoverbar3 = 0
        Changeoverbar4 = 0
        Changeoverbar5 = 0
        Changeoverbar6 = 0
        Changeoverbar7 = 0
        Changeover1_PRloss = 0
        Changeover2_PRloss = 0
        Changeover3_PRloss = 0
        Changeover4_PRloss = 0
        Changeover5_PRloss = 0
        Changeover6_PRloss = 0
        Changeover7_PRloss = 0
        Changeovertime1 = 0
        Changeovertime2 = 0
        Changeovertime3 = 0
        Changeovertime4 = 0
        Changeovertime5 = 0
        Changeovertime6 = 0
        Changeovertime7 = 0
        changeoverNo_of_Events1 = 0
        changeoverNo_of_Events2 = 0
        changeoverNo_of_Events3 = 0
        changeoverNo_of_Events4 = 0
        changeoverNo_of_Events5 = 0
        changeoverNo_of_Events6 = 0
        changeoverNo_of_Events7 = 0



        EquipmentLabel1string = ""
        EquipmentLabel2string = ""
        EquipmentLabel3string = ""
        EquipmentLabel4string = ""
        EquipmentLabel5string = ""
        EquipmentLabel6string = ""
        EquipMainbar1 = 0
        EquipMainbar2 = 0
        EquipMainbar3 = 0
        EquipMainbar4 = 0
        EquipMainbar5 = 0
        EquipMainbar6 = 0
        EquipMain1_PRloss = 0
        EquipMain2_PRloss = 0
        EquipMain3_PRloss = 0
        EquipMain4_PRloss = 0
        EquipMain5_PRloss = 0
        EquipMain6_PRloss = 0



        Equip1Label1string = ""
        Equip1Label2string = ""
        Equip1Label3string = ""
        Equip1bar1 = 0
        Equip1bar2 = 0
        Equip1bar3 = 0
        Equip1_1_PRloss = 0
        Equip1_2_PRloss = 0
        Equip1_3_PRloss = 0


        Equip2Label1string = ""
        Equip2Label2string = ""
        Equip2Label3string = ""
        Equip2bar1 = 0
        Equip2bar2 = 0
        Equip2bar3 = 0
        Equip2_1_PRloss = 0
        Equip2_2_PRloss = 0
        Equip2_3_PRloss = 0



        Equip3Label1string = ""
        Equip3Label2string = ""
        Equip3Label3string = ""
        Equip3bar1 = 0
        Equip3bar2 = 0
        Equip3bar3 = 0
        Equip3_1_PRloss = 0
        Equip3_2_PRloss = 0
        Equip3_3_PRloss = 0



        Equip21label1string = ""
        Equip21label2string = ""
        Equip21bar1 = 0
        Equip21bar2 = 0
        Equip21_1_PRloss = 0
        Equip21_2_PRloss = 0




        Equip22label1string = ""
        Equip22label2string = ""
        Equip22bar1 = 0
        Equip22bar2 = 0
        Equip22_1_PRloss = 0
        Equip22_2_PRloss = 0

        MaxPRindataset = 0
        MaxPRindataset_planned = 0
        MaxStopsindataset = 0
        MaxPRlossindataset_stops = 0
        MaxChangeoverBubble_COtime = 0

        selectedDTGroup = ""
        selectedRL4bar = ""
        selectedRLcolumn = 0

        selectedfailuremode = ""
        datalabelcontent = ""
    End Sub

    Private Sub Initialize_UseTrack_Variables()

        UseTrack_UPDTview = False
        UseTrack_PDTview = False
        UseTrack_PROverallTrends = False
        UseTrack_RawDatawindow_Main = False
        UseTrack_RawDatawindow_Paretos = False
        UseTrack_RawDatawindow_Variance = False
        UseTrack_WeibullMain = False
        UseTrack_WeibullMain_failuremodes = False
        UseTrack_IncontrolMain = False
        UseTrack_IncontrolControlChart = False
        UseTrack_IncontrolControlShift = False
        UseTrack_TopStopsMain = False
        UseTrack_StopsWatchMain = False
        UseTrack_TopStopsTrends = False
        UseTrack_ChangeMapping = False
        UseTrack_Filter = False
        UseTrack_ExportLossTree = False
        UseTrack_ExportDowntime = False
        UseTrack_ExportProduction = False
        UseTrack_ExportDependency = False
        UseTrack_Notes = False
        UseTrack_Simulation = False
        UseTrack_Notes_PickaLoss = False
        UseTrack_Notes_ExporttoExcel = False
        UseTrack_TargetsMain = False
        UseTrack_RawDataWindow_Trends = False
    End Sub

#Region "LiveLine"
    Private Sub ShowLiveLineControl()
        '  hidemenu'(true)
        ' SplashRectangle.Visibility = Visibility.Hidden
        '  livelinecontrol = new UserControls.Control_LiveLine()
        LiveLineControl.initialize(prStoryReport.MainLEDSReport, prStoryReport.MainLEDSReport)
        LiveLineControl.Visibility = Visibility.Visible
    End Sub
    Private Sub LiveLineLaunchButton_MouseDown(sender As Object, e As MouseButtonEventArgs)
        ShowLiveLineControl()
    End Sub
#End Region
End Class
