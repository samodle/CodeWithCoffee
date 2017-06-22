Imports System.Threading
Imports System.Net

Public Class WindowTargets
    Private Sub Targets_Loaded()
        TargetsLanguageSet()
       ' StartSync() 'dont need to do this now that targets are an embedded image
        SyncLabel.Visibility = Windows.Visibility.Hidden
        SyncImage.Visibility = Windows.Visibility.Hidden
    End Sub
    Private Sub TargetsLanguageSet()
        Select Case My.Settings.LanguageActive
            Case Lang.Chinese_Simplified
                TitleLabel.Content = "设定生产线目标"
                SyncLabel.Content = "云端同步"
                SaveTargetsButton.Content = "储存目标"

            Case Else
                TitleLabel.Content = "Set Line Targets"
                SyncLabel.Content = "Sync Targets with Cloud"
                SaveTargetsButton.Content = "Save Targets"

        End Select


    End Sub
#Region "Variables & Properties"
    Private parentLine As ProdLine
    Private AllCurrentTargets As New List(Of PRTarget)

    Private Unplanned_Tier1_Targets As New List(Of PRTarget)
    Private Unplanned_Tier2_Targets As New List(Of PRTarget)
    Private Unplanned_Tier3_Targets As New List(Of PRTarget)

    Private Planned_Tier1_Targets As New List(Of PRTarget)
    Private Planned_Tier2_Targets As New List(Of PRTarget)

    Private HaveIChangedAnything As Boolean = False
    Private UploadToCloud As Boolean = False

    'Private Const MAX_TARGETS_PERCARD_COUNT As Integer = 14
#End Region

#Region "Construction / Destruction"
    Public Sub New(IparentLine As ProdLine)
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        LineLabel.Content = IparentLine.ToString
        parentLine = IparentLine

        'initialize the dropdown
        With CardSelectDropdown.Items
            .Add("Card Selection")
            .Add("Unplanned Top Card")
            .Add("Planned Top Card")
            .Add("Unplanned Second Card")
            .Add("Unplanned Other")
            .Add("Planned Other")
        End With
        CardSelectDropdown.SelectedIndex = 0 'fyi this also triggers the selection changed event

        'bring in the targets
        If parentLine.doIhaveTargets Then findAllTargets()
    End Sub

    Private Sub findAllTargets()
        With parentLine.DowntimePercentTargets
            For i As Integer = 0 To .RawTargetList.Count - 1
                Select Case .RawTargetList(i).Card
                    Case prStoryCard.Unplanned
                        Unplanned_Tier1_Targets.Add(.RawTargetList(i))
                    Case prStoryCard.Equipment
                        Unplanned_Tier2_Targets.Add(.RawTargetList(i))
                    Case prStoryCard.Equipment_One
                        Unplanned_Tier3_Targets.Add(.RawTargetList(i))
                    Case prStoryCard.Planned
                        Planned_Tier1_Targets.Add(.RawTargetList(i))
                    Case prStoryCard.Changeover
                        Planned_Tier2_Targets.Add(.RawTargetList(i))
                End Select
            Next
        End With
    End Sub

    Private Sub writeThatDataOnClose() Handles Me.Closed
        Dim writeDataThread As Thread
        If HaveIChangedAnything Then
            writeDataThread = New Thread(AddressOf writeAllTargets)

            writeDataThread.Start()
        End If

    End Sub
#End Region

#Region "Pull / Push Targets"
    Sub CardDropdown_SelectionChanged()
        Dim cardSelected As Integer
        saveCheck.Visibility = Windows.Visibility.Hidden
        Select Case CardSelectDropdown.SelectedIndex
            Case 1 'unplanned
                cardSelected = prStoryCard.Unplanned
            Case 2 'planned
                cardSelected = prStoryCard.Planned
            Case 3 'unplanned 2
                cardSelected = prStoryCard.Equipment
            Case 4 'unplanned other
                cardSelected = prStoryCard.Equipment_One
            Case 5 'planned other
                cardSelected = prStoryCard.Changeover
            Case Else
                cardSelected = -1
        End Select
        resetAllInputs() 'clear it
        If cardSelected > 0 Then pullTargets(cardSelected) 'write it - only if you want it
    End Sub

    Private Sub pullTargets(cardNum As Integer)
        Dim targetsFound As Integer
        Select Case cardNum
            Case prStoryCard.Unplanned
                For targetsFound = 0 To Unplanned_Tier1_Targets.Count - 1
                    sendTargetToUI(targetsFound + 1, Unplanned_Tier1_Targets(targetsFound).targetName, Unplanned_Tier1_Targets(targetsFound).targetValue * 100)
                Next
            Case prStoryCard.Equipment
                For targetsFound = 0 To Unplanned_Tier2_Targets.Count - 1
                    sendTargetToUI(targetsFound + 1, Unplanned_Tier2_Targets(targetsFound).targetName, Unplanned_Tier2_Targets(targetsFound).targetValue * 100)
                Next
            Case prStoryCard.Equipment_One
                For targetsFound = 0 To Unplanned_Tier3_Targets.Count - 1
                    sendTargetToUI(targetsFound + 1, Unplanned_Tier3_Targets(targetsFound).targetName, Unplanned_Tier3_Targets(targetsFound).targetValue * 100)
                Next
            Case prStoryCard.Planned
                For targetsFound = 0 To Planned_Tier1_Targets.Count - 1
                    sendTargetToUI(targetsFound + 1, Planned_Tier1_Targets(targetsFound).targetName, Planned_Tier1_Targets(targetsFound).targetValue * 100)
                Next
            Case prStoryCard.Changeover
                For targetsFound = 0 To Planned_Tier2_Targets.Count - 1
                    sendTargetToUI(targetsFound + 1, Planned_Tier2_Targets(targetsFound).targetName, Planned_Tier2_Targets(targetsFound).targetValue * 100)
                Next
        End Select
    End Sub
    Private Sub sendTargetToUI(targetNum As Integer, fieldName As String, tgtVal As String)
        Select Case targetNum
            Case 1
                A_FieldName.Text = fieldName
                A_Value.Text = tgtVal
                A_FieldName.FontStyle = FontStyles.Normal
                A_Value.FontStyle = FontStyles.Normal
            Case 2
                B_FieldName.Text = fieldName
                B_Value.Text = tgtVal
                B_FieldName.FontStyle = FontStyles.Normal
                B_Value.FontStyle = FontStyles.Normal
            Case 3
                C_FieldName.Text = fieldName
                C_Value.Text = tgtVal
                C_FieldName.FontStyle = FontStyles.Normal
                C_Value.FontStyle = FontStyles.Normal
            Case 4
                D_FieldName.Text = fieldName
                D_Value.Text = tgtVal
                D_FieldName.FontStyle = FontStyles.Normal
                D_Value.FontStyle = FontStyles.Normal
            Case 5
                E_FieldName.Text = fieldName
                E_Value.Text = tgtVal
                E_FieldName.FontStyle = FontStyles.Normal
                E_Value.FontStyle = FontStyles.Normal
            Case 6
                F_FieldName.Text = fieldName
                F_Value.Text = tgtVal
                F_FieldName.FontStyle = FontStyles.Normal
                F_Value.FontStyle = FontStyles.Normal
            Case 7
                G_FieldName.Text = fieldName
                G_Value.Text = tgtVal
                G_FieldName.FontStyle = FontStyles.Normal
                G_Value.FontStyle = FontStyles.Normal
            Case 8
                A_FieldName_Copy.Text = fieldName
                A_Value_Copy.Text = tgtVal
                A_FieldName_Copy.FontStyle = FontStyles.Normal
                A_Value_Copy.FontStyle = FontStyles.Normal
            Case 9
                B_FieldName_Copy.Text = fieldName
                B_Value_Copy.Text = tgtVal
                B_FieldName_Copy.FontStyle = FontStyles.Normal
                B_Value_Copy.FontStyle = FontStyles.Normal
            Case 10
                C_FieldName_Copy.Text = fieldName
                C_Value_Copy.Text = tgtVal
                C_FieldName_Copy.FontStyle = FontStyles.Normal
                C_Value_Copy.FontStyle = FontStyles.Normal
            Case 11
                D_FieldName_Copy.Text = fieldName
                D_Value_Copy.Text = tgtVal
                D_FieldName_Copy.FontStyle = FontStyles.Normal
                D_Value_Copy.FontStyle = FontStyles.Normal
            Case 12
                E_FieldName_Copy.Text = fieldName
                E_Value_Copy.Text = tgtVal
                E_FieldName_Copy.FontStyle = FontStyles.Normal
                E_Value_Copy.FontStyle = FontStyles.Normal
            Case 13
                F_FieldName_Copy.Text = fieldName
                F_Value_Copy.Text = tgtVal
                F_FieldName_Copy.FontStyle = FontStyles.Normal
                F_Value_Copy.FontStyle = FontStyles.Normal
            Case 14
                G_FieldName_Copy.Text = fieldName
                G_Value_Copy.Text = tgtVal
                G_FieldName_Copy.FontStyle = FontStyles.Normal
                G_Value_Copy.FontStyle = FontStyles.Normal
            Case Else
                'this is too many targets for now...
        End Select
    End Sub



    Private Sub pushTargets(cardNum As Integer)
        Dim tmpFieldName As String, tmpVal As String, tmpValDBL As Double, i As Integer
        Dim tmpList As New List(Of PRTarget)

        'do this for each of our text boxes
        tmpFieldName = A_FieldName.Text
        tmpVal = A_Value.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = B_FieldName.Text
        tmpVal = B_Value.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = C_FieldName.Text
        tmpVal = C_Value.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = D_FieldName.Text
        tmpVal = D_Value.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = E_FieldName.Text
        tmpVal = E_Value.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = F_FieldName.Text
        tmpVal = F_Value.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = G_FieldName.Text
        tmpVal = G_Value.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = A_FieldName_Copy.Text
        tmpVal = A_Value_Copy.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = B_FieldName_Copy.Text
        tmpVal = B_Value_Copy.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = C_FieldName_Copy.Text
        tmpVal = C_Value_Copy.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = D_FieldName_Copy.Text
        tmpVal = D_Value_Copy.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = E_FieldName_Copy.Text
        tmpVal = E_Value_Copy.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = F_FieldName_Copy.Text
        tmpVal = F_Value_Copy.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        tmpFieldName = G_FieldName_Copy.Text
        tmpVal = G_Value_Copy.Text
        If Not tmpFieldName.Equals("field") And Double.TryParse(tmpVal, tmpValDBL) Then
            If tmpValDBL >= 0.1 And tmpValDBL < 100 Then tmpList.Add(New PRTarget(tmpValDBL / 100, tmpFieldName, cardNum))
        End If

        'alright, now lets actually push it!!
        Select Case cardNum
            Case prStoryCard.Unplanned
                For i = 0 To tmpList.Count - 1
                    Unplanned_Tier1_Targets.Add(tmpList(i))
                Next
            Case prStoryCard.Equipment
                For i = 0 To tmpList.Count - 1
                    Unplanned_Tier2_Targets.Add(tmpList(i))
                Next
            Case prStoryCard.Equipment_One
                For i = 0 To tmpList.Count - 1
                    Unplanned_Tier3_Targets.Add(tmpList(i))
                Next
            Case prStoryCard.Planned
                For i = 0 To tmpList.Count - 1
                    Planned_Tier1_Targets.Add(tmpList(i))
                Next
            Case prStoryCard.Changeover
                For i = 0 To tmpList.Count - 1
                    Planned_Tier2_Targets.Add(tmpList(i))
                Next
        End Select

    End Sub

    'send targets from lists to parent
    Private Sub saveTargets()
        UseTrack_TargetsMain = True
        Dim newTargets As DTPct_Targets, i As Integer
        newTargets = New DTPct_Targets(parentLine.Name, parentLine.SiteName)

        For i = 0 To Unplanned_Tier1_Targets.Count - 1
            With Unplanned_Tier1_Targets(i)
                newTargets.addNewTarget(.targetName, .Card, .targetValue)
            End With
        Next

        For i = 0 To Unplanned_Tier2_Targets.Count - 1
            With Unplanned_Tier2_Targets(i)
                newTargets.addNewTarget(.targetName, .Card, .targetValue)
            End With
        Next

        For i = 0 To Unplanned_Tier3_Targets.Count - 1
            With Unplanned_Tier3_Targets(i)
                newTargets.addNewTarget(.targetName, .Card, .targetValue)
            End With
        Next

        For i = 0 To Planned_Tier1_Targets.Count - 1
            With Planned_Tier1_Targets(i)
                newTargets.addNewTarget(.targetName, .Card, .targetValue)
            End With
        Next

        For i = 0 To Planned_Tier2_Targets.Count - 1
            With Planned_Tier2_Targets(i)
                newTargets.addNewTarget(.targetName, .Card, .targetValue)
            End With
        Next

        parentLine.DowntimePercentTargets = newTargets
    End Sub
#End Region

    Sub resetAllInputs()
        A_FieldName.Text = "field"
        A_Value.Text = "%"
        A_FieldName.FontStyle = FontStyles.Italic
        A_Value.FontStyle = FontStyles.Italic

        B_FieldName.Text = "field"
        B_Value.Text = "%"
        B_FieldName.FontStyle = FontStyles.Italic
        B_Value.FontStyle = FontStyles.Italic

        C_FieldName.Text = "field"
        C_Value.Text = "%"
        C_FieldName.FontStyle = FontStyles.Italic
        C_Value.FontStyle = FontStyles.Italic

        D_FieldName.Text = "field"
        D_Value.Text = "%"
        D_FieldName.FontStyle = FontStyles.Italic
        D_Value.FontStyle = FontStyles.Italic

        E_FieldName.Text = "field"
        E_Value.Text = "%"
        E_FieldName.FontStyle = FontStyles.Italic
        E_Value.FontStyle = FontStyles.Italic

        F_FieldName.Text = "field"
        F_Value.Text = "%"
        F_FieldName.FontStyle = FontStyles.Italic
        F_Value.FontStyle = FontStyles.Italic

        G_FieldName.Text = "field"
        G_Value.Text = "%"
        G_FieldName.FontStyle = FontStyles.Italic
        G_Value.FontStyle = FontStyles.Italic

        A_FieldName_Copy.Text = "field"
        A_Value_Copy.Text = "%"
        A_FieldName_Copy.FontStyle = FontStyles.Italic
        A_Value_Copy.FontStyle = FontStyles.Italic

        B_FieldName_Copy.Text = "field"
        B_Value_Copy.Text = "%"
        B_FieldName_Copy.FontStyle = FontStyles.Italic
        B_Value_Copy.FontStyle = FontStyles.Italic

        C_FieldName_Copy.Text = "field"
        C_Value_Copy.Text = "%"
        C_FieldName_Copy.FontStyle = FontStyles.Italic
        C_Value_Copy.FontStyle = FontStyles.Italic

        D_FieldName_Copy.Text = "field"
        D_Value_Copy.Text = "%"
        D_FieldName_Copy.FontStyle = FontStyles.Italic
        D_Value_Copy.FontStyle = FontStyles.Italic

        E_FieldName_Copy.Text = "field"
        E_Value_Copy.Text = "%"
        E_FieldName_Copy.FontStyle = FontStyles.Italic
        E_Value_Copy.FontStyle = FontStyles.Italic

        F_FieldName_Copy.Text = "field"
        F_Value_Copy.Text = "%"
        F_FieldName_Copy.FontStyle = FontStyles.Italic
        F_Value_Copy.FontStyle = FontStyles.Italic

        G_FieldName_Copy.Text = "field"
        G_Value_Copy.Text = "%"
        G_FieldName_Copy.FontStyle = FontStyles.Italic
        G_Value_Copy.FontStyle = FontStyles.Italic
    End Sub

    Sub Button_SaveTgts()
        Dim cardSelected As Integer
        Select Case CardSelectDropdown.SelectedIndex
            Case 1 'unplanned
                cardSelected = prStoryCard.Unplanned
            Case 2 'planned
                cardSelected = prStoryCard.Planned
            Case 3 'unplanned 2
                cardSelected = prStoryCard.Equipment
            Case 4 'unplanned other
                cardSelected = prStoryCard.Equipment_One
            Case 5 'planned other
                cardSelected = prStoryCard.Changeover
            Case Else
                cardSelected = -1
        End Select
        If cardSelected > -1 Then
            HaveIChangedAnything = True
            pushTargets(cardSelected)
            saveTargets()
            saveCheck.Visibility = Windows.Visibility.Visible
            If My.Settings.AdvancedSettings_isPasswordCorrect Then
              '  SyncLabel.Visibility = Windows.Visibility.Visible
              '  SyncImage.Visibility = Windows.Visibility.Visible
            End If
        End If
    End Sub

#Region "Mouse Move / Click"
    Private Sub IconMouseMove(sender As Object, e As MouseEventArgs)
        CardSelectDropdown.Opacity = 0.9
    End Sub

    Private Sub IconMouseLeave(sender As Object, e As MouseEventArgs)
        CardSelectDropdown.Opacity = 0.9
    End Sub

    Private Sub GeneralMouseMove(sender As Object, e As MouseEventArgs)
        sender.Opacity = 0.8
    End Sub

    Private Sub GeneralMouseLeave(sender As Object, e As MouseEventArgs)
        sender.Opacity = 1.0
    End Sub


#End Region
#Region "Textbox Select Event Handlers"
    Sub A_EventTextboxSelect() Handles A_FieldName.GotFocus
        A_FieldName.FontStyle = FontStyles.Normal
        A_FieldName.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub B_EventTextboxSelect() Handles B_FieldName.GotFocus
        B_FieldName.FontStyle = FontStyles.Normal
        B_FieldName.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub C_EventTextboxSelect() Handles C_FieldName.GotFocus
        C_FieldName.FontStyle = FontStyles.Normal
        C_FieldName.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub D_EventTextboxSelect() Handles D_FieldName.GotFocus
        D_FieldName.FontStyle = FontStyles.Normal
        D_FieldName.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub E_EventTextboxSelect() Handles E_FieldName.GotFocus
        E_FieldName.FontStyle = FontStyles.Normal
        E_FieldName.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub F_EventTextboxSelect() Handles F_FieldName.GotFocus
        F_FieldName.FontStyle = FontStyles.Normal
        F_FieldName.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub G_EventTextboxSelect() Handles G_FieldName.GotFocus
        G_FieldName.FontStyle = FontStyles.Normal
        G_FieldName.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub

    Sub A_EventTextboxSelect_Val() Handles A_Value.GotFocus
        A_Value.FontStyle = FontStyles.Normal
        A_Value.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub B_EventTextboxSelect_Val() Handles B_Value.GotFocus
        B_Value.FontStyle = FontStyles.Normal
        B_Value.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub C_EventTextboxSelect_Val() Handles C_Value.GotFocus
        C_Value.FontStyle = FontStyles.Normal
        C_Value.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub D_EventTextboxSelect_Val() Handles D_Value.GotFocus
        D_Value.FontStyle = FontStyles.Normal
        D_Value.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub E_EventTextboxSelect_Val() Handles E_Value.GotFocus
        E_Value.FontStyle = FontStyles.Normal
        E_Value.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub F_EventTextboxSelect_Val() Handles F_Value.GotFocus
        F_Value.FontStyle = FontStyles.Normal
        F_Value.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub G_EventTextboxSelect_Val() Handles G_Value.GotFocus
        G_Value.FontStyle = FontStyles.Normal
        G_Value.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub

    Sub A_EventTextboxSelect_Copy() Handles A_FieldName_Copy.GotFocus
        A_FieldName_Copy.FontStyle = FontStyles.Normal
        A_FieldName_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub B_EventTextboxSelect_Copy() Handles B_FieldName_Copy.GotFocus
        B_FieldName_Copy.FontStyle = FontStyles.Normal
        B_FieldName_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub C_EventTextboxSelect_Copy() Handles C_FieldName_Copy.GotFocus
        C_FieldName_Copy.FontStyle = FontStyles.Normal
        C_FieldName_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub D_EventTextboxSelect_Copy() Handles D_FieldName_Copy.GotFocus
        D_FieldName_Copy.FontStyle = FontStyles.Normal
        D_FieldName_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub E_EventTextboxSelect_Copy() Handles E_FieldName_Copy.GotFocus
        E_FieldName_Copy.FontStyle = FontStyles.Normal
        E_FieldName_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub F_EventTextboxSelect_Copy() Handles F_FieldName_Copy.GotFocus
        F_FieldName_Copy.FontStyle = FontStyles.Normal
        F_FieldName_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub G_EventTextboxSelect_Copy() Handles G_FieldName_Copy.GotFocus
        G_FieldName_Copy.FontStyle = FontStyles.Normal
        G_FieldName_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub

    Sub A_EventTextboxSelect_Copy_Val() Handles A_Value_Copy.GotFocus
        A_Value_Copy.FontStyle = FontStyles.Normal
        A_Value_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub B_EventTextboxSelect_Copy_Val() Handles B_Value_Copy.GotFocus
        B_Value_Copy.FontStyle = FontStyles.Normal
        B_Value_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub C_EventTextboxSelect_Copy_Val() Handles C_Value_Copy.GotFocus
        C_Value_Copy.FontStyle = FontStyles.Normal
        C_Value_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub D_EventTextboxSelect_Copy_Val() Handles D_Value_Copy.GotFocus
        D_Value_Copy.FontStyle = FontStyles.Normal
        D_Value_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub E_EventTextboxSelect_Copy_Val() Handles E_Value_Copy.GotFocus
        E_Value_Copy.FontStyle = FontStyles.Normal
        E_Value_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub F_EventTextboxSelect_Copy_Val() Handles F_Value_Copy.GotFocus
        F_Value_Copy.FontStyle = FontStyles.Normal
        F_Value_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub G_EventTextboxSelect_Copy_Val() Handles G_Value_Copy.GotFocus
        G_Value_Copy.FontStyle = FontStyles.Normal
        G_Value_Copy.Text = ""
        saveCheck.Visibility = Windows.Visibility.Hidden
    End Sub
#End Region

#Region "Target Creation"
    'send targets from all parents to list
    Sub writeAllTargets()
        Dim fsT As Object
        Dim lineIncrementer As Integer, targetIncrementer As Integer, tmpWriteString As String

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        For lineIncrementer = 0 To AllProdLines.Count - 1
            If AllProdLines(lineIncrementer).doIhaveTargets Then
                With AllProdLines(lineIncrementer).DowntimePercentTargets
                    For targetIncrementer = 0 To .RawTargetList.Count - 1
                        tmpWriteString = AllProdLines(lineIncrementer).Name & "::" & AllProdLines(lineIncrementer).SiteName & "," & .RawTargetList(targetIncrementer).Card & "," & .RawTargetList(targetIncrementer).targetName & "," & .RawTargetList(targetIncrementer).targetValue
                        fsT.writetext(tmpWriteString & vbCrLf)
                    Next
                End With
            End If
        Next

        'fin
        fsT.SaveToFile(PATH_PRSTORY_TARGETS & FILE_RAWTARGETS_CSV, 2) 'Save binary data To disk
        fsT = Nothing

        CSV_readTargetsFile()
      '  If UploadToCloud Then UploadSync()
    End Sub
#End Region

#Region "Sync"
    Sub StartSync()
        Try
            If My.Computer.Network.Ping("prstory.pg.com") Then
                If CheckIfFtpFileExists("ftp://prstory.pg.com/log.txt", "normalusers", "pgdigitalfactory412") Then
                    Dim sourcewebaddress As Uri = New Uri("http://prstory.pg.com/prstory_targets/prstory_dtpct_targets.csv")
                    Dim destinationfolderaddress As String = "C:/Users/Public/prstory/targets/prstory_dtpct_targets.csv"

                    Dim myWebClient As New WebClient()

                    My.Computer.Network.DownloadFile(sourcewebaddress, destinationfolderaddress, "", "", False, 10000, True)

                End If
            End If
        Catch ex As Exception
            MsgBox("Target Download Error " + ex.Message)
            Exit Sub

        End Try

    End Sub

    Sub TriggerUpload()
        UploadToCloud = True
    End Sub


    Sub UploadSync()
        Try
            If My.Computer.Network.Ping("prstory.pg.com") Then
                If CheckIfFtpFileExists("ftp://prstory.pg.com/log.txt", "normalusers", "pgdigitalfactory412") Then
                    Dim myWebClient As New WebClient()
                    Dim sourcewebaddress As Uri = New Uri(" ftp://prstory.pg.com/prstory_targets/prstory_dtpct_targets.csv")
                    Dim destinationfolderaddress As String = "C:/Users/Public/prstory/targets/prstory_dtpct_targets.csv"

                    Dim request As System.Net.FtpWebRequest = DirectCast(WebRequest.Create("ftp://prstory.pg.com/prstory_targets/prstory_dtpct_targets.csv"), System.Net.FtpWebRequest)
                    request.Credentials = New System.Net.NetworkCredential("normalusers", "pgdigitalfactory412")
                    request.Method = System.Net.WebRequestMethods.Ftp.UploadFile

                    Dim files() As Byte = System.IO.File.ReadAllBytes(destinationfolderaddress)

                    Dim strz As System.IO.Stream = request.GetRequestStream()
                    strz.Write(files, 0, files.Length)
                    strz.Close()
                    strz.Dispose()

                Else
                    MsgBox("We tried to sync the targets with the cloud, but it seems the cloud is not available. We have retained the target file locally, and you can try syncing it later.")
                End If
            Else
                MsgBox("We tried to sync the targets with the cloud, but it seems the cloud is not available. We have retained the target file locally, and you can try syncing it later. (Unable to ping prstory.pg.com)")


            End If

        Catch ex As Exception
            MsgBox("Cloud sync for targets is currently unavailable. We are working to get this back up as quickly as possible. " + ex.Message)
            Exit Sub
        End Try
    End Sub
#End Region

End Class
