Imports System.Collections.ObjectModel
Imports System.Threading

Imports System.Deployment.Application
Imports System.Net
Imports System.DirectoryServices
Imports System.ComponentModel
Imports System.Security.Cryptography
Imports System.Configuration
Imports System.Device.Location
Imports Newtonsoft.Json
Imports System.Threading.Tasks
Imports System.IO
Imports System.Xml


Class WindowMain_prstory
    Private prstoryReport As prStoryMainPageReport
    Private UseDemoData As Boolean = False

    Private Sub ShowFamilyCareVideo()
        Dim url As String = "https://pgtube.pg.com/pgtube/play.php?id=71171-f070f31f5125feee"

        Process.Start(url)
    End Sub

#Region "Threads / Raw Data"
    Dim importTargetsThread As Thread

    Dim snakeThread As Thread
    Dim analysisThread As Thread
    Dim getDTdataThread As Thread
    Dim getDTdataThread2 As Thread
    Dim getDTdataThread3 As Thread
    Dim getRateLossdataThread As Thread
    Dim getProdDataThread As Thread
    Dim progressBarThread As Thread
    Dim inControlThread As Thread
    Dim motionStopsThread As Thread
    Dim motionPRThread As Thread
    '  Dim motionprstoryThread As Thread
    Dim uptimeViewerThread As Thread
    Dim congratulationmessageThread As Thread
    'Partial Data Arrays
    Dim first10daysDT(,) As Object 'Array
    Dim second10daysDT(,) As Object 'Array
    Dim third10daysDT(,) As Object 'Array
    Dim rateLossDT(,) As Object
    Dim tmpProdArray As Array
    'data for rate loss
    Dim finalRateLossData(,) As Object
    Dim multilinethread As Thread
    Dim useThreadingForDataPulling As Boolean = True
    Dim isDTcutoffchanged As Boolean = False
    Dim doAnalyzeCheckpoint As Boolean = False
    Public _LinesList As New ObservableCollection(Of String)()
    Dim DoWeAlwaysRepullDataNoMatterWhatElseHasHappened As Boolean = False

    Dim individualRateData As New List(Of Object(,))
    Public Function getRandomLineName() As String
        Dim rnd = New Random()
        Dim randomLineName = _LinesList(rnd.Next(0, _LinesList.Count))
        'needs to be < 20 charts
        While randomLineName.Length > 19
            randomLineName = _LinesList(rnd.Next(0, _LinesList.Count))
        End While

        Return randomLineName
    End Function

    Public Function getRandomAPRILFOOLSLineName() As String

        Dim _LinesList_APR As New List(Of String)
        _LinesList_APR.Add("X-Wing Assembly Bay 1")
        _LinesList_APR.Add("Death Star Reactor Room")
        _LinesList_APR.Add("LexCorp Converting 17C")
        '_LinesList_APR.Add("Stark Industries Line 7")
        _LinesList_APR.Add("Stark Industries Line 9")
        _LinesList_APR.Add("Area 51")
        _LinesList_APR.Add("Buy n Large Batteries")
        _LinesList_APR.Add("SPECTRE Line 007")
        _LinesList_APR.Add("Acme Corp Anvil Assembly")
        _LinesList_APR.Add("Wonka Gobstoppers 13")
        _LinesList_APR.Add("Platform 9 3/4")
        _LinesList_APR.Add("Wayne Enterprises BioTech")
        _LinesList_APR.Add("Globex Corp Windows")
        ' _LinesList_APR.Add("Nova Corps")


        Dim rnd = New Random()
        Dim randomLineName = _LinesList_APR(rnd.Next(0, _LinesList_APR.Count))
        'needs to be < 20 charts
        '  While randomLineName.Length > 23
        '  randomLineName = _LinesList(rnd.Next(0, _LinesList_APR.Count))
        '  End While

        Return randomLineName
    End Function


#End Region

#Region "Sector/Site Selection"
    Dim tmpSector As BusinessUnit 'As New BusinessUnit("All Sectors")
    Dim tmpSite As ProdSite 'New productionSite("All Sites", "", "")
    Dim activeLineIndeces As New List(Of Integer)
    Dim PROF_connectionError As Boolean
    Dim PROF_secondaryConnectionError As Boolean
    Private _SectorCollection As New ObservableCollection(Of BusinessUnit)()
    Public ReadOnly Property SectorCollection
        Get
            Return _SectorCollection
        End Get
    End Property
    Private _SiteCollection As New ObservableCollection(Of ProdSite)()
    Public ReadOnly Property SiteCollection
        Get
            Return _SiteCollection
        End Get
    End Property
#Region "Splash Canvas"

    Private Sub CloseSplashCanvas()
        SplashCanvas.Visibility = Visibility.Hidden
        My.Settings.ShowUpdateNews = False
    End Sub

    Private Sub BackButtonSplashPressed()

        IntroCanvas.Visibility = Visibility.Visible
        BackButtonSplashCanvas.Visibility = Visibility.Hidden

        WhatsNewCanvas.Visibility = Visibility.Hidden
        SuggestionCanvas.Visibility = Visibility.Hidden
        PRcalculationCanvas.Visibility = Visibility.Hidden

    End Sub

    Private Sub LoadWhatsNewCanvas()
        BackButtonSplashCanvas.Visibility = Visibility.Visible
        IntroCanvas.Visibility = Visibility.Hidden

        WhatsNewCanvas.Visibility = Visibility.Visible

    End Sub
    Private Sub LoadSuggestionCanvas()
        BackButtonSplashCanvas.Visibility = Visibility.Visible
        IntroCanvas.Visibility = Visibility.Hidden


        SuggestionCanvas.Visibility = Visibility.Visible
    End Sub
    Private Sub LoadPRcalculationCanvas()
        BackButtonSplashCanvas.Visibility = Visibility.Visible
        IntroCanvas.Visibility = Visibility.Hidden

        PRcalculationCanvas.Visibility = Visibility.Visible

    End Sub
    Private Sub LoadSplashCanvas()

        Dim webAddress As String = "mailto:das.l@pg.com;odle.so.1@pg.com"
        Process.Start(webAddress)

    End Sub

#End Region
    Private Sub calculateSectorCollection()
        _SectorCollection.Clear()
        _SectorCollection.Add(tmpSector)
        Dim i As Integer
        For i = 0 To AllProductionSectors.Count - 1
            _SectorCollection.Add(AllProductionSectors(i))
        Next
    End Sub
    Private Sub calculateSiteCollection()
        _SiteCollection.Clear()
        _SiteCollection.Add(tmpSite)
        Dim i As Integer
        For i = 0 To AllProductionSites.Count - 1
            _SiteCollection.Add(AllProductionSites(i))
        Next
    End Sub
    Private Sub ShowSeachBarandHideDropDown()
        UseLineDropDown_textblock.Visibility = Visibility.Visible
        UseLineSearchBar_textblock.Visibility = Visibility.Hidden
        LineDropDown_Canvas.Visibility = Visibility.Hidden
        LineSelection_AutoComplete.Visibility = Visibility.Visible
        '   If isAPRILFOOLS Then
        '       SectorLabel.Content = SECTOR_APRILFOOLS
        '       SiteLabel.Content = "All Sites"
        '  Else
        SectorLabel.Content = "All Sectors"
        SiteLabel.Content = "All Sites"
        ' End If

        populatelineslist()
        LineSelection_icongreencheck.Visibility = Visibility.Hidden
        LineSelection_iconrederror.Visibility = Visibility.Hidden
    End Sub
    Private Sub ShowDropDownandHideSeachBar()
        UseLineDropDown_textblock.Visibility = Visibility.Hidden
        UseLineSearchBar_textblock.Visibility = Visibility.Visible
        LineDropDown_Canvas.Visibility = Visibility.Visible
        LineSelection_AutoComplete.Visibility = Visibility.Hidden

        figureOutWhichLinesToShow()

        LineSelection_AutoComplete.SelectedItem = ""
        LineSelection_icongreencheck.Visibility = Visibility.Hidden
        LineSelection_iconrederror.Visibility = Visibility.Hidden
    End Sub
    Private Sub figureOutWhichLinesToShow()
        Dim lineIncrementer As Integer
        Try
            prstory_linedropdown.Items.Clear()
        Catch ex As unknownMappingException
            'we good
        End Try
        activeLineIndeces.Clear()
        prstory_linedropdown.Items.Clear()
        Select Case My.Settings.LanguageActive
            Case Lang.English
                prstory_linedropdown.Items.Add("Line Selection")
            Case Lang.German
                prstory_linedropdown.Items.Add("wahlen Sie eine Leitung")
            Case Lang.Spanish
                prstory_linedropdown.Items.Add("Seleccione una linea")
            Case Lang.French
                prstory_linedropdown.Items.Add("Selectionner une ligne")
            Case Lang.Portuguese
                prstory_linedropdown.Items.Add("Seleção de Linha")
            Case Lang.Chinese_Simplified
                prstory_linedropdown.Items.Add("选择生产线")
        End Select
        Try


            '  If Not isAPRILFOOLS Then

            SectorLabel.Content = My.Settings.DefaultBU
            SiteLabel.Content = My.Settings.DefaultSite
            '  End If
        Catch ex As Exception
            SectorLabel.Content = "All Sectors"
            SiteLabel.Content = "All Sites"
            My.Settings.DefaultBU = "All Sectors"
            My.Settings.DefaultSite = "All Sites"
        End Try

        For lineIncrementer = 0 To AllProdLines.Count - 1
            With AllProdLines(lineIncrementer)
                If .SiteName.Equals(SiteLabel.Content) Or SiteLabel.Content.Equals(tmpSite.Name) Then
                    If .Sector.Equals(SectorLabel.Content) Or SectorLabel.Content.Equals(tmpSector.Name) Then
                        If .Sector = BETA_TESTING Then
                            If SectorLabel.Content = BETA_TESTING Then
                                prstory_linedropdown.Items.Add(.ToString)
                                activeLineIndeces.Add(lineIncrementer)
                            End If
                        Else
                            prstory_linedropdown.Items.Add(.ToString)
                            activeLineIndeces.Add(lineIncrementer)
                        End If
                    End If
                End If
            End With
        Next
        prstory_linedropdown.SelectedIndex = 0
    End Sub




    Private WithEvents watcher As GeoCoordinateWatcher
    Public Sub GetLocationDataEvent()
        watcher = New GeoCoordinateWatcher()
        AddHandler watcher.PositionChanged, AddressOf watcher_PositionChanged
        watcher.Start()
    End Sub

    Private Sub watcher_PositionChanged(ByVal sender As Object, ByVal e As GeoPositionChangedEventArgs(Of GeoCoordinate))
        MsgBox(e.Position.Location.Latitude.ToString & ", " &
               e.Position.Location.Longitude.ToString)
        ' Stop receiving updates after the first one.
        'watcher.Stop()
    End Sub

    Public Sub siteSelectedSimple(ByVal sender As ListBox, ByVal e As SelectionChangedEventArgs)
        Dim lbsender As ListBox
        Dim stringString As String

        lbsender = CType(sender, ListBox)
        If Not IsNothing(lbsender.SelectedItem) Then
            stringString = lbsender.SelectedItem.ToString
            SiteMenu.SelectedItems.Clear()
            If (stringString = "All Sites") Then
                SectorLabel.Content = "All Sectors"
                SiteSelected(stringString)
                calculateSiteCollection()
                calculateSectorCollection()
                SectorLabel.Content = "All Sectors"
            Else

                SiteSelected(stringString)
            End If


        End If
    End Sub
    Private Sub SiteSelected(siteName As String)
        Dim tmpSiteName As String, i As Integer, testSector As BusinessUnit
        Dim selectedSite As ProdSite
        Dim moduleIncrementer As Integer
        SiteMenu.Visibility = Visibility.Hidden
        ' get the label right
        tmpSiteName = siteName 'lbsender.SelectedItem.ToString
        SiteLabel.Content = tmpSiteName
        My.Settings.DefaultSite = tmpSiteName
        'find the right sectors
        If tmpSiteName.Equals(tmpSite.Name) Then
            calculateSectorCollection()
            calculateSiteCollection()
        Else
            selectedSite = AllProductionSites(getSiteIndexFromName(tmpSiteName))
            _SectorCollection.Clear()
            _SectorCollection.Add(tmpSector)
            '  For siteIncrementer = 0 To AllProductionSites.Count - 1
            For moduleIncrementer = 0 To selectedSite.ModulesList.Count - 1 'AllProductionSites(siteIncrementer).ModulesList.Count - 1
                testSector = selectedSite.ModulesList(moduleIncrementer).parentSector
                i = _SectorCollection.IndexOf(testSector)
                If i = -1 Then _SectorCollection.Add(testSector)
            Next

            _SiteCollection.Clear()
            _SiteCollection.Add(tmpSite)
            _SiteCollection.Add(selectedSite)
            'below added trying to fix some bugs
            For siteIncrementer = 0 To AllProductionSites.Count - 1
                For moduleIncrementer = 0 To AllProductionSites(siteIncrementer).ModulesList.Count - 1
                    testSector = AllProductionSites(siteIncrementer).ModulesList(moduleIncrementer).parentSector
                    If testSector.Name.Equals(SectorLabel.Name) Or SectorLabel.Name.Equals(tmpSector.Name) Then
                        i = _SiteCollection.IndexOf(AllProductionSites(siteIncrementer))
                        If i = -1 Then _SiteCollection.Add(AllProductionSites(siteIncrementer))
                    End If
                Next
            Next
            'fin

        End If
        figureOutWhichLinesToShow()
        SiteLabel_Close()

        prstory_linedropdown.Background = mybrushNOTESblue
        System.Windows.Forms.Application.DoEvents()
        Thread.Sleep(200)
        prstory_linedropdown.Background = mybrushdefaultbackgroundgray
    End Sub
    Public Sub SectorSelected(ByVal sender As Object, ByVal e As SelectionChangedEventArgs)
        Dim lbsender As ListBox, tmpSectorName As String, testSector As BusinessUnit, i As Integer
        Dim siteIncrementer As Integer, moduleIncrementer As Integer, selectedSector As BusinessUnit
        lbsender = CType(sender, ListBox)
        ' get the label right
        If Not IsNothing(lbsender.SelectedItem) Then
            tmpSectorName = lbsender.SelectedItem.ToString
            SectorLabel.Content = tmpSectorName
            My.Settings.DefaultBU = tmpSectorName

            If tmpSectorName.Equals(BETA_TESTING) Then MsgBox("Warning! You have selected lines that are still under Beta Testing. They have Not been validated For accuracy And may cause your computer To crash. Proceed at your own risk.", vbCritical, "beta test warning")
            'find the right sites
            If tmpSectorName.Equals(tmpSector.Name) Then
                calculateSiteCollection()
                calculateSectorCollection()
            Else
                _SiteCollection.Clear()
                _SiteCollection.Add(tmpSite)
                selectedSector = AllProductionSectors(AllProductionSectors.IndexOf(New BusinessUnit(tmpSectorName)))
                For siteIncrementer = 0 To AllProductionSites.Count - 1
                    For moduleIncrementer = 0 To AllProductionSites(siteIncrementer).ModulesList.Count - 1
                        testSector = AllProductionSites(siteIncrementer).ModulesList(moduleIncrementer).parentSector
                        If testSector.Name.Equals(tmpSectorName) Then
                            i = _SiteCollection.IndexOf(AllProductionSites(siteIncrementer))
                            If i = -1 Then _SiteCollection.Add(AllProductionSites(siteIncrementer))
                        End If
                    Next
                Next
                _SectorCollection.Clear()
                _SectorCollection.Add(tmpSector)
                _SectorCollection.Add(selectedSector)
                For sectorIncrementer As Integer = 0 To AllProductionSectors.Count - 1
                    If SiteLabel.Content.Equals(tmpSite.Name) Or AllProductionSectors(sectorIncrementer).isSectorAtSite(SiteLabel.Content) Then
                        If _SectorCollection.IndexOf(AllProductionSectors(sectorIncrementer)) = -1 Then _SectorCollection.Add(AllProductionSectors(sectorIncrementer))
                    End If
                Next
            End If

            SectorMenu.Visibility = Visibility.Hidden
            figureOutWhichLinesToShow()
            SectorLabel_Close()

            prstory_linedropdown.Background = mybrushNOTESblue
            System.Windows.Forms.Application.DoEvents()
            Thread.Sleep(200)
            prstory_linedropdown.Background = mybrushdefaultbackgroundgray
        End If
    End Sub

    'Line Selection in Auto Complete

    Sub LineSelection_AutoCompletechanged(sender As Object, e As RoutedEventArgs)
        If Not IsInitialized Then
            Return
        End If

        tempreasonlevel = ""
        prstory_linedropdown.SelectedItem = LineSelection_AutoComplete.SelectedItem
        If prstory_linedropdown.SelectedIndex <> -1 Then
            LineSelection_icongreencheck.Visibility = Visibility.Visible
            LineSelection_iconrederror.Visibility = Visibility.Hidden
        Else
            LineSelection_icongreencheck.Visibility = Visibility.Hidden
            LineSelection_iconrederror.Visibility = Visibility.Visible
        End If
    End Sub

    'SITE AND SECTOR LABELS
    Sub SectorLabel_Click()
        If SectorMenu.Visibility = Visibility.Visible Then
            SectorLabel_Close()
        Else
            SectorMenu.Visibility = Visibility.Visible
            SectorMenu_DropDownbtn.Visibility = Visibility.Hidden
            SectorMenu_DropUpbtn.Visibility = Visibility.Visible
        End If


    End Sub
    Sub SiteLabel_Click()
        If SiteMenu.Visibility = Visibility.Visible Then
            SiteLabel_Close()
        Else
            SiteMenu.Visibility = Visibility.Visible
            SiteMenu_DropDownbtn.Visibility = Visibility.Hidden
            SiteMenu_DropUpbtn.Visibility = Visibility.Visible
        End If
    End Sub
    Sub SectorLabel_Close()
        SectorMenu.Visibility = Visibility.Hidden
        SectorMenu_DropDownbtn.Visibility = Visibility.Visible
        SectorMenu_DropUpbtn.Visibility = Visibility.Hidden
    End Sub
    Sub SiteLabel_Close()
        SiteMenu.Visibility = Visibility.Hidden
        SiteMenu_DropDownbtn.Visibility = Visibility.Visible
        SiteMenu_DropUpbtn.Visibility = Visibility.Hidden
    End Sub

    Sub CleanUI()
        SectorLabel_Close()
        SiteLabel_Close()
    End Sub
#End Region

#Region "Init"
    Public Sub PlaySnake()
        Dim snakeS As New Snake
        snakeS.ShowDialog()
    End Sub
    Private Sub ShowNotificationDetails()
        NotificationCloseButton.Visibility = Visibility.Visible
        NotificationDetailsLabel.Visibility = Visibility.Visible
        GetUpdateButton.Visibility = Visibility.Visible
    End Sub
    Private Sub HideNotificationDetails()
        NotificationCloseButton.Visibility = Visibility.Hidden
        NotificationDetailsLabel.Visibility = Visibility.Hidden
        GetUpdateButton.Visibility = Visibility.Hidden
    End Sub
    Private Sub HideActiveNotifications()
        NotificationIcon.Visibility = Visibility.Hidden
        NotificationCountLabel.Visibility = Visibility.Hidden
        NotificationIcon_inactive.Visibility = Visibility.Visible
    End Sub
    Private Sub ShowctiveNotifications()
        NotificationIcon.Visibility = Visibility.Visible
        NotificationCountLabel.Visibility = Visibility.Visible
        NotificationIcon_inactive.Visibility = Visibility.Hidden
    End Sub
    Private Sub CheckScreenResolution()
        Dim screenWidth As Integer = My.Computer.Screen.Bounds.Width
        Dim screenHeight As Integer = My.Computer.Screen.Bounds.Height

        If screenWidth < 1200 Or screenHeight < 700 Then Me.WindowState = Windows.WindowState.Maximized

    End Sub

    Private Sub DecidetoShowUpdateNews()
        If My.Settings.ShowUpdateNews = True Then
            SplashCanvas.Visibility = Visibility.Visible
            BackButtonSplashPressed()
        Else
            SplashCanvas.Visibility = Visibility.Hidden
        End If
    End Sub



    Private Sub DisplayLoginScreen()
        Dim win As New WinLogin()
        '  win.parentwindow = me
        win.Owner = Me
        win.ShowDialog()
        If win.DialogResult.HasValue And win.DialogResult.Value Then
            BusyIndicator.IsBusy = True

            Dim bw = New BackgroundWorker()
            bw.WorkerReportsProgress = True
            bw.WorkerSupportsCancellation = True
            AddHandler bw.DoWork, AddressOf bw_DoWork
            AddHandler bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted
            bw.RunWorkerAsync()


        Else
            Me.Close()
        End If
    End Sub


    Sub prstory_onstart()

        DecidetoShowUpdateNews()
        CheckScreenResolution()
        HideActiveNotifications()

        Dim tempString As String
        importTargetsThread = New Thread(AddressOf CSV_readTargetsFile)

        My.Settings.Reload()
        HideNotificationDetails()
        verifyFolderStructure()
        If isApplicationUpdateNeeded() Then ShowctiveNotifications()
        bargraphreportwindow_Open = False
        IsRemappingDoneOnce = False
        tempString = ""
        SiteMenu.Visibility = Visibility.Hidden
        SectorMenu.Visibility = Visibility.Hidden

        hideprstorysettings()
        hideaboutwebcontrol()
        hideauxmenu()

        HideDateSelectionAlert()
        initializeMenuTextFromLanguage()

        initializeAllSites()

        importTargetsThread.Start()

        'figureOutWhichLinesToShow()  ' this shows lines for selected sector-site combination
        populatelineslist()
        ShowSeachBarandHideDropDown()
        calculateSectorCollection()
        calculateSiteCollection()

        LineSelection_AutoComplete.WatermarkContent = "Enter a line name -  ex. " & getRandomLineName() 'Albany Fam AK09"

        HideLineDefaultQuery()
        If My.Settings.areDatesSaved Then
            prstory_dateselectionLabel.Content = Format(My.Settings.LastStartDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(My.Settings.LastEndDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine
            prstory_dateselectionLabel.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
            starttimeselected = My.Settings.LastStartDate
            endtimeselected = My.Settings.LastEndDate
        End If

        populateLanguageDropdown()
        populatestartandendtimehourandmin()
        My.Settings.AdvancedSettings_isAvailabilityMode = False

        PRSTORY_VERSION_NUMBER = VersionLabel.Content

#If DEBUG Then
        'dateSelectionShortcut()
        ' DisplayLoginScreen()
#Else
                DisplayLoginScreen()
#End If


    End Sub

    Private Function setDefaultLineIndexFromName() As Integer
        Dim BetaCounter As Integer = 0
        For i As Integer = 0 To AllProdLines.Count - 1
            With AllProdLines(i)
                If .Sector.Equals(BETA_TESTING) Then
                    BetaCounter += 1
                ElseIf .parentSite.ThreeLetterID.Equals(My.Settings.DefaultSiteAcronym) And My.Settings.DefaultLineName.Equals(.Name) Then
                    My.Settings.DefaultLineIndex = i + 1 ' - BetaCounter
                    Return BetaCounter
                End If
            End With
        Next
        Return BetaCounter
    End Function

    Private Sub verifyFolderStructure()
        createFolder(PATH_PRSTORY)
        createFolder(PATH_PRSTORY_SETTINGS)
        createFolder(PATH_PRSTORY_RAWDATA)
        createFolder(SERVER_FOLDER_PATH)

        ImportMultilineGroups()  'JSONJSON
    End Sub
    Private Sub createFolder(folderName As String)
        Try
            If (Not System.IO.Directory.Exists(folderName)) Then
                System.IO.Directory.CreateDirectory(folderName)
            End If
        Catch ex As Exception
            MessageBox.Show("prstory has been unable to establish the required directories due to the following exception: " & ex.Message, "Folder Creation Error", vbCritical)
        End Try
    End Sub

    Private multilineGroups As List(Of ProductionLineGroup)

    Private Sub ImportMultilineGroups()
        Try
            multilineGroups = JSON_Import_LineGroup()
        Catch ex As Exception
            multilineGroups = New List(Of ProductionLineGroup)
        End Try
    End Sub

    Private Function JSON_Import_LineGroup(Optional FileNameX As String = "C:\\Users\\Public\\prstory\\prstoryLineGroupsJSON", Optional FileType As String = ".txt") As List(Of ProductionLineGroup)
        Try
            Dim json As String = File.ReadAllText(FileNameX & FileType)
            Dim o As List(Of ProductionLineGroup) = JsonConvert.DeserializeObject(Of List(Of ProductionLineGroup))(json)
            Return o
        Catch
            Throw
        End Try
    End Function

    Sub populatestartandendtimehourandmin()

        Dim k As Integer

        For k = 0 To 23
            If k > 9 Then
                starthour.Items.Add(CStr(k))
                endhour.Items.Add(CStr(k))
            Else
                starthour.Items.Add("0" & CStr(k))
                endhour.Items.Add("0" & CStr(k))
            End If
        Next

        For k = 0 To 59

            If k > 9 Then
                startmin.Items.Add(CStr(k))
                endmin.Items.Add(CStr(k))
            Else
                startmin.Items.Add("0" & CStr(k))
                endmin.Items.Add("0" & CStr(k))
            End If
        Next

        For k = 0 To 23
            If k > 9 Then
                Estarthour.Items.Add(CStr(k))
                Eendhour.Items.Add(CStr(k))
            Else
                Estarthour.Items.Add("0" & CStr(k))
                Eendhour.Items.Add("0" & CStr(k))
            End If
        Next

        For k = 0 To 59

            If k > 9 Then
                Estartmin.Items.Add(CStr(k))
                Eendmin.Items.Add(CStr(k))
            Else
                Estartmin.Items.Add("0" & CStr(k))
                Eendmin.Items.Add("0" & CStr(k))
            End If
        Next

    End Sub

    Private Sub SetLanguage_Main(sender As Object, e As MouseButtonEventArgs)
        If InStr(sender.name, "English") > 0 Then
            LanguageSelectionCombo.SelectedValue = "English"
            My.Settings.LanguageActive = LanguageSelectionCombo.SelectedIndex
            LanguageButton_Chinese.Foreground = mybrushlanguagewhite
            LanguageButton_English.Foreground = mybrushlanguagegreen

        ElseIf InStr(sender.name, "Chinese") > 0 Then
            LanguageSelectionCombo.SelectedValue = "中文"
            My.Settings.LanguageActive = LanguageSelectionCombo.SelectedIndex
            LanguageButton_Chinese.Foreground = mybrushlanguagegreen
            LanguageButton_English.Foreground = mybrushlanguagewhite
        Else
            LanguageSelectionCombo.SelectedValue = "English"
            My.Settings.LanguageActive = LanguageSelectionCombo.SelectedIndex
            LanguageButton_Chinese.Foreground = mybrushlanguagewhite
            LanguageButton_English.Foreground = mybrushlanguagegreen
        End If

        initializeMenuTextFromLanguage()

    End Sub
    Sub populatelineslist()
        activeLineIndeces.Clear()
        prstory_linedropdown.Items.Clear()
        'lists all lines in the drop down 
        Dim i As Integer
        Select Case My.Settings.LanguageActive
            Case Lang.English
                prstory_linedropdown.Items.Add("Line Selection")
            Case Lang.German
                prstory_linedropdown.Items.Add("wahlen Sie eine Leitung")
            Case Lang.Spanish
                prstory_linedropdown.Items.Add("Seleccione una linea")
            Case Lang.French
                prstory_linedropdown.Items.Add("Selectionner une ligne")
            Case Lang.Portuguese
                prstory_linedropdown.Items.Add("Seleção de Linha")
            Case Lang.Chinese_Simplified
                prstory_linedropdown.Items.Add("选择生产线")
        End Select
        LinesList.Clear()
        'prstory_linedropdown.Items.Clear()

        For i = 0 To AllProdLines.Count - 1
            If Not AllProdLines(i).Sector = BETA_TESTING Then
                prstory_linedropdown.Items.Add(AllProdLines(i).ToString) 'AllProductionLines(i).parentSite.Name & " " & AllProductionLines(i).parentModule.Name & " " & AllProductionLines(i).Name)
                LinesList.Add(AllProdLines(i).ToString)
                activeLineIndeces.Add(i)
            End If
        Next

        For i = 60 To 400
            QueryDaysBox.Items.Add(i)
        Next
        For i = 0 To 23
            FSSTHourBox.Items.Add(i)
        Next
        For i = 0 To 59
            FSSTMinBox.Items.Add(i)
        Next

        FSSTMinBox.SelectedItem = 0
        FSSTHourBox.SelectedItem = 0
        QueryDaysBox.SelectedItem = 99
    End Sub

    Public ReadOnly Property LinesList() As ObservableCollection(Of String)
        Get
            Return _LinesList
        End Get
    End Property

#End Region

#Region "Settings"
    Sub hideprstorysettings()

        prstorySettingsForm.Visibility = Visibility.Hidden
        prstorydayoption.Visibility = Visibility.Hidden

        prstoryCancelButton.Visibility = Visibility.Hidden
        prstoryGoButton.Visibility = Visibility.Hidden
        TargetLaunchButton.Visibility = Visibility.Hidden
        AdvancedSettingsButton.Visibility = Visibility.Hidden
        LineConfigExportButton.Visibility = Visibility.Hidden
        hideauxmenu()

    End Sub

    Sub launchDateExclusionSettings()

        If My.Settings.Exclude_StartHour < 10 Then
            Estarthour.SelectedValue = "0" + CStr(My.Settings.Exclude_StartHour)
        Else
            Estarthour.SelectedValue = CStr(My.Settings.Exclude_StartHour)
        End If

        If My.Settings.Exclude_EndHour < 10 Then
            Eendhour.SelectedValue = "0" + CStr(My.Settings.Exclude_EndHour)
        Else
            Eendhour.SelectedValue = CStr(My.Settings.Exclude_EndHour)
        End If

        If My.Settings.Exclude_StartMinutes < 10 Then
            Estartmin.SelectedValue = "0" + CStr(My.Settings.Exclude_StartMinutes)
        Else
            Estartmin.SelectedValue = CStr(My.Settings.Exclude_StartMinutes)
        End If

        If My.Settings.Exclude_EndMinutes < 10 Then
            Eendmin.SelectedValue = "0" + CStr(My.Settings.Exclude_EndMinutes)
        Else
            Eendmin.SelectedValue = CStr(My.Settings.Exclude_EndMinutes)
        End If

        DateSettings.Visibility = Visibility.Hidden
        DateExclusion.Visibility = Visibility.Visible

        EnableDateExcludeBox.IsChecked = My.Settings.EnableTimeSpanExclusion
    End Sub

    Sub hideDateExclusionSettings()
        DateSettings.Visibility = Visibility.Visible
        DateExclusion.Visibility = Visibility.Hidden
    End Sub

    Sub DateExcludeConfirmed()

        My.Settings.EnableTimeSpanExclusion = EnableDateExcludeBox.IsChecked

        If My.Settings.EnableTimeSpanExclusion Then
            My.Settings.Exclude_StartMinutes = CInt(Estartmin.SelectedValue)
            My.Settings.Exclude_EndMinutes = CInt(Eendmin.SelectedValue)
            My.Settings.Exclude_StartHour = CInt(Estarthour.SelectedValue)
            My.Settings.Exclude_EndHour = CInt(Eendhour.SelectedValue)
            DoWeAlwaysRepullDataNoMatterWhatElseHasHappened = True
        End If

        hideDateExclusionSettings()
    End Sub

    Sub DateExcludeCanceled()


        EnableDateExcludeBox.IsChecked = False
        My.Settings.EnableTimeSpanExclusion = False
        hideDateExclusionSettings()
    End Sub

    Sub launchprstorysettings()
        prstorySettingsForm.Visibility = Visibility.Visible
        prstorydayoption.Visibility = Visibility.Visible
        TargetLaunchButton.Visibility = Visibility.Visible
        LineConfigExportButton.Visibility = Visibility.Visible
        AdvancedSettingsButton.Visibility = Visibility.Visible
        prstorySettingsForm_Aux.Visibility = Visibility.Visible
        prstoryCancelButton.Visibility = Visibility.Visible

        With My.Settings
            inControlRuleOneBox.IsChecked = .inControl_useRule1
            inControlRuleTwoBox.IsChecked = .inControl_useRule2
            inControlRuleThreeBox.IsChecked = .inControl_useRule3
            inControlRuleFourBox.IsChecked = .inControl_useRule4
            inControlRuleFiveBox.IsChecked = .inControl_useRule5
            inControlRuleSixBox.IsChecked = .inControl_useRule6
            SAPnameCheckBox.IsChecked = .displaySAPnames
            SnakeCheckBox.IsChecked = .AdvancedSettings_PlaySnake
            TargetsEnabledTrueBox.IsChecked = .AdvancedSettings_isTargetsEnabled
            AvailabilityTrueBox.IsChecked = .AdvancedSettings_isAvailabilityMode
            NotesBox.IsChecked = .AdvancedSettings_UseNotes

            TimeSettingsProdBox.IsChecked = .AdvancedSettings_ProdUseMinutes
            TimeSettingsAvailBox.IsChecked = .AdvancedSettings_AvailUseMinutes

            EnableRemapBox.IsChecked = .PostMap_Enable
        End With

        '        querydaysbox.Items.Cast<int>().Select(item => item).ToList().GetRange(0,100)


        inControlBoxesCheckUncheck()
        Checkforavailabilitymode()
        prstorydayoption_MouseDown()
    End Sub
    Sub Checkforavailabilitymode()
        If My.Settings.AdvancedSettings_isAvailabilityMode = True Then AvailabilityTrueBox.IsChecked = True

    End Sub
    Sub launchprstorydaterange()
        hideauxmenu()
        DateSettings.Visibility = Visibility.Visible
        prstorySettingsForm_Aux.Visibility = Visibility.Visible
        closeauxbutton.Visibility = Visibility.Visible
        prstoryGoButton.Visibility = Visibility.Visible
        SetStartandEndTime()
    End Sub
    Sub SetStartandEndTime()
        If prstory_linedropdown.SelectedIndex = 0 Then
            starthour.SelectedValue = "06"
            endhour.SelectedValue = "06"
            startmin.SelectedValue = "00"
            endmin.SelectedValue = "00"

            Exit Sub
        End If
        Try
            If My.Settings.AdvancedSettings_isAvailabilityMode And areTimesSaved() And My.Settings.AdvancedSettings_AvailUseMinutes Then
                If My.Settings.DefaultStartTimeHour < 10 Then
                    starthour.SelectedValue = "0" + CStr(My.Settings.DefaultStartTimeHour)
                Else
                    starthour.SelectedValue = CStr(My.Settings.DefaultStartTimeHour)
                End If

                If My.Settings.DefaultEndTimeHour < 10 Then
                    endhour.SelectedValue = "0" + CStr(My.Settings.DefaultEndTimeHour)
                Else
                    endhour.SelectedValue = CStr(My.Settings.DefaultEndTimeHour)
                End If

                If My.Settings.DefaultStartTimeMinutes < 10 Then
                    startmin.SelectedValue = "0" + CStr(My.Settings.DefaultStartTimeMinutes)
                Else
                    startmin.SelectedValue = CStr(My.Settings.DefaultStartTimeMinutes)
                End If

                If My.Settings.DefaultEndTimeMinutes < 10 Then
                    endmin.SelectedValue = "0" + CStr(My.Settings.DefaultEndTimeMinutes)
                Else
                    endmin.SelectedValue = CStr(My.Settings.DefaultEndTimeMinutes)
                End If

            ElseIf Not My.Settings.AdvancedSettings_isAvailabilityMode And areTimesSaved() And My.Settings.AdvancedSettings_ProdUseMinutes Then
                If My.Settings.DefaultStartTimeHour < 10 Then
                    starthour.SelectedValue = "0" + CStr(My.Settings.DefaultStartTimeHour)
                Else
                    starthour.SelectedValue = CStr(My.Settings.DefaultStartTimeHour)
                End If

                If My.Settings.DefaultEndTimeHour < 10 Then
                    endhour.SelectedValue = "0" + CStr(My.Settings.DefaultEndTimeHour)
                Else
                    endhour.SelectedValue = CStr(My.Settings.DefaultEndTimeHour)
                End If

                If My.Settings.DefaultStartTimeMinutes < 10 Then
                    startmin.SelectedValue = "0" + CStr(My.Settings.DefaultStartTimeMinutes)
                Else
                    startmin.SelectedValue = CStr(My.Settings.DefaultStartTimeMinutes)
                End If

                If My.Settings.DefaultEndTimeMinutes < 10 Then
                    endmin.SelectedValue = "0" + CStr(My.Settings.DefaultEndTimeMinutes)
                Else
                    endmin.SelectedValue = CStr(My.Settings.DefaultEndTimeMinutes)
                End If
            Else
                starthour.SelectedValue = "0" & CStr(Int(AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr))
                endhour.SelectedValue = "0" & CStr(Int(AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr))

                If (60 * (AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr - Int(AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr))) = 0 Then
                    startmin.SelectedValue = "00"
                    endmin.SelectedValue = "00"
                Else
                    startmin.SelectedValue = "" & CStr(Int(60 * (AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr - Int(AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr))))
                    endmin.SelectedValue = "" & CStr(Int(60 * (AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr - Int(AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).ShiftStartFirst_Hr))))
                End If
            End If
        Catch ex As Exception
            starthour.SelectedValue = "06"
            endhour.SelectedValue = "06"
            startmin.SelectedValue = "00"
            endmin.SelectedValue = "00"
            Exit Sub
        End Try
    End Sub

    Private Function areTimesSaved() As Boolean
        Return My.Settings.DefaultStartTimeMinutes > -1 And My.Settings.DefaultEndTimeMinutes > -1 And My.Settings.DefaultStartTimeHour > -1 And My.Settings.DefaultEndTimeHour > -1
    End Function


    Sub launchprstorycombo_settingsaux()

        hideauxmenu()
        prstorySettingsForm_Aux.Visibility = Visibility.Visible
        closeauxbutton.Visibility = Visibility.Visible
        prstoryGoButton.Visibility = Visibility.Visible
        prstorySettingsForm_Aux.Visibility = Visibility.Visible
    End Sub
    Sub launchprstoryAdvancedOptions_settingsaux()
        hideauxmenu()
        AdvancedSettings.Visibility = Visibility.Visible
        prstorySettingsForm_Aux.Visibility = Visibility.Visible
        PDTbox.Text = My.Settings.PDT_maxMinutesBetweenEvents
        DTcutoffbox.Text = My.Settings.AdvancedSettings_DTcutoff
        UTcutoffbox.Text = My.Settings.AdvancedSettings_UTcutoff

        prstoryGoButton.Visibility = Visibility.Visible
    End Sub
    Sub launchprstoryAdvancedOptionsMC_settingsaux()
        'hide it all
        hideauxmenu()
        'show the background/std stuff
        closeauxbutton.Visibility = Visibility.Visible
        prstoryGoButton.Visibility = Visibility.Visible
        prstorySettingsForm_Aux.Visibility = Visibility.Visible
        'show me the monay!
        AdvancedSettings_MultiConstraint_Return.Visibility = Visibility.Visible
        MultiSettings.Visibility = Visibility.Visible
    End Sub

    Sub launchinctonrolOptions_settingsaux()
        'hide it all
        hideauxmenu()
        'show the background/std stuff
        closeauxbutton.Visibility = Visibility.Visible
        prstoryGoButton.Visibility = Visibility.Visible
        prstorySettingsForm_Aux.Visibility = Visibility.Visible
        inControlSettings.Visibility = Visibility.Visible
        AdvancedSettings.Visibility = Visibility.Hidden
        AdvancedSettings2.Visibility = Visibility.Visible

        inControlSettingsButton2.Content = "< inControl Settings"
    End Sub

    Sub launchinctonrolOptions_settingsaux2()
        AdvancedSettings.Visibility = Visibility.Visible
        AdvancedSettings.Visibility = Visibility.Visible
        inControlSettings.Visibility = Visibility.Hidden
    End Sub

    Sub SliderAct()
        Dim DefaultPosition As Integer = 1
        Dim CustomPosition As Integer = 2
        Dim WesterElectricPosition As Integer = 3
        Select Case CInt(inControlSlider.Value)
            Case DefaultPosition
                inControlRuleOneBox.IsChecked = True
                inControlRuleTwoBox.IsChecked = True
                inControlRuleThreeBox.IsChecked = True
                inControlRuleFourBox.IsChecked = True
                inControlRuleFiveBox.IsChecked = False
                inControlRuleSixBox.IsChecked = False
            Case CustomPosition

            Case WesterElectricPosition
                inControlRuleOneBox.IsChecked = True
                inControlRuleTwoBox.IsChecked = True
                inControlRuleThreeBox.IsChecked = True
                inControlRuleFourBox.IsChecked = False
                inControlRuleFiveBox.IsChecked = True
                inControlRuleSixBox.IsChecked = False
            Case Else
                Debugger.Break()
        End Select
    End Sub

    Sub inControlBoxesCheckUncheck()
        Dim DefaultPosition As Integer = 1
        Dim CustomPosition As Integer = 2
        Dim WesterElectricPosition As Integer = 3
        If inControlRuleOneBox.IsChecked = True And
            inControlRuleTwoBox.IsChecked = True And
            inControlRuleThreeBox.IsChecked = True And
            inControlRuleFourBox.IsChecked = False And
            inControlRuleFiveBox.IsChecked = True And
            inControlRuleSixBox.IsChecked = False Then
            inControlSlider.Value = WesterElectricPosition
        ElseIf inControlRuleOneBox.IsChecked = True And
            inControlRuleTwoBox.IsChecked = True And
            inControlRuleThreeBox.IsChecked = True And
            inControlRuleFourBox.IsChecked = True And
            inControlRuleFiveBox.IsChecked = False And
            inControlRuleSixBox.IsChecked = False Then
            inControlSlider.Value = DefaultPosition
        Else
            inControlSlider.Value = CustomPosition
        End If
    End Sub

    Sub hideauxmenu()
        AdvancedSettings.Visibility = Visibility.Hidden
        AdvancedSettings2.Visibility = Visibility.Hidden
        inControlSettings.Visibility = Visibility.Hidden
        DateSettings.Visibility = Visibility.Hidden
        MultiSettings.Visibility = Visibility.Hidden
        prstorySettingsForm_Aux.Visibility = Visibility.Hidden
        prstoryGoButton.Visibility = Visibility.Hidden
        closeauxbutton.Visibility = Visibility.Hidden
        DateExclusion.Visibility = Visibility.Hidden

        HideDateSelectionAlert()
    End Sub

    Private Sub settingsDONE_dates()
        If IsNothing(prstory_datepicker_enddate.SelectedDate) And IsNothing(prstory_datepicker_startdate.SelectedDate) Then
            Exit Sub
        ElseIf IsNothing(prstory_datepicker_enddate.SelectedDate) Or IsNothing(prstory_datepicker_startdate.SelectedDate) Then

            MsgBox("Date field cannot be left blank", vbCritical)

            If prstory_datepicker_startdate.SelectedDate = prstory_datepicker_enddate.SelectedDate Then

                MsgBox("Start date/time cannot be same as end date/time.", vbCritical)

                prstory_datepicker_startdate.SelectedDate = Nothing
                prstory_datepicker_enddate.SelectedDate = Nothing
                Exit Sub

            End If

            Exit Sub

        End If

        prstory_datepicker_startdate.SelectedDate = Format(prstory_datepicker_startdate.SelectedDate, "Short Date") & " " & starthour.SelectedValue.ToString & ":" & startmin.SelectedValue.ToString
        prstory_datepicker_enddate.SelectedDate = Format(prstory_datepicker_enddate.SelectedDate, "Short Date") & " " & endhour.SelectedValue.ToString & ":" & endmin.SelectedValue.ToString

        If prstory_datepicker_startdate.SelectedDate > prstory_datepicker_enddate.SelectedDate Then
            MsgBox("Start date/time cannot be later than end date/time.", vbCritical)
            Exit Sub
        End If

        If prstory_datepicker_enddate.SelectedDate > Now Then
            MsgBox("Warning: You have selected a time period that ends after the current time in your time zone. If this was not your intention, please reselect dates.")
        End If

        If DateDiff("d", prstory_datepicker_startdate.SelectedDate, prstory_datepicker_enddate.SelectedDate) > datapull_duration - 1 Then
            MsgBox("Sorry, we are not there yet." & vbNewLine & vbNewLine & "We are still working on getting prstory work for date ranges greater than 99 days.")
            Exit Sub
        End If

        'prstory_dateselectionLabel.Content = Format(prstory_datepicker_startdate.SelectedDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(prstory_datepicker_enddate.SelectedDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine
        prstory_dateselectionLabel.Content = Format(prstory_datepicker_startdate.SelectedDate).ToString & vbNewLine & Format(prstory_datepicker_enddate.SelectedDate).ToString & vbNewLine
        prstory_dateselectionLabel.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
        starttimeselected = prstory_datepicker_startdate.SelectedDate
        endtimeselected = prstory_datepicker_enddate.SelectedDate

        'save dates and line selection
        '   My.Settings.DefaultLineIndex = prstory_linedropdown.SelectedIndex

        If My.Settings.AdvancedSettings_isAvailabilityMode And My.Settings.AdvancedSettings_AvailUseMinutes Then
            My.Settings.DefaultStartTimeMinutes = CInt(startmin.SelectedValue)
            My.Settings.DefaultEndTimeMinutes = CInt(endmin.SelectedValue)
            My.Settings.DefaultStartTimeHour = CInt(starthour.SelectedValue)
            My.Settings.DefaultEndTimeHour = CInt(endhour.SelectedValue)
        ElseIf Not My.Settings.AdvancedSettings_isAvailabilityMode And My.Settings.AdvancedSettings_ProdUseMinutes Then
            My.Settings.DefaultStartTimeMinutes = CInt(startmin.SelectedValue)
            My.Settings.DefaultEndTimeMinutes = CInt(endmin.SelectedValue)
            My.Settings.DefaultStartTimeHour = CInt(starthour.SelectedValue)
            My.Settings.DefaultEndTimeHour = CInt(endhour.SelectedValue)
        End If



        My.Settings.LastStartDate = starttimeselected
        My.Settings.LastEndDate = endtimeselected
        My.Settings.areDatesSaved = True
        CheckifwecanrunDO_Analyze_directly()
    End Sub
    Private Sub settingsDONE_settings()
        'if we need to make a change to how we are showing names then:
        If SAPnameCheckBox.IsChecked And My.Settings.displaySAPnames = False Then
            My.Settings.displaySAPnames = True
            figureOutWhichLinesToShow()
        ElseIf Not SAPnameCheckBox.IsChecked And My.Settings.displaySAPnames = True Then
            My.Settings.displaySAPnames = False
            figureOutWhichLinesToShow()
        End If

        'if we want to run in availability mode and NOT in PR mode
        If AvailabilityTrueBox.IsChecked Then
            My.Settings.AdvancedSettings_isAvailabilityMode = True

        ElseIf Not AvailabilityTrueBox.IsChecked Then
            My.Settings.AdvancedSettings_isAvailabilityMode = False

        End If


        'if we want targets or not
        If TargetsEnabledTrueBox.IsChecked Then
            My.Settings.AdvancedSettings_isTargetsEnabled = True
        ElseIf Not TargetsEnabledTrueBox.IsChecked Then
            My.Settings.AdvancedSettings_isTargetsEnabled = False
        End If

        ''''

        'if we want to show notes or not
        If NotesBox.IsChecked Then
            My.Settings.AdvancedSettings_UseNotes = True
        ElseIf Not NotesBox.IsChecked Then
            My.Settings.AdvancedSettings_UseNotes = False
        End If


        My.Settings.AdvancedSettings_ProdUseMinutes = TimeSettingsProdBox.IsChecked
        My.Settings.AdvancedSettings_AvailUseMinutes = TimeSettingsAvailBox.IsChecked


        ''''

        My.Settings.AdvancedSettings_UseSimulation = True

        'if we want to include PR Out
        If IncludeExcludeBox.IsChecked Then
            IsExcludedEventsIncluded = True
        ElseIf Not IncludeExcludeBox.IsChecked Then
            IsExcludedEventsIncluded = False
        End If

        ''''


        If EnableRemapBox.IsChecked Then
            My.Settings.PostMap_Enable = True
            Try
                Dim x As String = Tier1RemapBox.SelectedItem.content
                If x = "Reason 1" Then
                    My.Settings.PostMap_Field1 = DowntimeField.Reason1
                ElseIf x = "Reason 2" Then
                    My.Settings.PostMap_Field1 = DowntimeField.Reason2
                ElseIf x = "Reason 3" Then
                    My.Settings.PostMap_Field1 = DowntimeField.Reason3
                ElseIf x = "Reason 4" Then
                    My.Settings.PostMap_Field1 = DowntimeField.Reason4
                ElseIf x = "Fault" Then
                    My.Settings.PostMap_Field1 = DowntimeField.Fault
                End If


                x = Tier2RemapBox.SelectedItem.content
                If x = "Reason 1" Then
                    My.Settings.PostMap_Field2 = DowntimeField.Reason1
                ElseIf x = "Reason 2" Then
                    My.Settings.PostMap_Field2 = DowntimeField.Reason2
                ElseIf x = "Reason 3" Then
                    My.Settings.PostMap_Field2 = DowntimeField.Reason3
                ElseIf x = "Reason 4" Then
                    My.Settings.PostMap_Field2 = DowntimeField.Reason4
                ElseIf x = "Fault" Then
                    My.Settings.PostMap_Field2 = DowntimeField.Fault
                End If

                x = Tier3RemapBox.SelectedItem.content
                If x = "Reason 1" Then
                    My.Settings.PostMap_Field3 = DowntimeField.Reason1
                ElseIf x = "Reason 2" Then
                    My.Settings.PostMap_Field3 = DowntimeField.Reason2
                ElseIf x = "Reason 3" Then
                    My.Settings.PostMap_Field3 = DowntimeField.Reason3
                ElseIf x = "Reason 4" Then
                    My.Settings.PostMap_Field3 = DowntimeField.Reason4
                ElseIf x = "Fault" Then
                    My.Settings.PostMap_Field3 = DowntimeField.Fault
                End If
            Catch
            End Try
            My.Settings.PostMap_Enable = True
        Else
            My.Settings.PostMap_Enable = False
        End If


        Try
            My.Settings.PDT_maxMinutesBetweenEvents = CDbl(PDTbox.Text)
            My.Settings.AdvancedSettings_DTcutoff = CDbl(DTcutoffbox.Text)
            My.Settings.AdvancedSettings_UTcutoff = CDbl(UTcutoffbox.Text)
        Catch ex As Exception

        End Try
        inControl_saveSettings()
        multiConstraint_saveSettings()

        My.Settings.Save()
    End Sub


    Sub settingsDONE()
        HideDateSelectionAlert()
        settingsDONE_dates()
        settingsDONE_settings()



        hideprstorysettings()


        GreenCheck.Visibility = Visibility.Hidden

    End Sub
    Private Sub CheckifwecanrunDO_Analyze_directly()
        If IsAnalyzeButtonClickSource_Analyze = True Then
            IsAnalyzeButtonClickSource_Analyze = False
            Do_Analyze()
        End If

    End Sub

    Private Sub multiConstraint_saveSettings()
        If MultiConstraint_ModeAButton.IsChecked Then
            My.Settings.AdvancedSettings_MultiConstraintAnalysisMode = MultiConstraintAnalysis.SingleConstraint
        ElseIf MultiConstraint_ModeBButton.IsChecked Then
            My.Settings.AdvancedSettings_MultiConstraintAnalysisMode = MultiConstraintAnalysis.NoRateLossStops
        ElseIf MultiConstraint_ModeCButton.IsChecked Then
            My.Settings.AdvancedSettings_MultiConstraintAnalysisMode = MultiConstraintAnalysis.RateLossAsStops
        End If
    End Sub
    Private Sub inControl_saveSettings()
        With My.Settings
            .inControl_useRule1 = inControlRuleOneBox.IsChecked
            .inControl_useRule2 = inControlRuleTwoBox.IsChecked
            .inControl_useRule3 = inControlRuleThreeBox.IsChecked
            .inControl_useRule4 = inControlRuleFourBox.IsChecked
            .inControl_useRule5 = inControlRuleFiveBox.IsChecked
            .inControl_useRule6 = inControlRuleSixBox.IsChecked
        End With
    End Sub

    Sub launchprstorytrend()
        hideprstorysettings()
    End Sub

    Private Function isApplicationUpdateNeeded() As Boolean
        Dim info As UpdateCheckInfo = Nothing
        If (ApplicationDeployment.IsNetworkDeployed) Then
            Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
            Try
                info = AD.CheckForDetailedUpdate()
            Catch dde As DeploymentDownloadException
                MessageBox.Show("Error. The new version cannot be downloaded at this time. " + ControlChars.Lf & ControlChars.Lf & "Please check your network connection, or try again later. Error: " + dde.Message)
                Return False
            Catch ioe As InvalidOperationException
                MessageBox.Show("Uh oh - application is likely not a ClickOnce application. Error: " & ioe.Message)
                Return False
            End Try

            If (info.UpdateAvailable) Then Return True

        End If
        Return False
    End Function

    Private Sub InstallUpdateSyncWithInfo()
        HideNotificationDetails()
        HideActiveNotifications()

        Dim info As UpdateCheckInfo = Nothing

        If (ApplicationDeployment.IsNetworkDeployed) Then
            Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment

            Try
                info = AD.CheckForDetailedUpdate()
            Catch dde As DeploymentDownloadException
                MessageBox.Show("The new version of the application cannot be downloaded at this time. " + ControlChars.Lf & ControlChars.Lf & "Please check your network connection, or try again later. Error: " + dde.Message)
                Return
            Catch ioe As InvalidOperationException
                MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " & ioe.Message)
                Return
            End Try

            If (info.UpdateAvailable) Then
                Dim doUpdate As Boolean = True

                If (Not info.IsUpdateRequired) Then
                    Dim dr As Forms.DialogResult = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available", Forms.MessageBoxButtons.OKCancel)
                    If (Not System.Windows.Forms.DialogResult.OK = dr) Then
                        doUpdate = False
                    End If
                Else
                    ' Display a message that the app MUST reboot. Display the minimum required version.
                    MessageBox.Show("This application has detected a mandatory update from your current " &
                        "version to version " & info.MinimumRequiredVersion.ToString() &
                        ". The application will now install the update and restart.",
                        "Update Available", Forms.MessageBoxButtons.OK,
                        Forms.MessageBoxIcon.Information)
                End If

                If (doUpdate) Then
                    Try
                        AD.Update()
                        MessageBox.Show("The application has been upgraded, please exit prstory then restart!")
                        'Application.Restart()

                    Catch dde As DeploymentDownloadException
                        MessageBox.Show("Cannot install the latest version of the application. " & ControlChars.Lf & ControlChars.Lf & "Please check your network connection, or try again later.")
                        Return
                    End Try
                End If
            Else
                MsgBox("No update available!")
            End If

        Else
            MsgBox("you're debugging right now, this wont work")
        End If
    End Sub



#End Region

#Region "Mouse Move/Leave/Down"
    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        sender.Opacity = 0.7
    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        sender.Opacity = 1.0
    End Sub


    Private Sub prstorymapping_Mousedown()
        launchprstorycombo_settingsaux()
    End Sub

    Private Sub IconMouseMove(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Hand

        If sender Is prstory_Analyze_Icon Then
            prstory_Analyze_Icon.Opacity = 0.8
        End If

        If sender Is prstory_Settings_Icon Then
            prstory_Settings_Icon.Opacity = 0.8
        End If

        If sender Is prstory_Target_Icon Then
            prstory_Target_Icon.Opacity = 0.8
        End If

        If sender Is prstory_linedropdown Then
            prstory_linedropdown.Opacity = 0.9
        End If

    End Sub

    Private Sub IconMouseLeave(sender As Object, e As MouseEventArgs)

        Cursor = Cursors.Arrow
        If sender Is prstory_Analyze_Icon Then
            prstory_Analyze_Icon.Opacity = 1.0
        End If
        If sender Is prstory_Settings_Icon Then
            prstory_Settings_Icon.Opacity = 1.0
        End If

        If sender Is prstory_Target_Icon Then
            prstory_Target_Icon.Opacity = 1.0
        End If

        If sender Is prstory_linedropdown Then
            prstory_linedropdown.Opacity = 0.9
        End If

    End Sub

    Private Sub prstoryCancelButton_MouseDown()
        If InStr(prstory_dateselectionLabel.Content, "sele", vbTextCompare) > 0 Or InStr(prstory_dateselectionLabel.Content, "Reisedatum", vbTextCompare) > 0 Then  ' sele is the common word for selection in 4 / 5 languages
            prstory_datepicker_startdate.SelectedDate = Nothing
            prstory_datepicker_enddate.SelectedDate = Nothing
        End If

        hideprstorysettings()
    End Sub

    Private Sub prstory_Settings_Icon_MouseDown()
        launchprstorysettings()

    End Sub

    Private Sub prstorydayoption_MouseDown()
        launchprstorydaterange()
    End Sub

    Private Sub prstorymonthoption_MouseDown()
        'prstory_datepicker_startdate.SelectedDate = DateTime.
        prstory_datepicker_enddate.SelectedDate = DateTime.Today.ToString
    End Sub
    Private Sub prstorymtdoption_MouseDown()

        prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-Day(DateTime.Today) + 1)
        'prstory_datepicker_startdate.SelectedDate = DateTime.Today.ToString("MM") & "/01/" & DateTime.Today.ToString("yyyy")
        prstory_datepicker_enddate.SelectedDate = DateTime.Today.ToString
    End Sub
    Private Sub prstorylast7daysoption_MouseDown()
        prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-7)
        prstory_datepicker_enddate.SelectedDate = DateTime.Today

    End Sub
    Private Sub prstoryyesterdayoption_MouseDown()
        If Hour(Now) > 6 Then

            prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-1)
            prstory_datepicker_enddate.SelectedDate = DateTime.Today
        Else
            prstory_datepicker_startdate.SelectedDate = DateTime.Today.AddDays(-2)
            prstory_datepicker_enddate.SelectedDate = DateTime.Today.AddDays(-1)
        End If

    End Sub




#End Region

#Region "Menu"
    Sub ShowMenu()

    End Sub
    Sub AppClick_targets()
        If prstory_linedropdown.SelectedIndex < 1 Then
            LineSelection_Alert.Content = "A line needs to be selected to enter targets"
            LineSelection_Alert.Visibility = Visibility.Visible
        Else
            LineSelection_Alert.Visibility = Visibility.Hidden
            HideMenu()
            Dim tgtWin As New WindowTargets(AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)))
            MessageBox.Show("Notice: Targets are only able to be set by the SPOC for your business unit. If need to update your targets, please contact your SPOC.", "Targets By prstory", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Warning)
            tgtWin.ShowDialog()
        End If
    End Sub

    Private Sub HideMenu()

        StockBlockA.Visibility = Visibility.Visible
    End Sub
#End Region

    Private Sub prstory_linedropdown_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If Not IsInitialized Then
            Return
        End If

        tempreasonlevel = ""
    End Sub

#Region "Progress Bar"
    Private Delegate Sub DelegateUpdateProgressBar()
    Private Sub updateProgressBar_TimeDriven()
        MainProgressBar.Visibility = Visibility.Visible

        With MainProgressBar
            .Value = 0
            While .Value < .Maximum - 2 And Not PROF_connectionError
                .Value += 1.8
                System.Windows.Forms.Application.DoEvents()
                Thread.Sleep(200)
            End While
            .Visibility = Visibility.Hidden
            'System.Windows.Forms.Application.DoEvents()
        End With
    End Sub

    Private Delegate Sub DelegateUpdateProgressBar_Fast()
    Private Sub updateProgressBar_TimeDriven_Fast()
        MainProgressBar.Visibility = Visibility.Visible

        With MainProgressBar
            .Value = 25
            Thread.Sleep(70)
            .Value = 60
            While .Value < .Maximum - 2 And Not PROF_connectionError
                .Value += 10
                System.Windows.Forms.Application.DoEvents()
                Thread.Sleep(150)
            End While
            .Visibility = Visibility.Hidden
        End With
    End Sub


    Private Sub updateProgressBar_Slow()
        MainProgressBar.Visibility = Visibility.Visible

        With MainProgressBar
            .Value = 0
            While .Value < .Maximum - 2 And Not PROF_connectionError
                .Value += 0.8
                System.Windows.Forms.Application.DoEvents()
                Thread.Sleep(200)
            End While
            .Visibility = Visibility.Hidden
        End With
    End Sub


#End Region

#Region "Do Analyze"
    'Handles prstory_Analyze_Icon.
    Private Sub bw_DoWork()
        System.Threading.Thread.Sleep(600)
    End Sub
    Private Sub bw_RunWorkerCompleted()
        BusyIndicator.IsBusy = False
        ' Dim x = New Window_Celebrations(AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Name)
        '     x.Show()
    End Sub

    Private bargraphReportWin As bargraphreportwindow
    Public Sub Do_Analyze()
#Region "April Fools"
        'If isAPRILFOOLS And AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).prStoryMapping = prStoryMapping.APRILFOOLS Then
        ' Dim BusyContent As New List(Of String)
        ' Dim rnd = New Random()

        'BusyContent.Add("Programming Flux Capacitor...")
        'BusyContent.Add("Warming Hyperdrive...")
        'BusyContent.Add("Spinning up the hamster...")
        'BusyContent.Add("Shovelling coal into the server...")
        'BusyContent.Add("Gremlins frantically finding data...")
        'BusyContent.Add("Waiting for Godot...")
        'BusyContent.Add("Replacing the vacuum tubes...")
        'BusyContent.Add("Determining Universal Physical Constants...")

        'Dim randomLineName = BusyContent(rnd.Next(0, BusyContent.Count))

        'BusyIndicator.BusyContent = randomLineName
        'BusyIndicator.IsBusy = True

        '       Dim bw = New BackgroundWorker()
        '       bw.WorkerReportsProgress = True
        '       bw.WorkerSupportsCancellation = True
        '       AddHandler bw.DoWork, AddressOf bw_DoWork
        'AddHandler bw.ProgressChanged, AddressOf bw_ProgressChanged
        '      AddHandler bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted
        '     bw.RunWorkerAsync()
#End Region
        '  Else
        '   Try

        If prepareForDoAnalyze() Then

            Dim i As Integer, rawDTdataColumns As Integer
            Dim _startTime As Date
            Dim _endTime As Date
            Dim lineToAnalyze As ProdLine
            Dim netEvents As Long
            Dim progBar As DelegateUpdateProgressBar = AddressOf updateProgressBar_TimeDriven
            Dim progBarFast As DelegateUpdateProgressBar = AddressOf updateProgressBar_TimeDriven_Fast
            Dim progBarSlow As DelegateUpdateProgressBar = AddressOf updateProgressBar_Slow

            Dim blockedStarvedError As Boolean
            blockedStarvedError = False

            rateLossDT = Nothing
            shouldSnakeClose = False
            doAnalyzeCheckpoint = False
            snakeThread = New Thread(AddressOf PlaySnake)
            PROF_connectionError = False
            PROF_secondaryConnectionError = False
            getProdDataThread = New Thread(AddressOf Do_Analyze_Prod)
            getRateLossdataThread = New Thread(AddressOf Do_Analyze_RateLoss)
            inControlThread = New Thread(AddressOf createInControlReport_Thread)
            motionStopsThread = New Thread(AddressOf createMotionStops_Thread)
            motionPRThread = New Thread(AddressOf createMotionPR_Thread)

            Dim paramObj_One(3) As Object 'DOWNTIME
            Dim paramObj_Two(3) As Object 'PRODUCTION
            Dim paramObj_Three(3) As Object ' TEMPORARY Only for lines which are less than 3 months old
            Dim paramObj_RL(3) As Object

            selectedindexofLine_temp = activeLineIndeces(selectedindexofLine_temp - 1)
            'check if we need to override the default first shift start time
            If EnableFSSTOverrideBox.IsChecked Then
                Dim x As Double
                x = FSSTHourBox.SelectedItem + (FSSTMinBox.SelectedItem / 60) * 100
                AllProdLines(selectedindexofLine_temp)._DayStartTimeHrs = 0
            End If

            ReInitializeAllPublicVariables()
            If AllProdLines(selectedindexofLine_temp).SiteName = "Demo" Then
                UseDemoData = True
            End If
            lineToAnalyze = AllProdLines(selectedindexofLine_temp)
            If lineToAnalyze.parentModule.SQLprocedurePROD = DefaultProficyProductionProcedure.NA Then
                My.Settings.AdvancedSettings_isAvailabilityMode = True
            End If
            My.Settings.defaultDownTimeField = lineToAnalyze.MappingLevelA
            My.Settings.defaultDownTimeField_Secondary = lineToAnalyze.MappingLevelB
            If My.Settings.AdvancedSettings_isAvailabilityMode Then AvailabilityTrueBox.IsChecked = True

            _endTime = prstory_datepicker_enddate.SelectedDate

            If Not UseDemoData Then



                getDTdataThread = New Thread(AddressOf Do_Analyze_DT)
                getDTdataThread2 = New Thread(AddressOf Do_Analyze_DT1)
                getDTdataThread3 = New Thread(AddressOf Do_Analyze_DT2)

                If lineToAnalyze.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.GLEDS Or lineToAnalyze.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.RE_CentralServer Or AllProdLines(selectedindexofLine_temp).Name.Contains("DIHU") Then
                    useThreadingForDataPulling = False
                End If

                If lineToAnalyze.IsStartupMode Then useThreadingForDataPulling = False




                If DoAnalyze_DoWeNeedData(_endTime, lineToAnalyze) Then
                    _startTime = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, _endTime) '-30

                    paramObj_One(0) = selectedindexofLine_temp
                    paramObj_One(1) = _startTime
                    paramObj_One(2) = _endTime

                    paramObj_Two(0) = selectedindexofLine_temp
                    paramObj_Two(1) = DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime)
                    paramObj_Two(2) = _endTime

                    paramObj_RL(0) = selectedindexofLine_temp
                    paramObj_RL(1) = prstory_datepicker_startdate.SelectedDate
                    paramObj_RL(2) = _endTime

                    If My.Settings.AdvancedSettings_PlaySnake Then snakeThread.Start()

                    If Not My.Settings.AdvancedSettings_isAvailabilityMode Then getProdDataThread.Start(paramObj_Two)

                    If useThreadingForDataPulling Then
                        getDTdataThread.Start(paramObj_One)
                        getDTdataThread2.Start(paramObj_One)
                        getDTdataThread3.Start(paramObj_One)
                        If lineToAnalyze.RateLoss_Mode = RateLossMode.Separate Then getRateLossdataThread.Start(paramObj_RL)
                    Else
                        If lineToAnalyze.IsStartupMode Then
                            paramObj_Three(0) = selectedindexofLine_temp
                            paramObj_Three(1) = prstory_datepicker_startdate.SelectedDate
                            paramObj_Three(2) = _endTime
                            getDTdataThread.Start(paramObj_Three)
                            _startTime = prstory_datepicker_startdate.SelectedDate
                        Else
                            getDTdataThread.Start(paramObj_Two)
                        End If
                    End If

                    If lineToAnalyze._isDualConstraint Then
                        progBarSlow.Invoke()
                    Else
                        progBar.Invoke()
                    End If


                    If useThreadingForDataPulling Then
                        While IsNothing(first10daysDT) And Not PROF_connectionError
                            Thread.Sleep(500)
                        End While
                    End If




                End If
            End If

            If Not PROF_connectionError Then
                If DoAnalyze_DoWeNeedData(_endTime, lineToAnalyze) Then
                    While IsNothing(third10daysDT)
                        If Not PROF_connectionError Then
                            Thread.Sleep(500)
                        Else
                            GoTo WEREGOINGTONEEDABIGGERBOAT
                        End If
                    End While
                    If useThreadingForDataPulling Then
                        While IsNothing(second10daysDT)
                            If Not PROF_connectionError Then
                                Thread.Sleep(500)
                            Else
                                GoTo WEREGOINGTONEEDABIGGERBOAT
                            End If
                        End While
                        Select Case lineToAnalyze.SQLdowntimeProcedure
                            Case DefaultProficyDowntimeProcedure.OneClick
                                rawDTdataColumns = 20
                            Case DefaultProficyDowntimeProcedure.QuickQuery
                                rawDTdataColumns = 29
                            Case DefaultProficyDowntimeProcedure.GLEDS
                                rawDTdataColumns = 29
                            Case DefaultProficyDowntimeProcedure.Maple
                                rawDTdataColumns = 20
                        End Select
                    End If
                    If useThreadingForDataPulling Then
                        netEvents = first10daysDT.GetLength(1) + second10daysDT.GetLength(1) + third10daysDT.GetLength(1)
                    Else
                        netEvents = third10daysDT.GetLength(1)
                    End If

                    Dim completeDTarray(rawDTdataColumns, netEvents - 1) As Object
                    If useThreadingForDataPulling Then
                        For i = 0 To rawDTdataColumns
                            Try
                                Array.Copy(first10daysDT, i * (first10daysDT.GetLength(1)), completeDTarray, i * netEvents, first10daysDT.GetLength(1))
                                System.Array.Copy(second10daysDT, i * (second10daysDT.GetLength(1)), completeDTarray, i * netEvents + first10daysDT.GetLength(1), second10daysDT.GetLength(1))
                                System.Array.Copy(third10daysDT, i * (third10daysDT.GetLength(1)), completeDTarray, i * netEvents + first10daysDT.GetLength(1) + second10daysDT.GetLength(1), third10daysDT.GetLength(1))
                            Catch e As Exception

                            End Try
                        Next
                    ElseIf lineToAnalyze._isDualConstraint Then
                        completeDTarray = third10daysDT
                    End If
                    'CHECK FOR DUAL CONSTRAINT
                    Dim rateLossData(,) As Object = New Object
                    Dim rateLossData2(,) As Object = New Object
                    If lineToAnalyze._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode <> MultiConstraintAnalysis.SingleConstraint Then
                        If lineToAnalyze.mainProdUnits.Count > 0 Then

                            '''''START>>>>SECTION THAT NEEDS TO BE OPTIMIZED FOR DOVER
                            'figure out what procedure we need

                            Dim fastPull As Boolean
                            fastPull = True

                            Dim isQQ As Boolean
                            If lineToAnalyze.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.QuickQuery Then
                                isQQ = True
                            Else
                                isQQ = False
                            End If

                            If fastPull Then

                                Dim threadList As List(Of Thread) = New List(Of Thread)
                                'find raw data
                                For ix As Integer = 0 To lineToAnalyze.mainProdUnits.Count - 1
                                    Dim paramObj(7) As Object 'DOWNTIME
                                    paramObj(0) = lineToAnalyze.mainProdUnits(ix)
                                    paramObj(1) = DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime)
                                    paramObj(2) = _endTime
                                    paramObj(3) = lineToAnalyze.ProficyServer_Name
                                    paramObj(4) = lineToAnalyze.ProficyServer_Username
                                    paramObj(5) = lineToAnalyze.ProficyServer_Password
                                    paramObj(6) = isQQ

                                    Dim rawThread As Thread

                                    rawThread = New Thread(AddressOf Do_Analyze_DT_Multiunit2)
                                    rawThread.Start(paramObj)
                                    threadList.Add(rawThread)
                                Next ix
                                For Each t In threadList
                                    t.Join()
                                Next

                                'combine raw data
                                If individualRateData.Count > 0 Then
                                    If isQQ Then
                                        rateLossData2 = PROF_mergeRateLossWithMain_Cairo(individualRateData(0), completeDTarray)
                                        For ij As Integer = 1 To individualRateData.Count - 1
                                            rateLossData2 = PROF_mergeRateLossWithMain_Cairo(individualRateData(ij), rateLossData2)
                                        Next
                                    Else
                                        rateLossData2 = PROF_mergeRateLossWithMain_Cairo_OneClick(individualRateData(0), completeDTarray)
                                        For ij As Integer = 1 To individualRateData.Count - 1
                                            rateLossData2 = PROF_mergeRateLossWithMain_Cairo_OneClick(individualRateData(ij), rateLossData2)
                                        Next
                                    End If

                                    lineToAnalyze.BabyWipesData = individualRateData

                                    finalRateLossData = rateLossData2
                                Else
                                    MessageBox.Show("Unable to gather supplementary data.")
                                End If

                            Else
                                For ix As Integer = 0 To lineToAnalyze.mainProdUnits.Count - 1
                                    Try
                                        If isQQ Then
                                            rateLossData = getRawProficyData(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze.mainProdUnits(ix), lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                                        Else 'assumes one click, ie dont have this for maple yet
                                            rateLossData = getRawProficyData_OneClick(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze.mainProdUnits(ix), lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                                        End If

                                        If ix = 0 Or IsNothing(rateLossData2) Then
                                            If isQQ Then
                                                rateLossData2 = PROF_mergeRateLossWithMain_Cairo(rateLossData, completeDTarray)
                                            Else
                                                rateLossData2 = PROF_mergeRateLossWithMain_Cairo_OneClick(rateLossData, completeDTarray)
                                            End If
                                        Else
                                            If isQQ Then
                                                rateLossData2 = PROF_mergeRateLossWithMain_Cairo(rateLossData, rateLossData2)
                                            Else
                                                rateLossData2 = PROF_mergeRateLossWithMain_Cairo_OneClick(rateLossData, rateLossData2)
                                            End If
                                        End If
                                    Catch e As Exception
                                        MessageBox.Show("Unable to pull data for MPU " & lineToAnalyze.mainProdUnits(ix) & ". We will try to continue without that data. Error message: " & e.Message)
                                    End Try
                                Next
                                finalRateLossData = rateLossData2
                            End If
                            '''''END SECTION THAT NEEDS TO BE OPTIMIZED FOR DOVER





                        ElseIf lineToAnalyze.SiteName = "Cairo" Then
                            rateLossData = getRawProficyData(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                            finalRateLossData = PROF_mergeRateLossWithMain_Cairo(rateLossData, completeDTarray)
                        ElseIf My.Settings.AdvancedSettings_isAvailabilityMode Then  ' the assumption here is Family Care is only sector that would run availability mode and dual constraint (Blocked/Starved + ...)
                            Try
                                If Not lineToAnalyze.IsStartupMode Then
                                    rateLossData = getRawProficyData_OneClick(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                                Else 'is IS IN STARTUP MODE!!!
                                    rateLossData = getRawProficyData_OneClick(_endTime, _startTime, lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                                End If

                                If IsNothing(rateLossData) Then
                                    lineToAnalyze._isDualConstraint = False
                                Else
                                    finalRateLossData = PROF_mergeRateLossWithMain_OneClick(rateLossData, completeDTarray)
                                End If
                            Catch ex As Exception
                                lineToAnalyze._isDualConstraint = False
                                blockedStarvedError = True
                            End Try
                        Else 'this is for one health rate loss
                            If lineToAnalyze.SQLdowntimeProcedure = DefaultProficyProductionProcedure.QuickQuery Then
                                rateLossData = getRawProficyData(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                                finalRateLossData = PROF_mergeRateLossWithMain(rateLossData, completeDTarray)
                            Else 'assumes otherwise one click
                                rateLossData = getRawProficyData_OneClick(_endTime, DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime), lineToAnalyze._rateLossDisplay, lineToAnalyze.ProficyServer_Name, lineToAnalyze.ProficyServer_Username, lineToAnalyze.ProficyServer_Password)
                                finalRateLossData = PROF_mergeRateLossWithMain_OneClick(rateLossData, completeDTarray)

                            End If
                        End If
                    End If
                    ''''''''''''''''''''''''

                    While IsNothing(tmpProdArray) And Not My.Settings.AdvancedSettings_isAvailabilityMode
                        If Not PROF_connectionError Then
                            Thread.Sleep(500)
                        Else
                            GoTo WEREGOINGTONEEDABIGGERBOAT
                        End If
                    End While

                    With AllProdLines(selectedindexofLine_temp) 'lineToAnalyze
                        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
                            .rawProficyProductionData = tmpProdArray
                            .rawProductionData = New ProductionDataset(AllProdLines(selectedindexofLine_temp), True)
                        End If

                        .rawProfStartTime = DateAdd(DateInterval.Day, -1 * datapull_duration, _endTime)
                        .rawProfEndTime = _endTime

                        If ._isDualConstraint And My.Settings.AdvancedSettings_MultiConstraintAnalysisMode <> MultiConstraintAnalysis.SingleConstraint Then
                            .rawProficyData = finalRateLossData
                            .rawRateLossData = rateLossData
                            If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
                                PROF_setDTteamsFromProd(selectedindexofLine_temp)

                                If AllProdLines(selectedindexofLine_temp).parentSite.Name = SITE_BROWNS_SUMMIT And AllProdLines(selectedindexofLine_temp).parentModule.Name = BS_SC Then PROF_setDTGroupInProdFromDT(selectedindexofLine_temp)

                                .rawProductionData = New ProductionDataset(AllProdLines(selectedindexofLine_temp), True)
                            End If

                            If .BabyWipesData.Count > 0 Then
                                For Each xx As Object(,) In .BabyWipesData
                                    .BabyWipesPRSTORYData.Add(New DowntimeDataset(AllProdLines(selectedindexofLine_temp), xx))
                                Next
                            End If

                            .rawDowntimeData = New DowntimeDataset(AllProdLines(selectedindexofLine_temp), finalRateLossData)
                        Else
                            If useThreadingForDataPulling Then
                                .rawProficyData = completeDTarray
                            Else
                                .rawProficyData = third10daysDT
                            End If

                            If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
                                PROF_setDTteamsFromProd(selectedindexofLine_temp)
                                If AllProdLines(selectedindexofLine_temp).parentSite.Name = SITE_BROWNS_SUMMIT And AllProdLines(selectedindexofLine_temp).parentModule.Name = BS_SC Then PROF_setDTGroupInProdFromDT(selectedindexofLine_temp)

                                .rawProductionData = New ProductionDataset(AllProdLines(selectedindexofLine_temp), True)

                            End If
                            If useThreadingForDataPulling Then
                                .rawDowntimeData = New DowntimeDataset(AllProdLines(selectedindexofLine_temp), completeDTarray)
                            Else
                                .rawDowntimeData = New DowntimeDataset(AllProdLines(selectedindexofLine_temp), third10daysDT)

                            End If
                        End If

                    End With

                    tmpProdArray = Nothing
                    first10daysDT = Nothing
                    second10daysDT = Nothing
                    third10daysDT = Nothing
                Else ' we didn't need to pull data!


                    If UseDemoData Then
                        Dim x = New DemoData()
                        AllProdLines(selectedindexofLine_temp).rawDowntimeData = x.CreateDowntimeData(selectedindexofLine_temp, _endTime.AddDays(-1 * datapull_duration), _endTime)
                    End If


                    'first restore default mapping
                    My.Settings.defaultDownTimeField = AllProdLines(selectedindexofLine_temp).MappingLevelA
                    My.Settings.defaultDownTimeField_Secondary = AllProdLines(selectedindexofLine_temp).MappingLevelB
                    AllProdLines(selectedindexofLine_temp).reMapRawData()
                    'do the progress bar!
                    progBarFast.Invoke()
                End If

                If AllProdLines(selectedindexofLine_temp).BabyWipesPRSTORYData.Count > 0 Then
                    prstoryReport = New prStoryMainPageReport(selectedindexofLine_temp, prstory_datepicker_startdate.SelectedDate, prstory_datepicker_enddate.SelectedDate, True)
                Else
                    prstoryReport = New prStoryMainPageReport(selectedindexofLine_temp, prstory_datepicker_startdate.SelectedDate, prstory_datepicker_enddate.SelectedDate)
                End If
                bargraphReportWin = New bargraphreportwindow(New prStoryMainPageReport(selectedindexofLine_temp, prstory_datepicker_startdate.SelectedDate, prstory_datepicker_enddate.SelectedDate))
                doAnalyzeCheckpoint = True

                'update the new adjusted dates
                prstory_dateselectionLabel.Content = Format(prstory_datepicker_startdate.SelectedDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine & Format(prstory_datepicker_enddate.SelectedDate, "MMMM dd, yyyy HH:mm").ToString & vbNewLine
                prstory_dateselectionLabel.HorizontalContentAlignment = Windows.HorizontalAlignment.Center

                bargraphReportWin.multilineGroups = multilineGroups

                bargraphReportWin.Owner = Me

                datalabelcontent = prstory_dateselectionLabel.Content
                DeactivateIncontrol = True  ' by default deactivated

                If Not IsNothing(rateLossDT) Then
                    AllProdLines(selectedindexofLine_temp).RawRateLossDataArray = rateLossDT
                    bargraphReportWin.RateLossReport_Icon.Visibility = Visibility.Visible
                    bargraphReportWin.RateLossReport_Label.Visibility = Visibility.Visible
                End If

                bargraphreportwindow_Open = True
                bargraphReportWin.Show()
                If Me.WindowState = Windows.WindowState.Maximized Then bargraphReportWin.WindowState = Windows.WindowState.Maximized
                Me.Visibility = Visibility.Hidden

                If Not My.Computer.Keyboard.CtrlKeyDown Then shouldSnakeClose = True
                If Not lineToAnalyze.IsStartupMode And lineToAnalyze.parentModule.Name <> "Baby Wipes" And lineToAnalyze.Name <> "IP74" And lineToAnalyze.Name <> "DIBH11" And lineToAnalyze.Name <> "S" And lineToAnalyze.Name <> "W Front" And lineToAnalyze.Name <> "W Back" And AllProdLines(selectedindexofLine_temp).parentModule.Name <> BS_APDO Then
                    inControlThread.Start()
                    motionPRThread.Start()
                    motionStopsThread.Start()
                End If

                If blockedStarvedError Then
                    bargraphReportWin.showRateLossError()
                End If

            Else 'this means there was a proficy connection error!
WEREGOINGTONEEDABIGGERBOAT:
                shouldSnakeClose = True
                Thread.Sleep(200)
                ReInitializeAllPublicVariables()
            End If
        End If

    End Sub



    Private Sub Do_Analyze_DT_Multiunit2(ByVal paramObj As Object)
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineName As String
        Dim serverName As String, serverPassword As String, serverUsername As String
        Dim isQQ As Boolean
        lineName = paramObj(0)
        STARTx = paramObj(1)
        ENDx = paramObj(2)
        serverName = paramObj(3)
        serverUsername = paramObj(4)
        serverPassword = paramObj(5)
        isQQ = paramObj(6)

        Dim retData As Object(,)

        Try
            If Not isQQ Then
                retData = getRawProficyData_OneClick(ENDx, STARTx, lineName, serverName, serverUsername, serverPassword)
            Else
                retData = getRawProficyData(ENDx, STARTx, lineName, serverName, serverUsername, serverPassword)
            End If
            individualRateData.Add(retData)
        Catch ex As System.Runtime.InteropServices.COMException
            PROF_connectionError = True
            MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
        Catch ex As Exception
            If Not PROF_connectionError Then
                PROF_connectionError = True

                MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)

            End If
        Finally
        End Try
    End Sub

    Private Function DoAnalyze_DoWeNeedData(endTime As Date, parentLine As ProdLine) As Boolean
        If UseDemoData Then
            Return False
        End If
        Dim startTime As Date = DateAdd(DateInterval.Day, -1 * (datapull_duration - 1), endTime)
        If (Not My.Settings.AdvancedSettings_isAvailabilityMode) And IsNothing(AllProdLines(selectedindexofLine_temp).rawProficyProductionData) Then Return True
        If IsExcludedEventsIncluded = True Then Return True  'LG Code addition to accomodate all events inclusion (PR out inclusion)
        If My.Settings.EnableTimeSpanExclusion Then Return True
        If DoWeAlwaysRepullDataNoMatterWhatElseHasHappened Then Return True
        If startTime >= parentLine.rawProfStartTime And endTime <= parentLine.rawProfEndTime Then Return False
        Return True
    End Function


    Private Sub createInControlReport_Thread()

        MasterDataSet = New inControlReport(AllProdLines(selectedindexofLine_temp), endtimeselected.AddDays(-1), endtimeselected) ' LG Code

    End Sub
    Private Sub imagedownload(imagefilename As String)
        Try
            If My.Computer.Network.Ping("prstory.pg.com") Then
                If CheckIfFtpFileExists("ftp://prstory.pg.com/prStory_img/" & imagefilename, "normalusers", "pgdigitalfactory412") Then
                    Dim sourcewebaddress As Uri = New Uri("http://prstory.pg.com/prStory_img/" & imagefilename)

                    Dim destinationfolderaddress As String = SERVER_FOLDER_PATH & imagefilename

                    Dim myWebClient As New WebClient()
                    My.Computer.Network.DownloadFile(sourcewebaddress, destinationfolderaddress, "", "", False, 10000, True)
                Else

                End If
            End If
        Catch e As Exception
            Exit Sub
        End Try

    End Sub

    Private Sub createMotionStops_Thread()
        Dim stopsMotionReport As MotionReport = New MotionReport(AllProdLines(selectedindexofLine_temp), starttimeselected, endtimeselected, prstoryReport.MainLEDSReport.DT_Report.UnplannedEventDirectory, 1)
        Dim k As Integer

        For k = 0 To Math.Min(14, stopsMotionReport.motionEvents_All15.Count - 1)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode(stopsMotionReport, False, k)

            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Weekly(stopsMotionReport, False, k)

            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_Monthly(stopsMotionReport, False, k)

            exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Weekly(stopsMotionReport, True, k)
            exportMotion_PR_HTML_AMCHART_selectedfailuremode_MTBF_Monthly(stopsMotionReport, True, k)

        Next k

        'Check if images have been downloaded for trend charts
        If System.IO.File.Exists(SERVER_FOLDER_PATH & "dragIconRoundBig.png") Then
            'The file exists
        Else
            'the file doesn't exist
            imagedownload("dragIconRoundBig.png")
        End If
        If System.IO.File.Exists(SERVER_FOLDER_PATH & "lens.png") Then
            'The file exists
        Else
            'the file doesn't exist
            imagedownload("lens.png")
        End If
        'we'll try to sneak this in here and see how it goes
        CSV_exportAllRawDataFromArrayObject(AllProdLines(selectedindexofLine_temp))

    End Sub
    Private Sub createMotionPR_Thread()
        Try
            Dim prMotionReport As Motion_LinePRReport = New Motion_LinePRReport(AllProdLines(selectedindexofLine_temp), 1)
            'exportMotion_PR_HTML(prMotionReport)
            exportMotion_PR_HTML_AMCHART(prMotionReport)
            exportMotion_PR_HTML_AMCHART_Weekly(prMotionReport)
            exportMotion_PR_HTML_AMCHART_Monthly(prMotionReport)

            exportMotion_SPD_HTML_AMCHART(prMotionReport)
            exportMotion_SPD_HTML_AMCHART_Weekly(prMotionReport)
            exportMotion_SPD_HTML_AMCHART_Monthly(prMotionReport)
        Catch ex As Exception
            MsgBox("Trend Error: " & ex.Message)
        End Try

    End Sub

    Private Function prepareForDoAnalyze() As Boolean
        Dim msgTextTmp As String
        selectedindexofLine_temp = prstory_linedropdown.SelectedIndex

        If selectedindexofLine_temp < 1 Then                                   'handles user generated no-line selection error

            Select Case My.Settings.LanguageActive
                Case Lang.English
                    msgTextTmp = "A line needs to be selected to initialize analysis"
                Case Lang.German
                    msgTextTmp = "für die Analyse ausgewählt werden, eine Linie muss"
                Case Lang.Spanish
                    msgTextTmp = "una línea necesita ser seleccionado para el análisis"
                Case Lang.French
                    msgTextTmp = "une ligne doit être sélectionnée pour l'analyse"
                Case Lang.Portuguese
                    msgTextTmp = "A linha tem de ser seleccionado para análise"
                Case Lang.Chinese_Simplified
                    msgTextTmp = "选择生产线"
                Case Else
                    msgTextTmp = "A line needs to be selected to initialize analysis"
            End Select
            LineSelection_Alert.Content = msgTextTmp
            LineSelection_Alert.Visibility = Visibility.Visible
            Return False
        End If

        If IsNothing(prstory_datepicker_enddate.SelectedDate) Or IsNothing(prstory_datepicker_startdate.SelectedDate) Then    'handles user generated  no-date selection error
            launchprstorysettings()
            launchprstorydaterange()
            prstorymtdoption_MouseDown()
            IsAnalyzeButtonClickSource_Analyze = True
            DateSelection_Alert.Visibility = Visibility.Visible
            Return False
        End If

        LineSelection_Alert.Visibility = Visibility.Hidden
        Return True
    End Function


#Region "Single Unit Data Collection"
    Private Sub Do_Analyze_DT(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As String, databaseName As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnit
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With
        STARTx = paramObj(1)
        ENDx = paramObj(2)

        Try
            Select Case preferredDtQuery
                Case DefaultProficyDowntimeProcedure.OneClick
                    third10daysDT = getRawProficyData_OneClick(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.OneClick_V27
                    third10daysDT = getRawProficyData_OneClick_v27(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    third10daysDT = getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.GLEDS
                    third10daysDT = getGLEDSData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.Maple
                    third10daysDT = getMapleData(ENDx, STARTx, tmpString, prodUnit, serverName, serverUsername, serverPassword, databaseName)
            End Select
        Catch ex As System.Runtime.InteropServices.COMException

            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally
        End Try
    End Sub
    Private Sub Do_Analyze_DT1(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As String, databaseName As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnit
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With
        STARTx = DateAdd(DateInterval.Day, -1 * (datapull_duration - (datapull_duration / 3)), paramObj(1)) '-60
        ENDx = DateAdd(DateInterval.Day, -1 * (datapull_duration - (datapull_duration / 3)), paramObj(2)) '-60

        Try
            Select Case preferredDtQuery
                Case DefaultProficyDowntimeProcedure.OneClick
                    first10daysDT = getRawProficyData_OneClick(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    first10daysDT = getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.OneClick_V27
                    third10daysDT = getRawProficyData_OneClick_v27(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)

                Case DefaultProficyDowntimeProcedure.GLEDS
                    first10daysDT = getGLEDSData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.Maple
                    first10daysDT = getMapleData(ENDx, STARTx, tmpString, prodUnit, serverName, serverUsername, serverPassword, databaseName)
            End Select
        Catch ex As System.Runtime.InteropServices.COMException
            ' Debugger.Break()
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally
        End Try
    End Sub
    Private Sub Do_Analyze_DT2(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date

        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As String, databaseName As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnit
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With

        STARTx = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, paramObj(1)) '-30
        ENDx = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, paramObj(2))    '-30

        Try
            Select Case preferredDtQuery
                Case DefaultProficyDowntimeProcedure.OneClick
                    second10daysDT = getRawProficyData_OneClick(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.QuickQuery
                    second10daysDT = getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.GLEDS
                    second10daysDT = getGLEDSData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyDowntimeProcedure.Maple
                    second10daysDT = getMapleData(ENDx, STARTx, tmpString, prodUnit, serverName, serverUsername, serverPassword, databaseName)
                Case DefaultProficyDowntimeProcedure.OneClick_V27
                    third10daysDT = getRawProficyData_OneClick_v27(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
            End Select
        Catch ex As System.Runtime.InteropServices.COMException
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally

        End Try
    End Sub
#End Region

#Region "Multi Unit Data Collection"
    Private Sub Do_Analyze_DT_MultiUnit(ByVal paramObj As Object)
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As List(Of String), databaseName As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnits
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With
        STARTx = paramObj(1)
        ENDx = paramObj(2)

        Try
            If Not isLineUsingMOT Then
                Select Case preferredDtQuery
                    Case DefaultProficyDowntimeProcedure.OneClick_MultiUnit
                        third10daysDT = getRawProficyData_OneClick_MultiUnit(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                    Case DefaultProficyDowntimeProcedure.QuickQuery_MultiUnit
                        third10daysDT = getRawProficyData_OneClick_MultiUnit(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                End Select
            Else
                MsgBox("Error - Line Not Configured For MOT SDK Data Access!")
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            ' Debugger.Break()
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally
        End Try
    End Sub
    Private Sub Do_Analyze_DT1_MultiUnit(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As List(Of String), databaseName As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnits
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With
        STARTx = DateAdd(DateInterval.Day, -1 * (datapull_duration - (datapull_duration / 3)), paramObj(1)) '-60
        ENDx = DateAdd(DateInterval.Day, -1 * (datapull_duration - (datapull_duration / 3)), paramObj(2)) '-60

        Try
            If Not isLineUsingMOT Then
                Select Case preferredDtQuery
                    Case DefaultProficyDowntimeProcedure.OneClick_MultiUnit
                        first10daysDT = getRawProficyData_OneClick_MultiUnit(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                    Case DefaultProficyDowntimeProcedure.QuickQuery_MultiUnit
                        first10daysDT = getRawProficyData_OneClick_MultiUnit(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                End Select
            Else
                MsgBox("Error - Line Not Configured For MOT SDK Data Access!")
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            ' Debugger.Break()
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally
        End Try
    End Sub
    Private Sub Do_Analyze_DT2_MultiUnit(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date
        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredDtQuery As Integer, prodUnit As List(Of String), databaseName As String
        Dim tmpString As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredDtQuery = .SQLdowntimeProcedure
            prodUnit = .mainProdUnits
            tmpString = .Name_MAPLE
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With

        STARTx = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, paramObj(1))
        ENDx = DateAdd(DateInterval.Day, -1 * datapull_duration / 3, paramObj(2))

        Try
            If Not isLineUsingMOT Then
                Select Case preferredDtQuery
                    Case DefaultProficyDowntimeProcedure.OneClick_MultiUnit
                        second10daysDT = getRawProficyData_OneClick_MultiUnit(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                    Case DefaultProficyDowntimeProcedure.QuickQuery_MultiUnit
                        second10daysDT = getRawProficyData_OneClick_MultiUnit(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                End Select
            Else
                MsgBox("Error - Line Not Configured For MOT SDK Data Access!")
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_connectionError Then
                PROF_connectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally

        End Try
    End Sub
#End Region
    Private Sub Do_Analyze_RateLoss(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date) 'As Array
        Dim STARTx As Date
        Dim ENDx As Date

        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, preferredDtQuery As Integer, serverUsername As String, prodUnit As String
        Dim tmpString As String
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            prodUnit = ._rateLossDisplay '.mainProdUnit
            preferredDtQuery = .SQLdowntimeProcedure
            tmpString = .Name_MAPLE
        End With

        STARTx = paramObj(1)
        ENDx = paramObj(2)

        If Math.Abs(DateDiff(DateInterval.Day, STARTx, ENDx)) > 80 Then STARTx = ENDx.AddDays(-70)

        Try
            If AllProdLines(lineIndex).SiteName = SITE_CRUX Then
                rateLossDT = getRawProficyData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
            Else
                rateLossDT = getRawProficyData_OneClick_v27(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            If Not PROF_secondaryConnectionError Then
                PROF_secondaryConnectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server for rate loss data. Please check your internet connection and try again." & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_secondaryConnectionError Then
                PROF_secondaryConnectionError = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to connect to server. Please check your internet connection and try again." & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally
            Dispatcher.Invoke(Windows.Threading.DispatcherPriority.Normal, New Action(AddressOf showRateLossIcon))
        End Try
    End Sub
    Private Sub showRateLossIcon()
        Try
            bargraphReportWin.RateLossReport_Icon.Visibility = Visibility.Visible
            bargraphReportWin.RateLossReport_Label.Visibility = Visibility.Visible
            AllProdLines(selectedindexofLine_temp).RawRateLossDataArray = rateLossDT
        Catch e As Exception
            '         If (doAnalyzeCheckpoint) Then
            '         MessageBox.Show("Rateloss Failure. " & e.Message)
            '         End If
        End Try
    End Sub

    Private Sub Do_Analyze_Prod(ByVal paramObj As Object) 'ByVal serverName As String, ByVal prodUnit As String, ByVal startDate As Date, ByVal endDate As Date)
        Dim STARTx As Date
        Dim ENDx As Date

        Dim lineIndex As Integer
        Dim serverName As String, serverPassword As String, serverUsername As String, preferredProdQuery As Integer, prodUnit As String, databaseName As String
        Dim isLineUsingMOT As Boolean
        lineIndex = paramObj(0)
        With AllProdLines(lineIndex)
            serverName = .ProficyServer_Name
            serverPassword = .ProficyServer_Password
            serverUsername = .ProficyServer_Username
            preferredProdQuery = .SQLproductionProcedure
            prodUnit = .mainProfProd
            isLineUsingMOT = .IS_MOT
            databaseName = .ServerDatabase
        End With

        STARTx = paramObj(1)
        ENDx = paramObj(2)

        Try
            Select Case preferredProdQuery
                Case DefaultProficyProductionProcedure.QuickQuery
                    tmpProdArray = getRawProficyProductionData(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0)) 'paramObj(3), paramObj(2), paramObj(1), paramObj(0))
                Case DefaultProficyProductionProcedure.SwingRoad
                    tmpProdArray = getRawProficyProductionData_SwingRoad(ENDx, STARTx, prodUnit, serverName, serverUsername, serverPassword)
                Case DefaultProficyProductionProcedure.Maple
                    tmpProdArray = getMaplePRODData(ENDx, STARTx, prodUnit, prodUnit, serverName, serverUsername, serverPassword, databaseName)
                Case DefaultProficyProductionProcedure.Maple_New
                    tmpProdArray = getMaplePRODData(ENDx, STARTx, prodUnit, prodUnit, serverName, serverUsername, serverPassword, databaseName)

            End Select
        Catch ex As System.Runtime.InteropServices.COMException
            If Not PROF_connectionError Then
                ' PROF_connectionError = True
                My.Settings.AdvancedSettings_isAvailabilityMode = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to retrieve production data from server. We will attempt to run in availability mode. If the error persists please contact your SPOC.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Catch ex As Exception
            If Not PROF_connectionError Then
                ' PROF_connectionError = True
                My.Settings.AdvancedSettings_isAvailabilityMode = True
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("Unable to retrieve production data from server. We will attempt to run in availability mode. If the error persists please contact your SPOC.   " & ex.Message, vbCritical)
                    Case Lang.German
                        MsgBox("keine Verbindung zum Server zu verbinden, überprüfen Sie bitte Ihre Internetverbindung und versuchen Sie es erneut", vbCritical)
                    Case Lang.Spanish
                        MsgBox("no puede conectarse al servidor, compruebe su conexión a Internet y vuelva a intentarlo", vbCritical)
                    Case Lang.French
                        MsgBox("incapable de se connecter au serveur, s'il vous plaît vérifier votre connexion Internet et essayez à nouveau", vbCritical)
                    Case Lang.Portuguese
                        MsgBox("incapaz de se conectar ao servidor, verifique a sua ligação à Internet e tente novamente", vbCritical)
                    Case Lang.Chinese_Simplified
                        MsgBox("无法连接到服务器", vbCritical)
                End Select
            End If
        Finally
        End Try
    End Sub
#End Region

    Private Sub MappingSelectionChange()
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
        End Select
        tempreasonlevel = MappingSelectionCombo.SelectedValue
    End Sub
    Private Sub LaunchAboutWindow()
        HideMenu()
        Dim webAddress As String = "http://prstory.pg.com/#one"
        Process.Start(webAddress)
    End Sub
    Private Sub hideaboutwebcontrol()
        gobackbutton.Visibility = Visibility.Hidden
    End Sub
    Public Function WhichMapping(mappingno As Integer) As String
        If prstory_linedropdown.SelectedIndex > 0 Then
            Select Case mappingno
                Case DowntimeField.Reason1
                    Return AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason1Name

                Case DowntimeField.Reason2
                    Return AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason2Name

                Case DowntimeField.Reason3
                    Return AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason3Name

                Case DowntimeField.Reason4
                    Return AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason4Name

                Case DowntimeField.Fault
                    Return AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).FaultCodeName

                Case DowntimeField.DTGroup
                    Return AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).DTgroupName
                Case Else
                    Return ""
            End Select
        Else
            Select Case mappingno
                Case DowntimeField.Reason1
                    Return "Reason 1" 'AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason1Name

                Case DowntimeField.Reason2
                    Return "Reason 2" 'AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason2Name

                Case DowntimeField.Reason3
                    Return "Reason 3" 'AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason3Name

                Case DowntimeField.Reason4
                    Return "Reason 4" 'AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Reason4Name

                Case DowntimeField.Fault
                    Return "Fault" 'AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).FaultCodeName

                Case DowntimeField.DTGroup
                    Return "DTGroup" 'AllProductionLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).DTgroupName
                Case Else
                    Throw New unknownMappingException
            End Select
        End If
    End Function
    Sub populatemappingdropdown()
        With MappingSelectionCombo.Items
            .Clear()
            .Add(AllProdLines(selectedindexofLine_temp).Reason1Name) ' 0
            .Add(AllProdLines(selectedindexofLine_temp).Reason2Name) '1
            .Add(AllProdLines(selectedindexofLine_temp).Reason3Name) '2
            .Add(AllProdLines(selectedindexofLine_temp).Reason4Name) '3
            .Add(AllProdLines(selectedindexofLine_temp).DTgroupName) '4
            .Add(AllProdLines(selectedindexofLine_temp).FaultCodeName) '5
        End With
        MappingSelectionCombo.SelectedValue = WhichMapping(My.Settings.defaultDownTimeField)
        tempreasonlevel = WhichMapping(My.Settings.defaultDownTimeField)
    End Sub
#Region "LanguageOptions"

    Sub populateLanguageDropdown()
        With LanguageSelectionCombo.Items
            .Clear()
            .Add("English")
            .Add("Deutsch")
            .Add("Español")
            .Add("Français")
            .Add("Português")
            '.Add("漢字簡化爭論")
            .Add("中文")
        End With
        LanguageSelectionCombo.SelectedIndex = My.Settings.LanguageActive
    End Sub
    Private Sub LanguageSelectionChange()
        My.Settings.LanguageActive = LanguageSelectionCombo.SelectedIndex
        initializeMenuTextFromLanguage()
    End Sub

    Private Sub initializeMenuTextFromLanguage()
        Select Case My.Settings.LanguageActive
            Case Lang.English
                prstory_dateselectionLabel.Content = "[Dates not selected yet]"
                tmpSector = New BusinessUnit("All Sectors")
                tmpSite = New ProdSite("All Sites", "", "", "", "", "")
                prstory_settingsLabel.Content = "Settings"
                prstory_analyzeLabel.Content = "Analyze"
                prstory_targetLabel.Content = "About"
                prstorydayoption.Content = "Pick Date Range"
                TargetLaunchButton.Content = "Targets"
                inControlSettingsButton.Content = "inControl Settings"
                AdvancedSettingsButton.Content = "Advanced Options"
                prstoryGoButton.Content = "Confirm"
                prstoryCancelButton.Content = "Cancel"
                prstorystartdate_label.Content = "Start date & time"
                prstoryenddate_label.Content = "End date & time"
                prstorymtdoption.Content = "Month to Date"
                prstorylast7daysoption.Content = "Last 7 Days"
                prstoryyesterdayopion.Content = "Yesterday"
                SAPnameCheckBox.Content = "Display Line SAP Names"
                SnakeCheckBox.Content = "Play Snake While Loading"
                AvailabilityTrueBox.Content = "Availability Mode"
                TargetsEnabledTrueBox.Content = "Targets Enabled"
                inControlRuleOneBox.Content = "Rule 1"
                inControlRuleTwoBox.Content = "Rule 2"
                inControlRuleThreeBox.Content = "Rule 3"
                inControlRuleFourBox.Content = "Rule 4"
                inControlRuleFiveBox.Content = "Rule 5"
                inControlRuleSixBox.Content = "Rule 6"
                inControlSliderLabelOne.Content = "Default"
                inControlSliderLabelTwo.Content = "Custom"
                inControlTitleLabel.Content = "inControl Stability Rules"
                AdvancedSettings_MultiConstraint_Go.Content = "Multi-Constraint Options"
                Set_Default_Line_Button.Content = "Set Selected Line as Default"
                LineDefaultQueryLabel.Content = "Do you want prstory to remember this line as the default selection?"
                LineDefaultYesButton.Content = "Yes"
                LineDefaultNoButton.Content = "No"
                LineDefaultCancelButton.Content = "I'll decide later"


                NotesBox.Content = "Notes Enabled"
                MultiConstraint_ModeAButton.Content = "Single Constraint Model"
                MultiConstraint_ModeBButton.Content = "Standard Multi-Constraint Model"
                MultiConstraint_ModeCButton.Content = "Model rate loss events as stops"
                prstoryCancelButton.Content = "Cancel"

                Multilinelaunchlabel.Content = "Want to analyze multiple lines? Click here."
            Case Lang.Chinese_Simplified
                prstory_dateselectionLabel.Content = "[没有选择]"

                prstory_settingsLabel.Content = "设定"
                prstory_analyzeLabel.Content = "分析"
                prstory_targetLabel.Content = "关于"
                prstorydayoption.Content = "选择日期"
                TargetLaunchButton.Content = "目标"
                inControlSettingsButton.Content = "inControl 设定"
                AdvancedSettingsButton.Content = "高级选项"
                prstoryGoButton.Content = "确认"
                prstoryCancelButton.Content = "取消"
                prstorystartdate_label.Content = "开始日期和时间"
                prstoryenddate_label.Content = "结束日期和时间"
                prstorymtdoption.Content = "本月"
                prstorylast7daysoption.Content = "过去7天"
                prstoryyesterdayopion.Content = "昨天"
                SAPnameCheckBox.Content = "显示生产线SAP名称"
                SnakeCheckBox.Content = "prstory 在载入时播放蛇型线"
                AvailabilityTrueBox.Content = "可靠性模式"
                TargetsEnabledTrueBox.Content = "启用目标"
                inControlRuleOneBox.Content = "规则  1"
                inControlRuleTwoBox.Content = "规则  2"
                inControlRuleThreeBox.Content = "规则  3"
                inControlRuleFourBox.Content = "规则  4"
                inControlRuleFiveBox.Content = "规则  5"
                inControlRuleSixBox.Content = "规则  6"
                inControlSliderLabelOne.Content = "默认"
                inControlSliderLabelTwo.Content = "顾客"
                inControlTitleLabel.Content = "inControl 稳定性规则 "
                AdvancedSettings_MultiConstraint_Go.Content = "多种限制选项"
                Set_Default_Line_Button.Content = "设置选中的行为默认"
                LineDefaultQueryLabel.Content = "希望prstory记住这条生产线?"
                LineDefaultYesButton.Content = "是"
                LineDefaultNoButton.Content = "不是"
                LineDefaultCancelButton.Content = "迟些"
                NotesBox.Content = "注意启用"
                MultiConstraint_ModeAButton.Content = "单一限制模式"
                MultiConstraint_ModeBButton.Content = "标准多重限制模式"
                MultiConstraint_ModeCButton.Content = "设置速度损失为停机"
                prstoryCancelButton.Content = "取消"
                Multilinelaunchlabel.Content = "你要分析多行？ 点击这里."
        End Select
        initializeTooltipTextFromLanguage() 'set the incontrol settings text
        initializeVersionTooltipTextFromLanguage() 'set the version no tooltip text
    End Sub

    Private Sub initializeTooltipTextFromLanguage()
        Dim Pnl As New StackPanel
        Dim tbk As New TextBlock

        Select Case My.Settings.LanguageActive
            Case Lang.English

                'RULE 1
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Rule #1"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Any single data point falls outside the 3σ limit"
                Pnl.Children.Add(tbk)
                inControlRuleOneBox.ToolTip = Pnl
                'RULE 2
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Rule #2"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Two out of three consecutive points fall beyond the 2σ limit on the same side of the centerline"
                Pnl.Children.Add(tbk)
                inControlRuleTwoBox.ToolTip = Pnl
                'RULE 3
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Rule #3"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Four out of five consecutive points fall beyond the 1σ limit on the same side of centerline"
                Pnl.Children.Add(tbk)
                inControlRuleThreeBox.ToolTip = Pnl
                'RULE 4
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Rule #4"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Five points in a row are increasing"
                Pnl.Children.Add(tbk)
                inControlRuleFourBox.ToolTip = Pnl
                'RULE 5
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Rule #5"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Nine points fall on the same side of the centerline"
                Pnl.Children.Add(tbk)
                inControlRuleFiveBox.ToolTip = Pnl
                'RULE 6
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Rule #6"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Eight consecutive points exist with none within the 1σ limit"
                Pnl.Children.Add(tbk)
                inControlRuleSixBox.ToolTip = Pnl


                'default label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Default Rules"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Use the prstory default rules to detmine stability"
                Pnl.Children.Add(tbk)
                inControlSliderLabelOne.ToolTip = Pnl
                'custom label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Custom Rules"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Configure your own stability rules!"
                Pnl.Children.Add(tbk)
                inControlSliderLabelTwo.ToolTip = Pnl
                'we label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Western Electric Rules"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Industry standard set of rules developed at the Western Electric Company in the 1950s"
                Pnl.Children.Add(tbk)
                inControlSliderLabelThree.ToolTip = Pnl
                inControlSliderLabelThree_Copy.ToolTip = Pnl

            'get the rest of the labels brah
            '  inControlSliderLabelOne.ToolTip = "Select the default prstory stability rules"
            '  inControlSliderLabelTwo.ToolTip = "Pick your own stability rules"
            '  inControlSliderLabelThree.ToolTip = "Use the Western Electric Rules to determine phenomena stability"
            '  inControlSliderLabelThree_Copy.ToolTip = "Use the Western Electric Rules to determine phenomena stability"
            Case Lang.German
                'RULE 1
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regel #1"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "jeder einzelnen Datenpunkt außerhalb des 3σ Grenze"
                Pnl.Children.Add(tbk)
                inControlRuleOneBox.ToolTip = Pnl
                'RULE 2
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regel #2"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "zwei von drei aufeinanderfolgenden Punkten fallen außerhalb des 2σ Grenze auf der gleichen Seite der Mittellinie"
                Pnl.Children.Add(tbk)
                inControlRuleTwoBox.ToolTip = Pnl
                'RULE 3
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regel #3"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Vier von fünf aufeinanderfolgenden Punkten über die fallen 1σ Grenze auf der gleichen Seite der Mittellinie"
                Pnl.Children.Add(tbk)
                inControlRuleThreeBox.ToolTip = Pnl
                'RULE 4
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regel #4"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Fünf Punkte in Folge steigen"
                Pnl.Children.Add(tbk)
                inControlRuleFourBox.ToolTip = Pnl
                'RULE 5
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regel #5"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Neun Punkte fallen auf der gleichen Seite der Mittellinie"
                Pnl.Children.Add(tbk)
                inControlRuleFiveBox.ToolTip = Pnl
                'RULE 6
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regel #6"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Acht aufeinander folgenden Punkten bestehen keine im 1σ Grenze"
                Pnl.Children.Add(tbk)
                inControlRuleSixBox.ToolTip = Pnl


                'default label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Standardregeln"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "verwenden Sie die prstory Standardregeln, um die Stabilität zu bestimmen"
                Pnl.Children.Add(tbk)
                inControlSliderLabelOne.ToolTip = Pnl
                'custom label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "benutzerdefinierte Regeln"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "konfigurieren Sie Ihre eigenen Regelwerke"
                Pnl.Children.Add(tbk)
                inControlSliderLabelTwo.ToolTip = Pnl
                'we label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Western Electric-Regeln"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Industriestandard-Satz von Regeln an der westlichen Electric Company in den 1950er Jahren entwickelt,"
                Pnl.Children.Add(tbk)
                inControlSliderLabelThree.ToolTip = Pnl
                inControlSliderLabelThree_Copy.ToolTip = Pnl
            Case Lang.Spanish
                'RULE 1
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regla #1"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "cualquier punto de datos único está fuera del límite de 3σ"
                Pnl.Children.Add(tbk)
                inControlRuleOneBox.ToolTip = Pnl
                'RULE 2
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regla #2"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "dos de cada tres puntos consecutivos caen más allá del límite 2σ en el mismo lado de la línea central"
                Pnl.Children.Add(tbk)
                inControlRuleTwoBox.ToolTip = Pnl
                'RULE 3
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regla #3"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Cuatro de cinco puntos consecutivos caen más allá del límite 1σ en el mismo lado de la línea central"
                Pnl.Children.Add(tbk)
                inControlRuleThreeBox.ToolTip = Pnl
                'RULE 4
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regla #4"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Cinco puntos consecutivos están aumentando"
                Pnl.Children.Add(tbk)
                inControlRuleFourBox.ToolTip = Pnl
                'RULE 5
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regla #5"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Nueve puntos caen en el mismo lado de la línea central"
                Pnl.Children.Add(tbk)
                inControlRuleFiveBox.ToolTip = Pnl
                'RULE 6
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regla #6"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Existen ocho puntos consecutivos con ninguno en el plazo 1σ"
                Pnl.Children.Add(tbk)
                inControlRuleSixBox.ToolTip = Pnl


                'default label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Reglas predeterminadas"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "utilizar las reglas prstory por defecto para determinar la estabilidad"
                Pnl.Children.Add(tbk)
                inControlSliderLabelOne.ToolTip = Pnl
                'custom label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "reglas personalizadas"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "configurar sus propios conjuntos de reglas!"
                Pnl.Children.Add(tbk)
                inControlSliderLabelTwo.ToolTip = Pnl
                'we label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Reglas de Western Electric"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "conjunto estándar de reglas desarrollado en la Western Electric Company en la década de 1950"
                Pnl.Children.Add(tbk)
                inControlSliderLabelThree.ToolTip = Pnl
                inControlSliderLabelThree_Copy.ToolTip = Pnl
            Case Lang.French
                'RULE 1
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "règle #1"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "tout point de données unique est en dehors de la limite de 3σ"
                Pnl.Children.Add(tbk)
                inControlRuleOneBox.ToolTip = Pnl
                'RULE 2
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "règle #2"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "deux sur trois points consécutifs tombe au-delà de la limite de 2σ sur le même côté de l'axe"
                Pnl.Children.Add(tbk)
                inControlRuleTwoBox.ToolTip = Pnl
                'RULE 3
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "règle #3"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Quatre des cinq points consécutifs tomber au-delà de la limite de 1σ sur le même côté de l'axe"
                Pnl.Children.Add(tbk)
                inControlRuleThreeBox.ToolTip = Pnl
                'RULE 4
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "règle #4"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Cinq points dans une rangée sont en augmentation"
                Pnl.Children.Add(tbk)
                inControlRuleFourBox.ToolTip = Pnl
                'RULE 5
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "règle #5"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Neuf points se situent sur le même côté de l'axe"
                Pnl.Children.Add(tbk)
                inControlRuleFiveBox.ToolTip = Pnl
                'RULE 6
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "règle #6"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Huit points consécutifs existe avec aucun dans la limite de 1σ"
                Pnl.Children.Add(tbk)
                inControlRuleSixBox.ToolTip = Pnl


                'default label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Règles par défaut"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "utiliser les règles prstory par défaut pour déterminer la stabilité"
                Pnl.Children.Add(tbk)
                inControlSliderLabelOne.ToolTip = Pnl
                'custom label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "des règles personnalisées"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "configurer vos propres ensembles de règles de stabilité"
                Pnl.Children.Add(tbk)
                inControlSliderLabelTwo.ToolTip = Pnl
                'we label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Règles Western Electric"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "industrie ensemble standard de règles développées au Western Electric dans les années 1950"
                Pnl.Children.Add(tbk)
                inControlSliderLabelThree.ToolTip = Pnl
                inControlSliderLabelThree_Copy.ToolTip = Pnl
            Case Lang.Portuguese
                'RULE 1
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "regra #1"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "qualquer ponto único de dados está fora do limite 3σ"
                Pnl.Children.Add(tbk)
                inControlRuleOneBox.ToolTip = Pnl
                'RULE 2
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "regra #2"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "dois dos três pontos consecutivos cair para além do limite 2σ no mesmo lado da linha central"
                Pnl.Children.Add(tbk)
                inControlRuleTwoBox.ToolTip = Pnl
                'RULE 3
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "regra #3"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Quatro de cinco pontos consecutivos cair para além do limite 1σ no mesmo lado da linha central"
                Pnl.Children.Add(tbk)
                inControlRuleThreeBox.ToolTip = Pnl
                'RULE 4
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "regra #4"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Cinco pontos em uma fileira estão a aumentar"
                Pnl.Children.Add(tbk)
                inControlRuleFourBox.ToolTip = Pnl
                'RULE 5
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "regra #5"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Nove pontos caem no mesmo lado da linha central"
                Pnl.Children.Add(tbk)
                inControlRuleFiveBox.ToolTip = Pnl
                'RULE 6
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "regra #6"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "Oito pontos consecutivos existir com nenhum dentro do limite 1σ"
                Pnl.Children.Add(tbk)
                inControlRuleSixBox.ToolTip = Pnl


                'default label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Regras padrão"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "usar as regras padrão prstory"
                Pnl.Children.Add(tbk)
                inControlSliderLabelOne.ToolTip = Pnl
                'custom label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "regras personalizadas"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "fazer suas próprias regras!"
                Pnl.Children.Add(tbk)
                inControlSliderLabelTwo.ToolTip = Pnl
                'we label
                Pnl = New StackPanel
                tbk = New TextBlock
                tbk.FontWeight = FontWeights.Bold
                tbk.Text = "Western Electric Regras"
                Pnl.Children.Add(tbk)
                tbk = New TextBlock
                tbk.Text = "conjunto de regras padrão da indústria desenvolvido no Western Electric Company na década de 1950"
                Pnl.Children.Add(tbk)
                inControlSliderLabelThree.ToolTip = Pnl
                inControlSliderLabelThree_Copy.ToolTip = Pnl
        End Select



    End Sub
    Private Sub initializeVersionTooltipTextFromLanguage()
        Dim Pnl As New StackPanel
        Dim tbk As New TextBlock

        Dim headerString As String = "", BugA As String = "", BugB As String = "", BugC As String = ""
        Dim subtitleString As String = ""

        BugA = "- Some new features are taking shape. Look out for an update about prstory on your email in a few weeks!" & vbCrLf '& "Huangpu" & vbCrLf
        BugB = "- All family care converting lines are on prstory now!!! Thanks Amanda Gonzalez! Many more baby care lines too! Thanks Aditi!" & vbCrLf '& "[Paretos in Raw Data]" & vbCrLf
        BugC = "- Several minor bug fixes and improvements." & vbCrLf & "Thanks to everyone who provided feedback to help make prstory better!" & vbCrLf

        Select Case My.Settings.LanguageActive
            Case Lang.English
                headerString = "Whats new with prstory?"
                subtitleString = "Last Update: 5th February 2016" & vbCrLf

            Case Lang.German
                headerString = "was ist neu mit prstory?"
                subtitleString = "Letzte Aktualisierung:  2015"
                BugA = "- All new inControl configuration in Advanced Options." & vbCrLf & "Customize your control rules to your operation" & vbCrLf
            Case Lang.Spanish
                headerString = "cuál es nuevo con prstory?"
                subtitleString = "Última actualización:  2015"
            '   BugA = "- All new inControl configuration in Advanced Options." & vbCrLf & "Customize your control rules to your operation" & vbCrLf
            Case Lang.French
                headerString = "quoi de neuf avec prstory?"
                subtitleString = "Dernière mise à jour  2015"
            ' BugA = "- All new inControl configuration in Advanced Options." & vbCrLf & "Customize your control rules to your operation" & vbCrLf
            Case Lang.Portuguese
                headerString = "o que é novo com prstory?"
                subtitleString = "Última atualização:  2015"
                ' BugA = "- All new inControl configuration in Advanced Options." & vbCrLf & "Customize your control rules to your operation" & vbCrLf
        End Select
        tbk.FontWeight = FontWeights.Bold
        tbk.Text = headerString
        Pnl.Children.Add(tbk)
        tbk = New TextBlock
        tbk.FontStyle = FontStyles.Italic
        tbk.Text = subtitleString
        Pnl.Children.Add(tbk)
        tbk = New TextBlock
        tbk.Text = BugA
        Pnl.Children.Add(tbk)
        tbk = New TextBlock
        tbk.Text = BugB
        Pnl.Children.Add(tbk)
        tbk = New TextBlock
        tbk.Text = BugC
        Pnl.Children.Add(tbk)
        ''VersionLabel.ToolTip = Pnl

    End Sub
#End Region
    Sub ReInitializeAllPublicVariables()
        IsRemappingDoneOnce = False

        bargraphreportwindow_Open = False

        AllProdLines(selectedindexofLine_temp).isFilterByBrandcode = False

        datalabelcontent = ""

        motionchartsource = 1
        bubblenumberpublic = 1

        MasterDataSet = Nothing
        shouldSnakeClose = False
        PROF_connectionError = False
        UseDemoData = False
    End Sub
    Private Sub HideLineDefaultQuery()
        LineDefaultQueryLabel.Visibility = Visibility.Hidden
        LineDefaultYesButton.Visibility = Visibility.Hidden
        LineDefaultNoButton.Visibility = Visibility.Hidden
        LineDefaultCancelButton.Visibility = Visibility.Hidden
    End Sub
    Private Sub ShowLineDefaultQuery()
        LineDefaultQueryLabel.Visibility = Visibility.Visible
        LineDefaultYesButton.Visibility = Visibility.Visible
        LineDefaultNoButton.Visibility = Visibility.Visible
        LineDefaultCancelButton.Visibility = Visibility.Visible
        LineSelection_Alert.Visibility = Visibility.Hidden
    End Sub
    Private Sub dateSelectionShortcut()
        launchprstorysettings()
        launchprstorydaterange()
        prstorymtdoption_MouseDown()
        settingsDONE()
    End Sub
#Region "Saving / Importing Default Settings"
    Private Sub LineDefaultYesClicked()
        My.Settings.AdvancedSettings_UTcutoff = LineSelection_AutoComplete.SelectedItem

        MessageBox.Show("Default Line Set!")

        If False Then
            If prstory_linedropdown.SelectedIndex > 0 Then

                My.Settings.DefaultLineIndex = prstory_linedropdown.SelectedIndex
                My.Settings.DefaultSiteAcronym = AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).parentSite.ThreeLetterID
                My.Settings.DefaultLineName = AllProdLines(activeLineIndeces(prstory_linedropdown.SelectedIndex - 1)).Name
                My.Settings.WantstoSetDefaultLine = False
                If LineDefaultQueryLabel.Visibility = Windows.Visibility.Hidden Then GreenCheck.Visibility = Visibility.Visible
                My.Settings.Save()
                HideLineDefaultQuery()
            Else
                Select Case My.Settings.LanguageActive
                    Case Lang.English
                        MsgBox("No line selected")
                    Case Lang.Chinese_Simplified
                        MsgBox("选择生产线")
                    Case Else
                        MsgBox("No line selected")
                End Select
            End If
        End If
    End Sub

    Private Sub HideDateSelectionAlert()
        DateSelection_Alert.Visibility = Visibility.Hidden
    End Sub
#End Region

    Public Sub maincanvasmousedown(sender As Object, e As MouseButtonEventArgs)
        SiteMenu.Visibility = Visibility.Hidden
        SectorMenu.Visibility = Visibility.Hidden
    End Sub

    Public Sub ExportAllLineData(sender As Object, e As MouseButtonEventArgs)

        Dim appPath As String
        Dim dialog As New System.Windows.Forms.FolderBrowserDialog()

        Dim fsT As Object
        Dim fileName As String
        dialog.RootFolder = Environment.SpecialFolder.Desktop 'lg
        dialog.SelectedPath = "C:\"
        dialog.Description = "Select Path To Save prStory Raw Data Files"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            appPath = dialog.SelectedPath
            fileName = appPath & "\" & "LineDataExport" & Now.Second & ".csv"

            fsT = CreateObject("ADODB.Stream")
            fsT.Type = 2 'Specify stream type - we want To save text/string data.
            fsT.Charset = "utf-8" 'Specify charset For the source text data.
            fsT.Open() 'Open the stream And write binary data To the object


            fsT.WriteText("Site Name" & "," & "Module Name" & "," & "Colloq. Name" & "," & "DT Prod Unit" & "," & "Production Prod Unit" & "," & "Rate Loss/Secondary Unit" & "," & "Server Type" & "," & "Stored Proc For Downtime Data,Server Name, Database" & "," & "Start Time (Hrs), # Shifts [length (hrs)]")
            fsT.WriteText(vbCrLf)
            'actually export the data



            For i As Integer = 0 To AllProdLines.Count - 1
                '  If AllProductionLines(i).parentmodule.parentsector.name = sector_family Then

                Dim serverType As String = ""
                Dim storedProc As String = ""

                Select Case AllProdLines(i).SQLdowntimeProcedure
                    Case DefaultProficyDowntimeProcedure.QuickQuery
                        storedProc = "QuickQuery"
                    Case DefaultProficyDowntimeProcedure.OneClick
                        storedProc = "OneClick v26"
                    Case DefaultProficyDowntimeProcedure.Maple
                        storedProc = "OneClick Maple"
                    Case DefaultProficyDowntimeProcedure.GLEDS
                        storedProc = "GLEDS"
                    Case Else
                        storedProc = "Custom SQL"
                End Select

                Select Case AllProdLines(i).ProficyServer_Username
                    Case "comxclient"
                        serverType = "Proficy v5"
                    Case "PRStory"
                        serverType = "Proficy v6"
                    Case "One_Click"
                        serverType = "Maple"
                    Case Else
                        serverType = "Custom Configuration"
                End Select

                fsT.WriteText(AllProdLines(i).SiteName & "," & AllProdLines(i).parentModule.Name & "," & AllProdLines(i).Name & "," & AllProdLines(i).mainProdUnit & "," & AllProdLines(i).mainProfProd & "," & AllProdLines(i)._rateLossDisplay & "," & serverType & "," & storedProc & "," & AllProdLines(i).ProficyServer_Name & "," & AllProdLines(i).ServerDatabase & "," & AllProdLines(i)._DayStartTimeHrs & "," & AllProdLines(i).NumberOfShifts & "[" & 24 / Math.Max(1, AllProdLines(i).NumberOfShifts) & "]")

                fsT.WriteText(vbCrLf)
                ' End If
            Next

            'fin
            Try
                fsT.SaveToFile(fileName, 2) 'Save binary data To disk
            Catch ex As Exception
                MsgBox("A file with same name is open. Please close that file and export again.")
                'fsT.SavetoFile(fileName & "_2", 2)
                fsT = Nothing
                Exit Sub

            End Try

            fsT = Nothing
            'show the folder
            Process.Start(appPath)
        End If
    End Sub



#Region "Multiline"

    Private Sub LaunchMultiLineWindow()


        Dim multilinewindow As New Window_Multiline
        multilinewindow.InitializeMultilineGroups(multilineGroups)
        multilinewindow.Owner = Me
        Me.Visibility = Visibility.Hidden
        multilinewindow.Show()
    End Sub
    Private Sub MultilineMouseMove(sender As Object, e As MouseEventArgs)
        sender.background = mybrushdarkgray
        sender.foreground = mybrushlanguagewhite
        sender.opacity = 0.9
    End Sub

    Private Sub MultilineMouseLeave(sender As Object, e As MouseEventArgs)

        sender.foreground = mybrushdefaultfontgray
        sender.background = mybrushdefaultbackgroundgray
        sender.opacity = 1.0
    End Sub

    Private Sub NewsFlashLabel_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Process.Start("http://prstory.pg.com")
    End Sub

    Private Sub Page1GoButton_MouseDown(sender As Object, e As MouseButtonEventArgs)
        Page1GoButton.Content = "< Back To First Page"
        AdvancedSettings.Visibility = Visibility.Visible
        AdvancedSettings2.Visibility = Visibility.Hidden
    End Sub

    Private Sub Page2GoButton_MouseDown(sender As Object, e As MouseButtonEventArgs)
        AdvancedSettings.Visibility = Visibility.Hidden
        AdvancedSettings2.Visibility = Visibility.Visible
    End Sub

    Private Sub QueryDaysBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        datapull_duration = QueryDaysBox.SelectedItem
    End Sub

#End Region

End Class

