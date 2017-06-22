Imports System.IO
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports System.Security.Cryptography

Module ProductionLevels
    Public Const LINE_USING_MOT As Boolean = True

    Public AllProdLines As New List(Of ProdLine)()
    Public AllProdModules As New List(Of ProdModule)()
    Public AllProductionSites As New List(Of ProdSite)()
    Public AllProductionSectors As New List(Of BusinessUnit)()

    Public Sub JSON_Export(exportObject As List(Of ProdLine), Optional FileNameX As String = "C:\\Users\\odle.so.1\\Desktop\\prstoryLinesJSON", Optional FileType As String = ".txt")

        Dim jsonData As String = JsonConvert.SerializeObject(exportObject) '0))
        Dim FileName As String = FileNameX & FileType
        Dim fcreate As FileStream = File.Open(FileName, FileMode.Create)
        Dim writer As StreamWriter = New StreamWriter(fcreate)

        writer.Write(jsonData)
        writer.Close()
    End Sub

    Public Function getBUIndexFromName(sectorName As String) As Integer
        Dim i As Integer ', tempSite As productionSite
        For i = 0 To AllProductionSectors.Count - 1
            If AllProductionSectors(i).Name = sectorName Then Return i
        Next
        Return -1
    End Function

    Public Function getSiteIndexFromName(siteName As String) As Integer
        Dim i As Integer ', tempSite As productionSite
        For i = 0 To AllProductionSites.Count - 1
            If AllProductionSites(i).Name = siteName Then Return i
        Next
        Return -1
    End Function

    Public Function getModuleIndexFromID(moduleName As Guid) As Integer
        Dim i As Integer ', tempSite As productionSite
        For i = 0 To AllProdModules.Count - 1
            If AllProdModules(i).ID = moduleName Then Return i
        Next
        Return -1
    End Function

End Module

Module AllKnownSitesModules
    '#Region "Constants"
    'sectors
    Public Const SECTOR_HEALTH As String = "One Health"
    Public Const SECTOR_HOME As String = "Home Care"
    Public Const SECTOR_BEAUTY As String = "Beauty Care"
    Public Const SECTOR_FAMILY As String = "Family Care"
    Public Const SECTOR_BABY As String = "Baby Care"
    Public Const SECTOR_FEM As String = "Fem Care"
    Public Const SECTOR_FHC As String = "F&HC"
    Public Const SECTOR_APRILFOOLS As String = "April 1"
    Public Const BETA_TESTING As String = "Beta Lines"

    'sites
    Public Const SITE_BROWNS_SUMMIT As String = "GBO" 'BS
    Public Const SITE_SWING_ROAD As String = "Swing Road" 'SR
    Public Const SITE_IOWA_CITY As String = "IC"
    Public Const SITE_IOWA_CITY_ORALCARE As String = "ICOC"
    Public Const SITE_GROSS_GERAU As String = "GRO"
    Public Const SITE_HUANGPU As String = "Huangpu"
    Public Const SITE_NAUCALPAN As String = "NAUP"
    Public Const SITE_XIQING As String = "Xiqing"
    Public Const SITE_CRUX As String = "Crux"
    Public Const SITE_MANDIDEEP = "MDP"
    Public Const SITE_PHOENIX = "PHX"
    Public Const SITE_BORYSPIL = "Boryspil"
    Public Const SITE_ALEXANDRIA = "Alex"
    Public Const SITE_LIMA = "Lima"
    Public Const SITE_BELLEVILLE = "Belleville"
    Public Const SITE_MARISCALA = "Mariscala"
    Public Const SITE_AKASHI = "Akashi"
    Public Const SITE_ALBANY = "Albany"
    Public Const SITE_HYDERABAD = "Hyderabad"
    Public Const SITE_BUDAPEST = "Budapest"
    Public Const SITE_OXNARD = "Oxnard"
    Public Const SITE_MANCHESTER = "Manchester"
    Public Const SITE_TEPEJI = "Tepeji"
    Public Const SITE_LOUVERIA = "Louveira"
    Public Const SITE_CAPEGIRARDEAU = "Cape Girardeau"
    Public Const SITE_MEHOOPANY = "Mehoopany"
    Public Const SITE_JIJONA = "Jijona"
    Public Const SITE_OCTOBER6 = "October 6"
    Public Const SITE_LAGOS = "Lagos"
    Public Const SITE_GREENBAY = "Greenbay"
    Public Const SITE_BOXELDER = "Box Elder"
    Public Const SITE_BENCAT = "Ben Cat"
    Public Const SITE_GUATIRE = "Guatire"
    Public Const SITE_MATERIALES = "Materiales"
    Public Const SITE_GYONGYOS = "Gyongyos"
    Public Const SITE_NOVO = "Novo"
    Public Const SITE_GEBZE = "Gebze"
    Public Const SITE_JEDDAH = "Jeddah"
    Public Const SITE_TAICANG = "Taicang"
    Public Const SITE_BANGKOK = "Bangkok"
    Public Const SITE_SINGAPOREPIONEER = "Singapore"
    Public Const SITE_VILLAMERCEDES = "Villa Mercedes"
    Public Const SITE_SANTIAGO = "Santiago"

    Public Const SITE_RIO = "RIO"

    Public Const SITE_APRILFOOLS = " "

    'Browns Summit Modules
    Public Const BS_APDO As String = "APDO"
    Public Const BS_OC As String = "Oral Care"
    Public Const BS_SC As String = "Skin Care"
End Module

Module initializeSites
    Sub initializeAllSites()
        AllProductionSectors.Add(New BusinessUnit(SECTOR_HEALTH))
        AllProductionSectors.Add(New BusinessUnit(SECTOR_BABY))
        AllProductionSectors.Add(New BusinessUnit(SECTOR_BEAUTY))
        AllProductionSectors.Add(New BusinessUnit(SECTOR_FEM))
        AllProductionSectors.Add(New BusinessUnit(SECTOR_FHC))
        AllProductionSectors.Add(New BusinessUnit(SECTOR_FAMILY))
        AllProductionSectors.Add(New BusinessUnit(SECTOR_HOME))
        AllProductionSectors.Add(New BusinessUnit("Hair Care"))
        'If isAPRILFOOLS Then AllProductionSectors.Add(New BusinessUnit(SECTOR_APRILFOOLS))
        AllProductionSectors.Add(New BusinessUnit(BETA_TESTING))

        '   initializeMultiUnitTestLines()
        Dim y As String = "01000000D08C9DDF0115D1118C7A00C04FC297EB0100000072B8CC6CC3EE294C8E8148C25BCC19100000000002000000000003660000C000000010000000EBA27DBAAE8C7455CED02C980BC68A4A0000000004800000A0000000100000001018D102CF6DECE1545287674D279BD010000000A9FF538A7799784155E43A8B44512670140000002443E5755126B53C80F26A781CF2F97988F09BFF"
        'If isAPRILFOOLS Then initialize_APRIL_FOOLS()

        'site w/ one health
        initializeSite_BrownsSummit(y)

        initializeSite_GrossGerau()
        initializeSite_HuangPu()
        initializeSite_Naucalpan()
        initializeSite_Dammam()
        initializeSite_Montornes()
        'initializeSite_Xiqing()
        initializeSite_SwingRoad()
        initializeSite_Crux()
        initializeSite_PHCPhoenix()

        initializeSite_IowaCity()
        initializeSite_IowaCity_OC()
        initializeSite_Mandideep()
        initializeSite_Baddi()
        initializeSite_Borysil()
        initializeSite_Belleville()
        initializeSite_Belleville_Test()

        initializeSite_Budapest()

        initializeSite_Mariscala()

        initializeSite_Albany()

        initializeSite_Oxnard()
        initializeSite_BabyGLEDS()

        initializeSite_Hyderabad()
        initializeSite_Manchester()
        initializeSite_Tepeji()
        initializeSite_HuangpuBaby()
        initializeSite_Louveira()

        initializeSite_Jijona()
        initializeSites_MoreBabyCare()
        initializeSite_Akashi()
        'initializeSite_Lima()
        initializeSite_October6()
        initializeSite_Lagos()
        initializeSite_Rakona()
        initializeSite_Amiens()
        initializeSite_Greenbay()
        initializeSite_BoxElder()
        initializeSite_CapeGirardeau()
        initializeSite_Mehoopany()
        initializeSite_BenCat()
        initializeSite_Guatire()
        initializeSite_MATERIALES()
        initializeSite_GYONGYOS()
        initializeSite_NOVO()
        initializeSite_GEBZE()
        initializeSite_Ibadan()
        initializeSite_JEDDAH()
        initializeSite_Taicang()
        initializeSite_Bangkok()
        initializeSite_Xiqing()
        initializeSite_Singapore()
        initializeSite_Santiago()
        initializeSite_VillaMercedes()
        initializeSite_Xiquing()
        initializeSite_Rio()
        initializeSite_JKT()
        initializeSite_Jakarta()
        initializeSite_Reading()
        initializeSite_Ordzho()

        initializeTEMP()
        initializeSite_STL()
        InitializeSite_KansasCity()
        InitializeSite_BinhDuong()
        initializeSite_XQ()
        initializeSite_Cairo()

        initializeSite_Auburn()
        initializeSite_Crailshiem()
        initializeSite_Hub()

        initializeSite_Dover()
        initializeSite_Mequinenza()
        initializeSite_Alexandria()
        initializeSite_Euskirchen()

        Dim sitename As String = "Demo"
        AllProductionSites.Add(New ProdSite(sitename, "prstory.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "PRS"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Line", sitename, SECTOR_FAMILY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String
        lineName = "I"
        AddLine(lineName, ModuleID, lineName & " Converter")
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True




        AllProductionSites.Sort() 'this puts the site names in alphabetical order for the dropdown menu

        SERVER_PW_V6 = "comxclient"

    End Sub

#Region "April Fools"
    ''  Private Sub initialize_APRIL_FOOLS()
    'Dim MODULE_APRILFOOLS = " "
    '     AllProductionSites.Add(New productionSite(SITE_APRILFOOLS, "", "", "", "", "APR"))
    ' Dim ModuleID As Guid = Guid.NewGuid()
    ''     AllProductionModules.Add(New productionModule(ModuleID, MODULE_APRILFOOLS, SITE_APRILFOOLS, SECTOR_APRILFOOLS, prStoryMapping.APRILFOOLS, DefaultProficyDowntimeProcedure.QuickQuery_MOT, DefaultProficyProductionProcedure.SwingRoad, "", "", "", "", DowntimeField.Reason1, DowntimeField.Reason2, DTsched_Mapping.APRILFOOLS))
    'initialize the lines
    '
    '   AllProductionLines.Add(New productionLine("X-Wing Assembly Bay 1", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    '  AllProductionLines.Add(New productionLine("Death Star Reactor Room", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("LexCorp Converting 17C", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    'AllProductionLines.Add(New productionLine("Stark Industries Line 7", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("Stark Industries Line 9", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("Area 51", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("Buy n Large Batteries", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("SPECTRE Line 007", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    'AllProductionLines.Add(New productionLine("Acme Corp Anvil Assembly", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    'AllProductionLines.Add(New productionLine("Wonka Gobstoppers 13", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("Platform 9 3/4", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("Wayne Enterprises BioTech", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' AllProductionLines.Add(New productionLine("Globex Corp Windows", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    ' ' AllProductionLines.Add(New productionLine("Nova Corps", "3", SITE_APRILFOOLS, MODULE_APRILFOOLS, 2, 12, 6, "", ""))
    '
    '    End Sub

#End Region

    Private Sub initializeSite_Reading()
        Dim siteName As String = "Reading"
        AllProductionSites.Add(New ProdSite(siteName, "RGM-MESDATABE.NA.PG.COM", "", SERVER_PW_V6, SERVER_UN_V6, "JKT"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Grooming", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHTPLANNEDPlusOne, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        'initialize the lines
        Dim lineName As String

        lineName = "RMC L1"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6.5, lineName & " Main", lineName))
        lineName = "RMC L2"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6.5, lineName & " Main", lineName))

    End Sub


    'ord-mesdtahc
    Private Sub initializeSite_Ordzho()
        Dim siteName As String = "Ordzho"
        AllProductionSites.Add(New ProdSite(siteName, "ord-mesdtahc", "", SERVER_PW_V6, SERVER_UN_V6, "ORD"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC", siteName, SECTOR_FHC, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        'initialize the lines
        AllProdLines.Add(New ProdLine("LP01", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD LP01 Main", ""))
        AllProdLines.Add(New ProdLine("LP02", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD LP02 Main", ""))
        AllProdLines.Add(New ProdLine("SOAP02", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD SOAP02 Main", ""))
        AllProdLines.Add(New ProdLine("PSG01", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD PSG01 Main", ""))
        AllProdLines.Add(New ProdLine("PSG02", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD PSG02 Main", ""))
        AllProdLines.Add(New ProdLine("PSG03", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD PSG03 Main", ""))
        AllProdLines.Add(New ProdLine("PSG04", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD PSG04 Main", ""))
        AllProdLines.Add(New ProdLine("PSG05", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "ORD PSG05 Main", ""))
        AllProdLines.Add(New ProdLine("PSG06", "", siteName, ModuleID, 2, 12, 6, "ORD PSG06 Filler1", ""))

    End Sub



    Private Sub initializeSite_Jakarta()
        Dim lineName As String, collName As String
        Dim siteName As String = "Jakarta"
        AllProductionSites.Add(New ProdSite(siteName, "jkt-mesdatabc", "", SERVER_PW_V6, SERVER_UN_V6, "JAK"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC", siteName, SECTOR_FHC, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        'initialize the lines
        AllProdLines.Add(New ProdLine("LLJK 111", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "LLJK 111 Main", "LLJK 111"))
        AllProdLines.Add(New ProdLine("LLJK 112", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "LLJK 112 Main", "LLJK 112"))
        AllProdLines.Add(New ProdLine("LLJK 123", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "LLJK 123 Mespack3", "LLJK 123"))
        ' allproductionlines(allproductionlines.count - 1).isstartupmode = true

        lineName = "LLJK-012"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Main", lineName))
        lineName = "LLJK-013"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Main", lineName))
        lineName = "LLJK-014"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Main", lineName))

        lineName = "PLJK 101"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Main", lineName))
        lineName = "PLJK 102"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Main", lineName))
        lineName = "PLJK 103"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Main", lineName))
        lineName = "PLJK 104"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Main", lineName))


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Beauty", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHTPLANNEDPlusTwo, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason2, -1, DTsched_Mapping.Greensboro, 2, 12, 6))

        lineName = "JKTHCSL1"
        collName = "Arjunior Filler 1"
        AddLine(collName, ModuleID2, lineName & " " & collName, lineName)

        lineName = "JKTHCSL1"
        collName = "Arjunior Filler 2"
       AddLine(collName, ModuleID2, lineName & " " & collName, lineName)

        lineName = "JKTHCSL2"
        collName = "Arjunior Filler 3"
        AddLine(collName, ModuleID2, lineName & " " & collName, lineName)

        lineName = "JKTHCSL2"
        collName = "Arjunior Filler 4"
        AddLine(collName, ModuleID2, lineName & " " & collName, lineName)

        lineName = "JKTHCSL2"
        collName = "Arjunior Filler 5"
        AddLine("MC5", ModuleID2, lineName & " " & collName, lineName)

        lineName = "JKTHCBL1"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 6, lineName & " Main", lineName))

    End Sub


    Private Sub initializeSite_JKT()
        Dim siteName As String = "JKT"
        AllProductionSites.Add(New ProdSite(siteName, "jkt-mesdatabc", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "JKT"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "HC", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        'initialize the lines
        AllProdLines.Add(New ProdLine("Bottle Line 1", "JKTHCBL1", siteName, ModuleID, 2, 12, 6, "JKTHCBL1 Main", "JKTHCBL1"))
        ' allproductionlines(allproductionlines.count - 1).isstartupmode = true

    End Sub

    Private Sub initializeSite_Rio()
        AllProductionSites.Add(New ProdSite(SITE_RIO, "BRIO-MESDATABE", "Hair Care", SERVER_PW_V6, SERVER_UN_V6, "RIO"))
        Dim ModuleID As Guid = Guid.NewGuid()
        '  AllProductionModules.Add(New productionModule(ModuleID, "CPRJ", SITE_RIO, SECTOR_BABY, prStoryMapping.GENERIC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        'initialize the lines
        '  AllProductionLines.Add(New productionLine("10", "10", SITE_RIO, ModuleID, 2, 12, 6, "CPRJ010 Main", "CPRJ010"))



        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "HCRJ", SITE_RIO, "Hair Care", prStoryMapping.Rio, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"3M", "CERMEX",
            "NGS", "Bundles", "Frascos", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"3M", "CERMEX",
            "NGS", "Bundles", "Frascos", OTHERS_STRING}

        'initialize the lines
        AllProdLines.Add(New ProdLine("001", "10", SITE_RIO, ModuleID2, 2, 12, 6, "HCRJ001 Main", "HCRJ001"))
        AllProdLines.Add(New ProdLine("003", "10", SITE_RIO, ModuleID2, 2, 12, 6, "HCRJ003 Main", "HCRJ003"))
        AllProdLines.Add(New ProdLine("004", "10", SITE_RIO, ModuleID2, 2, 12, 6, "HCRJ004 Main", "HCRJ004"))
        AllProdLines.Add(New ProdLine("005", "10", SITE_RIO, ModuleID2, 2, 12, 6, "HCRJ005 Main Encartuchadeira", "HCRJ005"))

    End Sub


#Region "PHC"
    Private Sub initializeSite_SwingRoad()
        AllProductionSites.Add(New ProdSite(SITE_SWING_ROAD, "gsr-mesdtahhc", PROFICY_SERVER_PASSWORD_SWINGROAD, PROFICY_SERVER_USERNAME_SWINGROAD, "SWR"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "PHC", SITE_SWING_ROAD, SECTOR_HEALTH, prStoryMapping.SwingRoad, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Reason2, DTsched_Mapping.SwingRoad))
        'initialize the lines
        AllProdLines.Add(New ProdLine("3", "3", SITE_SWING_ROAD, ModuleID, 2, 12, 6, "Line 3 Filler Reliability", "OTC Line 3"))
        AllProdLines.Add(New ProdLine("4", "4", SITE_SWING_ROAD, ModuleID, 2, 12, 6, "Line 4 Filler Reliability", "OTC Line 4"))
        AllProdLines.Add(New ProdLine("5", "5", SITE_SWING_ROAD, ModuleID, 2, 12, 6, "Line 5 Filler Reliability", "OTC Line 5"))
        AllProdLines.Add(New ProdLine("6", "6", SITE_SWING_ROAD, ModuleID, 2, 12, 6, "Line 6 Thermoformer Reliability", "OTC Line 6", prStoryMapping.SwingRoad_6))
        AllProdLines.Add(New ProdLine("7", "7", SITE_SWING_ROAD, ModuleID, 2, 12, 6, "Line 7 Blisterformer Reliability", "OTC Line 7", prStoryMapping.SwingRoad_7))

    End Sub

    Private Sub initializeSite_PHCPhoenix()
        AllProductionSites.Add(New ProdSite(SITE_PHOENIX, "PHX-MESDTAHHC", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "PHX"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "One Health", SITE_PHOENIX, SECTOR_HEALTH, prStoryMapping.Phoenix, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Phenoix))
        'initialize the lines
        AllProdLines.Add(New ProdLine("Line A", "Line A", SITE_PHOENIX, ModuleID, 2, 12, 6.5, "PHX Pack Line A Filler Capper", "PHX Pack Line A"))

        '   AllProductionModules.Add(New productionModule("One HealthX", SITE_PHOENIX, SECTOR_HEALTH, prStoryMapping.Phoenix, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Reason3, DTsched_Mapping.Phenoix))
        AllProdLines.Add(New ProdLine("Line D", "Line D", SITE_PHOENIX, ModuleID, 2, 12, 6.5, "PHX Pack Line D Packet Filler", "PHX Pack Line D", prStoryMapping.Pheonix_D))
        AllProdLines.Add(New ProdLine("Line G", "Line G", SITE_PHOENIX, ModuleID, 2, 12, 6.5, "PHX Pack Line G Filler", "PHX Pack Line G", prStoryMapping.Pheonix_D))

    End Sub
#End Region
#Region "Oral Care"
    Private Sub initializeSite_GrossGerau()
        AllProductionSites.Add(New ProdSite(SITE_GROSS_GERAU, "gge-mesdatahc", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "GRO"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "OC", SITE_GROSS_GERAU, SECTOR_HEALTH, prStoryMapping.OralCareGross, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Reason2, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("0", "0", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 0 Füller", "GRO Line 0"))
        AllProdLines.Add(New ProdLine("1", "1", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 1 Füller", "GRO Line 1"))
        AllProdLines.Add(New ProdLine("2", "2", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 2 Filler", "GRO Line 2"))
        AllProdLines.Add(New ProdLine("3", "3", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 3 Filler", "GRO Line 3"))
        AllProdLines.Add(New ProdLine("4", "4", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 4 Filler", "GRO Line 4"))
        AllProdLines.Add(New ProdLine("5", "5", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 5 Filler", "GRO Line 5"))
        AllProdLines.Add(New ProdLine("6", "6", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 6 Füller", "GRO Line 6"))
        AllProdLines.Add(New ProdLine("7", "7", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 7 Füller", "GRO Line 7"))
        AllProdLines.Add(New ProdLine("8", "8", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 8 Filler", "GRO Line 8"))
        AllProdLines.Add(New ProdLine("9", "9", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line 9 Filler", "GRO Line 9"))

        AllProdLines.Add(New ProdLine("EOL", "", SITE_GROSS_GERAU, ModuleID, 2, 12, 6, "Line EOL Wickler", "GRO Line EOL"))

        Dim ModuleIDx As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleIDx, "OC", SITE_GROSS_GERAU, SECTOR_HEALTH, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Reason2, DTsched_Mapping.Greensboro))

        Dim lineName As String
        lineName = "Line TM1"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM1 Main", "GRO Line TM1"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True


        lineName = "Line TM2"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM2 Main", "GRO " & lineName))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "Line TM3"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM3 Main", "GRO " & lineName))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "Line TM4"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM4 Main", "GRO " & lineName))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "Line TM5"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM5 Main", "GRO " & lineName))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True


        lineName = "Line TM6"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM6 Main", "GRO " & lineName))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True


        lineName = "Line TM7"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM7 Main", "GRO " & lineName))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "Line TM8"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TM8 Main", "GRO Line TM8"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "Line TMEOL"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GROSS_GERAU, ModuleIDx, 2, 12, 6, "Line TMEOL Main", "GRO Line TMEOL"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

    End Sub

    Private Sub initializeSite_Crux()
        AllProductionSites.Add(New ProdSite(SITE_CRUX, "bcru-mesdatabc", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "CRU"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Oral Care", SITE_CRUX, SECTOR_HEALTH, prStoryMapping.OralCareCrux, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("2", SITE_CRUX, ModuleID, 2, 12, 6, "Linha B Cartoner", "Linha B", -1, False, "Linha 2 Rate Loss", RateLossMode.Separate))
        AllProdLines.Add(New ProdLine("3", SITE_CRUX, ModuleID, 2, 12, 6, "Linha 3 Cartoner", "Linha C", -1, False, "Linha 3 Rate Loss", RateLossMode.Separate))
    End Sub

    Private Sub initializeSite_HuangPu2()
        Dim siteName As String = SITE_HUANGPU
        AllProductionSites.Add(New ProdSite(SITE_HUANGPU, "", "", SERVER_PW_V6, SERVER_UN_V6, "HPU"))
        ' AllProductionModules.Add(New productionModule("Oral Care", SITE_HUANGPU, SECTOR_HEALTH, prStoryMapping.NoMappingAvailable, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Hair Care", SITE_HUANGPU, SECTOR_BEAUTY, prStoryMapping.HuangPu, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.HuangpuHC))

        AllProdLines.Add(New ProdLine("B", "B", SITE_HUANGPU, ModuleID, 2, 12, 8, "New Line B Main", "New Line B"))
        AllProdLines.Add(New ProdLine("C", "C", SITE_HUANGPU, ModuleID, 2, 12, 8, "New LineC Main", "New LineC"))
        AllProdLines.Add(New ProdLine("F", "F", SITE_HUANGPU, ModuleID, 2, 12, 8, "HC LineF Main", "HC LineF"))
        AllProdLines.Add(New ProdLine("G", "G", SITE_HUANGPU, ModuleID, 2, 12, 8, "HC LineG Main", "HC LineG"))
        AllProdLines.Add(New ProdLine("K", "K", SITE_HUANGPU, ModuleID, 2, 12, 8, "HC LineK Main", "HC LineK"))
        AllProdLines.Add(New ProdLine("R", "R", SITE_HUANGPU, ModuleID, 2, 12, 8, "HC Line R Main", "HC LineR"))
        AllProdLines.Add(New ProdLine("A", "A", SITE_HUANGPU, ModuleID, 2, 12, 8, "SC LineA Main", "SC LineA"))
        ' AllProductionLines.Add(New productionLine("B", "B", SITE_HUANGPU, ModuleID, 2, 12, 8, "SC LineB Main", "SC LineB"))
        ' AllProductionLines.Add(New productionLine("C", "C", SITE_HUANGPU, ModuleID, 2, 12, 8, "SC LineC Main", "SC LineC"))


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Oral Care", siteName, SECTOR_HEALTH, prStoryMapping.OralCare, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("Line D", "Line D", SITE_HUANGPU, ModuleID2, 2, 12, 8, "OC LineD Main", "OC LineD"))

        '  Dim ModuleID2 As Guid = Guid.NewGuid()
        '  AllProductionModules.Add(New productionModule(ModuleID2, "Fem Care", siteName, SECTOR_FEM, prStoryMapping.HuangPu, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.HuangpuHC))
        '  Dim lineName As String

        '        lineName = "QAHU001"
        '        AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, "New Line B Main", "New Line B"))

    End Sub

    Private Sub initializeSite_Xiqing()
        Dim siteName As String = SITE_XIQING
        AllProductionSites.Add(New ProdSite(siteName, "xq-mesdatabc.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "XIQ"))
        'AllProductionSites.Add(New productionSite(siteName, "Prct-mesdatabc", proficy_SERVER_PASSWORD_V6, PROFICY_SERVER_USERNAME_V6, "XIQ"))
        '  AllProductionSites.Add(New productionSite(sitename, "Prct-proficy002.ap.pg.com", proficy_SERVER_PASSWORD_QQ, PROFICY_SERVER_USERNAME_QQ, "HPU"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Hair Care", siteName, SECTOR_BEAUTY, prStoryMapping.HuangPu, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.HuangpuHC))

        AllProdLines.Add(New ProdLine("Line 8", "Line 8", siteName, ModuleID, 2, 12, 8, "CWAH-010 Main", "CWAH-010"))
        AllProdLines.Add(New ProdLine("Line 4", "Line 4", siteName, ModuleID, 2, 12, 8, "CWAH-005 Main", "CWAH-005"))
        AllProdLines.Add(New ProdLine("Line 5", "Line 5", siteName, ModuleID, 2, 12, 8, "CWAH-006 Main", "CWAH-006"))
        AllProdLines.Add(New ProdLine("Line 6", "Line 6", siteName, ModuleID, 2, 12, 8, "CWAH-007 Main", "CWAH-007"))

        Dim lineName As String
        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem Care", siteName, SECTOR_FEM, prStoryMapping.Fem_LCC_HPU, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville, 2, 12, 8))

        lineName = "QAAH001"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QAAH002"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QAAH005"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QAAH006"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QAAH009"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QAAH010"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QAAH011"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QBAH001"
        AddLine(lineName, ModuleID2, lineName & " Converter")
        lineName = "QBAH002"
        AddLine(lineName, ModuleID2, lineName & " Converter")

    End Sub
    Private Sub initializeSite_Bangkok()
        AllProductionSites.Add(New ProdSite(SITE_BANGKOK, "BKK-MESDATABE", "", SERVER_PW_V6, SERVER_UN_V6, "BKK"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Hair Care", SITE_BANGKOK, SECTOR_BEAUTY, prStoryMapping.HuangPu, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.HuangpuHC))

        AllProdLines.Add(New ProdLine("B1", "B1", SITE_BANGKOK, ModuleID, 2, 12, 8, "BKK B1 Main", "BKK B1"))
        AllProdLines.Add(New ProdLine("B2", "B2", SITE_BANGKOK, ModuleID, 2, 12, 8, "BKK B2 Main", "BKK B2"))
        AllProdLines.Add(New ProdLine("B3", "B3", SITE_BANGKOK, ModuleID, 2, 12, 8, "BKK B3 Main", "BKK B3"))
        AllProdLines.Add(New ProdLine("B9", "B9", SITE_BANGKOK, ModuleID, 2, 12, 8, "BKK B9 Main", "BKK B9"))
    End Sub

    Private Sub initializeSite_Taicang()
        AllProductionSites.Add(New ProdSite(SITE_TAICANG, "tac-mesdtahc", SERVER_PW_V6, SERVER_UN_V6, "TAI"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Hair", SITE_TAICANG, SECTOR_BEAUTY, prStoryMapping.HuangPu, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.HuangpuHC))

        AllProdLines.Add(New ProdLine("Line 1", "Line 1", SITE_TAICANG, ModuleID, 2, 12, 8, "CWTI-010 Main", "CWTI-010"))
        AllProdLines.Add(New ProdLine("Line 2", "Line 2", SITE_TAICANG, ModuleID, 2, 12, 8, "CWTI-020 Main", "CWTI-020"))
        AllProdLines.Add(New ProdLine("Line 3", "Line 3", SITE_TAICANG, ModuleID, 2, 12, 8, "CWTI-030 Main", "CWTI-030"))
        AllProdLines.Add(New ProdLine("Line 4", "Line 4", SITE_TAICANG, ModuleID, 2, 12, 8, "CWTI-040 Main", "CWTI-040"))
        AllProdLines.Add(New ProdLine("Line S", "Line S", SITE_TAICANG, ModuleID, 2, 12, 8, "CWTI-050 Main", "CWTI-050"))
    End Sub



    Private Sub initializeSite_Naucalpan()
        Dim siteName As String = SITE_NAUCALPAN
        AllProductionSites.Add(New ProdSite(siteName, "NAUP-MESDTAHC", "", SERVER_PW_V6, SERVER_UN_V6, "NAU"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Oral", siteName, SECTOR_HEALTH, prStoryMapping.OralCareNau, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("Oasis", "Oasis", siteName, ModuleID, 2, 12, 6, "NAUP OASIS Main", "NAUP OASIS", prStoryMapping.OralCareNau_DF, True, "NAUP OASIS Perdida de velocidad")) ', True, "Perdida de velocidad", prStoryMapping.OralCare_DF))
        AllProdLines.Add(New ProdLine("1", "1", siteName, ModuleID, 2, 12, 6, "NAUP IWK1 Llenadora", "NAUP IWK1"))
        AllProdLines.Add(New ProdLine("2", "2", siteName, ModuleID, 2, 12, 6, "NAUP IWK2 Llenadora", "NAUP IWK2"))
        AllProdLines.Add(New ProdLine("3", "3", siteName, ModuleID, 2, 12, 6, "NAUP IWK3 Main", "NAUP IWK3"))


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "PHC PB", siteName, SECTOR_HEALTH, prStoryMapping.NaucalpanPHC_B, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("Pepto B", "Pepto B", siteName, ModuleID2, 2, 12, 6.75, "NAUP PEPTO B Main", "NAUP PEPTO B"))

        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "PHC PJ", siteName, SECTOR_HEALTH, prStoryMapping.NaucalpanPHC_J, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("Pepto J", "Pepto J", siteName, ModuleID3, 2, 12, 6.75, "NAUP PEPTO J Main", "NAUP PEPTO J"))

        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "PHC PM", siteName, SECTOR_HEALTH, prStoryMapping.NaucalpanPHC_Mex, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("Pepto Mexico", "Pepto Mexico", siteName, ModuleID4, 2, 12, 6.75, "NAUP PEPTO MEX Main", "NAUP PEPTO MEX"))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "PHC", siteName, SECTOR_HEALTH, prStoryMapping.NaucalpanPHC_Vita1, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("Vita 1", "Vita 1", siteName, ModuleID5, 2, 12, 6.75, "NAUP Vita 1 Enflex", "NAUP Vita 1"))
        AllProdLines.Add(New ProdLine("Vita 2", "Vita 2", siteName, ModuleID5, 2, 12, 6.75, "NAUP Vita 2 Enflex", "NAUP Vita 2"))

        Dim ModuleID6 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID6, "PHC", siteName, SECTOR_HEALTH, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        Dim x = New List(Of String) From {"Bundlera", "Cartoneta", "Casepacker", "Etiquetadora", "Llenadora", "No relacionados con el equipo", "Sorter de Tarro", "Tapadora", "Túnel", OTHERS_STRING}
        '   Dim x2 = New List(Of String) From {"AJUSTES", "EMERGENCIAS", "EQUIPO", "INTERVENCIONES", "MAKING", "MATERIALES", "RCO", "SERVICIOS", "TERMINO DE CEDULA", OTHERS_STRING}

        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = x
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = x

        Dim lineName As String
        lineName = "TARRO"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID6, 2, 12, 6.75, "NAUP " & lineName & " Main", "NAUP " & lineName))

        Dim ModuleID7 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID7, "PHC", siteName, SECTOR_HEALTH, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        Dim xy = New List(Of String) From {"Axon", "Cerradora", "Contadoras", "Llenadora", "No relacionados con el equipo", "Multifilas", "Sorter", "Sorter de Tapa", "Tapadora", "Túnel", OTHERS_STRING}
        ' Dim xy2 = New List(Of String) From {"AJUSTES", "EMERGENCIAS", "EQUIPO", "INTERVENCIONES", "MAKING", "MATERIALES", "RCO", "SERVICIOS", "TERMINO DE CEDULA", OTHERS_STRING}

        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = xy
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = xy

        lineName = "VVR Latas"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID7, 2, 12, 6.75, "NAUP " & lineName & " Main", "NAUP " & lineName))
    End Sub


    Private Sub initializeSite_BrownsSummit(y As String)
        SERVER_PW_V6 = "02PrStory" 'Text.Encoding.Default.GetString(u(s(y)))

        AllProductionSites.Add(New ProdSite(SITE_BROWNS_SUMMIT, "GBS-MESDATAHC", "gbs-hist001r", SERVER_PW_V6, SERVER_UN_V6, "GBO"))

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, BS_OC, SITE_BROWNS_SUMMIT, SECTOR_HEALTH, prStoryMapping.OralCare, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Machine", "Machine Section", "Stop Code", "Root Cause", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        '  AllProductionLines.Add(New productionLine("A", "BSA", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line A Main", "BS OC Line A"))
        AllProdLines.Add(New ProdLine("B", "BSB", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS B Line Reliability (JETN002)", "BS B Line (JETN002)"))
        AllProdLines.Add(New ProdLine("C", "BSC", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line C Main", "BS OC Line C"))
        AllProdLines.Add(New ProdLine("D", "BSD", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line D Main", "BS OC Line D"))
        '  AllProductionLines.Add(New productionLine("E", "BSE", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line E Main", "BS OC Line E"))
        AllProdLines.Add(New ProdLine("F", "BSF", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line F DF Main", "BS OC Line F DF", prStoryMapping.OralCare_DF, True, "BS OC Line F DF RateLoss"))
        AllProdLines.Add(New ProdLine("R", "BSR", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line R Main", "BS OC Line R"))

        AllProdLines.Add(New ProdLine("S", "BSS", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line S Main", "BS OC Line S"))

        AllProdLines.Add(New ProdLine("T", "BST", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line T DF Main", "BS OC Line T DF", prStoryMapping.OralCare_DF, True, "BS OC Line T DF RateLoss"))
        AllProdLines.Add(New ProdLine("Z", "BSZ", SITE_BROWNS_SUMMIT, ModuleID, 2, 12, 6, "BS OC Line Z Main", "BS OC Line Z"))

        'initialize the module
        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, BS_SC, SITE_BROWNS_SUMMIT, SECTOR_BEAUTY, prStoryMapping.SkinCare, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Reason3, DTsched_Mapping.SkinCare))
        AllProdLines.Add(New ProdLine("LSTL", "LSTL", SITE_BROWNS_SUMMIT, ModuleID2, 2, 12, 6, "BS1 Main", "BS1"))
        AllProdLines.Add(New ProdLine("3", "3", SITE_BROWNS_SUMMIT, ModuleID2, 2, 12, 6, "BS3 Main", "BS3"))
        AllProdLines.Add(New ProdLine("4", "4", SITE_BROWNS_SUMMIT, ModuleID2, 2, 12, 6, "BS4 Main", "BS4"))
        AllProdLines(AllProdLines.Count - 1).formatMapping = MappingByFormat.SkinCare
        AllProdLines(AllProdLines.Count - 1).shapeMapping = MappingByShape.SkinCare

        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "MC", SITE_BROWNS_SUMMIT, SECTOR_HEALTH, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Reason3, DTsched_Mapping.SkinCare))
        AllProdLines.Add(New ProdLine("X", "", SITE_BROWNS_SUMMIT, ModuleID3, 2, 12, 6, "BS MC Line X Main", "BS MC Line X"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
        AllProdLines.Add(New ProdLine("Y", "", SITE_BROWNS_SUMMIT, ModuleID3, 2, 12, 6, "BS MC Line Y Main", "BS MC Line Y"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True


    End Sub

#End Region

#Region "F&HC"
    Private Sub initializeSite_Hyderabad()
        AllProductionSites.Add(New ProdSite(SITE_HYDERABAD, "143.35.53.51", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "HYD"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC Akash", SITE_HYDERABAD, SECTOR_FHC, prStoryMapping.Hyderabad, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Tier2, DTsched_Mapping.Hyderabad))
        'initialize the lines
        AllProdLines.Add(New ProdLine("A3", "A3", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A3 Akash", "HYD A3"))
        AllProdLines.Add(New ProdLine("A4", "A4", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A4 Akash", "HYD A4"))
        AllProdLines.Add(New ProdLine("A9", "A9", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A9 Akash", "HYD A9"))
        AllProdLines.Add(New ProdLine("A10", "A10", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A10 Akash", "HYD A10"))
        AllProdLines.Add(New ProdLine("A11", "A11", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A11 Akash", "HYD A11"))
        AllProdLines.Add(New ProdLine("A12", "A12", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A12 Akash", "HYD A12"))
        AllProdLines.Add(New ProdLine("A13", "A13", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A13 Akash", "HYD A13"))
        AllProdLines.Add(New ProdLine("A14", "A14", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A14 Akash", "HYD A14"))
        AllProdLines.Add(New ProdLine("A15", "A15", SITE_HYDERABAD, ModuleID, 3, 8, 8, "HYD A15 Akash", "HYD A15"))

        ' AllProductionModules.Add(New productionModule("Jinyi", SITE_HYDERABAD, SECTOR_FHC, prStoryMapping.Hyderabad, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, DowntimeField.Reason3, DTsched_Mapping.Hyderabad))
        ' AllProductionLines.Add(New productionLine("J1", "J1", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J1_J2 J1", "HYD J1_J2"))
        ' AllProductionLines.Add(New productionLine("J2", "J2", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J1_J2 J2", "HYD J1_J2"))
        '
        '        AllProductionLines.Add(New productionLine("J3", "J3", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J3_J4 J3", "HYD J3_J4"))
        '        AllProductionLines.Add(New productionLine("J4", "J4", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J3_J4 J4", "HYD J3_J4"))
        '
        '        AllProductionLines.Add(New productionLine("J5", "J5", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J5_J6 J5", "HYD J5_J6"))
        '        AllProductionLines.Add(New productionLine("J6", "J6", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J5_J6 J6", "HYD J5_J6"))
        '
        '        AllProductionLines.Add(New productionLine("J7", "J7", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J7_J8 J7", "HYD J7_J8"))
        '        AllProductionLines.Add(New productionLine("J8", "J8", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J7_J8 J8", "HYD J7_J8"))
        '
        '        AllProductionLines.Add(New productionLine("J9", "J9", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J9_J10 J9", "HYD J9_J10"))
        '        AllProductionLines.Add(New productionLine("J10", "J10", SITE_HYDERABAD, "Jinyi", 3, 8, 8, "HYD J9_J10 J10", "HYD J9_J10"))
    End Sub


    Private Sub initializeSite_Lima()
        AllProductionSites.Add(New ProdSite(SITE_ALEXANDRIA, "137.179.105.202\LMSUDDB001", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "ALX"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC", SITE_ALEXANDRIA, SECTOR_FHC, prStoryMapping.GENERIC, DefaultProficyDowntimeProcedure.Maple, DefaultProficyProductionProcedure.Maple, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("212", "212", SITE_ALEXANDRIA, ModuleID, 2, 12, 6, "CONVERTER", "Line 212"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 212"
    End Sub





    Private Sub initializeSite_Amiens()
        Dim siteName As String = "Amiens"
        AllProductionSites.Add(New ProdSite(siteName, "ami-maple013.eu.pg.com", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "AMI"))
        AllProductionSites(AllProductionSites.Count - 1).ServerDatabase = "MES"

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "HDL&FE", siteName, SECTOR_FHC, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.Maple, DefaultProficyProductionProcedure.Maple, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Rakona))

        AllProdLines.Add(New ProdLine("Line 64", "Line 64", siteName, ModuleID, 2, 12, 6, "Remplisseuse", "Line 64"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 64"


    End Sub

    Private Sub initializeSite_Rakona()
        AllProductionSites.Add(New ProdSite("Rakona", "143.3.3.39\RAKMP101", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "RAK"))
        AllProductionSites(AllProductionSites.Count - 1).ServerDatabase = "RAKMES2"

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC LIQ", "Rakona", SECTOR_FHC, prStoryMapping.Rakona, DefaultProficyDowntimeProcedure.Maple, DefaultProficyProductionProcedure.Maple_New, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Rakona))

        AllProdLines.Add(New ProdLine("Line 1", "Line 1", "Rakona", ModuleID, 2, 12, 6, "FILLER", "Line 1"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 1"
        AllProdLines.Add(New ProdLine("Line 2", "Line 2", "Rakona", ModuleID, 2, 12, 6, "FILLER", "Line 2"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 2"
        AllProdLines.Add(New ProdLine("Line 7", "Line 7", "Rakona", ModuleID, 2, 12, 6, "FILLER", "Line 7"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 7"
        AllProdLines.Add(New ProdLine("Line 8", "Line 8", "Rakona", ModuleID, 2, 12, 6, "FILLER", "Line 8"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 8"

        '   AllProductionLines.Add(New productionLine("Linka H", "Linka H", "Rakona", ModuleID, 2, 12, 6, "FILLER", "Linka H"))
        '  AllProductionLines(AllProductionLines.Count - 1).Name_MAPLE = "Linka H"


        initializeSite_Rakona2()

    End Sub

    Private Sub initializeSite_Rakona2()
        AllProductionSites.Add(New ProdSite("Rakona 2", "RAK-PROD031\RAKMP101", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "RAK"))
        AllProductionSites(AllProductionSites.Count - 1).ServerDatabase = "RAKMES2"

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC", "Rakona", SECTOR_FHC, prStoryMapping.Rakona, DefaultProficyDowntimeProcedure.Maple, DefaultProficyProductionProcedure.Maple_New, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Rakona))

        AllProdLines.Add(New ProdLine("Linka H", "Linka H", "Rakona", ModuleID, 2, 12, 6, "FILLER", "Linka H"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Linka H"


    End Sub




    Private Sub initializeSite_STL()
        Dim siteName As String = "St. Louis"
        AllProductionSites.Add(New ProdSite(siteName, "stl-mespakdb", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "STL"))
        AllProductionSites(AllProductionSites.Count - 1).ServerDatabase = "Packing"

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Packing", siteName, SECTOR_FHC, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.Maple, DefaultProficyProductionProcedure.Maple_New, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Rakona))

        Dim lineName As String
        Dim inputProdUnitA As String = "filler"

        lineName = "Line 1"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, inputProdUnitA, lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = lineName

        lineName = "Line 2"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, inputProdUnitA, lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = lineName

        lineName = "Line 4"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, inputProdUnitA, lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = lineName

        lineName = "Line 5"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, inputProdUnitA, lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = lineName

        lineName = "Line 6"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, inputProdUnitA, lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = lineName

        lineName = "Line 10"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, "Converter", lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = lineName

        AllProdLines.Add(New ProdLine("Line 12", "Line 12", siteName, ModuleID, 2, 12, 6, "Converter", "Line 12"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 12"

        AllProdLines.Add(New ProdLine("Line 14", "Line 14", siteName, ModuleID, 2, 12, 6, "Converter", "Line 14"))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Line 14"

        lineName = "Line 20"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, inputProdUnitA, lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = lineName



    End Sub
#End Region


    Private Sub initializeSite_Baddi()

        Dim siteName As String = "Baddi"

        AllProductionSites.Add(New ProdSite(siteName, "IBP-MESDATAHHC2", "", SERVER_PW_V6, SERVER_UN_V6, "BAD"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Hair Care", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        ' BDIHCSL1
        'BHIHCSL2
        'BDIHCSL4
        'BDIHCBL1
        'BDIHCRL1


        lineName = "BDIHCSL1"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CCBSL1", ""))


    End Sub

    Private Sub initializeSite_Alexandria()
        Dim siteName As String = "Aleksandrow"

        AllProductionSites.Add(New ProdSite(siteName, "ale-mesdatabe2", "", SERVER_PW_V6, SERVER_UN_V6, "ALW"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Beauty", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String
        ' : Jx, J2, Ty, T3, and CG
        lineName = "JX"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "ALE" & lineName & " Main", ""))

        lineName = "J2"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "ALE" & lineName & " Main", ""))

        lineName = "TY"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "ALE" & lineName & " Main", ""))

        ' lineName = "T3"
        '  AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "ALE" & lineName & " Main", ""))

        lineName = "CG"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "ALE" & lineName & " Main", ""))
    End Sub

#Region "Baby Care"
    Private Sub initializeSite_Dover()
        Dim siteName As String = "Dover"

        AllProductionSites.Add(New ProdSite(siteName, "dvr-mesdatabc2", "", SERVER_PW_V6, SERVER_UN_V6, "DVR"))
        Dim lineName As String
        If False Then
            Dim ModuleID As Guid = Guid.NewGuid()
            AllProdModules.Add(New ProdModule(ModuleID, "Baby Wipes Constraint ", siteName, SECTOR_BABY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))



            lineName = "QPDO Conv"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO L21"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Wrapper", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO L22"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Wrapper", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO L23"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Wrapper", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO L24"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Wrapper", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO L25"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Wrapper", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO L26"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Wrapper", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO LP3"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " L5 ShrinkWrapper5", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

            lineName = "QPDO LP3 L6"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " ShrinkWrapper6", ""))
            AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
            ' QPDO Conv
            'QPDO L21
            'QPDO L22
            'QPDO L23
            'QPDO L24
            'QPDO L25
            'QPDO L26
            'QPDO LP3'
            'QPDO LP3 L6
        End If


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Baby Wipes", siteName, SECTOR_FAMILY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))


        lineName = "QPDO Conv"

        Dim prodList = New List(Of String) From {"QPDO Conv Diverter1", "QPDO Conv Diverter2", "QPDO Conv Diverter4", "QPDO Conv Diverter3"}
        '  AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE A MIELE 27", "", prStoryMapping.STRAIGHT, True, theVar))
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, "QPDO Conv Converter", "", prStoryMapping.STRAIGHT, True, prodList))
        '  allproductionlines(allproductionlines.count - 1).isstartupmode = true

        lineName = "QPDO L25"
        prodList = New List(Of String) From {"QPDO L25 Case Packer", "QPDO L25 Bagger", "QPDO L25 X_Conv E FW25-500", "QPDO L25 Secondary Packaging Conveyor", "QPDO L25 Case Erector", "QPDO L25 Fitment Applicator", "QPDO L25 Case Sealer"}
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, "QPDO L25 Wrapper", "", prStoryMapping.STRAIGHT, True, prodList))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "QPDO L21"
        prodList = New List(Of String) From {"QPDO L21 Secondary Packaging Conveyor", "QPDO L21 Case Packer", "QPDO L21 Case Erector", "QPDO L21 X_Conv A FW21-300", "QPDO L21 Case Sealer", "QPDO L21 Fitment Applicator", "QPDO L21 Fitment Applicator (Infeed)"}
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, "QPDO L21 Wrapper", "", prStoryMapping.STRAIGHT, True, prodList))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "QPDO L22"
        prodList = New List(Of String) From {"QPDO L22 Case Erector", "QPDO L22 Secondary Packaging Conveyor", "QPDO L22 Case Packer", "QPDO L22  X_Conv B FW22-200", "QPDO L22 Fitment Applicator", "QPDO L22 Case Sealer", "QPDO L22 FITMENT APPLICATOR (INFEED)"}
        '  AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE A MIELE 27", "", prStoryMapping.STRAIGHT, True, theVar))
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, "QPDO L22 Wrapper", "", prStoryMapping.STRAIGHT, True, prodList))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "QPDO L24"
        prodList = New List(Of String) From {"QPDO L24 Bagger", "QPDO L24 Fitment Applicator", "QPDO L24 Case Packer", "QPDO L24 Case Erector", "QPDO L24 Secondary Packaging Conveyor", "QPDO L24 X_Conv D FW24-100", "QPDO L24 Case Sealer"}
        '  AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE A MIELE 27", "", prStoryMapping.STRAIGHT, True, theVar))
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, "QPDO L24 Wrapper", "", prStoryMapping.STRAIGHT, True, prodList))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "QPDO L23"
        prodList = New List(Of String) From {"QPDO L23 Secondary Packaging Conveyor", "QPDO L23 Case Erector", "QPDO L23 Case Packer", "QPDO L23 X_Conv C FW23-400", "QPDO L23 Case Sealer", "QPDO L23 Fitment Applicator (Infeed)", "QPDO L23 Fitment Applicator"}
        '  AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE A MIELE 27", "", prStoryMapping.STRAIGHT, True, theVar))
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, "QPDO L23 Wrapper", "", prStoryMapping.STRAIGHT, True, prodList))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        lineName = "QPDO L26"
        prodList = New List(Of String) From {"QPDO L26 Wrapper", "QPDO L26 Case Erector", "QPDO L26 Bagger", "QPDO L26 Fitment Applicator", "QPDO L26 Case Packer", "QPDO L26 Secondary Packaging Conveyor", "QPDO L26 X_Conv F FW26-600"}
        '  AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE A MIELE 27", "", prStoryMapping.STRAIGHT, True, theVar))
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, "QPDO L26 Wrapper", "", prStoryMapping.STRAIGHT, True, prodList))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

    End Sub

    Private Sub initializeSite_Mequinenza()
        Dim siteName As String = "Mequinenza"

        AllProductionSites.Add(New ProdSite(siteName, "meq-mesdatabc", "", SERVER_PW_V6, SERVER_UN_V6, "MEQ"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Wipes", siteName, SECTOR_BABY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        lineName = "QPAZ100"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", ""))
        lineName = "QPAZ200"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Flow Wrapper", ""))
        lineName = "QPAZ201"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Flow Wrapper", ""))
        lineName = "QPAZ202"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Flow Wrapper", ""))
        lineName = "QPAZ203"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Flow Wrapper", ""))
        lineName = "QPAZ204"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Flow Wrapper", ""))
        lineName = "QPAZ205"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Flow Wrapper", ""))

        initializeSite_Urlati()
    End Sub

    Private Sub initializeSite_Urlati()
        Dim siteName As String = "Urlati"

        AllProductionSites.Add(New ProdSite(siteName, "URL-MESDTAHC", "", SERVER_PW_V6, SERVER_UN_V6, "URL"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Hair Care", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        lineName = "UN L1 Main"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 Unscrambler"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 Orientator"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 Divizor"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 Etichetare"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 infoliator Baxuri"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 Paletizor"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 BMS"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 PAGO"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 Monobloc"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L1 Sticker-Pack"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Main"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Unscrambler"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Orientator"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Divizor"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Etichetare"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Infoliator Baxuri"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Paletizor"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 BMS"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Monobloc"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L2 Sticker-Pack"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Main"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Unscrambler"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Orientator"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Divizor"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Etichetare"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Infoliator Baxuri"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Paletizor"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 BMS"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Monobloc"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L3 Sticker-Pack"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L4 Main"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L4 NGRU"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L4 Etichetare"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L4 NGS"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L4 BMS"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L4 Monobloc"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "UN L4 Sticker-Pack"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "FPHU Main"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))

        lineName = "FPHU Stretch"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName, ""))
    End Sub

    Private Sub initializeSite_Euskirchen()
        Dim siteName As String = "Euskirchen"

        AllProductionSites.Add(New ProdSite(siteName, "EUS-MESDTABC", "", SERVER_PW_V6, SERVER_UN_V6, "EUS"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Wipes", siteName, SECTOR_BABY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        lineName = "QPEU082"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Pack Leg", ""))
        lineName = "QPEU083"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Pack Leg", ""))
        lineName = "QPEU084"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Pack Leg", ""))

        Dim modName As String = "Baby"
        Dim tmpLineName As String
        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, modName, siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        tmpLineName = "DIEU131"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU132"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU133"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU134"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU135"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU136"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU137"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU138"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU139"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU171"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU172"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU173"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU174"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU175"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU176"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU177"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU178"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))
        tmpLineName = "DIEU179"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName & " Converter", tmpLineName))



    End Sub

    Private Sub initializeSite_Santiago()
        Dim siteName As String = SITE_SANTIAGO
        Dim lineName As String
        Dim ModuleID As Guid = Guid.NewGuid()

        AllProductionSites.Add(New ProdSite(siteName, "CHPA-MESDTABC", "", SERVER_PW_V6, SERVER_UN_V6, "SAN"))
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "DISI110"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "DISI800"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Central Conveyor", lineName))
    End Sub

    Private Sub initializeSite_VillaMercedes()
        Dim siteName As String = SITE_VILLAMERCEDES
        Dim lineName As String
        Dim ModuleID As Guid = Guid.NewGuid()

        AllProductionSites.Add(New ProdSite(siteName, "AVM-MESDTABC.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "VIM"))
        AllProdModules.Add(New ProdModule(ModuleID, "Baby", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "DIMR110"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "DIMR111"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "DIMR112"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "DIMR113"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "DIMR005"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "DIMR004"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " CONVERTER", lineName))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem", siteName, SECTOR_FEM, prStoryMapping.Fem_LCC_HPU, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "QAMR001"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " CONVERTER", lineName))

    End Sub

    Private Sub initializeSite_Mandideep()
        AllProductionSites.Add(New ProdSite(SITE_MANDIDEEP, "mdp-mesdtabc", "", SERVER_PW_V6, SERVER_UN_V6, "MDP"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_MANDIDEEP, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("IP71", "IP71", SITE_MANDIDEEP, ModuleID, 3, 8, 7, "DIMX171 CONVERTER", "DIMX171"))
        AllProdLines.Add(New ProdLine("IP72", "IP72", SITE_MANDIDEEP, ModuleID, 3, 8, 7, "DIMX172 CONVERTER", "DIMX172"))
        AllProdLines.Add(New ProdLine("IP73", "IP73", SITE_MANDIDEEP, ModuleID, 3, 8, 7, "DIMX173 CONVERTER", "DIMX173"))
        AllProdLines.Add(New ProdLine("IP74", "IP74", SITE_MANDIDEEP, ModuleID, 3, 8, 7, "DIMX174 CONVERTER", "DIMX174"))

        AllProdLines.Add(New ProdLine("IP02", "IP02", SITE_MANDIDEEP, ModuleID, 3, 8, 7, "DIMX101 CONVERTER", "DIMX101"))
        AllProdLines.Add(New ProdLine("IP03", "IP03", SITE_MANDIDEEP, ModuleID, 3, 8, 7, "DIMX102 CONVERTER", "DIMX102"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem Care", SITE_MANDIDEEP, SECTOR_FEM, prStoryMapping.Mandideep_Fem, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        Dim lineName As String
        lineName = "QAMX001"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_MANDIDEEP, ModuleID2, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "QAMX002"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_MANDIDEEP, ModuleID2, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "QAMX003"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_MANDIDEEP, ModuleID2, 3, 8, 7, lineName & " CONVERTER", lineName))

        initializeSite_Mandideep2()
    End Sub

    Private Sub initializeSite_Mandideep2()
        Dim siteName As String = "MPD "
        AllProductionSites.Add(New ProdSite(siteName, "mdp-mesdatafc", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "MD2"))
        'initialize the lines
        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem Care", siteName, SECTOR_FEM, prStoryMapping.Mandideep_Fem, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        Dim lineName As String
        lineName = "QAMX004"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "QAMX005"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " CONVERTER", lineName))

    End Sub

    Private Sub initializeSite_XQ()
        Dim siteName As String

        siteName = "Xiqing Oral"
        AllProductionSites.Add(New ProdSite(siteName, "XQ-MESdataBE", "", SERVER_PW_V6, SERVER_UN_V6, "XQQ"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Care", siteName, SECTOR_HEALTH, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        Dim lineName As String

        lineName = "Line1"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "Jeah-010 main", lineName))
        lineName = "Line2"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "Jeah-020 main", lineName))
        lineName = "Line3"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "Jeah-030 main", lineName))
        lineName = "Line4"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "Jeah-040 main", lineName))

    End Sub

    Private Sub initializeSite_Xiquing()
        Dim siteName As String

        siteName = "Xiquing"
        AllProductionSites.Add(New ProdSite(siteName, "XQ-MESDATABC", "", SERVER_PW_V6, SERVER_UN_V6, "XQU"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        Dim lineName As String

        lineName = "DIAH111"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH112"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH113"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH114"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH120"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH121"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH122"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH123"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH124"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH161"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIAH162"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
    End Sub

    Private Sub initializeSite_Manchester()
        AllProductionSites.Add(New ProdSite(SITE_MANCHESTER, "MAN-MESDATABC2.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "MAN"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_MANCHESTER, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIMA101", "DIMA101", SITE_MANCHESTER, ModuleID, 3, 8, 7, "DIMA101 CONVERTER", "DIMA101"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
        AllProdLines.Add(New ProdLine("DIMA105", "DIMA105", SITE_MANCHESTER, ModuleID, 3, 8, 7, "DIMA105 CONVERTER", "DIMA105"))
        AllProdLines.Add(New ProdLine("DIMA106", "DIMA106", SITE_MANCHESTER, ModuleID, 3, 8, 7, "DIMA106 CONVERTER", "DIMA106"))
        AllProdLines.Add(New ProdLine("DIMA107", "DIMA107", SITE_MANCHESTER, ModuleID, 3, 8, 7, "DIMA107 CONVERTER", "DIMA107"))
        AllProdLines.Add(New ProdLine("DIMA108", "DIMA108", SITE_MANCHESTER, ModuleID, 3, 8, 7, "DIMA108 CONVERTER", "DIMA108"))

        AllProdLines.Add(New ProdLine("DIMA109", "DIMA109", SITE_MANCHESTER, ModuleID, 3, 8, 7, "DIMA109 CONVERTER", "DIMA109"))
        AllProdLines.Add(New ProdLine("DIMA110", "DIMA110", SITE_MANCHESTER, ModuleID, 3, 8, 7, "DIMA110 CONVERTER", "DIMA110"))


    End Sub

    Private Sub initializeSite_Tepeji()
        AllProductionSites.Add(New ProdSite(SITE_TEPEJI, "MTEP-MESDTABC2", "", SERVER_PW_V6, SERVER_UN_V6, "TPE"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_TEPEJI, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIBH112", "DIBH112", SITE_TEPEJI, ModuleID, 3, 8, 7.5, "DIBH112 CONVERTER", "DIBH112"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
        AllProdLines.Add(New ProdLine("DIBH104", "DIBH104", SITE_TEPEJI, ModuleID, 3, 8, 7.5, "DIBH104 CONVERTER", "DIBH104"))
        AllProdLines.Add(New ProdLine("DIBH106", "DIBH106", SITE_TEPEJI, ModuleID, 3, 8, 7.5, "DIBH106 CONVERTER", "DIBH106"))
        AllProdLines.Add(New ProdLine("DIBH108", "DIBH108", SITE_TEPEJI, ModuleID, 3, 8, 7.5, "DIBH108 CONVERTER", "DIBH108"))
        AllProdLines.Add(New ProdLine("DIBH109", "DIBH109", SITE_TEPEJI, ModuleID, 3, 8, 7.5, "DIBH109 CONVERTER", "DIBH109"))
        AllProdLines.Add(New ProdLine("DIBH110", "DIBH110", SITE_TEPEJI, ModuleID, 3, 8, 7.5, "DIBH110 CONVERTER", "DIBH110"))
        AllProdLines.Add(New ProdLine("DIBH111", "DIBH111", SITE_TEPEJI, ModuleID, 3, 8, 7.5, "DIBH111 CONVERTER", "DIBH111"))


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem", SITE_TEPEJI, SECTOR_FEM, prStoryMapping.TepejiFem, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville))
        'initialize the lines
        AllProdLines.Add(New ProdLine("QABH101", "QABH101", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QABH101 CONVERTER", "QABH101"))
        AllProdLines.Add(New ProdLine("QABH102", "QABH102", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QABH102 CONVERTER", "QABH102"))
        AllProdLines.Add(New ProdLine("QABH013", "QABH013", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QABH013 CONVERTER", "QABH013"))
        AllProdLines.Add(New ProdLine("QABH004", "QABH004", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QABH004 CONVERTER", "QABH004"))
        AllProdLines.Add(New ProdLine("QABH007", "QABH007", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QABH007 CONVERTER", "QABH007"))
        AllProdLines.Add(New ProdLine("QABH009", "QABH009", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QABH009 CONVERTER", "QABH009"))
        AllProdLines.Add(New ProdLine("QABH010", "QABH010", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QABH010 CONVERTER", "QABH010"))
        AllProdLines.Add(New ProdLine("QBBH006", "QBBH006", SITE_TEPEJI, ModuleID2, 3, 8, 7.5, "QBBH006 CONVERTER", "QBBH006"))

    End Sub
    Private Sub initializeSite_Louveira()
        Dim siteName As String = SITE_LOUVERIA
        Dim lineName As String
        Dim ModuleID As Guid = Guid.NewGuid()


        AllProductionSites.Add(New ProdSite(siteName, "BLOV-MESDTABC", "", SERVER_PW_V6, SERVER_UN_V6, "LOU"))
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DILU101", "DILU101", siteName, ModuleID, 3, 8, 7, "DILU101 Converter", "DILU101"))
        AllProdLines.Add(New ProdLine("DILU111", "DILU111", siteName, ModuleID, 3, 8, 7, "DILU111 CONVERTER", "DILU111"))
        AllProdLines.Add(New ProdLine("DILU112", "DILU112", siteName, ModuleID, 3, 8, 7, "DILU112 CONVERTER", "DILU112"))
        AllProdLines.Add(New ProdLine("DILU113", "DILU113", siteName, ModuleID, 3, 8, 7, "DILU113 CONVERTER", "DILU113"))
        AllProdLines.Add(New ProdLine("DILU114", "DILU114", siteName, ModuleID, 3, 8, 7, "DILU114 CONVERTER", "DILU114"))
        AllProdLines.Add(New ProdLine("DILU115", "DILU115", siteName, ModuleID, 3, 8, 7, "DILU115 CONVERTER", "DILU115"))
        AllProdLines.Add(New ProdLine("DILU116", "DILU116", siteName, ModuleID, 3, 8, 7, "DILU116 CONVERTER", "DILU116"))
        AllProdLines.Add(New ProdLine("DILU117", "DILU117", siteName, ModuleID, 3, 8, 7, "DILU117 CONVERTER", "DILU117"))
        AllProdLines.Add(New ProdLine("DILU118", "DILU118", siteName, ModuleID, 3, 8, 7, "DILU118 CONVERTER", "DILU118"))
        AllProdLines.Add(New ProdLine("DILU119", "DILU119", siteName, ModuleID, 3, 8, 7, "DILU119 CONVERTER", "DILU119"))
        AllProdLines.Add(New ProdLine("DILU144", "DILU144", siteName, ModuleID, 3, 8, 7, "DILU144 CONVERTER", "DILU144"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        Dim ModuleID2 = New Guid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem", siteName, SECTOR_FEM, prStoryMapping.Fem_LuisCustom, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Fem_LuisCustom))

        lineName = "QALU001"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QALU005"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QALU007"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QALU012"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QALU013"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QALU014"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QALU015"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QALU017"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "LLLU001"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Main", lineName))


    End Sub
    Private Sub initializeSite_Jijona()
        AllProductionSites.Add(New ProdSite(SITE_JIJONA, "jij-mesdtabc.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "MHP"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_JIJONA, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIAB101", "DIAB101", SITE_JIJONA, ModuleID, 3, 8, 7, "DIAB101 CONVERTER", "DIAB101"))
        AllProdLines.Add(New ProdLine("DIAB102", "DIAB102", SITE_JIJONA, ModuleID, 3, 8, 7, "DIAB102 CONVERTER", "DIAB102"))
        AllProdLines.Add(New ProdLine("DIAB103", "DIAB103", SITE_JIJONA, ModuleID, 3, 8, 7, "DIAB103 CONVERTER", "DIAB103"))
        AllProdLines.Add(New ProdLine("DIAB104", "DIAB104", SITE_JIJONA, ModuleID, 3, 8, 7, "DIAB104 CONVERTER", "DIAB104"))
        AllProdLines.Add(New ProdLine("DIAB105", "DIAB105", SITE_JIJONA, ModuleID, 3, 8, 7, "DIAB105 CONVERTER", "DIAB105"))
        AllProdLines.Add(New ProdLine("DIAB107", "DIAB107", SITE_JIJONA, ModuleID, 3, 8, 7, "DIAB107 CONVERTER", "DIAB107"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem Care", SITE_JIJONA, SECTOR_FEM, prStoryMapping.JijonaUltra, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville))
        'initialize the lines
        Dim lineName As String

        lineName = "QAAB003"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_JIJONA, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "QAAB005"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_JIJONA, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAAB008"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_JIJONA, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))

        initializeSite_Goa()
    End Sub

    Private Sub initializeSite_Goa()
        Dim siteName As String = "Goa"
        AllProductionSites.Add(New ProdSite(siteName, "155.124.82.63", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "GOA"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem Care", siteName, SECTOR_FEM, prStoryMapping.Fem_LCC_HPU, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville))
        'initialize the lines
        Dim lineName As String

        lineName = "QAGO003"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO004"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO005"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO006"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO007"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO008"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO009"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO010"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAGO011"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))


    End Sub

    Private Sub initializeSite_Akashi()
        AllProductionSites.Add(New ProdSite(SITE_AKASHI, "AKA-MESDATABC2", "", SERVER_PW_V6, SERVER_UN_V6, "AKA"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_AKASHI, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIAK131", "DIAK131", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK131 CONVERTER", "DIAK131"))
        AllProdLines.Add(New ProdLine("DIAK132", "DIAK132", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK132 CONVERTER", "DIAK132"))
        AllProdLines.Add(New ProdLine("DIAK133", "DIAK133", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK133 CONVERTER", "DIAK133"))
        AllProdLines.Add(New ProdLine("DIAK134", "DIAK134", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK134 CONVERTER", "DIAK134"))
        AllProdLines.Add(New ProdLine("DIAK135", "DIAK135", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK135 CONVERTER", "DIAK135"))
        AllProdLines.Add(New ProdLine("DIAK136", "DIAK136", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK136 CONVERTER", "DIAK136"))
        AllProdLines.Add(New ProdLine("DIAK137", "DIAK137", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK137 Converter", "DIAK137"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        AllProdLines.Add(New ProdLine("DIAK171", "DIAK171", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK171 CONVERTER", "DIAK171"))
        AllProdLines.Add(New ProdLine("DIAK172", "DIAK172", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK172 CONVERTER", "DIAK172"))
        AllProdLines.Add(New ProdLine("DIAK173", "DIAK173", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK173 CONVERTER", "DIAK173"))
        AllProdLines.Add(New ProdLine("DIAK174", "DIAK174", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK174 CONVERTER", "DIAK174"))
        AllProdLines.Add(New ProdLine("DIAK175", "DIAK175", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK175 CONVERTER", "DIAK175"))
        AllProdLines.Add(New ProdLine("DIAK176", "DIAK176", SITE_AKASHI, ModuleID, 3, 8, 7, "DIAK176 CONVERTER", "DIAK176"))
    End Sub

    Private Sub initializeSite_Montornes()
        Dim siteName As String = "Montornes"
        AllProductionSites.Add(New ProdSite(siteName, "mon-mesdtabc2", "", SERVER_PW_V6, SERVER_UN_V6, "MNT"))

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, " ", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHTPlusOne, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        lineName = "QIMS017"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Converter", lineName))

        lineName = "QIMS018"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Converter", lineName))

        lineName = "QIMS019"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Converter", lineName))

        lineName = "QBMS038"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Converter", lineName))

        lineName = "QAMS002"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Converter", lineName))
    End Sub

    Private Sub initializeSite_Dammam()
        Dim siteName As String = "Dammam"
        AllProductionSites.Add(New ProdSite(siteName, "DAM-MESDTAHHC", "", SERVER_PW_V6, SERVER_UN_V6, "DAM"))

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "DCF", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHTPlusOne, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        lineName = "DCF Line 01"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Main", lineName))
        lineName = "DCF Line 02"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Main", lineName))
        lineName = "DCF Line 03"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Main", lineName))
        lineName = "DCF Line 05"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Main", lineName))
        lineName = "DCF Line 07"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Main", lineName))
        lineName = "DCF PDC 03"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 8, lineName & " Main", lineName))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "PHC", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHTPlusOne, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        lineName = "PHC Line B01"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PHC Line B02"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PHC Line B03"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PHC Line B04"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PHC Line B05"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 8, lineName & " Main", lineName))
        lineName = "SAB6HAC"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 8, lineName & " Main", lineName))

        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "PSG", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHTPlusOne, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))

        lineName = "PSG Line B01"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line B02"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line B04"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, "SA20LND Main", lineName))

        lineName = "PSG Line C01"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C02"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C03"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C04"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C05"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C06"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C07"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C08"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C09"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C10"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C11"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C12"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C13"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "PSG Line C14"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))
        lineName = "SA19LND"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 8, lineName & " Main", lineName))


    End Sub

    Private Sub initializeSite_HuangPu()
        Dim siteName As String = SITE_HUANGPU & " Hair"
        AllProductionSites.Add(New ProdSite(siteName, "prch-mesdatabe.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "HPU"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Care", siteName, SECTOR_BEAUTY, prStoryMapping.HuangPu, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.HuangpuHC))

        AllProdLines.Add(New ProdLine("B", "B", siteName, ModuleID, 2, 12, 8, "New Line B Main", "New Line B"))
        AllProdLines.Add(New ProdLine("C", "C", siteName, ModuleID, 2, 12, 8, "New LineC Main", "New LineC"))
        AllProdLines.Add(New ProdLine("F", "F", siteName, ModuleID, 2, 12, 8, "HC LineF Main", "HC LineF"))
        AllProdLines.Add(New ProdLine("G", "G", siteName, ModuleID, 2, 12, 8, "HC LineG Main", "HC LineG"))
        AllProdLines.Add(New ProdLine("K", "K", siteName, ModuleID, 2, 12, 8, "HC LineK Main", "HC LineK"))
        AllProdLines.Add(New ProdLine("R", "R", siteName, ModuleID, 2, 12, 8, "HC Line R Main", "HC LineR"))
        AllProdLines.Add(New ProdLine("A", "A", siteName, ModuleID, 2, 12, 8, "SC LineA Main", "SC LineA"))
        AllProdLines.Add(New ProdLine("B", "B", siteName, ModuleID, 2, 12, 8, "SC LineB Main", "SC LineB"))
        AllProdLines.Add(New ProdLine("C", "C", siteName, ModuleID, 2, 12, 8, "SC LineC Main", "SC LineC"))

        initializeSite_HuangPu_Oral()
    End Sub

    Private Sub initializeSite_HuangPu_Oral()
        Dim siteName As String = SITE_HUANGPU & " Oral"
        AllProductionSites.Add(New ProdSite(siteName, "prch-mesdatabe.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "HPO"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Care", siteName, SECTOR_HEALTH, prStoryMapping.OralCare, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.DTGroup, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("Line D", "Line D", siteName, ModuleID2, 2, 12, 8, "OC LineD Main", "OC LineD"))

    End Sub

    Private Sub initializeSite_HuangpuBaby()
        Dim siteName As String = SITE_HUANGPU

        AllProductionSites.Add(New ProdSite(siteName, "PRCH-MESDTABC.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "HPB"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIHU101", "DIHU101", siteName, ModuleID, 3, 8, 7, "DIHU101 CONVERTER", "DIHU101"))
        AllProdLines.Add(New ProdLine("DIHU102", "DIHU102", siteName, ModuleID, 3, 8, 7, "DIHU102 CONVERTER", "DIHU102"))
        AllProdLines.Add(New ProdLine("DIHU103", "DIHU103", siteName, ModuleID, 3, 8, 7, "DIHU103 CONVERTER", "DIHU103"))
        AllProdLines.Add(New ProdLine("DIHU104", "DIHU104", siteName, ModuleID, 3, 8, 7, "DIHU104 CONVERTER", "DIHU104"))
        AllProdLines.Add(New ProdLine("DIHU105", "DIHU105", siteName, ModuleID, 3, 8, 7, "DIHU105 CONVERTER", "DIHU105"))
        AllProdLines.Add(New ProdLine("DIHU106", "DIHU106", siteName, ModuleID, 3, 8, 7, "DIHU106 CONVERTER", "DIHU106"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem Care", siteName, SECTOR_FEM, prStoryMapping.FemCare_Pads_Huangpu, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        Dim lineName As String
        lineName = "QAHU001"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAHU003"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAHU010"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAHU011"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAHU012"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QAHU013"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QBHU001"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))




    End Sub
    Private Sub initializeSite_October6()
        AllProductionSites.Add(New ProdSite(SITE_OCTOBER6, "OCT-MESDATABC", "", SERVER_PW_V6, SERVER_UN_V6, "OCT"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_OCTOBER6, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIOC101", "DIOC101", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC101 CONVERTER", "DIOC101"))
        AllProdLines.Add(New ProdLine("DIOC102", "DIOC102", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC102 CONVERTER", "DIOC102"))
        AllProdLines.Add(New ProdLine("DIOC103", "DIOC103", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC103 CONVERTER", "DIOC103"))
        AllProdLines.Add(New ProdLine("DIOC104", "DIOC104", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC104 CONVERTER", "DIOC104"))
        AllProdLines.Add(New ProdLine("DIOC105", "DIOC105", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC105 CONVERTER", "DIOC105"))
        AllProdLines.Add(New ProdLine("DIOC106", "DIOC106", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC106 CONVERTER", "DIOC106"))
        AllProdLines.Add(New ProdLine("DIOC107", "DIOC107", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC107 CONVERTER", "DIOC107"))
        AllProdLines.Add(New ProdLine("DIOC108", "DIOC108", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC108 CONVERTER", "DIOC108"))
        AllProdLines.Add(New ProdLine("DIOC109", "DIOC109", SITE_OCTOBER6, ModuleID, 3, 8, 7, "DIOC109 CONVERTER", "DIOC109"))


    End Sub

    Private Sub initializeSite_Cairo()
        Dim siteName As String = "Cairo"


        AllProductionSites.Add(New ProdSite(siteName, "cai-mesdatahc", "", SERVER_PW_V6, SERVER_UN_V6, "CAI"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "CPLN", siteName, SECTOR_FHC, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        If False Then

            lineName = "MIELE A"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE A MIELE 27", "", prStoryMapping.STRAIGHT, True, "CPLN " & "MIELE A MIELE 28"))

            lineName = "MIELE B"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE B MIELE 29", "", prStoryMapping.STRAIGHT, True, "CPLN " & "MIELE B MIELE 30"))


            lineName = "MIELE A MIELE 27"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & lineName, ""))

            lineName = "MIELE A MIELE 28"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & lineName, ""))

            lineName = "MIELE B MIELE 29"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & lineName, ""))

            lineName = "MIELE B MIELE 30"
            AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & lineName, ""))

        End If

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem", siteName, SECTOR_FEM, prStoryMapping.Fem_LCC_HPU, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        lineName = "QACA008"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", ""))
        lineName = "QACA009"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", ""))
        lineName = "QACA013"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", ""))
        lineName = "QACA015"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", ""))
        lineName = "QACA017"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", ""))
        lineName = "QACA019"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", ""))
        lineName = "QACA054"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", ""))

        ' lineName = "MIELE X"
        '  AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "CPLN " & "MIELE A MIELE 27", "", prStoryMapping.STRAIGHT, True, theVar))
    End Sub

    Private Sub initializeSite_Lagos()
        AllProductionSites.Add(New ProdSite(SITE_LAGOS, "lag-mesdatabc", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "OCT"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_LAGOS, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DILS101", "DILS101", SITE_LAGOS, ModuleID, 3, 8, 7, "DILS101 CONVERTER", "DILS101"))
        AllProdLines.Add(New ProdLine("DILS102", "DILS102", SITE_LAGOS, ModuleID, 3, 8, 7, "DILS102 CONVERTER", "DILS102"))


    End Sub

    Private Sub initializeSite_BenCat()
        AllProductionSites.Add(New ProdSite(SITE_BENCAT, "BC-Mesdatabc", "", SERVER_PW_V6, SERVER_UN_V6, "BEN"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_BENCAT, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIVP101", "DIVP101", SITE_BENCAT, ModuleID, 3, 8, 7, "DIVP101 CONVERTER", "DIVP101"))
        AllProdLines.Add(New ProdLine("DIVP171", "DIVP171", SITE_BENCAT, ModuleID, 3, 8, 7, "DIVP171 CONVERTER", "DIVP171"))
    End Sub

    Private Sub initializeSite_Guatire()
        Dim sitename As String = SITE_GUATIRE
        AllProductionSites.Add(New ProdSite(SITE_GUATIRE, "VGUA-MESDATABC.la.pg.com", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "GUA"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_GUATIRE, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIGU001", "DIGU001", SITE_GUATIRE, ModuleID, 3, 8, 7, "DIGU001 Converter", "DIGU001"))
        AllProdLines.Add(New ProdLine("DIGU002", "DIGU002", SITE_GUATIRE, ModuleID, 3, 8, 7, "DIGU002 Converter", "DIGU002"))
        AllProdLines.Add(New ProdLine("DIGU003", "DIGU003", SITE_GUATIRE, ModuleID, 3, 8, 7, "DIGU003 Converter", "DIGU003"))
        AllProdLines.Add(New ProdLine("DIGU104", "DIGU104", SITE_GUATIRE, ModuleID, 3, 8, 7, "DIGU104 Converter", "DIGU104"))
        AllProdLines.Add(New ProdLine("DIGU105", "DIGU105", SITE_GUATIRE, ModuleID, 3, 8, 7, "DIGU105 Converter", "DIGU105"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem", sitename, SECTOR_FEM, prStoryMapping.Fem_LCC_HPU, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        Dim lineName As String

        lineName = "QAGU001"
        AllProdLines.Add(New ProdLine(lineName, lineName, SITE_GUATIRE, ModuleID2, 3, 8, 7, lineName & " Converter", ""))


    End Sub

    Private Sub initializeSite_MATERIALES()
        AllProductionSites.Add(New ProdSite(SITE_MATERIALES, "PEMA-MESDATABC.la.pg.com", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "MAT"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_MATERIALES, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIML102", "DIML102", SITE_MATERIALES, ModuleID, 3, 8, 7, "DIML102 CONVERTER", "DIML102"))
    End Sub

    Private Sub initializeSite_GYONGYOS()
        AllProductionSites.Add(New ProdSite(SITE_GYONGYOS, "hugp-mesdatabc", "", SERVER_PW_V6, SERVER_UN_V6, "GYO"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_GYONGYOS, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIGY101", "DIGY101", SITE_GYONGYOS, ModuleID, 3, 8, 7, "DIGY101 CONVERTER", "DIGY101"))
        AllProdLines.Add(New ProdLine("DIGY102", "DIGY102", SITE_GYONGYOS, ModuleID, 3, 8, 7, "DIGY102 CONVERTER", "DIGY102"))
        AllProdLines.Add(New ProdLine("DIGY103", "DIGY103", SITE_GYONGYOS, ModuleID, 3, 8, 7, "DIGY103 CONVERTER", "DIGY103"))

    End Sub
    Private Sub initializeSite_NOVO()
        AllProductionSites.Add(New ProdSite(SITE_NOVO, "nov-mesdtabc.na.pg.com", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "NOV"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_NOVO, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DINK101", "DINK101", SITE_NOVO, ModuleID, 3, 8, 7, "DINK101 CONVERTER", "DINK101"))
        AllProdLines.Add(New ProdLine("DINK102", "DINK102", SITE_NOVO, ModuleID, 3, 8, 7, "DINK102 CONVERTER", "DINK102"))
        AllProdLines.Add(New ProdLine("DINK103", "DINK103", SITE_NOVO, ModuleID, 3, 8, 7, "DINK103 CONVERTER", "DINK103"))
        AllProdLines.Add(New ProdLine("DINK104", "DINK104", SITE_NOVO, ModuleID, 3, 8, 7, "DINK104 CONVERTER", "DINK104"))
        AllProdLines.Add(New ProdLine("DINK105", "DINK105", SITE_NOVO, ModuleID, 3, 8, 7, "DINK105 CONVERTER", "DINK105"))

    End Sub
    Private Sub initializeSite_GEBZE()

        AllProductionSites.Add(New ProdSite(SITE_GEBZE, "geb-mesdtabc.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "GBZ"))
        Dim ModuleID As Guid = Guid.NewGuid()
        Dim lineName As String
        Dim siteName As String = SITE_GEBZE
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", SITE_GEBZE, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIAD102", "DIAD102", SITE_GEBZE, ModuleID, 3, 8, 7, "DIAD102 CONVERTER", "DIAD102"))
        AllProdLines.Add(New ProdLine("DIAD103", "DIAD103", SITE_GEBZE, ModuleID, 3, 8, 7, "DIAD103 CONVERTER", "DIAD103"))
        AllProdLines.Add(New ProdLine("DIAD104", "DIAD104", SITE_GEBZE, ModuleID, 3, 8, 7, "DIAD104 CONVERTER", "DIAD104"))
        AllProdLines.Add(New ProdLine("DIAD105", "DIAD105", SITE_GEBZE, ModuleID, 3, 8, 7, "DIAD105 CONVERTER", "DIAD105"))
        AllProdLines.Add(New ProdLine("DIAD106", "DIAD106", SITE_GEBZE, ModuleID, 3, 8, 7, "DIAD106 CONVERTER", "DIAD106"))
        AllProdLines.Add(New ProdLine("DIAD107", "DIAD107", SITE_GEBZE, ModuleID, 3, 8, 7, "DIAD107 CONVERTER", "DIAD107"))


        Dim ModuleIDx As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleIDx, "Fem Care", SITE_GEBZE, SECTOR_FEM, prStoryMapping.FemCare_Pads, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        'QAAD005 (line N5) and QAAD007 (line N7)
        lineName = "QAAD005"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleIDx, 3, 8, 7, lineName & " CONVERTER", lineName))
        lineName = "QAAD007"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleIDx, 3, 8, 7, lineName & " CONVERTER", lineName))


        Dim ModuleIDx2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleIDx2, "F&HC", SITE_GEBZE, SECTOR_FHC, prStoryMapping.Rakona, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Rakona))

        'QAAD005 (line N5) and QAAD007 (line N7)
        lineName = "FPAD001"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleIDx2, 3, 8, 7, lineName & " Main", lineName))



    End Sub

    Private Sub initializeSite_Ibadan()
        Dim siteName As String = "Ibadan"
        Dim lineName As String


        AllProductionSites.Add(New ProdSite(siteName, "Iba-mesdtahhc", "", SERVER_PW_V6, SERVER_UN_V6, "IBA"))

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC", SITE_MANDIDEEP, SECTOR_FHC, prStoryMapping.STRAIGHTPLANNEDPlusOne, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        lineName = "UVA A"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L1 UVA 1A Main", "IB L1 UVA 1A"))

        lineName = "UVA B"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L1 UVA 1B Filler", "IB L1 UVA 1B"))

        lineName = "UVA C"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L1 UVA 1C Filler", "IB L1 UVA 1C"))

        lineName = "UVA D"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L1 UVA 1D Filler", "IB L1 UVA 1D"))

        lineName = "UVA E"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L2 UVA 2E Filler", "IB L2 UVA 2E"))

        lineName = "UVA F"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L2 UVA 2F Filler", "IB L2 UVA 2F"))

        lineName = "UVA G"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L2 UVA 2G Filler", "IB L2 UVA 2G"))

        lineName = "UVA H"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, "IB L2 UVA 2H Filler", "IB L2 UVA 2H"))

    End Sub

    Private Sub initializeSite_JEDDAH()
        Dim siteName As String = SITE_JEDDAH

        AllProductionSites.Add(New ProdSite(siteName, "JED-MESDATABC2", SERVER_PW_V6, SERVER_UN_V6, "JDH"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Care", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("DIJE111", "DIJE111", siteName, ModuleID, 3, 8, 7, "DIJE111 CONVERTER", "DIJE111"))
        AllProdLines.Add(New ProdLine("DIJE112", "DIJE112", siteName, ModuleID, 3, 8, 7, "DIJE112 CONVERTER", "DIJE112"))
        AllProdLines.Add(New ProdLine("DIJE113", "DIJE113", siteName, ModuleID, 3, 8, 7, "DIJE113 CONVERTER", "DIJE113"))
        AllProdLines.Add(New ProdLine("DIJE114", "DIJE114", siteName, ModuleID, 3, 8, 7, "DIJE114 CONVERTER", "DIJE114"))
        AllProdLines.Add(New ProdLine("DIJE115", "DIJE115", siteName, ModuleID, 3, 8, 7, "DIJE115 CONVERTER", "DIJE115"))
        AllProdLines.Add(New ProdLine("DIJE116", "DIJE116", siteName, ModuleID, 3, 8, 7, "DIJE116 CONVERTER", "DIJE116"))
        AllProdLines.Add(New ProdLine("DIJE117", "DIJE117", siteName, ModuleID, 3, 8, 7, "DIJE117 CONVERTER", "DIJE117"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem Care", siteName, SECTOR_BABY, prStoryMapping.FemCare_Pads, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines
        AllProdLines.Add(New ProdLine("LCC1", "QAJE010", siteName, ModuleID2, 3, 8, 7, "QAJE010 CONVERTER", "QAJE010"))
        AllProdLines.Add(New ProdLine("LCC2", "QAJE011", siteName, ModuleID2, 3, 8, 7, "QAJE011 CONVERTER", "QAJE011"))
        AllProdLines.Add(New ProdLine("LCC3", "QAJE012", siteName, ModuleID2, 3, 8, 7, "QAJE012 CONVERTER", "QAJE012"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

    End Sub
    Private Sub initializeSite_BabyGLEDS()
        Dim siteName As String = "GLEDS"
        Dim tmpLineName As String
        AllProductionSites.Add(New ProdSite(siteName, "EUS-GLEDS001\eusp101", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "AKA"))
        ' Dim ModuleID As Guid = Guid.NewGuid()
        '  AllProductionModules.Add(New productionModule(ModuleID, SITE_AKASHI, "Baby", SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.GLEDS, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        'initialize the lines
        ' AllProductionLines.Add(New productionLine("DIAK-135", "DIAK-135", "Baby", ModuleID, 3, 8, 8, "DIAK-135", "DIAK-135"))


        Dim modName As String = "Jijona"

        '  siteName = "Jijona"

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, modName, siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.GLEDS, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        tmpLineName = "DIAB-101"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID2, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIAB-102"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID2, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIAB-103"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID2, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIAB-104"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID2, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIAB-105"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID2, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIAB-107"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID2, 3, 8, 8, tmpLineName, tmpLineName))

        '    tmpLineName = "QAAB-003"
        '    AllProductionLines.Add(New productionLine(tmpLineName, tmpLineName, sitename, moduleID2, 3, 8, 8, tmpLineName, tmpLineName))



        modName = "Cape"
        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, modName, siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.GLEDS, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        tmpLineName = "DICG-152"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-153"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-154"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-155"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-156"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-157"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-158"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-159"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-160"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-161"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-162"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-163"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-164"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-165"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-166"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-167"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-168"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DICG-169"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID3, 2, 12, 6, tmpLineName, tmpLineName))







        modName = "Euskirchen"
        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, modName, siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.GLEDS, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        tmpLineName = "DIEU-131"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-132"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-133"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-134"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-135"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-136"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-137"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-138"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-139"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-171"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-172"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-173"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-174"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-175"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-176"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-177"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-178"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))
        tmpLineName = "DIEU-179"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID4, 3, 8, 8, tmpLineName, tmpLineName))


        '    tmpSiteName = "Pescara"
        '    AllProductionModules.Add(New productionModule(tmpSiteName, modName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        '    tmpLineName = "DIFA-105"
        '    AllProductionLines.Add(New productionLine(tmpLineName, tmpLineName, modName, tmpSiteName, 2, 12, 6, tmpLineName, tmpLineName))
        '    With AllProductionLines(AllProductionLines.Count - 1)
        '    .ProficyServer_Name = "172.22.6.50"
        '    .ProficyServer_Password = PROFICY_SERVER_PASSWORD_QQ
        '    .ProficyServer_Username = PROFICY_SERVER_USERNAME_QQ
        '    End With




        modName = "Targowek"
        Dim ModuleID6 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID6, modName, siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.GLEDS, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        tmpLineName = "DITR-101"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-102"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-103"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-104"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-105"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-106"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-107"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-108"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-109"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-110"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-111"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-112"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-113"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-114"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-115"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))
        tmpLineName = "DITR-116"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, siteName, ModuleID6, 3, 8, 6.5, tmpLineName, tmpLineName))




    End Sub


#End Region



    Private Sub initializeSite_Hub()
        Dim siteName As String = "Hub"
        AllProductionSites.Add(New ProdSite(siteName, "hub-mesdatabe.eu.pg.com", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "OXN"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "S&PC", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String = "PC_PK1"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Main", ""))

        lineName = "PC_PK2"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Main", ""))

    End Sub


#Region "Beauty Care"
    Private Sub initializeSite_Singapore()
        AllProductionSites.Add(New ProdSite(SITE_SINGAPOREPIONEER, "sgts-mesdataphc.na.pg.com", SERVER_PW_QQ, PROFICY_SERVER_USERNAME_QQ, "SGP"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Pioneer Beauty", SITE_SINGAPOREPIONEER, SECTOR_BEAUTY, prStoryMapping.SingaporePioneer, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.SingaporePioneer))
        'initialize the lines
        AllProdLines.Add(New ProdLine("PUP1 Shubham1", "PUP1 Shubham1", SITE_SINGAPOREPIONEER, ModuleID, 2, 12, 6, "PUP1 Shubham1", "PUP1"))
        AllProdLines(AllProdLines.Count - 1).doIincludeAllUptime = True
        AllProdLines.Add(New ProdLine("PUP1 Shubham2", "PUP1 Shubham2", SITE_SINGAPOREPIONEER, ModuleID, 2, 12, 6, "PUP1 Shubham2", "PUP1"))
        AllProdLines(AllProdLines.Count - 1).doIincludeAllUptime = True
        AllProdLines.Add(New ProdLine("PUP2 Shubham3", "PUP2 Shubham3", SITE_SINGAPOREPIONEER, ModuleID, 2, 12, 6, "PUP2 Shubham3", "PUP2"))
        AllProdLines(AllProdLines.Count - 1).doIincludeAllUptime = True
        AllProdLines.Add(New ProdLine("PUP2 Shubham4", "PUP2 Shubham4", SITE_SINGAPOREPIONEER, ModuleID, 2, 12, 6, "PUP2 Shubham4", "PUP2"))
        AllProdLines(AllProdLines.Count - 1).doIincludeAllUptime = True
    End Sub

    Private Sub initializeSite_Mariscala()
        Dim siteName As String = SITE_MARISCALA
        AllProductionSites.Add(New ProdSite(siteName, "MARP-MESdataBE2.na.pg.com", "", SERVER_PW_V6, "PRStory", "MAR"))
        Dim ModuleID As Guid = Guid.NewGuid()

        AllProdModules.Add(New ProdModule(ModuleID, "Beauty", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Mariscala))

        'initialize the lines
        AllProdLines.Add(New ProdLine("3", "3", siteName, ModuleID, 2, 12, 6, "HCMR003 Main", "HCMR003"))
        AllProdLines.Add(New ProdLine("5", "5", siteName, ModuleID, 2, 12, 6, "HCMR005 Main", "HCMR005"))
        AllProdLines.Add(New ProdLine("9", "9", siteName, ModuleID, 2, 12, 6, "HCMR009 Main", "HCMR009"))
        AllProdLines.Add(New ProdLine("10", "10", siteName, ModuleID, 2, 12, 6, "HCMR010 Main", "HCMR010"))
        AllProdLines.Add(New ProdLine("11", "11", siteName, ModuleID, 2, 12, 6, "HCMR011 Main", "HCMR011"))
        AllProdLines.Add(New ProdLine("12", "12", siteName, ModuleID, 2, 12, 6, "HCMR012 Main", "HCMR012"))

        AllProdLines.Add(New ProdLine("35", "35", siteName, ModuleID, 2, 12, 6, "CPMR035 Main", "CPMR035"))
        AllProdLines.Add(New ProdLine("39", "39", siteName, ModuleID, 2, 12, 6, "CPMR039 Main", "CPMR039"))
        AllProdLines.Add(New ProdLine("43", "43", siteName, ModuleID, 2, 12, 6, "CPMR043 Main", "CPMR043"))

        Dim ModuleIDx As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleIDx, "Beauty", siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Mariscala))

        AllProdLines.Add(New ProdLine("HAMR 3 L1", "HAMR 3 L1", siteName, ModuleIDx, 2, 12, 6, "HAMR003 Llenadora 1", "HAMR003 Llenadora 1 Production"))
        AllProdLines.Add(New ProdLine("HAMR 3 L2", "HAMR 3 L2", siteName, ModuleIDx, 2, 12, 6, "HAMR003 Llenadora 2", "HAMR003 Llenadora 2 Production"))
        AllProdLines.Add(New ProdLine("HAMR 3 L3", "HAMR 3 L3", siteName, ModuleIDx, 2, 12, 6, "HAMR003 Llenadora 3", "HAMR003 Llenadora 3 Production"))

        AllProdLines.Add(New ProdLine("HAMR 7 L1", "HAMR 7 L1", siteName, ModuleIDx, 2, 12, 6, "HAMR007 Llenadora 1", "HAMR007"))
        AllProdLines.Add(New ProdLine("HAMR 7 L2", "HAMR 7 L2", siteName, ModuleIDx, 2, 12, 6, "HAMR007 Llenadora 2", "HAMR008"))
        AllProdLines.Add(New ProdLine("HAMR 7 L3", "HAMR 7 L3", siteName, ModuleIDx, 2, 12, 6, "HAMR007 Llenadora 3", "HAMR009"))

        AllProdLines.Add(New ProdLine("HAMR 10", "HAMR 10", siteName, ModuleIDx, 2, 12, 6, "HAMR010 Main", "HAMR010"))
    End Sub


    Private Sub initializeSite_IowaCity()
        AllProductionSites.Add(New ProdSite(SITE_IOWA_CITY, "IC-MESDATABE", "", SERVER_PW_V6, SERVER_UN_V6, "IAC"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Oral Rinse", SITE_IOWA_CITY, SECTOR_HEALTH, prStoryMapping.IowaCity, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        AllProdModules(AllProdModules.Count - 1).MultilineGroup = "Oral Rinse"


        'initialize the lines
        AllProdLines.Add(New ProdLine("13", "13", SITE_IOWA_CITY, ModuleID, 2, 12, 6, "ICT L13 Main", "ICT L13"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True
        AllProdLines.Add(New ProdLine("14", "14", SITE_IOWA_CITY, ModuleID, 2, 12, 6, "ICU L14 Main", "ICU L14"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True
        AllProdLines.Add(New ProdLine("17", "17", SITE_IOWA_CITY, ModuleID, 2, 12, 6, "ICR Main", "ICR"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True
        AllProdLines(AllProdLines.Count - 1).MultilineGroupName = "A"


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Beauty", SITE_IOWA_CITY, SECTOR_BEAUTY, prStoryMapping.GENERIC, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("Line 2", "2", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICB L2 Main", "ICB L2"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True
        AllProdLines(AllProdLines.Count - 1).MultilineGroupName = "A"

        AllProdLines.Add(New ProdLine("Line 3", "3", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICC L3 Main", "ICC L3"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True
        AllProdLines(AllProdLines.Count - 1).MultilineGroupName = "B"

        '   AllProductionLines.Add(New productionLine("Line 4", "4", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICC L4 Main", "ICC L4"))
        '   AllProductionLines(AllProductionLines.Count - 1).FieldCheck_ProductGroup = True
        '   AllProductionLines(AllProductionLines.Count - 1).MultilineGroupName = "C"
        AllProdLines.Add(New ProdLine("Line 6", "6", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICH L6 Main", "ICH L6"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True
        AllProdLines.Add(New ProdLine("Line 8", "8", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICK L8 Main", "ICK L8"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True

        AllProdLines.Add(New ProdLine("Line 10", "10", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICN L10 Main", "ICN L10"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True


        AllProdLines.Add(New ProdLine("Line 16", "16", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICG L16 Main", "ICG L16"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True
        AllProdLines.Add(New ProdLine("Line 17", "17", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICR Main", "ICR"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True

        AllProdLines.Add(New ProdLine("Line 19", "19", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICJ L19 Main", "ICJ L19"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True


        AllProdLines.Add(New ProdLine("Line 20", "20", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICL L20 Main", "ICL L20"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True

        AllProdLines.Add(New ProdLine("Line 21", "21", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICV Main", "ICV"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True

        AllProdLines.Add(New ProdLine("Line 22", "22", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICW Main", "ICW"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True

        AllProdLines.Add(New ProdLine("Line 23A", "23A", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICY Filler A", "ICY"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True

        AllProdLines.Add(New ProdLine("Line 23B", "23B", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICY Filler B", "ICY"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True

        AllProdLines.Add(New ProdLine("Line 24", "24", SITE_IOWA_CITY, ModuleID2, 2, 12, 6, "ICA L1 Main", "ICA L1"))
        AllProdLines(AllProdLines.Count - 1).FieldCheck_ProductGroup = True


    End Sub

    Private Sub initializeSite_IowaCity_OC()
        AllProductionSites.Add(New ProdSite(SITE_IOWA_CITY_ORALCARE, "ICIA-MESDATAPHC", "", SERVER_PW_V6, SERVER_UN_V6, "ICO"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Packing", SITE_IOWA_CITY_ORALCARE, SECTOR_HEALTH, prStoryMapping.ICOC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'get those fixed fields!
        '*ATS
        'Blister machine
        '*brush transfer
        'Cartoner
        '*Case Packer
        '*Labeler
        '*Pick & Place 1
        '*Pick & Place 2
        '*Rollover
        '*Off Quality Matl
        'Other
        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"ATS", "Blister Machine",
            "Brush Transfer", "Cartoner", "Labeler", "Pick & Place", "Rollover", "Material", "Tray", "Case Packer", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"AM", "Changeover", "CILs",
            "Training", "SU/SD", "Maintenance", "Material", OTHERS_STRING}

        'initialize the lines
        Dim cName As String
        Dim shiftStartTime As Double = 7.25

        cName = "37-44"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-45"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-54"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-55"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-56"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Blister Machine", "ICOC " & cName))
        cName = "37-57"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-59"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-60"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-61"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-80"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "37-90"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "Line 1"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "Line 2"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        cName = "Line 4"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "PP01"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))


        'ICOC 291 Line
        'ICOC 291 Making
        'ICOC 285 Line
        'ICOC 285 Making
        'ICOC 287 Line
        'ICOC 287 Making
        'ICOC 288 Line
        'ICOC 286 Line

        cName = "291"
        AllProdLines.Add(New ProdLine(cName & " Line", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        AllProdLines.Add(New ProdLine(cName & " Making", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Making", "ICOC " & cName))

        cName = "285"
        AllProdLines.Add(New ProdLine(cName & " Line", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        AllProdLines.Add(New ProdLine(cName & " Making", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Making", "ICOC " & cName))

        cName = "287"
        AllProdLines.Add(New ProdLine(cName & " Line", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))
        AllProdLines.Add(New ProdLine(cName & " Making", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Making", "ICOC " & cName))

        cName = "288"
        AllProdLines.Add(New ProdLine(cName & " Line", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "286"
        AllProdLines.Add(New ProdLine(cName & " Line", cName, SITE_IOWA_CITY_ORALCARE, ModuleID, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))



        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Making", SITE_IOWA_CITY_ORALCARE, SECTOR_HEALTH, prStoryMapping.ICOC_Making, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        cName = "290"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "292"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "294"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "295"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "296"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "297"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "298"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "299"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "301"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))


        cName = "304"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "305"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))


        cName = "ER01"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "ER02"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "ER03"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "ER04"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "ER05"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "ER06"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "ER07"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))

        cName = "ER08"
        AllProdLines.Add(New ProdLine(cName, cName, SITE_IOWA_CITY_ORALCARE, ModuleID2, 2, 12, shiftStartTime, "ICOC " & cName & " Line", "ICOC " & cName))










    End Sub


    Private Sub initializeTEMP()
        Dim siteName As String = SITE_BROWNS_SUMMIT
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, BS_APDO, siteName, SECTOR_BEAUTY, prStoryMapping.APDO_I, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason3, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("I", "BSI", siteName, ModuleID, 2, 12, 6, "BS I Line Filler", ""))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, BS_APDO & " ", siteName, SECTOR_BEAUTY, prStoryMapping.APDO_J, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Machine", "Machine Section", "Stop Code", "Root Cause", DowntimeField.Reason3, -1, DTsched_Mapping.Greensboro))
        AllProdLines.Add(New ProdLine("J", "BSJ", siteName, ModuleID2, 2, 12, 6, "BS APDO J Line Main", "BS APDO J Line"))

        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, BS_APDO, siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason3, -1, DTsched_Mapping.Greensboro))

        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Case Code Dater", "Case Erector", "Casepacker (Douglas)", "EOL", "Excluded Time", "Material Quality", "Planned downtime", "Plant Systems & Others", "Supply Losses", "Twins Heat Tunnel", "Twins Sleever", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Case Code Dater", "Case Erector", "Casepacker (Douglas)", "EOL", "Excluded Time", "Material Quality", "Planned downtime", "Plant Systems & Others", "Supply Losses", "Twins Heat Tunnel", "Twins Sleever", OTHERS_STRING}

        AllProdLines.Add(New ProdLine("Twins 1", "BSI", siteName, ModuleID3, 2, 12, 6, "BS APDO Twins 1 Main", "BS APDO Twins 1"))
        AllProdLines.Add(New ProdLine("Twins 2", "BSI", siteName, ModuleID3, 2, 12, 6, "BS APDO Twins 2 Main", "BS APDO Twins 2"))


        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, BS_APDO, siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason3, -1, DTsched_Mapping.Greensboro))

        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Accumulator", "Bundler", "Can Bulk Hopper", "Capsticker", "Case Packer", "Cooling Tunnel", "EOL", "Excluded Time", "Filler", "Labeler", "Material Quality", "Packing Material Supply", "Planned downtime", "Plant Systems & Others", "Plugger/Depucker", "Product Quality", "Product Supply", "Puck Sorter", "Robotic Sorter", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Accumulator", "Bundler", "Can Bulk Hopper", "Capsticker", "Case Packer", "Cooling Tunnel", "EOL", "Excluded Time", "Filler", "Labeler", "Material Quality", "Packing Material Supply", "Planned downtime", "Plant Systems & Others", "Plugger/Depucker", "Product Quality", "Product Supply", "Puck Sorter", "Robotic Sorter", OTHERS_STRING}

        AllProdLines.Add(New ProdLine("W Front", "BSW1", siteName, ModuleID3, 2, 12, 6, "BS APDO W Line Main", "BS APDO W Line Front"))
        AllProdLines.Add(New ProdLine("W Back", "BSW2", siteName, ModuleID3, 2, 12, 6, "BS APDO W Line Plugger/DePucker", "BS APDO W Line Back"))


        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, BS_APDO, siteName, SECTOR_BEAUTY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.QuickQuery, DefaultProficyProductionProcedure.QuickQuery, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason3, -1, DTsched_Mapping.Greensboro))


        AllProdLines.Add(New ProdLine("Line 6", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO Line 6 Main", "BS APDO Line 6"))
        AllProdLines.Add(New ProdLine("G Line", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO G Line Main", "BS APDO G Line"))

        AllProdLines.Add(New ProdLine("I Line Back", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO I Line Trimmer", "BS APDO I Line Back"))
        AllProdLines.Add(New ProdLine("I Line Front", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO I Line Main", "BS APDO I Line Front"))
        AllProdLines.Add(New ProdLine("I Line Cannister Supply", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO I Line Labeler", "BS APDO I Line Cannister Supply"))

        AllProdLines.Add(New ProdLine("V Line Back", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO V Line Plugger/Depucker", "BS APDO V Line Back"))
        AllProdLines.Add(New ProdLine("V Line Canister Supply", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO V Line Labeler", "BS APDO V Line Cannister Supply"))
        AllProdLines.Add(New ProdLine("V Line Front", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO V Line Main", "BS APDO V Line Front"))

        AllProdLines.Add(New ProdLine("K Line Front", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO K Line Main", "BS APDO K Line Front"))
        AllProdLines.Add(New ProdLine("K Line Back", "BSW1", siteName, ModuleID5, 2, 12, 6, "BS APDO K Line DePucker", "BS APDO K Line Back"))


        AllProdLines.Add(New ProdLine("N Line Front", "", siteName, ModuleID5, 2, 12, 6, "BS APDO N Line Main", "BS APDO N Line Front"))
        AllProdLines.Add(New ProdLine("N Line Back", "", siteName, ModuleID5, 2, 12, 6, "BS APDO N Line Closing Section", "BS APDO N Line Back"))

        AllProdLines.Add(New ProdLine("L Line Front", "", siteName, ModuleID5, 2, 12, 6, "BS APDO L Line Main", "BS APDO L Line Front"))
        AllProdLines.Add(New ProdLine("L Line Back", "", siteName, ModuleID5, 2, 12, 6, "BS APDO L Line Closing Section", "BS APDO L Line Back"))

    End Sub
    Private s_aditionalEntropy As Byte() = {9, 8, 7, 6, 5}

#End Region
#Region "Fem Care"
    Private Sub initializeSite_Borysil()
        AllProductionSites.Add(New ProdSite(SITE_BORYSPIL, "BOR-MESDATABC.na.pg.com", SERVER_PW_V6, SERVER_UN_V6, "BRY"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fem Care", SITE_BORYSPIL, SECTOR_FEM, prStoryMapping.Boryspil, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Ukraine))
        'initialize the lines
        AllProdLines.Add(New ProdLine("QBBS001", "QBBS001", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QBBS001 Converter", "QBBS001"))
        AllProdLines.Add(New ProdLine("QBBS002", "QBBS002", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QBBS002 Converter", "QBBS002"))
        AllProdLines.Add(New ProdLine("QBBS003", "QBBS003", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QBBS003 Converter", "QBBS003"))
        AllProdLines.Add(New ProdLine("QBBS004", "QBBS004", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QBBS004 Converter", "QBBS004"))
        AllProdLines.Add(New ProdLine("QBBS005", "QBBS005", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QBBS005 Converter", "QBBS005"))
        AllProdLines.Add(New ProdLine("QBBS006", "QBBS006", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QBBS006 Converter", "QBBS006"))
        AllProdLines.Add(New ProdLine("QBBS007", "QBBS007", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QBBS007 Converter", "QBBS007"))
        AllProdLines.Add(New ProdLine("QABS011", "QABS0", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QABS011 Converter", "QABS011"))
        AllProdLines.Add(New ProdLine("QABS012", "QABS012", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QABS012 Converter", "QABS012"))
        AllProdLines.Add(New ProdLine("QABS013", "QABS013", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QABS013 Converter", "QABS013"))
        AllProdLines.Add(New ProdLine("QABS014", "QABS014", SITE_BORYSPIL, ModuleID, 2, 12, 6, "QABS014 Converter", "QABS014"))
    End Sub

    Private Function u(ByVal data() As Byte) As Byte()
        Try
            'Decrypt the data using DataProtectionScope.CurrentUser.
            Return ProtectedData.Unprotect(data, s_aditionalEntropy, DataProtectionScope.CurrentUser)
        Catch e As CryptographicException
            Console.WriteLine(e.ToString())
            Return Nothing
        End Try

    End Function

    '   Private Shared Function ByteArrayToString(ba As Byte()) As String
    '       Dim hex As String = BitConverter.ToString(ba)
    '       Return hex.Replace("-", "")
    '   End Function


    Private Function s(hex As [String]) As Byte()
        Dim NumberChars As Integer = hex.Length
        Dim bytes As Byte() = New Byte(NumberChars / 2 - 1) {}
        For i As Integer = 0 To NumberChars - 1 Step 2
            bytes(i / 2) = Convert.ToByte(hex.Substring(i, 2), 16)
        Next
        Return bytes
    End Function

    Private Sub initializeSite_Budapest()
        Dim siteName As String = SITE_BUDAPEST

        AllProductionSites.Add(New ProdSite(siteName, "hyg-mesdatabc.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "BDA"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "FGC", siteName, SECTOR_FEM, prStoryMapping.BudapestFGC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Budapest))
        'initialize the lines
        AllProdLines.Add(New ProdLine("1", "QAHYB0", siteName, ModuleID, 2, 12, 6, "QAHYB0 Converter", ""))
        AllProdLines.Add(New ProdLine("2", "QAHYC0", siteName, ModuleID, 2, 12, 6, "QAHYC0 Converter", ""))
        AllProdLines.Add(New ProdLine("3", "QAHYD0", siteName, ModuleID, 2, 12, 6, "QAHYD0 Converter", ""))
        AllProdLines.Add(New ProdLine("4", "QAHYE0", siteName, ModuleID, 2, 12, 6, "QAHYE0 Converter", ""))
        AllProdLines.Add(New ProdLine("5", "QAHYF0", siteName, ModuleID, 2, 12, 6, "QAHYF0 Converter", ""))
        AllProdLines.Add(New ProdLine("6", "QAHY016", siteName, ModuleID, 2, 12, 6, "QAHY016 Converter", ""))
        AllProdLines.Add(New ProdLine("7", "QAHY017", siteName, ModuleID, 2, 12, 6, "QAHY017 Converter", ""))
        AllProdLines.Add(New ProdLine("8", "QAHY018", siteName, ModuleID, 2, 12, 6, "QAHY018 Converter", ""))
        AllProdLines.Add(New ProdLine("9", "QAHY019", siteName, ModuleID, 2, 12, 6, "QAHY018 Converter", ""))
        AllProdLines.Add(New ProdLine("10", "QAHY026", siteName, ModuleID, 2, 12, 6, "QAHY026 Converter", ""))


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "LCC", siteName, SECTOR_FEM, prStoryMapping.BudapestLCC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Budapest))
        AllProdLines.Add(New ProdLine("1", "QAHY031", siteName, ModuleID2, 2, 12, 6, "QAHY031 Converter", ""))
        AllProdLines.Add(New ProdLine("2", "QAHY032", siteName, ModuleID2, 2, 12, 6, "QAHY032 Converter", ""))
        AllProdLines.Add(New ProdLine("3", "QAHY033", siteName, ModuleID2, 2, 12, 6, "QAHY033 Converter", ""))
        AllProdLines.Add(New ProdLine("4", "QAHY034", siteName, ModuleID2, 2, 12, 6, "QAHY034 Converter", ""))
        AllProdLines.Add(New ProdLine("5", "QAHY035", siteName, ModuleID2, 2, 12, 6, "QAHY035 Converter", ""))

        Dim lineName As String
        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "FPD", siteName, SECTOR_FEM, prStoryMapping.BudapestLCC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Budapest))

        lineName = "QRHY508"
        AllProdLines.Add(New ProdLine("LINE8", lineName, siteName, ModuleID3, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QRHY509"
        AllProdLines.Add(New ProdLine("LINE9", lineName, siteName, ModuleID3, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QRHY510"
        AllProdLines.Add(New ProdLine("LINE10", lineName, siteName, ModuleID3, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QRHY511"
        AllProdLines.Add(New ProdLine("LINE7", lineName, siteName, ModuleID3, 2, 12, 6, lineName & " Converter", ""))

        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "Pearl", siteName, SECTOR_FEM, prStoryMapping.BudapestLCC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Budapest))

        lineName = "QXHY030"
        AllProdLines.Add(New ProdLine("Pearl1 / V2 line", lineName, siteName, ModuleID4, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QXHY031"
        AllProdLines.Add(New ProdLine("Pearl2 / V3 line", lineName, siteName, ModuleID4, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QXHY032"
        AllProdLines.Add(New ProdLine("Pearl3 / V1 line", lineName, siteName, ModuleID4, 2, 12, 6, lineName & " Converter", ""))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Ivory", siteName, SECTOR_FEM, prStoryMapping.BudapestLCC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Budapest))

        lineName = "QAHY041 P1"
        AllProdLines.Add(New ProdLine("Ivory 1", lineName, siteName, ModuleID5, 2, 12, 6, lineName & " Converter", ""))

        lineName = "QAHY041 P2"
        AllProdLines.Add(New ProdLine("Ivory 2", lineName, siteName, ModuleID5, 2, 12, 6, lineName & " Converter", ""))



        Dim ModuleID6 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID6, "Always Ultra", siteName, SECTOR_FEM, prStoryMapping.BudapestLCC, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Budapest))

        lineName = "QAHYB0"
        AllProdLines.Add(New ProdLine("Always/FCG1 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHYC0"
        AllProdLines.Add(New ProdLine("Always/FCG2 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHYD0"
        AllProdLines.Add(New ProdLine("Always/FCG3 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHYE0"
        AllProdLines.Add(New ProdLine("Always/FCG4 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHYF0"
        AllProdLines.Add(New ProdLine("Always/FCG5 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))

        lineName = "QAHY016"
        AllProdLines.Add(New ProdLine("Always/FCG6 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHY017"
        AllProdLines.Add(New ProdLine("Always/FCG7 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHY018"
        AllProdLines.Add(New ProdLine("Always/FCG8 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHY019"
        AllProdLines.Add(New ProdLine("Always/FCG9 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QAHY026"
        AllProdLines.Add(New ProdLine("Always/FCG10 line", lineName, siteName, ModuleID6, 2, 12, 6, lineName & " Converter", ""))


        Dim ModuleID7 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID7, "Liners", siteName, SECTOR_FEM, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Budapest))


        lineName = "QBHYSA"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID7, 2, 12, 6, lineName & " Converter", ""))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
        lineName = "QBSYSB"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID7, 2, 12, 6, lineName & " Converter", ""))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

    End Sub
    Private Sub initializeSite_Belleville()
        Dim siteName As String = SITE_BELLEVILLE
        Dim lineName As String
        AllProductionSites.Add(New ProdSite(siteName, "BELL-MESDATABC2", "", SERVER_PW_V6, SERVER_UN_V6, "BEL"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fem Pads", siteName, SECTOR_FEM, prStoryMapping.FemCare_Pads, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville))
        'initialize the lines
        AllProdLines.Add(New ProdLine("QABV033", "QABV033", siteName, ModuleID, 2, 12, 6, "QABV033 Converter", "QABV033"))
        AllProdLines.Add(New ProdLine("QABV034", "QABV034", siteName, ModuleID, 2, 12, 6, "QABV034 Converter", "QABV034"))
        AllProdLines.Add(New ProdLine("QABV042", "QABV042", siteName, ModuleID, 2, 12, 6, "QABV042 Converter", "QABV042"))
        AllProdLines.Add(New ProdLine("QABV032", "QABV032", siteName, ModuleID, 2, 12, 6, "QABV032 Converter", "QABV032"))
        AllProdLines.Add(New ProdLine("QABV062", "QABV062", siteName, ModuleID, 2, 12, 6, "QABV062 Converter", "QABV062"))

        AllProdLines.Add(New ProdLine("PEBV058", "PEBV058", siteName, ModuleID, 2, 12, 6, "PEBV058 Converter", "PEBV058"))
        AllProdLines.Add(New ProdLine("PEBV060", "PEBV060", siteName, ModuleID, 2, 12, 6, "PEBV060 Converter", "PEBV060"))
        AllProdLines.Add(New ProdLine("PEBV080", "PEBV080", siteName, ModuleID, 2, 12, 6, "PEBV080 Converter", "PEBV080"))
        '   AllProductionLines.Add(New productionLine("PEBV081", "PEBV080", siteName, ModuleID, 2, 12, 6, "PEBV081 Converter", "PEBV081"))
        '   allproductionlines(allproductionlines.count - 1).isstartupmode = true
        AllProdLines.Add(New ProdLine("QABV038", "QABV038", siteName, ModuleID, 2, 12, 6, "QABV038 Converter", "QABV038"))
        AllProdLines.Add(New ProdLine("QABV039", "QABV039", siteName, ModuleID, 2, 12, 6, "QABV039 Converter", "QABV039"))
        AllProdLines.Add(New ProdLine("QABV043", "QABV043", siteName, ModuleID, 2, 12, 6, "QABV043 Converter", "QABV043"))
        AllProdLines.Add(New ProdLine("QABV063", "QABV063", siteName, ModuleID, 2, 12, 6, "QABV063 Converter", "QABV063"))
        AllProdLines.Add(New ProdLine("QABV064", "QABV064", siteName, ModuleID, 2, 12, 6, "QABV064 Converter", "QABV064"))
        AllProdLines.Add(New ProdLine("QABV065", "QABV065", siteName, ModuleID, 2, 12, 6, "QABV065 Converter", "QABV065"))
        AllProdLines.Add(New ProdLine("QABV066", "QABV066", siteName, ModuleID, 2, 12, 6, "QABV066 Converter", "QABV066"))
        AllProdLines.Add(New ProdLine("QABV067", "QABV067", siteName, ModuleID, 2, 12, 6, "QABV067 Converter", "QABV067"))
        AllProdLines.Add(New ProdLine("QABV068", "QABV068", siteName, ModuleID, 2, 12, 6, "QABV068 Converter", "QABV068"))
        AllProdLines.Add(New ProdLine("QBBV074", "QBBV074", siteName, ModuleID, 2, 12, 6, "QBBV074 Converter", "QBBV074"))
        AllProdLines.Add(New ProdLine("QBBV075", "QBBV075", siteName, ModuleID, 2, 12, 6, "QBBV075 Converter", "QBBV075"))
        AllProdLines.Add(New ProdLine("QBBV076", "QBBV076", siteName, ModuleID, 2, 12, 6, "QBBV076 Converter", "QBBV076"))
        AllProdLines.Add(New ProdLine("QBBV077", "QBBV077", siteName, ModuleID, 2, 12, 6, "QBBV077 Converter", "QBBV077"))
        AllProdLines.Add(New ProdLine("QBBV078", "QBBV078", siteName, ModuleID, 2, 12, 6, "QBBV078 Converter", "QBBV078"))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem ", SITE_BELLEVILLE, SECTOR_FEM, prStoryMapping.Fem_LCC_HPU, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville))

        lineName = "PEBV081"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 6, lineName & " Converter", ""))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "Emulsion Making", SITE_BELLEVILLE, SECTOR_FEM, prStoryMapping.STRAIGHTPlusOne, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville))

        lineName = "BV21"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID3, 2, 12, 6, lineName & " Production", ""))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
    End Sub

    Private Sub initializeSite_Belleville_Test()
        Dim siteName As String = SITE_BELLEVILLE & " Test"
        AllProductionSites.Add(New ProdSite(siteName, "BELL-MESTRNBC", "", SERVER_PW_V6, SERVER_UN_V6, "BLV"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fem Pads", siteName, SECTOR_FEM, prStoryMapping.FemCare_Pads, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Belleville))
        'initialize the lines
        AllProdLines.Add(New ProdLine("QABV033", "QABV033", siteName, ModuleID, 2, 12, 6, "QABV033 Converter", "QABV033"))
        AllProdLines.Add(New ProdLine("QABV034", "QABV034", siteName, ModuleID, 2, 12, 6, "QABV034 Converter", "QABV034"))
        AllProdLines.Add(New ProdLine("QABV042", "QABV042", siteName, ModuleID, 2, 12, 6, "QABV042 Converter", "QABV042"))
        AllProdLines.Add(New ProdLine("QABV032", "QABV032", siteName, ModuleID, 2, 12, 6, "QABV032 Converter", "QABV032"))
        AllProdLines.Add(New ProdLine("QABV062", "QABV062", siteName, ModuleID, 2, 12, 6, "QABV062 Converter", "QABV062"))

        AllProdLines.Add(New ProdLine("PEBV058", "PEBV058", siteName, ModuleID, 2, 12, 6, "PEBV058 Converter", "PEBV058"))
        AllProdLines.Add(New ProdLine("PEBV060", "PEBV060", siteName, ModuleID, 2, 12, 6, "PEBV060 Converter", "PEBV060"))
        AllProdLines.Add(New ProdLine("PEBV080", "PEBV080", siteName, ModuleID, 2, 12, 6, "PEBV080 Converter", "PEBV080"))
        AllProdLines.Add(New ProdLine("QABV038", "QABV038", siteName, ModuleID, 2, 12, 6, "QABV038 Converter", "QABV038"))
        AllProdLines.Add(New ProdLine("QABV039", "QABV039", siteName, ModuleID, 2, 12, 6, "QABV039 Converter", "QABV039"))
        AllProdLines.Add(New ProdLine("QABV043", "QABV043", siteName, ModuleID, 2, 12, 6, "QABV043 Converter", "QABV043"))
        AllProdLines.Add(New ProdLine("QABV063", "QABV063", siteName, ModuleID, 2, 12, 6, "QABV063 Converter", "QABV063"))
        AllProdLines.Add(New ProdLine("QABV064", "QABV064", siteName, ModuleID, 2, 12, 6, "QABV064 Converter", "QABV064"))
        AllProdLines.Add(New ProdLine("QABV065", "QABV065", siteName, ModuleID, 2, 12, 6, "QABV065 Converter", "QABV065"))
        AllProdLines.Add(New ProdLine("QABV066", "QABV066", siteName, ModuleID, 2, 12, 6, "QABV066 Converter", "QABV066"))
        AllProdLines.Add(New ProdLine("QABV067", "QABV067", siteName, ModuleID, 2, 12, 6, "QABV067 Converter", "QABV067"))
        AllProdLines.Add(New ProdLine("QABV068", "QABV068", siteName, ModuleID, 2, 12, 6, "QABV068 Converter", "QABV068"))
        AllProdLines.Add(New ProdLine("QBBV075", "QBBV075", siteName, ModuleID, 2, 12, 6, "QBBV075 Converter", "QBBV075"))
        AllProdLines.Add(New ProdLine("QBBV076", "QBBV076", siteName, ModuleID, 2, 12, 6, "QBBV076 Converter", "QBBV076"))
        AllProdLines.Add(New ProdLine("QBBV077", "QBBV077", siteName, ModuleID, 2, 12, 6, "QBBV077 Converter", "QBBV077"))
        AllProdLines.Add(New ProdLine("QBBV078", "QBBV078", siteName, ModuleID, 2, 12, 6, "QBBV078 Converter", "QBBV078"))


    End Sub


#End Region
#Region "Family Care"

    Private Sub initializeSite_Albany()
        Dim siteName As String = SITE_ALBANY
        AllProductionSites.Add(New ProdSite(siteName, "ay-mesdata002.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "ALB"))
        Dim ModuleIDx As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleIDx, "Fam", siteName, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        'initialize the lines

        Dim lineName As String = "AK05"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AK06"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AK08"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AK09"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AC1"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AC2"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AC3"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AC4"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))

        '  AllProductionLines.Add(New productionLine("AK05", "AK05", siteName, ModuleIDx, 3, 8, 7.5, "AK05 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AK06", "AK06", siteName, ModuleIDx, 2, 12, 7.5, "AK06 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AK08", "AK08", siteName, ModuleIDx, 2, 12, 7.5, "AK08 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AK09", "AK09", siteName, ModuleIDx, 2, 12, 7.5, "AK09 Converter Reliability", ""))
        '      AllProductionLines.Add(New productionLine("AC1", "AC1", siteName, ModuleIDx, 2, 12, 7.5, "AC1 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AC2", "AC2", siteName, ModuleIDx, 2, 12, 7.5, "AC2 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AC3", "AC3", siteName, ModuleIDx, 2, 12, 7.5, "AC3 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AC4", "AC4", siteName, ModuleIDx, 2, 12, 7.5, "AC4 Converter Reliability", ""))

        lineName = "AT10"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AT13"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AT14"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "AT16"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDx, 3, 8, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))



        '       AllProductionLines.Add(New productionLine("AT10", "AT10", siteName, ModuleIDx, 2, 12, 7.5, "AT10 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AT13", "AT13", siteName, ModuleIDx, 2, 12, 7.5, "AT13 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AT14", "AT14", siteName, ModuleIDx, 2, 12, 7.5, "AT14 Converter Reliability", ""))
        '       AllProductionLines.Add(New productionLine("AT16", "AT16", siteName, ModuleIDx, 2, 12, 7.5, "AT16 Converter Reliability", ""))

        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fam", siteName, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        'initialize the lines
        'MF
        AllProdLines.Add(New ProdLine("AK05 MultiFlow", "AK05MF", siteName, ModuleID, 2, 12, 7.5, "AK05 MF5 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AK05 MF5 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK06 MultiFlow", "AK06MF", siteName, ModuleID, 2, 12, 7.5, "AK06 MF5 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AK06 MF5 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK08 MultiFlow", "AK08MF", siteName, ModuleID, 2, 12, 7.5, "AK08 MF5 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AK08 MF5 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK09 MultiFlow", "AK09MF", siteName, ModuleID, 2, 12, 7.5, "AK09 MF5 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AK09 MF5 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC1 Bundler Ext", "AC1BundlerExt", siteName, ModuleID, 2, 12, 7.5, "AC1 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC1 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC1 Bundler Int", "AC1BundlerInt", siteName, ModuleID, 2, 12, 7.5, "AC1 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC1 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC2 Bundler Ext", "AC2BundlerExt", siteName, ModuleID, 2, 12, 7.5, "AC2 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC2 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC2 Bundler Int", "AC2BundlerInt", siteName, ModuleID, 2, 12, 7.5, "AC2 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC2 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC3 Bundler Ext", "AC3BundlerExt", siteName, ModuleID, 2, 12, 7.5, "AC3 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC3 Bundler External Blocked/Starved"))   ' Both One Click and PRSTORY wont work
        AllProdLines.Add(New ProdLine("AC3 Bundler Int", "AC3BundlerInt", siteName, ModuleID, 2, 12, 7.5, "AC3 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC3 Bundler Internal Blocked/Starved"))   ' Both One Click and PRSTORY wont work
        AllProdLines.Add(New ProdLine("AC4 Bundler Ext", "AC4BundlerExt", siteName, ModuleID, 2, 12, 7.5, "AC4 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC4 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC4 Bundler Int", "AC4BundlerInt", siteName, ModuleID, 2, 12, 7.5, "AC4 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AC4 Bundler Internal Blocked/Starved"))

        AllProdLines.Add(New ProdLine("AT10 MultiFlow", "AT10MF", siteName, ModuleID, 2, 12, 7.5, "MFE Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MFE Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT14 MultiFlow", "AT14MF", siteName, ModuleID, 2, 12, 7.5, "AT15 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AT15 MF Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT13 MultiFlow", "AT13MF", siteName, ModuleID, 2, 12, 7.5, "AT13 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AT13 MF Blocked/Starved"))


        'Wrapper
        AllProdLines.Add(New ProdLine("AK05 Wrapper", "AK05Wrapper", siteName, ModuleID, 2, 12, 7.5, "AK05 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AK05 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK06 Wrapper", "AK06Wrapper", siteName, ModuleID, 2, 12, 7.5, "AK06 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AK06 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK08 Wrapper", "AK08Wrapper", siteName, ModuleID, 2, 12, 7.5, "AK08 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AK08 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK09 Wrapper", "AK09Wrapper", siteName, ModuleID, 2, 12, 7.5, "AK09 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AK09 Wrapper Blocked/Starved"))

        AllProdLines.Add(New ProdLine("AC1 Wrapper Ext", "AC1WrapperExt", siteName, ModuleID, 2, 12, 7.5, "AC1 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC1 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC1 Wrapper Int", "AC1WrapperInt", siteName, ModuleID, 2, 12, 7.5, "AC1 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC1 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC2 Wrapper Ext", "AC2WrapperExt", siteName, ModuleID, 2, 12, 7.5, "AC2 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC2 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC2 Wrapper Int", "AC2WrapperInt", siteName, ModuleID, 2, 12, 7.5, "AC2 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC2 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC3 Wrapper Ext", "AC3WrapperExt", siteName, ModuleID, 2, 12, 7.5, "AC3 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC3 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC3 Wrapper Int", "AC3WrapperInt", siteName, ModuleID, 2, 12, 7.5, "AC3 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC3 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC4 Wrapper Ext", "AC4WrapperExt", siteName, ModuleID, 2, 12, 7.5, "AC4 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC4 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC4 Wrapper Int", "AC4WrapperInt", siteName, ModuleID, 2, 12, 7.5, "AC4 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AC4 Wrapper Internal Blocked/Starved"))

        AllProdLines.Add(New ProdLine("AT10 Wrapper", "AT10Wrapper", siteName, ModuleID, 2, 12, 7.5, "AT10 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AT10 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT13 Wrapper", "AT13Wrapper", siteName, ModuleID, 2, 12, 7.5, "AT13 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AT13 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT14 Wrapper", "AT14Wrapper", siteName, ModuleID, 2, 12, 7.5, "AT14 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AT14 Wrapper Blocked/Starved"))


        'AT16
        AllProdLines.Add(New ProdLine("AT16 ACP", "AT16 ACP", siteName, ModuleID, 2, 12, 7.5, "AT16 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AT16 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT16 LCP", "AT16 LCP", siteName, ModuleID, 2, 12, 7.5, "AT16 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "AT16 MF Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT16 Internal Wrapper", "AT16 Internal Wrapper", siteName, ModuleID, 2, 12, 7.5, "AT16 Internal Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AT16 Internal Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT16 External Wrapper", "AT16 External Wrapper", siteName, ModuleID, 2, 12, 7.5, "AT16 External Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AT16 External Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT16 Internal Logsaw", "AT16 Internal Logsaw", siteName, ModuleID, 2, 12, 7.5, "AT16 Internal Logsaw Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AT16 Internal Logsaw Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT16 External Logsaw", "AT16 External Logsaw", siteName, ModuleID, 2, 12, 7.5, "AT16 External Logsaw Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "AT16 External Logsaw Blocked/Starved"))



        'ACP
        AllProdLines.Add(New ProdLine("AK05 ACP", "AK05ACP", siteName, ModuleID, 2, 12, 7.5, "AK05 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, False, "AK05 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK06 ACP", "AK06ACP", siteName, ModuleID, 2, 12, 7.5, "AK06 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, False, "AK06 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK08 ACP", "AK08ACP", siteName, ModuleID, 2, 12, 7.5, "AK08 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, False, "AK08 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AK09 ACP", "AK09ACP", siteName, ModuleID, 2, 12, 7.5, "AK09 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, False, "AK09 ACP Blocked/Starved"))

        AllProdLines.Add(New ProdLine("AC1 Casepacker Ext", "AC1CasepackerExt", siteName, ModuleID, 2, 12, 7.5, "AC1 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC1 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC1 Casepacker Int", "AC1CasepackerInt", siteName, ModuleID, 2, 12, 7.5, "AC1 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC1 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC2 Casepacker Ext", "AC2CasepackerExt", siteName, ModuleID, 2, 12, 7.5, "AC2 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC2 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC2 Casepacker Int", "AC2CasepackerInt", siteName, ModuleID, 2, 12, 7.5, "AC2 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC2 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC3 Casepacker Ext", "AC3CasepackerExt", siteName, ModuleID, 2, 12, 7.5, "AC3 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC3 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC3 Casepacker Int", "AC3CasepackerInt", siteName, ModuleID, 2, 12, 7.5, "AC3 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC3 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC4 Casepacker Ext", "AC4CasepackerExt", siteName, ModuleID, 2, 12, 7.5, "AC4 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC4 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC4 Casepacker Int", "AC4CasepackerInt", siteName, ModuleID, 2, 12, 7.5, "AC4 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AC4 Casepacker Internal Blocked/Starved"))

        AllProdLines.Add(New ProdLine("AT09 ACP", "AT09ACP", siteName, ModuleID, 2, 12, 7.5, "AT09AT10 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AT09AT10 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT13 ACP", "AT13ACP", siteName, ModuleID, 2, 12, 7.5, "AT13 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AT13 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AT14 ACP", "AT14ACP", siteName, ModuleID, 2, 12, 7.5, "AT14 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "AT14 ACP Blocked/Starved"))


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fam", siteName, SECTOR_FAMILY, prStoryMapping.FamilyMaking, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Zone SB", "Zone DT", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Brand Change", "Down Day", "Soft Swing",
            "RLS", "DT RLS", OTHERS_STRING}


        'initialize the lines

        AllProdLines.Add(New ProdLine("1A", "1A", siteName, ModuleID2, 2, 12, 7.5, "AY1A Reliability", "", prStoryMapping.FamilyMaking, True, "AY1A Sheetbreak"))
        AllProdLines.Add(New ProdLine("2A", "2A", siteName, ModuleID2, 2, 12, 7.5, "AY2A Reliability", "", prStoryMapping.FamilyMaking, True, "AY2A Sheetbreak"))
        AllProdLines.Add(New ProdLine("3A", "3A", siteName, ModuleID2, 2, 12, 7.5, "AY3A Reliability", "", prStoryMapping.FamilyMaking, True, "AY3A Sheetbreak"))
        AllProdLines.Add(New ProdLine("4A", "4A", siteName, ModuleID2, 2, 12, 7.5, "AY4A Reliability", "", prStoryMapping.FamilyMaking, True, "AY4A Sheetbreak"))
        AllProdLines.Add(New ProdLine("5A", "5A", siteName, ModuleID2, 2, 12, 7.5, "AY5A Reliability", "", prStoryMapping.FamilyMaking, True, "AY5A Sheetbreak"))
        AllProdLines.Add(New ProdLine("6A", "6A", siteName, ModuleID2, 2, 12, 7.5, "AY6A Reliability", "", prStoryMapping.FamilyMaking, True, "AY6A Sheetbreak"))


        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "Fam", siteName, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Palletizer, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))


        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Quality", "Upper Level", "Lower Level", "Blocked/Starved",
            "Cycle Stop / Other", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Changeover", "Blocked/Starved",
            "Planned Intervention", "CIL/RLS", "Roll Change", OTHERS_STRING}

        AllProdLines.Add(New ProdLine("Albany Palletizer #1", "AYPAL01", siteName, ModuleID3, 2, 12, 7.5, "AYPAL01 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("Albany Palletizer #2", "AYPAL02", siteName, ModuleID3, 2, 12, 7.5, "AYPAL02 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL02 Blocked/Starved"))
        'AllProductionLines.Add(New productionLine("Albany Palletizer #6", "AYPAL06", siteName, ModuleID3, 2, 12, 7.5, "AYPAL06 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL06 Blocked/Starved"))
        'AllProductionLines.Add(New productionLine("Albany Palletizer #7", "AYPAL07", siteName, ModuleID3, 2, 12, 7.5, "AYPAL07 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL07 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #8", "AYPAL08", siteName, ModuleID3, 2, 12, 7.5, "AYPAL08 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL08 Blocked/Starved"))
        'AllProductionLines.Add(New productionLine("Albany Palletizer #12", "AYPAL12", siteName, ModuleID3, 2, 12, 7.5, "AYPAL12 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL12 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #13", "AYPAL13", siteName, ModuleID3, 2, 12, 7.5, "AYPAL13 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL13 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #14", "AYPAL14", siteName, ModuleID3, 2, 12, 7.5, "AYPAL14 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL14 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #16", "AYPAL16", siteName, ModuleID3, 2, 12, 7.5, "AYPAL16 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL16 Blocked/Starved"))
        'AllProductionLines.Add(New productionLine("Albany Palletizer #17", "AYPAL17", siteName, ModuleID3, 2, 12, 7.5, "AYPAL17 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL17 Blocked/Starved"))
        'AllProductionLines.Add(New productionLine("Albany Palletizer #18", "AYPAL18", siteName, ModuleID3, 2, 12, 7.5, "AYPAL18 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL18 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #19", "AYPAL19", siteName, ModuleID3, 2, 12, 7.5, "AYPAL19 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL19 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #20", "AYPAL20", siteName, ModuleID3, 2, 12, 7.5, "AYPAL20 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL20 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #21", "AYPAL21", siteName, ModuleID3, 2, 12, 7.5, "AYPAL21 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL21 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #22", "AYPAL22", siteName, ModuleID3, 2, 12, 7.5, "AYPAL22 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL22 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #23", "AYPAL23", siteName, ModuleID3, 2, 12, 7.5, "AYPAL23 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL23 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #24", "AYPAL24", siteName, ModuleID3, 2, 12, 7.5, "AYPAL24 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL24 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #25", "AYPAL25", siteName, ModuleID3, 2, 12, 7.5, "AYPAL25 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL25 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #26", "AYPAL26", siteName, ModuleID3, 2, 12, 7.5, "AYPAL26 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL26 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Albany Palletizer #27", "AYPAL27", siteName, ModuleID3, 2, 12, 7.5, "AYPAL27 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "AYPAL27 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC3 Palletizer", "AC3PAL", siteName, ModuleID3, 2, 12, 7.5, "AC3 Palletizer Reliability DT", ""))
        AllProdLines.Add(New ProdLine("AC4 Palletizer", "AC4PAL", siteName, ModuleID3, 2, 12, 7.5, "AC4 Palletizer Reliability DT", ""))

        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "Fam", siteName, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        AllProdLines.Add(New ProdLine("AC3_StretchWrapper", "AC3_StretchWrapper", siteName, ModuleID4, 2, 12, 7.5, "AC3 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "AC3 Stretchwrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC4_StretchWrapper", "AC4_StretchWrapper", siteName, ModuleID4, 2, 12, 7.5, "AC4 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "AC4 Stretchwrapper Blocked/Starved"))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Fam", siteName, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Stacker, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        AllProdLines.Add(New ProdLine("AC3_Stacker", "AC3_Stacker", siteName, ModuleID5, 2, 12, 7.5, "AC3 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "AC3 Palletizer Stacker Blocked/Starved"))
        AllProdLines.Add(New ProdLine("AC4_Stacker", "AC4_Stacker", siteName, ModuleID5, 2, 12, 7.5, "AC4 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "AC4 Palletizer Stacker Blocked/Starved"))

        Dim ModuleIDt As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleIDt, "Fam", siteName, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.ModPack))



        lineName = "ModPack"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleIDt, 2, 12, 7.5, "AK09 ModPack System Reliability", "", prStoryMapping.FamilyCareUnitOP_ModPACK, True, "AK09 ModPack Blocked/Starved"))
        AllProdLines(AllProdLines.Count - 1).doIincludeAllUptime = True
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True


    End Sub

    Private Sub initializeSite_Oxnard()
        AllProductionSites.Add(New ProdSite(SITE_OXNARD, "ox-mesdata002.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "OXN"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fam", SITE_OXNARD, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        'initialize the lines
        '  AllProductionLines.Add(New productionLine("OKK1", "OKK1", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK1 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("OKK2", "OKK2", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK2 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("OKK3", "OKK3", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK3 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("OTT4", "OTT4", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OTT4 Converter Reliability", ""))

        'AllProductionLines.Add(New productionLine("OTT5", "OTT5", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OTT5 Converter Reliability", ""))

        Dim lineName As String = "XC01"
        AllProdLines.Add(New ProdLine(lineName, SITE_OXNARD, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))

        lineName = "OKK1"
        AllProdLines.Add(New ProdLine(lineName, SITE_OXNARD, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
        lineName = "OKK2"
        AllProdLines.Add(New ProdLine(lineName, SITE_OXNARD, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
        lineName = "OKK3"
        AllProdLines.Add(New ProdLine(lineName, SITE_OXNARD, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
        lineName = "OTT4"
        AllProdLines.Add(New ProdLine(lineName, SITE_OXNARD, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
        lineName = "OTT5"
        AllProdLines.Add(New ProdLine(lineName, SITE_OXNARD, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))


        'Wrapper
        AllProdLines.Add(New ProdLine("OKK1 Wrapper", "OKK1Wrapper", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK1 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "OKK1 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OKK2 Wrapper", "OKK2Wrapper", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK2 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "OKK2 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OKK3 1 Wrapper", "OKK3_1Wrapper", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK3 1 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "OKK3 1 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OKK3 2 Wrapper", "OKK3_2Wrapper", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK3 2 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "OKK3 2 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OTT4 Wrapper", "OTT4Wrapper", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OTT4 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "OTT4 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OTT5 Wrapper", "OTT5Wrapper", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OTT5 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "OTT5 Wrapper Blocked/Starved"))

        AllProdLines.Add(New ProdLine("XC01 Wrapper Ext", "XC01 Wrapper Ext", SITE_OXNARD, ModuleID, 2, 12, 5.5, "XC01 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "XC01 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("XC01 Wrapper Int", "XC01 Wrapper Int", SITE_OXNARD, ModuleID, 2, 12, 5.5, "XC01 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "XC01 Wrapper Internal Blocked/Starved"))


        'ACP

        AllProdLines.Add(New ProdLine("OKK1 ACP", "OKK1ACP", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK1 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "OKK1 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OKK2 ACP", "OKK2ACP", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK2 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "OKK2 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OKK3 ACP", "OKK3ACP", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK3 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "OKK3 ACP Blocked/Starved"))


        AllProdLines.Add(New ProdLine("OTT5 ACP", "OTT5 ACP", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OTT5 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "OTT5 ACP Blocked/Starved"))

        AllProdLines.Add(New ProdLine("XC01 ACP Ext", "XC01 ACP Ext", SITE_OXNARD, ModuleID, 2, 12, 5.5, "XC01 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "XC01 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("XC01 ACP Int", "XC01 ACP Int", SITE_OXNARD, ModuleID, 2, 12, 5.5, "XC01 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "XC01 Casepacker Internal Blocked/Starved"))




        'Multiflow
        AllProdLines.Add(New ProdLine("OKK1 MultiFlow", "OKK1MF", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK1 Multiflow Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "OKK1 Multiflow Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OKK2 MultiFlow", "OKK2MF", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK2 Multiflow Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "OKK2 Multiflow Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OKK3 MultiFlow", "OKK3MF", SITE_OXNARD, ModuleID, 2, 12, 5.5, "OKK3 Multiflow Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "OKK3 Multiflow Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OTT4 MultiFlow", "OTT4MF", SITE_OXNARD, ModuleID, 2, 12, 5.5, "Tissue 4 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "Tissue 4 MF Blocked/Starved"))
        AllProdLines.Add(New ProdLine("OTT5 MultiFlow", "OTT5MF", SITE_OXNARD, ModuleID, 2, 12, 5.5, "Tissue 5 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "Tissue 5 MF Blocked/Starved"))

        AllProdLines.Add(New ProdLine("XC01 Bundler Ext", "XC01 Bundler Ext", SITE_OXNARD, ModuleID, 2, 12, 5.5, "XC01 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "XC01 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("XC01 Bundler Int", "XC01 Bundler Int", SITE_OXNARD, ModuleID, 2, 12, 5.5, "XC01 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "XC01 Bundler Internal Blocked/Starved"))


        'Making
        AllProdLines.Add(New ProdLine("1X", "1X", SITE_OXNARD, ModuleID, 2, 12, 5.5, "PC1X Reliability", "", prStoryMapping.FamilyMaking, True, "PC1X Sheetbreak"))
        AllProdLines.Add(New ProdLine("2X", "2X", SITE_OXNARD, ModuleID, 2, 12, 5.5, "PC2X Reliability", "", prStoryMapping.FamilyMaking, True, "PC2X Sheetbreak"))




        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "Fam", SITE_OXNARD, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Palletizer, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))


        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Quality", "Upper Level", "Lower Level", "Blocked/Starved",
            "Cycle Stop / Other", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Changeover", "Blocked/Starved",
            "Planned Intervention", "CIL/RLS", "Roll Change", OTHERS_STRING}


        AllProdLines.Add(New ProdLine("XC01 Palletizer", "XC01PAL", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "XC01 Palletizer DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer)) ', True, ""))


        AllProdLines.Add(New ProdLine("Oxnard P1", "PCPAL01", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL01 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL01 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Oxnard P2", "PCPAL02", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL01 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL02 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Oxnard P3", "PCPAL03", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL03 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL03 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Oxnard P4", "PCPAL04", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL04 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL04 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Oxnard P5", "PCPAL05", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL05 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL05 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Oxnard P6", "PCPAL06", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL06 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL06 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Oxnard P7", "PCPAL07", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL07 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL07 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Oxnard P8", "PCPAL08", SITE_OXNARD, ModuleID3, 2, 12, 5.5, "PCPAL08 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "PCPAL08 Blocked/Starved"))



        'AllProductionLines.Add(New productionLine("F", "BSF", SITE_BROWNS_SUMMIT, BS_OC, 2, 12, 6, "BS OC Line F DF Main", "BS OC Line F DF", prStoryMapping.OralCare_DF, True, "BS OC Line F DF RateLoss"))

        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "Fam", SITE_OXNARD, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("XC01_StretchWrapper", "XC01_StretchWrapper", SITE_OXNARD, ModuleID4, 2, 12, 5.5, "XC01 Palletizer Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, ""))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Fam", SITE_OXNARD, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("XC01_Stacker", "XC01_Stacker", SITE_OXNARD, ModuleID5, 2, 12, 5.5, "XC01 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, ""))
    End Sub

    Private Sub initializeSite_Greenbay()
        AllProductionSites.Add(New ProdSite(SITE_GREENBAY, "gbay-proficy002.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "GBY"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fam", SITE_GREENBAY, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        'initialize the lines

        Dim lineName As String = "FC01"
        Dim siteName As String = SITE_GREENBAY
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FC02"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FC03"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FC04"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FK66"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FK68"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FK69"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FK70"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FTT5"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FTT6"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FTT7"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "FTT8"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 5.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))

        '  AllProductionLines.Add(New productionLine("FC01", "FC01", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC01 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FC02", "FC02", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC02 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FC03", "FC03", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC03 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FC04", "FC04", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC04 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FK66", "FK66", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK66 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FK68", "FK68", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK68 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FK69", "FK69", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK69 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FK70", "FK70", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK70 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("FTT5", "FTT5", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT5 Converter Reliability", ""))
        ' AllProductionLines.Add(New productionLine("FTT6", "FTT6", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT6 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("FTT7", "FTT7", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT7 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("FTT8", "FTT8", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT8 Converter Reliability", ""))

        'Bundlers and Multiflow
        AllProdLines.Add(New ProdLine("BTYPACK Bundler 1", "BTYPACK Bundler 1", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "BTYPACK Bundler 1 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BTYPACK Bundler 1 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BTYPACK Bundler 2", "BTYPACK Bundler 2", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "BTYPACK Bundler 2 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BTYPACK Bundler 2 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BTYPACK Bundler 3", "BTYPACK Bundler 3", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "BTYPACK Bundler 3 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BTYPACK Bundler 3 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BTYPACK Bundler 4", "BTYPACK Bundler 4", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "BTYPACK Bundler 4 Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BTYPACK Bundler 4 Blocked/Starved"))

        AllProdLines.Add(New ProdLine("FTL4PACK Bundler", "FTL4PACK Bundler", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTL4PACK Bundler Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FTL4PACK Bundler Blocked/Starved"))
        AllProdLines.Add(New ProdLine("TTNPACK Bundler N", "TTNPACK Bundler North", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "TTNPACK Bundler North Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "TTNPACK Bundler North Blocked/Starved"))
        AllProdLines.Add(New ProdLine("TTNPACK Bundler S", "TTNPACK Bundler South", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTL4PACK Bundler Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FTL4PACK Bundler Blocked/Starved"))

        AllProdLines.Add(New ProdLine("FC01 Bundler Ext ", "FC01 Bundler Ext ", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC01 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC01 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC01 Bundler Int ", "FC01 Bundler Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC01 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC01 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC02 Bundler Ext", "FC02 Bundler Ext", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC02 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC02 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC02 Bundler Int", "FC02 Bundler Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC02 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC02 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03 Bundler Ext", "FC03 Bundler Ext", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC03 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC03 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03 Bundler Int", "FC03 Bundler Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC03 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC03 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04 Bundler Ext", "FC04 Bundler Ext", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC04 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC04 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04 Bundler Int", "FC04 Bundler Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC04 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "FC04 Bundler Internal Blocked/Starved"))



        'Wrapper
        AllProdLines.Add(New ProdLine("FK68 Wrapper", "FK68 Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK68 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FK68 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FK69 Wrapper", "FK69 Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK69 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FK69 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FK70 N Wrapper", "FK70 North Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK70 North Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FK70 North Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FK70 S Wrapper", "FK70 South Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FK70 South Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FK70 South Wrapper Blocked/Starved"))

        AllProdLines.Add(New ProdLine("FTT5 Wrapper", "FTT5 Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT5 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FTT5 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FTT6 Wrapper", "FTT6 Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT6 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FTT6 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FTT7 Wrapper", "FTT7 Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT7 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FTT7 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FTT8 Wrapper", "FTT8 Wrapper", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FTT8 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FTT8 Wrapper Blocked/Starved"))

        AllProdLines.Add(New ProdLine("FC01 Wrapper Ext", "FC01 Wrapper Ext", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC01 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC01 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC01 Wrapper Int", "FC01 Wrapper Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC01 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC01 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC02 Wrapper Ext", "FC02 Wrapper Ext", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC02 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC02 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC02 Wrapper Int", "FC02 Wrapper Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC02 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC02 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03 Wrapper Ext", "FC03 Wrapper Ext", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC03 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC03 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03 Wrapper Int", "FC03 Wrapper Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC03 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC03 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04 Wrapper Ext", "FC04 Wrapper Ext", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC04 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC04 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04 Wrapper Int", "FC04 Wrapper Int", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC04 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "FC04 Wrapper Internal Blocked/Starved"))

        'ACP Casepacker
        AllProdLines.Add(New ProdLine("FC01 ACP Ext", "FC01ACPExt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC01 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC01 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC01 ACP Int", "FC01ACPInt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC01 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC01 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC02 ACP Ext", "FC02ACPExt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC02 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC02 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC02 ACP Int", "FC02ACPInt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC02 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC02 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03 ACP Ext", "FC03ACPExt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC03 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC03 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03 ACP Int", "FC03ACPInt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC03 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC03 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04 ACP Ext", "FC04ACPExt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC04 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC04 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04 ACP Int", "FC04ACPInt", SITE_GREENBAY, ModuleID, 2, 12, 5.5, "FC04 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "FC04 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FT5 ACP", "FT5ACP", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "TTNPACK FTT5ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "TTNPACK FTT5ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FT6 ACP", "FT6ACP", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "TTNPACK FTT6ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "TTNPACK FTT6ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FT7_8 ACP", "FT7_8ACP", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "TTNPACK FTT78ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "TTNPACK FTT78ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BTYPack ACP Central", " BTYPackACPCentral ", SITE_GREENBAY, ModuleID, 2, 12, 6.5, " BTYPACK CACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, " BTYPACK CACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FK69 ACP", " FK69ACP ", SITE_GREENBAY, ModuleID, 2, 12, 6.5, " BTYPACK 69 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, " BTYPACK 69 ACP Blocked/Starved"))


        'Making
        AllProdLines.Add(New ProdLine("10F", "10F", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "FP10 Reliability", "", prStoryMapping.FamilyMaking, True, "FP10 Sheetbreak"))
        AllProdLines.Add(New ProdLine("11F", "11F", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "FP11 Reliability", "", prStoryMapping.FamilyMaking, True, "FP11 Sheetbreak"))
        AllProdLines.Add(New ProdLine("12F", "12F", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "FP12 Reliability", "", prStoryMapping.FamilyMaking, True, "FP12 Sheetbreak"))
        AllProdLines.Add(New ProdLine("13F", "13F", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "FP13 Reliability", "", prStoryMapping.FamilyMaking, True, "FP13 Sheetbreak"))
        AllProdLines.Add(New ProdLine("14F", "14F", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "FP14 Reliability", "", prStoryMapping.FamilyMaking, True, "FP14 Sheetbreak"))
        AllProdLines.Add(New ProdLine("15F", "15F", SITE_GREENBAY, ModuleID, 2, 12, 6.5, "FP15 Reliability", "", prStoryMapping.FamilyMaking, True, "FP15 Sheetbreak"))


        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "Fam", SITE_GREENBAY, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Palletizer, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))


        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Quality", "Upper Level", "Lower Level", "Blocked/Starved",
            "Cycle Stop / Other", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Changeover", "Blocked/Starved",
            "Planned Intervention", "CIL/RLS", "Roll Change", OTHERS_STRING}



        AllProdLines.Add(New ProdLine("Greenbay P06", "GBPAL06", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL06 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL06 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P09", "GBPAL09", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL09 Palletizer (070) DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL09 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P10", "GBPAL10", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL10 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL10 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P11", "GBPAL11", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL11 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL11 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P12", "GBPAL12", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL12 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL12 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P13", "GBPAL13", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL13 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL13 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P14", "GBPAL14", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL14 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL14 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P16", "GBPAL16", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL16 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL16 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P26", "GBPAL26", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL26 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL26 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P27", "GBPAL27", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL27 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL27 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P28", "GBPAL28", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL28 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL28 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P29", "GBPAL29", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL29 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL29 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P30", "GBPAL30", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL30 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL30 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("Greenbay P50", "GBPAL50", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "GBPAL50 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL50 Blocked/Starved"))

        AllProdLines.Add(New ProdLine("FC01 Palletizer", "FC01PAL", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "FC01 Palletizer DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "FC01 Palletizer Blocked/Starved"))
        ' AllProductionLines(AllProductionLines.Count - 1).IsStartupMode = True
        AllProdLines.Add(New ProdLine("FC02 Palletizer", "FC02PAL", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "FC02 Palletizer Reliability [285] DT", ""))
        AllProdLines.Add(New ProdLine("FC03 Palletizer", "FC03PAL", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "FC03 Palletizer Reliability [285] DT", ""))
        AllProdLines.Add(New ProdLine("FC04 Palletizer", "FC04PAL", SITE_GREENBAY, ModuleID3, 2, 12, 6.5, "FC04 Palletizer Reliability [285] DT", ""))




        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "Puffs", SITE_GREENBAY, SECTOR_FAMILY, prStoryMapping.Puffs, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        ' AllProductionLines.Add(New productionLine("Greenbay P12", "GBPAL12", SITE_GREENBAY, ModuleID4, 2, 12, 6.5, "GBPAL12 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "GBPAL12 Blocked/Starved"))

        AllProdLines.Add(New ProdLine("FFF7", SITE_GREENBAY, ModuleID4, 2, 12, 6.5, "FFF7 Converter Reliability", "", -1, False, "FFF7 Converter Rate Loss", RateLossMode.Separate))
        AllProdLines.Add(New ProdLine("FF7A", SITE_GREENBAY, ModuleID4, 2, 12, 6.5, "FF7A Converter Reliability", "", -1, False, "FF7A Converter Rate Loss", RateLossMode.Separate))

        '  AllProductionLines.Add(New productionLine("FFF1 East", "FFF1 East", SITE_GREENBAY, ModuleID4, 2, 12, 6.5, "FFF1 Converter East Reliability", "", prStoryMapping.straight, True, "FFF1 Converter East Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FFF1 East", "FFF1 East", SITE_GREENBAY, ModuleID4, 2, 12, 6.5, "FFF1 Converter East Reliability", ""))

        AllProdLines.Add(New ProdLine("FF1 West", SITE_GREENBAY, ModuleID4, 2, 12, 6.5, "FFF1 Converter West Reliability", "", -1, False, "FFF1 Converter West Rate Loss", RateLossMode.Separate))
        AllProdLines.Add(New ProdLine("FF1 Combined", SITE_GREENBAY, ModuleID4, 2, 12, 6.5, "FFF1 Converter Reliability", "", -1, False, "FFF1 Converter Rate Loss", RateLossMode.Separate))


        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Fam", SITE_GREENBAY, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("FC01_StretchWrapper", "FC01_StretchWrapper", SITE_GREENBAY, ModuleID5, 2, 12, 6.5, "FC01 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "FC01 Stretchwrapper Blocked/Starved"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
        AllProdLines.Add(New ProdLine("FC02_StretchWrapper", "FC02_StretchWrapper", SITE_GREENBAY, ModuleID5, 2, 12, 6.5, "FC02 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "FC02 Stretchwrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03_StretchWrapper", "FC03_StretchWrapper", SITE_GREENBAY, ModuleID5, 2, 12, 6.5, "FC03 Stretchwrapper Reliability DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "FC03 Stretchwrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04_StretchWrapper", "FC04_StretchWrapper", SITE_GREENBAY, ModuleID5, 2, 12, 6.5, "FC04 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "FC04 Stretchwrapper Blocked/Starved"))

        Dim ModuleID6 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID6, "Fam", SITE_GREENBAY, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Stacker, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        AllProdLines.Add(New ProdLine("FC01_Stacker", "FC01_Stacker", SITE_GREENBAY, ModuleID6, 2, 12, 6.5, "FC01 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "FC01 Stacker Blocked/Starved"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True
        AllProdLines.Add(New ProdLine("FC02_Stacker", "FC02_Stacker", SITE_GREENBAY, ModuleID6, 2, 12, 6.5, "FC02 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "FC02 Stacker Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC03_Stacker", "FC03_Stacker", SITE_GREENBAY, ModuleID6, 2, 12, 6.5, "FC03 Palletizer Stacker Reliability DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "FC03 Stacker Blocked/Starved"))
        AllProdLines.Add(New ProdLine("FC04_Stacker", "FC04_Stacker", SITE_GREENBAY, ModuleID6, 2, 12, 6.5, "FC04 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "FC04 Stacker Blocked/Starved"))

    End Sub

    Private Sub initializeSite_BoxElder()
        AllProductionSites.Add(New ProdSite(SITE_BOXELDER, "beu-mesdatafc.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "BOX"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fam", SITE_BOXELDER, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        'initialize the lines
        'initialize the lines

        Dim lineName As String = "BC01"
        Dim siteName As String = SITE_BOXELDER
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "BC02"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "BC03"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))







        '   AllProductionLines.Add(New productionLine("BC01", "BC01", SITE_BOXELDER, ModuleID, 2, 12, 5.5, "BC01 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("BC02", "BC02", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC02 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("BC03", "BC03", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC03 Converter Reliability", ""))

        'MF
        AllProdLines.Add(New ProdLine("BC01 Bundler Ext", "BC01BundlerExt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC01 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BC01 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC01 Bundler Int", "BC01BundlerInt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC01 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BC01 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02 Bundler Ext", "BC02BundlerExt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC02 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BC02 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02 Bundler Int", "BC02BundlerInt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC02 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BC02 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03 Bundler Ext", "BC03BundlerExt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC03 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BC03 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03 Bundler Int", "BC03BundlerInt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC03 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "BC03 Bundler Internal Blocked/Starved"))

        'WRAPPER
        AllProdLines.Add(New ProdLine("BC01 Wrapper Ext", "BC01Wrapper Ext", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC01 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "BC01 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC01 Wrapper Int", "BC01WrapperInt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC01 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "BC01 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02 Wrapper Ext", "BC02WrapperExt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC02 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "BC02 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02 Wrapper Int", "BC02WrapperInt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC02 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "BC02 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03 Wrapper Ext", "BC03WrapperExt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC03 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "BC03 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03 Wrapper Int", "BC03WrapperInt", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC03 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "BC03 Wrapper Internal Blocked/Starved"))


        'ACP
        AllProdLines.Add(New ProdLine("BC01_Ext ACP", "BC01_ExtACP", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC01 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "BC01 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC01_Int ACP", "BC01_IntACP", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC01 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "BC01 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02_Ext ACP", "BC02_ExtACP", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC02 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "BC02 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02_Int ACP", "BC02_IntACP", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC02 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "BC02 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03_Ext ACP", "BC03_ExtACP", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC03 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "BC03 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03_Int ACP", "BC03_IntACP", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BC03 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "BC03 Casepacker Internal Blocked/Starved"))


        'Making
        AllProdLines.Add(New ProdLine("15B", "15B", SITE_BOXELDER, ModuleID, 2, 12, 7.5, "BE15 Reliability", "", prStoryMapping.FamilyMaking, True, "BE15 Sheetbreak"))




        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "Fam", SITE_BOXELDER, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Palletizer, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))


        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Quality", "Upper Level", "Lower Level", "Blocked/Starved",
            "Cycle Stop / Other", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Changeover", "Blocked/Starved",
            "Planned Intervention", "CIL/RLS", "Roll Change", OTHERS_STRING}

        AllProdLines.Add(New ProdLine("BC01 Palletizer", "BC01PAL", SITE_BOXELDER, ModuleID3, 2, 12, 7.5, "BC01 Palletizer DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("BC02 Palletizer", "BC02PAL", SITE_BOXELDER, ModuleID3, 2, 12, 7.5, "BC02 Palletizer DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("BC03 Palletizer", "BC03PAL", SITE_BOXELDER, ModuleID3, 2, 12, 7.5, "BC03 Palletizer DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))

        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "Fam", SITE_BOXELDER, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("BC01_StretchWrapper", "BC01_StretchWrapper", SITE_BOXELDER, ModuleID4, 2, 12, 7.5, "BC01 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "BC01 Stretchwrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02_StretchWrapper", "BC02_StretchWrapper", SITE_BOXELDER, ModuleID4, 2, 12, 7.5, "BC02 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "BC02 Stretchwrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03_StretchWrapper", "BC03_StretchWrapper", SITE_BOXELDER, ModuleID4, 2, 12, 7.5, "BC03 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, "BC03 Stretchwrapper Blocked/Starved"))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Fam", SITE_BOXELDER, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Stacker, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("BC01_Stacker", "BC01_Stacker", SITE_BOXELDER, ModuleID5, 2, 12, 7.5, "BC01 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "BC01 Stacker Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC02_Stacker", "BC02_Stacker", SITE_BOXELDER, ModuleID5, 2, 12, 7.5, "BC02 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "BC02 Stacker Blocked/Starved"))
        AllProdLines.Add(New ProdLine("BC03_Stacker", "BC03_Stacker", SITE_BOXELDER, ModuleID5, 2, 12, 7.5, "BC03 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, "BC03 Stacker Blocked/Starved"))




    End Sub

    Private Sub initializeSite_CapeGirardeau()
        AllProductionSites.Add(New ProdSite(SITE_CAPEGIRARDEAU, "cape-mesdatafc", "", SERVER_PW_V6, SERVER_UN_V6, "CGU"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fam", SITE_CAPEGIRARDEAU, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        'initialize the lines

        Dim lineName As String = "GC07"
        Dim siteName As String = SITE_CAPEGIRARDEAU

        Dim doRateLoss As Boolean = True

        If doRateLoss Then
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "GC08"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
            lineName = "GK21"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
            lineName = "GK22"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
            lineName = "GK23"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
            lineName = "GK24"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Rate Loss", RateLossMode.Separate))
            lineName = "GT1"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT01 Converter Reliability", "", -1, False, "GT01 Rate Loss", RateLossMode.Separate))
            lineName = "GT2"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT02 Converter Reliability", "", -1, False, "GT02 Rate Loss", RateLossMode.Separate))
            lineName = "GT3"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT03 Converter Reliability", "", -1, False, "GT03 Rate Loss", RateLossMode.Separate))
            lineName = "GT4"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT04 Converter Reliability", "", -1, False, "GT04 Rate Loss", RateLossMode.Separate))
        Else
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "GC08"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "GK21"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False,))
            lineName = "GK22"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "GK23"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "GK24"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "GT1"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT01 Converter Reliability", "", -1, False))
            lineName = "GT2"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT02 Converter Reliability", "", -1, False))
            lineName = "GT3"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT03 Converter Reliability", "", -1, False))
            lineName = "GT4"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, "GT04 Converter Reliability", "", -1, False))
        End If


        'MF

        AllProdLines.Add(New ProdLine("GC07 Bundler Ext", "GC07BundlerExt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC07 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GC07 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC07 Bundler Int", "GC07BundlerInt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC07 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GC07 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC08 Bundler Ext", "GC08BundlerExt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC08 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GC08 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC08 Bundler Int", "GC08BundlerInt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC08 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GC08 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK21 MF Multiflow", "GK21MFMultiflow", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK21 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GK21 MF Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK22 MF Multiflow", "GK22MFMultiflow", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK22 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GK22 MF Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK23 MF Multiflow", "GK23MFMultiflow", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK23 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GK23 MF Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK24 MF Multiflow", "GK24MFMultiflow", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK24 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GK24 MF Blocked/Starved"))
        ' AllProductionLines.Add(New productionLine("MF_Berk Multiflow", "MF_BerkMultiflow", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "MF_Berk Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MF_Berk Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MF_GL Multiflow", "MF_GLMultiflow", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "MF_GL Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MF_GL Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GT03/GT04 Multiflow", "MF_BerkMultiflow", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT03/GT04 MF Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "GT03/GT04 MF Blocked/Starved"))
        AllProdLines(AllProdLines.Count - 1).IsStartupMode = True

        'Wrapper

        AllProdLines.Add(New ProdLine("GK21 Wrapper A", "GK21WrapperA", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK21 Wrapper A Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GK21 Wrapper A Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK21 Wrapper B", "GK21WrapperB", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK21 Wrapper B Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GK21 Wrapper B Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK22 Wrapper", "GK22Wrapper", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK22 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GK22 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK23 Wrapper B", "GK23WrapperA", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK23 Wrapper A Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GK23 Wrapper A Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GK23 Wrapper B", "GK23WrapperB", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK23 Wrapper B Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GK23 Wrapper B Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GT01 Wrapper", "GT01Wrapper", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT01 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GT01 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GT02 Wrapper", "GT02Wrapper", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT02 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GT02 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GT03 Wrapper", "GT03Wrapper", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT03 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GT03 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GT04 Wrapper", "GT04Wrapper", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT04 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GT04 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC07 Wrapper Ext", "GC07WrapperExt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC07 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GC07 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC07 Wrapper Int", "GC07WrapperInt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC07 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GC07 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC08 Wrapper Ext", "GC08WrapperExt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC08 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GC08 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC08 Wrapper Int", "GC08WrapperInt", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC08 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "GC08 Wrapper Internal Blocked/Starved"))


        'ACP

        AllProdLines.Add(New ProdLine("GKPK51_GK21 ACP", "GKPK51_GK21ACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK21 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, False, ""))
        AllProdLines.Add(New ProdLine("GKPK51_GK22 ACP", "GKPK51_GK22ACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK22 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, False, ""))
        AllProdLines.Add(New ProdLine("GKPK51_GK23 ACP", "GKPK51_GK23ACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GK23 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, False, ""))
        AllProdLines.Add(New ProdLine("GTPK52_GT01 ACP", "GTPK52_GT01ACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT01 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GT01 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GTPK52_GT02 ACP", "GTPK52_GT02ACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT02 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GT02 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GTPK52_GT03 ACP", "GTPK52_GT03ACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT03 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GT03 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GTPK52 ACP", "GTPK52ACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GT04 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GT04 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC07_Ext ACP", "GC07_ExtACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC07 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GC07 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC07_Int ACP", "GC07_IntACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC07 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GC07 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC08_Ext ACP", "GC08_ExtACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC08 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GC08 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("GC08_Int ACP", "GC08_IntACP", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GC08 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "GC08 Casepacker Internal Blocked/Starved"))


        'Making
        AllProdLines.Add(New ProdLine("5G", "5G", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GP05 Reliability", "", prStoryMapping.FamilyMaking, True, "GP05 Sheetbreak"))
        AllProdLines.Add(New ProdLine("6G", "6G", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GP06 Reliability", "", prStoryMapping.FamilyMaking, True, "GP06 Sheetbreak"))
        AllProdLines.Add(New ProdLine("7G", "7G", SITE_CAPEGIRARDEAU, ModuleID, 2, 12, 7.5, "GP07 Reliability", "", prStoryMapping.FamilyMaking, True, "GP07 Sheetbreak"))


        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "Fam", SITE_CAPEGIRARDEAU, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Palletizer, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))


        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Quality", "Upper Level", "Lower Level", "Blocked/Starved",
            "Cycle Stop / Other", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Changeover", "Blocked/Starved",
            "Planned Intervention", "CIL/RLS", "Roll Change", OTHERS_STRING}



        AllProdLines.Add(New ProdLine("GK21 Palletizer", "GK21", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "GK21 ULS Towel Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GK22 Palletizer", "GK22", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "GK22 ULS Towel Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GK23 Palletizer", "GK23", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "GK23 ULS Towel Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GK24 Palletizer", "GK24", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "GK24 ULS Towel Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GT01 Palletizer", "ULSAPalTissue", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "ULS A Tissue Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GT02 Palletizer", "ULSBPalTissue", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "ULS B Tissue Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GT03 Palletizer", "ULSCPalTissue", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "ULS C Tissue Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GT04 Palletizer", "ULSDPalTissue", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "ULS D Tissue Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GC07 Palletizer", "GC07PAL", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "GC07 Palletizer Reliability [285] DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))
        AllProdLines.Add(New ProdLine("GC08 Palletizer", "GC08PAL", SITE_CAPEGIRARDEAU, ModuleID3, 2, 12, 7.5, "GC08 Palletizer Reliability [285] DT", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, ""))

        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "Fam", SITE_CAPEGIRARDEAU, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("GC07_StretchWrapper", "GC07_StretchWrapper", SITE_CAPEGIRARDEAU, ModuleID4, 2, 12, 7.5, "GC07 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, ""))
        AllProdLines.Add(New ProdLine("GC08_StretchWrapper", "GC08_StretchWrapper", SITE_CAPEGIRARDEAU, ModuleID4, 2, 12, 7.5, "GC08 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, ""))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Fam", SITE_CAPEGIRARDEAU, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Stacker, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        AllProdLines.Add(New ProdLine("GC07_Stacker", "GC07_Stacker", SITE_CAPEGIRARDEAU, ModuleID5, 2, 12, 7.5, "GC07 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, ""))
        AllProdLines.Add(New ProdLine("GC08_Stacker", "GC08_Stacker", SITE_CAPEGIRARDEAU, ModuleID5, 2, 12, 7.5, "GC08 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, ""))





    End Sub

    Private Sub initializeSite_Mehoopany()
        AllProductionSites.Add(New ProdSite(SITE_MEHOOPANY, "mp-mesdtafc.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "MPN"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fam", SITE_MEHOOPANY, SECTOR_FAMILY, prStoryMapping.Albany, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))
        'initialize the lines

        Dim lineName As String = "MC01"
        Dim siteName As String = SITE_MEHOOPANY

        Dim doRateLoss As Boolean = True

        If doRateLoss Then
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MC02"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK71"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK72"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK74"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK75"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK77"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK78"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK79"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK80"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK83"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK84"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
            lineName = "MK85"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        Else
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MC02"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK71"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK72"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK74"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK75"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK77"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK78"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK79"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK80"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK83"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK84"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))
            lineName = "MK85"
            AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False))

        End If

        AllProdLines.Add(New ProdLine("MNN1", "MNN1", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MNN1 Converter Reliability", "", prStoryMapping.FamilyCareUnitOP_Napkins))
        AllProdLines.Add(New ProdLine("MNN2", "MNN2", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MNN2 Converter Reliability", "", prStoryMapping.FamilyCareUnitOP_Napkins))
        AllProdLines.Add(New ProdLine("MNN3", "MNN3", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MNN3 Converter Reliability", "", prStoryMapping.FamilyCareUnitOP_Napkins))
        AllProdLines.Add(New ProdLine("MNN4", "MNN4", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MNN4 Converter Reliability", "", prStoryMapping.FamilyCareUnitOP_Napkins))
        AllProdLines.Add(New ProdLine("MNN5", "MNN5", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MNN5 Converter Reliability", "", prStoryMapping.FamilyCareUnitOP_Napkins))





        '   AllProductionLines.Add(New productionLine("MC01", "MC01", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC01 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MC02", "MC02", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC02 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK71", "MK71", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK71 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK72", "MK72", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK72 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK74", "MK74", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK74 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK75", "MK75", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK75 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK77", "MK77", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK77 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK78", "MK78", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK78 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK79", "MK79", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK79 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK80", "MK80", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK80 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK83", "MK83", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK83 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MK84", "MK84", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK84 Converter Reliability", ""))
        '    AllProductionLines.Add(New productionLine("MK85", "MK85", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK85 Converter Reliability", ""))

        lineName = "MT60"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MT61"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MT62"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MT63"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MT65"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MT66"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MT67"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MK70"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MK81"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))
        lineName = "MK82"
        AllProdLines.Add(New ProdLine(lineName, siteName, ModuleID, 2, 12, 7.5, lineName & " Converter Reliability", "", -1, False, lineName & " Converter Rate Loss", RateLossMode.Separate))



        '   AllProductionLines.Add(New productionLine("MT60", "MT60", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT60 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MT61", "MT61", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT61 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MT62", "MT62", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT62 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MT63", "MT63", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT63 Converter Reliability", ""))
        '   AllProductionLines.Add(New productionLine("MT65", "MT65", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT65 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("MT66", "MT66", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT66 Converter Reliability", ""))
        ''  AllProductionLines.Add(New productionLine("MT67", "MT67", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT67 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("MK70", "MK70", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK70 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("MK81", "MK81", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK81 Converter Reliability", ""))
        '  AllProductionLines.Add(New productionLine("MK82", "MK82", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK82 Converter Reliability", ""))


        'MF
        AllProdLines.Add(New ProdLine("MC01 Bundler Ext", "MC01BundlerExt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC01 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MC01 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC01 Bundler Int", "MC01BundlerInt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC01 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MC01 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC02 Bundler Ext", "MC02BundlerExt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC02 Bundler External Reliability [256] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MC02 Bundler External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC02 Bundler Int", "MC02BundlerInt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC02 Bundler Internal Reliability [255] DT", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MC02 Bundler Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT65  Multiflow", "MT65Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT65 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MT65 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT66  Multiflow", "MT66Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT66 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MT66 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK81  Multiflow", "MK81Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK81 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK81 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK82  Multiflow", "MK82Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK82 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK82 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK79  Multiflow", "MK79Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK79 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK79 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK80  Multiflow", "MK80Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK80 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK80 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK83  Multiflow", "MK83Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK83 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK83 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK84  Multiflow", "MK84Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK84 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK84 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK85  Multiflow", "MK85Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK85 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK85 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK70/72 West  Multiflow", "MK70/72WestMultiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK70/72 West LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK70/72 West LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK74 East  Multiflow", "MK74EastMultiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK74 East LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK74 East LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK75  Multiflow", "MK75Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK75 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK75 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT60  Multiflow", "MT60Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT60 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MT60 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT61  Multiflow", "MT61Multiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT61 LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MT61 LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK71 East  Multiflow", "MK71EastMultiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK71 East LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK71 East LCP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK77 West  Multiflow", "MK77WestMultiflow", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK77 West LCP Reliability", "", prStoryMapping.FamilyCareUnitOP_mf, True, "MK77 West LCP Blocked/Starved"))



        'WRAPPER

        AllProdLines.Add(New ProdLine("MK72 Wrapper", "MK72Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK72 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK72 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK74 Wrapper", "MK74Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK74 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK74 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK75 Wrapper", "MK75Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK75 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK75 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK78 Wrapper", "MK78Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK78 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK78 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK79 Wrapper", "MK79Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK79 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK79 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK80 Wrapper", "MK80Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK80 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK80 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK83 Wrapper West", "MK83WrapperWest", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK83 West Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK83 West Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK83 Wrapper East", "MK83WrapperEast", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK83 East Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK83 East Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK84 Wrapper East", "MK84WrapperEast", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK84 East Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK84 East Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK85 Wrapper", "MK85Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK85 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK85 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT60 Wrapper", "MT60Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT60 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT60 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT61 Wrapper", "MT61Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT61 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT61 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT62 Wrapper", "MT62Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT62 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT62 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT63 Wrapper", "MT63Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT63 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT63 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT65 Wrapper East", "MT65WrapperEast", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT65 East Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT65 East Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT65 Wrapper West", "MT65WrapperWest", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT65 West Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT65 West Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT67 Wrapper", "MT67Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT67 Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT67 Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC01 Wrapper Ext", "MC01WrapperExt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC01 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MC01 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC01 Wrapper Int", "MC01WrapperInt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC01 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MC01 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC02 Wrapper Ext", "MC02WrapperExt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC02 Wrapper External Reliability [221] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MC02 Wrapper External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC02 Wrapper Int", "MC02WrapperInt", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC02 Wrapper Internal Reliability [220] DT", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MC02 Wrapper Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT66 Wrapper East", "MT66WrapperEast", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT66 East Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT66 East Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT66 Wrapper West", "MT66WrapperWest", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT66 West Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MT66 West Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK81 Wrapper East", "MK81WrapperEast", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK81 East Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK81 East Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK81 Wrapper West", "MK81WrapperWest", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK81 West Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK81 West Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK82 Wrapper East", "MK82WrapperEast", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK82 East Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK82 East Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK82 Wrapper West", "MK82WrapperWest", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK82 West Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK82 West Wrapper Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MK84 Wrapper", "MK84Wrapper", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK84 East Wrapper Reliability", "", prStoryMapping.FamilyCareUnitOP_Wrapper, True, "MK84 East Wrapper Blocked/Starved"))

        'ACP

        AllProdLines.Add(New ProdLine("MNN6 ACP", "MNN6ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MNN6 Casepacker Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MNN6 Casepacker Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK13 ACP", "PK13ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MNPK13 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MNPK13 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK16 Tissue ACP", "PK16TissueACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MTACP16 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MTACP16 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK17 ACP", "PK17ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK79/80 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MK79/80 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK23 Towel_MK72 ACP", "PK23Towel_MK72ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK72 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MK72 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK23 Towel_MK74 ACP", "PK23Towel_MK74ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK74 Center ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MK74 Center ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK26 Tissue_MT60 ACP", "PK26Tissue_MT60ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT60 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MT60 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK26 Tissue_MT62 ACP", "PK26Tissue_MT62ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT62 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MT62 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK26 Tissue_MT63 ACP", "PK26Tissue_MT63ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MT63 ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MT63 ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK26 Towel_MK77 ACP", "PK26Towel_MK77ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK77 North ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MK77 North ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("PK26 Towel_MK78 ACP", "PK26Towel_MK78ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MK78 South ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MK78 South ACP Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC01_Ext ACP", "MC01_ExtACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC01 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MC01 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC01_Int ACP", "MC01_IntACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC01 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MC01 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC02_Ext ACP", "MC02_ExtACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC02 Casepacker External Reliability [271] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MC02 Casepacker External Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC02_Int ACP", "MC02_IntACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MC02 Casepacker Internal Reliability [270] DT", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "MC02 Casepacker Internal Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MT67 ACP", "MT67 ACP", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "PK23 EAST ACP Reliability", "", prStoryMapping.FamilyCareUnitOP_ACP, True, "PK23 EAST ACP Blocked/Starved"))

        'Making
        AllProdLines.Add(New ProdLine("1M", "1M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP1M Reliability", "", prStoryMapping.FamilyMaking, True, "MP1M Sheetbreak"))
        AllProdLines.Add(New ProdLine("2M", "2M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP2M Reliability", "", prStoryMapping.FamilyMaking, True, "MP2M Sheetbreak"))
        AllProdLines.Add(New ProdLine("3M", "3M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP3M Reliability", "", prStoryMapping.FamilyMaking, True, "MP3M Sheetbreak"))
        AllProdLines.Add(New ProdLine("4M", "4M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP4M Reliability", "", prStoryMapping.FamilyMaking, True, "MP4M Sheetbreak"))
        AllProdLines.Add(New ProdLine("5M", "5M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP5M Reliability", "", prStoryMapping.FamilyMaking, True, "MP5M Sheetbreak"))
        AllProdLines.Add(New ProdLine("6M", "6M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP6M Reliability", "", prStoryMapping.FamilyMaking, True, "MP6M Sheetbreak"))
        AllProdLines.Add(New ProdLine("7M", "7M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP7M Reliability", "", prStoryMapping.FamilyMaking, True, "MP7M Sheetbreak"))
        AllProdLines.Add(New ProdLine("8M", "8M", SITE_MEHOOPANY, ModuleID, 2, 12, 7.5, "MP8M Reliability", "", prStoryMapping.FamilyMaking, True, "MP8M Sheetbreak"))



        Dim ModuleID3 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID3, "Fam", SITE_MEHOOPANY, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Palletizer, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))


        AllProdModules(AllProdModules.Count - 1).UnplannedT1List = New List(Of String) From {"Quality", "Upper Level", "Lower Level", "Blocked/Starved",
            "Cycle Stop / Other", OTHERS_STRING}
        AllProdModules(AllProdModules.Count - 1).PlannedT1List = New List(Of String) From {"Changeover", "Blocked/Starved",
            "Planned Intervention", "CIL/RLS", "Roll Change", OTHERS_STRING}




        AllProdLines.Add(New ProdLine("MPPAL03", "MPPAL03", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL03 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL03 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL04", "MPPAL04", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL04 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL04 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL05", "MPPAL05", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL05 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL05 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL07", "MPPAL07", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL07 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL07 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL08", "MPPAL08", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL08 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL08 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL10", "MPPAL10", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL10 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL10 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL11", "MPPAL11", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL11 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL11 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL12", "MPPAL12", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL12 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL12 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL13", "MPPAL13", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL13 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL13 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL14", "MPPAL14", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL14 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL14 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL15", "MPPAL15", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL15 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL15 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL16", "MPPAL16", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL16 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL16 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL17", "MPPAL17", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL17 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL17 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL18", "MPPAL18", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL18 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL18 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL19", "MPPAL19", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL19 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL19 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL20", "MPPAL20", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL20 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL20 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL21", "MPPAL21", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL21 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL21 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL22", "MPPAL22", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL22 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL22 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL23", "MPPAL23", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL23 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL23 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL24", "MPPAL24", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL24 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL24 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL25", "MPPAL25", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL25 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL25 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL26", "MPPAL26", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL26 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL26 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL27", "MPPAL27", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL27 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL27 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL28", "MPPAL28", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL28 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL28 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL29", "MPPAL29", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL29 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL29 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL30", "MPPAL30", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL30 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL30 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL31", "MPPAL31", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL31 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL31 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL36", "MPPAL36", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL36 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL36 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL37", "MPPAL37", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL37 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL37 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL38", "MPPAL38", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL38 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL38 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL60", "MPPAL60", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL60 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL60 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MPPAL61", "MPPAL61", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MPPAL61 Reliability", "", prStoryMapping.FamilyCareUnitOP_Palletizer, True, "MPPAL61 Blocked/Starved"))
        AllProdLines.Add(New ProdLine("MC01 Palletizer", "MC01PAL", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MC01 Palletizer Reliability DT", ""))
        AllProdLines.Add(New ProdLine("MC02 Palletizer", "MC02PAL", SITE_MEHOOPANY, ModuleID3, 2, 12, 7.5, "MC02 Palletizer Reliability DT", ""))

        Dim ModuleID4 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID4, "Fam", SITE_MEHOOPANY, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("MC01_StretchWrapper", "MC01_StretchWrapper", SITE_MEHOOPANY, ModuleID4, 2, 12, 7.5, "MC01 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, ""))
        AllProdLines.Add(New ProdLine("MC02_StretchWrapper", "MC02_StretchWrapper", SITE_MEHOOPANY, ModuleID4, 2, 12, 7.5, "MC02 Stretchwrapper DT", "", prStoryMapping.FamilyCareUnitOP_STRETCHWRAPPER, True, ""))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Fam", SITE_MEHOOPANY, SECTOR_FAMILY, prStoryMapping.FamilyCareUnitOP_Stacker, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Albany))

        AllProdLines.Add(New ProdLine("MC01_Stacker", "MC01_Stacker", SITE_MEHOOPANY, ModuleID5, 2, 12, 7.5, "MC01 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, ""))
        AllProdLines.Add(New ProdLine("MC02_Stacker", "MC02_Stacker", SITE_MEHOOPANY, ModuleID5, 2, 12, 7.5, "MC02 Palletizer Stacker DT", "", prStoryMapping.FamilyCareUnitOP_Stacker, True, ""))





        initializeSite_Mehoopany_Baby()


    End Sub

    Private Sub initializeSite_Mehoopany_Baby()
        Dim tmpSiteName As String = "Mehoopany Baby"
        AllProductionSites.Add(New ProdSite(tmpSiteName, "mp-mesdatabc.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "MPB"))

        Dim ModuleID5 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID5, "Baby", tmpSiteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.BabyCare))
        Dim tmpLineName As String = "DIMP-126"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-127"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-128"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-129"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-130"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-131"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-132"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-133"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-134"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-135"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-136"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-137"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-138"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-139"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-140"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-141"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-142"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-144"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))

        tmpLineName = "DIMP-145"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))
        tmpLineName = "DIMP-146"
        AllProdLines.Add(New ProdLine(tmpLineName, tmpLineName, tmpSiteName, ModuleID5, 2, 12, 6, tmpLineName, tmpLineName))

    End Sub


#End Region

    Private Sub initializeSites_MoreBabyCare()
        initializeSite_Cabuyao()
        initializeSite_Johannesburg()
        initializeSite_Luogang()
    End Sub

    Private Sub initializeSite_Luogang()
        Dim siteName As String = "Luogang"
        Dim lineName As String
        AllProductionSites.Add(New ProdSite(siteName, "luo-mesdatabc.ap.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "CAB"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "DILL171"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "DILL181"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "DILL182"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "DILL183"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "DILL184"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "DILL185"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "DILL186"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        lineName = "DILL187"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

    End Sub

    Private Sub initializeSite_Johannesburg()
        Dim siteName As String = "Johannesburg"
        Dim lineName As String
        AllProductionSites.Add(New ProdSite(siteName, "JOH-MESDATABC.eu.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "CAB"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "DIJH101"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIJH102"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DIJH103"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

    End Sub


    Private Sub initializeSite_Cabuyao()
        Dim siteName As String = "Cabuyao"
        AllProductionSites.Add(New ProdSite(siteName, "cab-mesdatabc", "", SERVER_PW_V6, SERVER_UN_V6, "CAB"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby", siteName, SECTOR_BABY, prStoryMapping.Mandideep, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        Dim lineName As String = "DICB107"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "DICB108"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        '   lineName = "DICB109"
        '   AllProductionLines.Add(New productionLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QACB004"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "Fem", siteName, SECTOR_FEM, prStoryMapping.Fem_LCC_HPU, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))


        lineName = "QACB003"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
        lineName = "QACB004"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 3, 8, 7, lineName & " Converter", lineName))
    End Sub

    Private Sub InitializeSite_KansasCity()
        Dim siteName As String = "Kansas City"
        Dim lineName As String
        AllProductionSites.Add(New ProdSite(siteName, "kc-mesdb101\mesdp101", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "KSC"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby", siteName, SECTOR_BABY, prStoryMapping.GENERIC, DefaultProficyDowntimeProcedure.Maple, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "DIJH101"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))

    End Sub

    Private Sub InitializeSite_BinhDuong()
        Dim siteName As String = "Binh Duong"
        Dim lineName As String
        AllProductionSites.Add(New ProdSite(siteName, "155.126.130.82", SERVER_PW_MAPLE, SERVER_UN_MAPLE, "BDO"))
        AllProductionSites(AllProductionSites.Count - 1).ServerDatabase = "MES"
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "F&HC", siteName, SECTOR_FHC, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.Maple, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "Chameleon"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Converter", lineName))
        AllProdLines(AllProdLines.Count - 1).Name_MAPLE = "Chame_Filler"

    End Sub

    Private Sub initializeSite_Crailshiem()
        Dim siteName As String = "Crailsheim"
        'cra-mesdtabc
        AllProductionSites.Add(New ProdSite(siteName, "cra-mesdtabc", "", SERVER_PW_V6, SERVER_UN_V6, "AUB"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fem Care", siteName, SECTOR_FEM, prStoryMapping.FemCare_Pads, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        Dim lineName As String


        lineName = "PECR018"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "PECR028"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR012"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        lineName = "QACR022"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR013"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR015"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR016"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR019"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR020"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR021"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QACR022"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        lineName = "QBCR031"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QBCR032"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QBCR033"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QBCR034"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QBCR035"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QBCR036"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))
        lineName = "QBCR038"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))


        Dim ModuleID2 As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID2, "HomeCare", siteName, SECTOR_HOME, prStoryMapping.FemCare_Pads, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        lineName = "FNCR041"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 6, lineName & " Converter", ""))

        lineName = "WWCR043"
        AllProdLines.Add(New ProdLine("FNCR043", lineName, siteName, ModuleID2, 2, 12, 6, "WWCR043" & " Converter", ""))

        lineName = "FMCR045"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 6, lineName & " Cartoner", ""))

        lineName = "FMCR046"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID2, 2, 12, 6, lineName & " Cartoner", ""))


    End Sub



    Private Sub initializeSite_Auburn()
        Dim siteName As String = "Auburn"

        AllProductionSites.Add(New ProdSite(siteName, "Abn-mesdatabc02.na.pg.com", "", SERVER_PW_V6, SERVER_UN_V6, "AUB"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Fem Care", siteName, SECTOR_FEM, prStoryMapping.TepejiFem, DefaultProficyDowntimeProcedure.OneClick, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Tier2, -1, DTsched_Mapping.Greensboro))
        'initialize the lines

        Dim lineName As String

        lineName = "QXAE42"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE42 Converter
        'QXAE43 Converter
        lineName = "QXAE43"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE45 Converter
        lineName = "QXAE45"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE46 Converter
        lineName = "QXAE46"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE47 Converter
        lineName = "QXAE47"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE48 Converter
        lineName = "QXAE48"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE49 Converter
        lineName = "QXAE49"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE61 Converter
        lineName = "QXAE61"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE62 Converter
        lineName = "QXAE62"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE63 Converter
        lineName = "QXAE63"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE64 Converter
        lineName = "QXAE64"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE65 Converter
        lineName = "QXAE65"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE66 Converter
        lineName = "QXAE66"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QXAE96 Converter
        lineName = "QXAE96"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))


        'QRAE71 Converter
        lineName = "QRAE71"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE72 Converter
        lineName = "QRAE72"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE73 Converter
        lineName = "QRAE73"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE74 Converter
        lineName = "QRAE74"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE75 Converter
        lineName = "QRAE75"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE76 Converter
        lineName = "QRAE76"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE77 Converter
        lineName = "QRAE77"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE78 Converter
        lineName = "QRAE78"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE79 Converter
        lineName = "QRAE79"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE80 Converter
        lineName = "QRAE80"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE81 Converter
        lineName = "QRAE81"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE82 Converter
        lineName = "QRAE82"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE84 Converter
        lineName = "QRAE84"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE85 Converter
        lineName = "QRAE85"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))

        'QRAE86 Converter
        lineName = "QRAE86"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 2, 12, 6, lineName & " Converter", ""))


    End Sub

    Private Sub initializeMultiUnitTestLines()
        Dim siteName As String = "Dover MultiUnit"

        AllProductionSites.Add(New ProdSite(siteName, "dvr-mesdatabc2", "", SERVER_PW_V6, SERVER_UN_V6, "DVR"))
        Dim ModuleID As Guid = Guid.NewGuid()
        AllProdModules.Add(New ProdModule(ModuleID, "Baby Wipes", siteName, SECTOR_BABY, prStoryMapping.STRAIGHT, DefaultProficyDowntimeProcedure.OneClick_MultiUnit, DefaultProficyProductionProcedure.NA, "Reason 1", "Reason 2", "Reason 3", "Reason 4", DowntimeField.Reason1, -1, DTsched_Mapping.Greensboro))

        Dim lineName As String

        Dim masterProdUnitList As List(Of String) = New List(Of String)
        masterProdUnitList.Add("QPDO L21 Wrapper")
        masterProdUnitList.Add("QPDO L22 Wrapper")

        lineName = "21 & 22"
        AllProdLines.Add(New ProdLine(lineName, lineName, siteName, ModuleID, 3, 8, 7, lineName & " Wrapper", ""))
        AllProdLines(AllProdLines.Count - 1).mainProdUnits = masterProdUnitList


    End Sub


    Private Sub AddLine(name As String, moduleID As Guid, dtMPU As String, Optional prodMPU As String = "")
        AllProdLines.Add(New ProdLine(name, moduleID, dtMPU))
    End Sub
End Module

