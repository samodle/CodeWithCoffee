Public Module AllEnumConsts
    Public Enum EventType
        Running
        Excluded
        Planned
        Unplanned
    End Enum
    Public Enum DowntimeMetrics
        DTpct
        DT
        SPD
        MTBF
        Stops
        MTTR
        PR
        PDTpct
        UPDTpct
        SKUs
        UnitsProduced
        NumChangeovers
        SchedTime
        Survivability
        Chronicity
        NA
    End Enum

    Public Enum DefaultProficyDowntimeProcedure
        QuickQuery = 0
        OneClick = 1
        Maple = 2
        GLEDS = 3
        RE_CentralServer = 6
        OneClick_V27
        OneClick_MultiUnit
        QuickQuery_MultiUnit
    End Enum
    Public Enum DefaultProficyProductionProcedure
        QuickQuery = 0
        SwingRoad = 1
        Maple = 2
        Maple_New = 3
        QuickQuery_MOT = 4
        NA = 5

    End Enum

    Public Enum MultiConstraintAnalysis
        SingleConstraint = 0
        RateLossAsStops = 1
        NoRateLossStops = 2
    End Enum

    Public Enum RateLossMode
        Interlaced = 0
        Separate = 1
    End Enum

    Public Enum Lang
        English = 0
        German = 1
        Spanish = 2
        French = 3
        Portuguese = 4
        'Chinese_Traditional = 5
        Chinese_Simplified = 5
    End Enum


    Public Enum DTsched_Mapping
        Greensboro = 0
        Phenoix = 1
        SwingRoad = 2
        SkinCare = 3
        HuangpuHC = 4
        APDO = 5
        Belleville = 6
        Ukraine = 7
        Mariscala = 8
        BabyCare = 9
        Albany = 10
        Hyderabad = 11
        Budapest = 12
        Rakona = 13
        SingaporePioneer = 14
        ModPack = 15
        Fem_LuisCustom
        ' APRILFOOLS = 15
    End Enum


    Public Enum MappingByFormat
        SkinCare = 0
        NoMapping = 1
    End Enum
    Public Enum MappingByShape
        SkinCare = 0
        NoMapping = 1
    End Enum
    Public Enum MappingOneClick
        OralCare
        IowaCity
        NoMapping
    End Enum
End Module

Public Module GlobalProficyConsts

    Public Const SQL_PROCEDURE_QUICKQUERY As String = "spLocal_PQQ_DowntimeExplorer"
    Public Const SQL_PROCEDURE_ONECLICK As String = "spLocal_ExtractOneClickData_v026"
    Public Const SQL_PROCEDURE_ONECLICK_FAMILY As String = "spLocal_ExtractOneClickData_v027"

    Public SERVER_PW_QQ As String = "comxclient"
    Public PROFICY_SERVER_USERNAME_QQ As String = "comxclient"

    Public SERVER_PW_V6 As String = "comxclient"
    Public SERVER_UN_V6 As String = "PRStory"

    Public SERVER_UN_MAPLE As String = "One_Click"
    Public SERVER_PW_MAPLE As String = "DarthVader6"

    Public Const PROFICY_SERVER_PASSWORD_SWINGROAD As String = "Reports1"
    Public Const PROFICY_SERVER_USERNAME_SWINGROAD As String = "profreports"

    Public Enum DowntimeField
        startTime = 0
        endTime = 1
        DT = 2
        UT = 3
        MasterProdUnit = 4
        Location = 5
        Fault = 6
        Reason1 = 7
        Reason2 = 8
        Reason3 = 9
        Reason4 = 10
        PR_inout = 11
        Team = 12
        PlannedUnplanned = 13
        DTGroup = 14
        Product = 15
        ProductCode = 16
        Comment = 17
        'MAPPED DATA
        Tier1 = 18
        Tier2 = 19
        Tier3 = 20
        Format = 21
        Shape = 22
        Classification = 23
        OneClick = 24
        Stopclass = 25 ' LG code
        ProductGroup = 26
        'added for rate loss
        '  RateActual = 30
        '  RateTarget = 31
        NA
    End Enum

    Public Enum ProductionColumn
        StartTime = 0
        EndTime = 1
        ProductionUnit = 2
        ProductCode = 3
        Product = 4
        ProductionStatus = 5
        Shift = 6
        Team = 7
        ActualUnits = 8
        ActualCases = 9
        AdjustedCases = 10
        AdjustedUnits = 11
        StatUnits = 12
        ActualRate = 13
        TargetRate = 14
        SchedTime = 15
        UnitsPerCase = 16
        StatConversion = 17
        ' UT = 18
        ' PR = 19
    End Enum

    'Other
    Public Const BLANK_INDICATOR As String = "<BLANK>"
End Module


Module HTML_Consts
    Public Const SERVER_FOLDER_PATH As String = PATH_PRSTORY & "html\" '"C:\Users\Public\"
    Public Const PATH_PRSTORY As String = "C:\Users\Public\prstory\"
    Public Const PATH_PRSTORY_SETTINGS As String = PATH_PRSTORY & "settings\"
    Public Const PATH_PRSTORY_TARGETS As String = PATH_PRSTORY & "targets\"
    Public Const PATH_PRSTORY_RAWDATA As String = PATH_PRSTORY & "rawdata\"

    Public Const FILE_RAWTARGETS_CSV As String = "prstory_dtpct_targets.csv"

    'names for individual HTML files
    Public Const HTML_UPTIME_VIEWER As String = "UptimeViewer"
    Public Const HTML_LOSS_TREE As String = "LossTree"
    Public Const HTML_LOSS_TREEMAP As String = "LossTreeMap"

    'colors
    Public Const HTMLCOLOR_BrightGreen As String = "'00FF00'"
    Public Const HTMLCOLOR_BrightRed As String = "'FF0000'"
    Public Const HTMLCOLOR_BrightYellow As String = "'FFFF00'"
    Public Const HTMLCOLOR_BrightBlue As String = "'0000FF'"
    Public Const HTMLCOLOR_LightGrey As String = "'D8D8D8'"
End Module



Module FixedStrings_for_AmCharts
    Public Const amchartJS As String = ""
End Module


Module FixedPositions_Target_Rectangles
    Public Const tier2_target_rect_top = 480
End Module

Module FixedHeight_StopsWatch_LinearUI_Rectangles
    Public Const linearUI_rectheight = 51
End Module