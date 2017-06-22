Module prstory_Consts
    Public Enum prStoryMapping
        NoMappingAvailable 'no mapping established
        OralCare
        OralCareNau
        OralCareGross
        OralCareCrux
        SwingRoad
        SwingRoad_6
        SwingRoad_7
        APDO_I
        APDO_J
        IowaCity
        OralCare_DF
        OralCareNau_DF
        Mandideep
        Mandideep_Fem
        SkinCare
        Phoenix
        Pheonix_D
        IowaCityBeauty
        HuangPu
        Boryspil
        FemCare_Pads
        FemCare_Pads_Huangpu
        Fem_LCC_HPU
        Mariscala
        Mariscala2
        Albany
        FamilyCareUnitOP_Wrapper
        FamilyCareUnitOP_mf
        FamilyCareUnitOP_ACP
        FamilyCareUnitOP_Palletizer
        FamilyCareUnitOP_ModPACK
        FamilyCareUnitOP_Napkins
        FamilyCareUnitOP_STRETCHWRAPPER
        FamilyCareUnitOP_Stacker
        Puffs
        GENERIC
        Hyderabad
        BudapestFGC
        BudapestLCC
        NaucalpanPHC_B
        NaucalpanPHC_J
        NaucalpanPHC_Mex
        NaucalpanPHC_Vita1
        Rakona
        TepejiFem
        JijonaUltra
        FamilyMaking
        SingaporePioneer
        ICOC
        ICOC_Making
        Rio
        APRILFOOLS
        Fem_LuisCustom
        STRAIGHT 'T1 -> R1, etc
        STRAIGHTPlusOne
        STRAIGHTPLANNEDPlusOne
        STRAIGHTPLANNEDPlusTwo
    End Enum
End Module

Module Mapping_prstory

    Public Sub getSTRAIGHTprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            .Tier1 = .Reason1
            .Tier2 = .Reason2
            .Tier3 = .Reason3
        End With
    End Sub
    Public Sub getSTRAIGHTPLUSONEprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            .Tier1 = .Reason2
            .Tier2 = .Reason3
            .Tier3 = .Fault

            If .Tier2 = BLANK_INDICATOR And Not .Reason1 = BLANK_INDICATOR Then
                .Tier1 = .Reason1
            End If
        End With
    End Sub
    Public Sub getSTRAIGHTPLANNEDPLUSONEprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                .Tier1 = .Reason1
                .Tier2 = .Reason2
                .Tier3 = .Reason3
            Else
                .Tier1 = .Reason2
                .Tier2 = .Team
                .Tier3 = .Fault
            End If
        End With
    End Sub
        Public Sub getSTRAIGHTPLANNEDPLUSTWOprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                .Tier1 = .Reason1
                .Tier2 = .Reason2
                .Tier3 = .Reason3
            Else
                .Tier1 = .Reason3
                .Tier2 = .Team
                .Tier3 = .Fault
            End If
        End With
    End Sub


    Public Sub getFemLCCHPUMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                .Tier2 = .Reason1
                .Tier3 = .Reason2

                If .Location.Equals("Area 1") Or .Location.Equals("AREA 1") Or .Location.Equals("Area1") Then
                    .Tier1 = "Area 1"
                ElseIf .Location.Equals("Area 2") Or .Location.Equals("AREA 2") Or .Location.Equals("Area2") Then
                    .Tier1 = "Area 2"
                ElseIf .Location.Equals("Area 3") Or .Location.Equals("AREA 3") Or .Location.Equals("Area3") Then
                    .Tier1 = "Area 3"
                ElseIf .Location.Equals("Area 4") Or .Location.Equals("AREA 4") Or .Location.Equals("Area4") Or .Location.Equals("Packaging") Then
                    .Tier1 = "Area 4"
                Else
                    .Tier1 = OTHERS_STRING
                End If
            Else
                .Tier2 = .Team
                .Tier3 = .Product

                'REASON 2
                If .Reason2.Contains("CIL") Or .Reason2.Contains("RLS") Then
                    .Tier1 = "CIL"
                ElseIf .Reason2.Contains("ChangeOver") Or .Reason2.Contains("CHANGEOVER") Or .Reason2.Contains("CHANGE OVER") Then
                    .Tier1 = "CO"
                ElseIf .Reason2.Contains("AM") Then
                    .Tier1 = "AM"
                ElseIf .Reason2.Contains("PM") Then
                    .Tier1 = "PM"
                ElseIf .Reason2.Contains("Project Work") Or .Reason2.Contains("PROJECT") Then
                    .Tier1 = "Project"
                ElseIf .Reason2.Contains("Lunch and Break") Then
                    .Tier1 = "Lunch and Break"
                ElseIf .Reason2.Contains("Planned Utility Outages") Then
                    .Tier1 = "Planned Utility Outages"
                ElseIf .Reason2.Contains("Training and Meeting") Or .Reason2.Contains("MEETING") Or .Reason2.Contains("TRAINING") Then
                    .Tier1 = "Training and Meeting"

                    'REASON 1
                ElseIf .Reason1.Contains("CIL") Or .Reason1.Contains("RLS") Then
                    .Tier1 = "CIL"
                ElseIf .Reason1.Contains("CHANGE") Or .Reason1.Equals("Changeover") Or .Reason1.Contains("SIZE CHANGE") Or .Reason1.Contains("Pack CHANGE") Or .Reason1.Contains("PRODUCT CHANGE") Then
                    .Tier1 = "CO"
                ElseIf .Reason1.Contains("ORGANIZATION") Then
                    .Tier1 = "Org"
                ElseIf .Reason1.Contains("PM") Or .Reason1.Contains("MAINTENANCE") Then
                    .Tier1 = "PM"

                ElseIf .Reason1.Equals("EO sellable") Or .Reason1.Contains("EO") Then
                    .Tier1 = "EO"
                ElseIf .Reason1.Contains("MEETING") Or .Reason1.Contains("TRAINING") Then
                    .Tier1 = "Training and Meeting"
                ElseIf .Reason1.Contains("PROJECT") Then
                    .Tier1 = .Reason1
                ElseIf .Reason1.Equals("Logistics") Then
                    .Tier1 = .Reason1


                    'REASON 3
                ElseIf .Reason3.Contains("CIL") Or .Reason3.Contains("RLS") Then
                    .Tier1 = "CIL"
                ElseIf .Reason3.Contains("ChangeOver") Or .Reason3.Contains("CHANGEOVER") Or .Reason3.Contains("CHANGE OVER") Then
                    .Tier1 = "CO"
                ElseIf .Reason3.Contains("AM") Then
                    .Tier1 = "AM"
                ElseIf .Reason3.Contains("PM") Then
                    .Tier1 = "PM"

                ElseIf .Reason3.Contains("PLANNED INTERVENTION") Then
                    .Tier1 = "PLANNED INTERVENTION"
                ElseIf .Reason3.Contains("Training") Then
                    .Tier1 = "Training and Meeting"

                ElseIf .Reason3.Contains("Product feed Max. level") Then
                    .Tier1 = "PLANNED INTERVENTION"


                Else
                    .Tier1 = OTHERS_STRING
                End If

            End If
        End With
    End Sub


#Region "FamilyCare"

    Public Sub getPuffsMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            .Tier1 = "Converter"
            .Tier2 = OTHERS_STRING
            .Tier3 = .Reason2

            If .isUnplanned Then
                If .Reason1.Contains("Special Causes") Then
                    If .Reason2.Contains("EO") Then
                        .Tier1 = "Converter"
                        .Tier2 = "Other Special Causes"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Reason1.Contains("Special Causes") Then
                    .Tier1 = "Converter"
                    .Tier2 = "Unscheduled Time"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason2.Contains("QP") Then
                    .Tier1 = "Converter"
                    .Tier2 = "ELP"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason2.Contains("Palletizer") Then
                    .Tier1 = "Blocked-Starved"
                    .Tier2 = "FPH"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Fault.Contains("Diverter") Then
                    .Tier1 = "Blocked-Starved"
                    .Tier2 = "Diverter"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Fault.Contains("LCP") Then
                    .Tier1 = "Blocked-Starved"
                    .Tier2 = "LCP"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Warehouse") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "FPH"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Casepacker") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Casepackers"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("EYE BLOCKED") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("PEC BLOCKED") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("[2050] CONVEYOR JAM - BTCs") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("[2682]-Jam Eye Blocked - BTU") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Loader/Bucket - Servo Fault (Cn)") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Loader/Bucket - Servo Fault (Cs)") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Overload Clutch - Transfer Belt Cn") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("PCMC Loader - Torque Limit (Cn)") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Bundle Quality"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Downstream Not Ready - Down Stream Cn") Then
                        If .Reason2.Contains("Downstream") Then
                            .Tier1 = "Blocked-Starved"
                            .Tier2 = "FPH"
                            .Tier3 = .Reason2
                            Exit Sub
                        End If
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Downstream Not Ready - Down Stream Cs") Then
                        If .Reason2.Contains("Downstream") Then
                            .Tier1 = "Blocked-Starved"
                            .Tier2 = "FPH"
                            .Tier3 = .Reason2
                            Exit Sub
                        End If
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Downstream Not Ready - Down Stream Cn") Then
                        If .Reason2.Contains("Cartoner") Then
                            .Tier1 = "Blocked-Starved"
                            .Tier2 = "Cartoners"
                            .Tier3 = .Reason2
                            Exit Sub
                        End If
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Fault.Contains("Downstream Not Ready - Down Stream Cs") Then
                        If .Reason2.Contains("Casepacker") Then
                            .Tier1 = "Blocked-Starved"
                            .Tier2 = "Casepackers"
                            .Tier3 = .Reason2
                            Exit Sub
                        End If
                    End If
                End If

                If .Location.Contains("1") Then
                    If .Reason2.Contains("BTU") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "BTU"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7a") Then
                    If .Reason2.Contains("CART") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Cartoners"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7a") Then
                    If .Reason2.Contains("ACP") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Casepackers"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7a") Then
                    If .Reason1.Contains("UWS99 Blocked") Then
                        If .Reason2.Contains("2x8 Down") Then
                            .Tier1 = "Blocked-Starved"
                            .Tier2 = "Diverter"
                            .Tier3 = .Reason2
                            Exit Sub
                        End If
                    End If
                End If

                If .Reason1.Contains("UWS99 Blocked") Then
                    If .Reason2.Contains("BTU") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "BTU"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Reason2.Contains("CART") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Cartoners"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Reason2.Contains("ACP") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Casepackers"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Fault.Contains("Softpack Wrapper") Then
                        .Tier1 = "Softpack"
                        .Tier2 = "Softpack Wrapper"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Fault.Contains("Softpack Erector") Then
                        .Tier1 = "Softpack"
                        .Tier2 = "Softpack Erector"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Fault.Contains("Softpack Case Sealer") Then
                        .Tier1 = "Softpack"
                        .Tier2 = "Softpack Casesealer"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Fault.Contains("Softpack") Then
                        .Tier1 = "Softpack"
                        .Tier2 = "Softpack Other"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Fault.Contains("[3]-NCart & SCart Down") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Cartoners"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Fault.Contains("[4]-ECSPR & WCSPR Down") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Casepackers"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("7") Then
                    If .Fault.Contains("[7] 2x8 Down") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "Diverter"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Reason2.Contains("Cartoner") Then
                    .Tier1 = "Blocked-Starved"
                    .Tier2 = "Cartoners"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason2.Contains("Casepacker") Then
                    .Tier1 = "Blocked-Starved"
                    .Tier2 = "Casepackers"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Location.Contains("Converter") Then
                    If .Reason1.Contains("Blocked") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "UWS Blocked"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("Converter") Then
                    .Tier1 = "Converter"
                    .Tier2 = "UWS"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Location.Contains("UWS") Then
                    If .Reason1.Contains("Blocked") Then
                        .Tier1 = "Blocked-Starved"
                        .Tier2 = "UWS Blocked"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("UWS") Then
                    .Tier1 = "Converter"
                    .Tier2 = "UWS"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Location.Contains("Pack Belt") Then
                    .Tier1 = "Converter"
                    .Tier2 = "UWS"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Location.Contains("Saw") Then
                    .Tier1 = "Converter"
                    .Tier2 = "Saw"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason2.Contains("Uncoded") Then
                    .Tier1 = "Converter"
                    .Tier2 = "Uncoded"
                    .Tier3 = .Reason2
                    Exit Sub
                End If



            Else ' planned
                If .Reason1.Contains("Planned Intervention") Then
                    If .Reason2.Contains("CIL") Then
                        .Tier1 = "CIL"
                        .Tier2 = "RL1"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    If .Reason2.Contains("Main") Then
                        .Tier1 = "Planned Maintenance"
                        .Tier2 = "RL1"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Reason2.Contains("Blowdown") Then
                    .Tier1 = "RLS"
                    .Tier2 = "RL1"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason2.Contains("RLS") Then
                    .Tier1 = "RLS"
                    .Tier2 = "RL1"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Contains("Special Causes") Then
                    If .Reason2.Contains("Meetings") Then
                        .Tier1 = "Meetings"
                        .Tier2 = "RL1"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Reason2.Contains("Changeover") Then
                    .Tier1 = "Changeover"
                    .Tier2 = "RL1"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    .Tier1 = "Planned Maintenance"
                    .Tier2 = "RL1"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Location.Contains("Converter") Then
                    If .Reason1.Contains("PRC") Then
                        .Tier1 = "Parent Roll Change"
                        .Tier2 = "RL1"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If

                If .Location.Contains("UWS") Then
                    If .Reason1.Contains("PRC") Then
                        .Tier1 = "Parent Roll Change"
                        .Tier2 = "RL1"
                        .Tier3 = .Reason2
                        Exit Sub
                    End If
                End If
            End If
        End With
    End Sub



    Public Sub getAlbanyprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason1.Contains("Blocked") Or .Reason1.Contains("Starved") Or .Reason1.Contains("downstream") Then
                    .Tier1 = "Block/Starved"
                    .Tier2 = .Reason2
                    If .Reason2.Contains("Stacker") Then
                        .Tier3 = "WH-Palletizer"
                    ElseIf .Reason2.Contains("Log") Then
                        .Tier3 = "Logsaw"
                    ElseIf .Reason2.Contains("Pallet") Then
                        .Tier3 = "Palletizer"
                    ElseIf .Reason2.Contains("Stretchwrapper") Then
                        .Tier3 = "Palletizer"
                    ElseIf .Reason2.Contains("Wrapper") Then
                        .Tier3 = "Wrapper"
                    ElseIf .Reason2.Contains("Multi") Then
                        .Tier3 = "MF-Bundler"
                    ElseIf .Reason2.Contains("bund") Then
                        .Tier3 = "ACP"
                    ElseIf .Reason2.Contains("ACP") Then
                        .Tier3 = "ACP"
                    ElseIf .Reason2.Contains("casepacker") Then
                        .Tier3 = "ACP"
                    ElseIf .Reason2.Contains("Case Conveyor") Then
                        .Tier3 = "Full Case Conveyor"
                    ElseIf .Reason2.Contains("Package Conveyor") Then
                        .Tier3 = "Pkg Conveyor"
                    ElseIf .Reason2.Contains("coupon") Then
                        .Tier3 = "Couponer"
                    ElseIf .Reason2.Contains("No Roll Delivered") Then
                        .Tier3 = "No Roll Delivered"
                    ElseIf .Reason2.Contains("No Paper") Then
                        .Tier3 = "No Paper"
                    Else
                        .Tier3 = "Others"
                    End If



                    .Tier3 = .Fault
                ElseIf .Reason2.Contains("QP") Then
                    .Tier1 = "ELP"
                    .Tier2 = "ELP"
                    If Left(.Reason2, 2).Equals("QP") Then
                        .Tier3 = Right(.Reason2, Len(.Reason2) - 3)

                    Else
                        .Tier2 = .Reason2
                    End If
                    '.Tier3 = .Reason1
                ElseIf .Reason1.Contains("UWS05") Then
                    If .Reason2.Contains("reject") Then
                        .Tier1 = "ELP"
                        .Tier2 = "ELP"
                        .Tier3 = "QP Rejected Roll"

                    End If
                ElseIf .Reason1.Contains("Blocked") Then
                    If .Reason2.Contains("paper") Then
                        .Tier1 = "ELP"
                        .Tier2 = "ELP"
                        .Tier3 = "QP paper in TS"

                    End If
                    'DSFA#OR#Magnet#OR#Karlinal
                ElseIf .Location.Contains("DSFA") Or .Location.Contains("Magnet") Or .Location.Contains("Karlinal") Then
                    .Tier1 = "Materials"
                    If .Reason1.Contains("dirty") Or .Reason1.Contains("Hygiene") Or .Reason2.Contains("dirty") Or .Reason2.Contains("Hygiene") Then
                        .Tier2 = "Hyg-SurfaceAdditives"
                        .Tier3 = .Fault

                    ElseIf .Reason1.Contains("Electrical") Then
                        .Tier2 = "EBD SurfaceAdditives"
                        .Tier3 = .Reason2
                    ElseIf .Reason1.Contains("Mechanical") Then
                        .Tier2 = "MBD SurfaceAdditives"
                        .Tier3 = .Reason2
                    Else
                        .Tier2 = .Reason1
                        .Tier3 = .Reason2
                    End If
                    '   ElseIf .PlannedUnplanned.Contains("Raw Material") Then
                    '       .Tier1 = "Materials"
                    '       .Tier2 = .Reason1
                    '       .Tier3 = .Reason2
                ElseIf .Reason1.Contains("AUX") Or .Reason1.Contains("Util") Then
                    .Tier1 = "Utilities"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                ElseIf .Location.Contains("Winder") Then
                    .Tier1 = "Converter"
                    .Tier2 = "Winder"
                    If .Reason1.Contains("dirty") Or .Reason1.Contains("Hygiene") Or .Reason2.Contains("dirty") Or .Reason2.Contains("Hygiene") Then
                        .Tier3 = "Hygiene"
                    ElseIf .Reason1.Contains("Electrical") Then
                        .Tier3 = "EBD"

                    ElseIf .Reason1.Contains("Mechanical") Then
                        .Tier3 = "MBD"

                    Else
                        If .Reason1.Contains("WND") Then
                            .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                        Else
                            .Tier3 = .Reason1
                        End If
                    End If
                ElseIf Left(.Reason1, 3) = "WND" Then
                    .Tier1 = "Converter"
                    .Tier2 = "Winder"
                    .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                    If .Reason2.Contains("Slitter") Then
                        .Tier3 = "Slitter Problems"
                    ElseIf .Reason2.Contains("Core") Then
                        .Tier3 = "Core Problems"
                    ElseIf .Reason2.Contains("Troubleshooting") Then
                        .Tier3 = "Troubleshoot-CL-Adjust"
                    ElseIf .Reason2.Contains("Re-centerlining") Then
                        .Tier3 = "Troubleshoot-CL-Adjust"
                    ElseIf .Reason2.Contains("Set-up error") Then
                        .Tier3 = "Troubleshoot-CL-Adjust"
                    ElseIf .Reason2.Contains("Adjust") Then
                        .Tier3 = "Troubleshoot-CL-Adjust"
                    ElseIf .Reason2.Contains("out of timing") Then
                        .Tier3 = "Troubleshoot-CL-Adjust"
                    ElseIf .Reason2.Contains("knowledge") Then
                        .Tier3 = "Lack of Knowledge"
                    Else
                        '  .Tier3 = "WND Others"
                    End If
                ElseIf .Location.Contains("Emb") Or .Location.Contains("UES") Or .Location.Contains("Combiner") Then
                    .Tier1 = "Converter"
                    .Tier2 = "E-L"

                    If .Reason1.Contains("dirty") Or .Reason1.Contains("Hygiene") Or .Reason2.Contains("dirty") Or .Reason2.Contains("Hygiene") Then
                        .Tier3 = "Hygiene"
                    ElseIf .Reason1.Contains("Electrical") Then
                        .Tier3 = "EBD"

                    ElseIf .Reason1.Contains("Mechanical") Then
                        .Tier3 = "MBD"

                    Else
                        If .Reason1.Contains("E_L") Then
                            .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                        Else
                            .Tier3 = .Reason1
                        End If
                    End If
                ElseIf Left(.Reason1, 3) = "E_L" Or Left(.Reason1, 3) = "CMB" Or Left(.Reason1, 3) = "UES" Then
                    .Tier1 = "Converter"
                    .Tier2 = "E_L"
                    .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                    If .Reason1.Contains("CMB30") Or .Reason1.Contains("E_L30") Then
                        .Tier3 = "Web Loss"
                    ElseIf .Reason1.Contains("UES01") Then
                        .Tier3 = "Adjust UES DOE"
                    ElseIf .Reason1.Contains("E_L40") Or .Reason1.Contains("E_L42") Or .Reason1.Contains("E_L44") Or .Reason1.Contains("E_L46") Or .Reason1.Contains("E_L48") Or .Reason1.Contains("E_L50") Or .Reason1.Contains("E_L60") Then
                        .Tier3 = "Roll Wrap"
                    Else
                        '.Tier3 = "E_L-CMB Others"
                    End If
                ElseIf Left(.Reason1, 3) = "MCD" Then
                    .Tier1 = "Converter"
                    .Tier2 = "MCD"
                    .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                    If .Reason1.Contains("MCD30") Then
                        .Tier3 = "Web Loss"
                    ElseIf .Reason1.Contains("MCD32") Or .Reason1.Contains("MCD40") Then
                        .Tier3 = "Roll Wrap"
                    End If

                ElseIf Left(.Reason1, 3) = "UWS" Then
                    .Tier1 = "Converter"
                    .Tier2 = "UWS"
                    .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                    If .Reason1.Contains("UWS30") Then
                        .Tier3 = "Web Loss"
                    ElseIf .Reason1.Contains("UWS34") Then
                        .Tier3 = "Roll Wrap"

                    ElseIf .Reason1.Contains("UWS36") Then
                        .Tier3 = "Splice Failure"
                    End If
                ElseIf Left(.Reason1, 3) = "TSL" Then
                    .Tier1 = "Converter"
                    .Tier2 = "TSL"
                    .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                ElseIf Left(.Reason1, 3) = "ACC" Then
                    .Tier1 = "Converter"
                    .Tier2 = "ACC"
                    .Tier3 = Right(.Reason1, Len(.Reason1) - 6)
                    '  ElseIf .Reason1.Contains() Then
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
                If .Tier3 <> "" And .Tier3 <> BLANK_INDICATOR Then
                    .DTGroup = .Tier1 & "-" & .Tier2 & "-" & .Tier3
                Else
                    .DTGroup = .Tier1 & "-" & .Tier2
                End If

            Else
                If .Reason2.Contains("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .DTGroup.Equals("Changeover") Then
                    .Tier1 = "CO"
                    .Tier3 = .Reason3
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("UWS05") Then
                    .Tier1 = "PR Changes"
                    .Tier3 = .Reason3
                    .Tier2 = .Team
                ElseIf .Reason2.Equals("Blowdown") Then
                    .Tier1 = .Reason2
                    .Tier3 = .Reason3
                    .Tier2 = .Team
                ElseIf .Reason2.Contains("Centerline") Then
                    .Tier1 = "AM"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .Reason2.Contains("Maintenance") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Team
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub

    Public Sub getFamilyCareUnitOP_Stacker(ByRef searchevent As DowntimeEvent)
        With searchevent
            .Tier1 = "Other"
            .Tier2 = "Other"
            .Tier3 = BLANK_INDICATOR


            If .isUnplanned Then
                If .Fault.Contains("FAULTS CLEARED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Adjust") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Pallet Debris") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Unit Fell") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Tracking") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Troubleshooting") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    If .Reason2.Contains("Training") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("MANUAL MODE") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("SCHMERSAL PB REQUEST DOOR OPEN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0003 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0004 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0005 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0006 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("AUTO MODE NOT STARTED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("BYPASS MODE NOT STARTED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET RETAINER NOT IN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET RETAINER NOT BACK") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET RETAINER NOT IN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET RETAINER NOT BACK") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT LOAD DETECT NOT CLEARING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT LOAD DETECT NOT CLEARING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("NO LOAD DETECTED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET STOP NOT IN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET STOP NOT BACK") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET STOP NOT IN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET STOP NOT BACK") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("PALLET CENTERING NOT IN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("PALLET CENTERING NOT BACK") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0022 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0023 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0024 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0025 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0026 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0027 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0028 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LOAD IN POSITION PEC NOT BLOCKED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LOAD IN POSITION PEC NOT CLEARING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0031 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0032 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("DISC RELAY DOES NOT MATCH DISCONNECT") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("E-STOP RELAY NOT ENERGIZED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0035 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("REMOTE E-STOP PB FAULT") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LOCAL E-STOP PB FAULT") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("MAIN MCP DISCONNECT OFF") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0039 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LOW AIR PRESSURE") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT SAFETY PIN NOT BACK") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT SAFETY PIN NOT BACK") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT CHAIN SPRING NOT HOME") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT CHAIN SPRING NOT HOME") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("CARRIAGE LIMIT PROX WITHOUT DECEL PROX SENSED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("BOTH CARRIAGE DECEL PROXES SENSED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0047 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("TOP OF LOAD NOT FOUND") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT LOAD DETECT NOT WORKING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT LOAD DETECT NOT WORKING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0051 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LOWER LIMIT SWITCH NOT SWITCHING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("UPPER LIMIT SWITCH NOT SWITCHING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("HAS A TALL LOAD") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("HAS MISMATCHED LOADS") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0056 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0057 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0058 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("MCR OR CONTACTOR STUCK ON") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INFEED LIGHT CURTAIN TRIPPED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT LIGHT CURTAIN TRIPPED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("E-STOP RELAY NOT CLOSING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("E-STOP RELAY STUCK CLOSED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT ACCESS DOOR IS OPEN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT ACCESS DOOR IS OPEN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("ROSS AIR VALVE FAULT - RESET AT VALVE") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("CARRIAGE VFD FAULT") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET RETAINER IN PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET RETAINER BACK PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET RETAINER IN PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET RETAINER BACK PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET STOP IN PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT PALLET STOP BACK PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET STOP IN PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT PALLET STOP BACK PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("PALLET CENTERING IN PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("PALLET CENTERING BACK PROX NOT CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT SAFETY PIN PROX NOT CLEARING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT SAFETY PIN PROX NOT CLEARING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("2ND LOAD HEIGHT DETECT PEC IS ALWAYS CLEAR") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LEFT SIDE FULL LAYER PEC202E NOT BLOCKED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("RIGHT SIDE FULL LAYER PEC202F NOT BLOCKED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0083 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("UNSTABLE LOAD FAULT") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("UNKNOWN LOAD FAULT") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0086 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("MAX STACK LIMIT REACHED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("LIFT THERMAL OVERLOAD DETECTED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT 2 VFD FAULTED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT 2 CONVEYOR VFD COMM FAULT") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("GUARD 1A - LS202B - UPSTREAM SLIDING DOOR OPEN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("GUARD 1B - LS202C - DOWNSTREAM SLIDING DOOR OPEN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("GUARD 2 - LS202D - DISCHARGE XFER DOOR OPEN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("GUARD 5 - LS202E - DRIVE SIDE/FPC ACCESS DOOR OPEN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT 2 / LABELER CONVEYOR JAM") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0096 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INFEED DECISION TAKING TOO LONG") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT DECISION TAKING TOO LONG") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0099 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INFEED SIDESHIFT 1 M202AA NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 1 M202AA") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 1 M202AA") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 1 M202AA") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 1 M202AA") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 1 M202AA") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 1 M202AA") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INFEED SIDESHIFT 2 M202AB NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 2 M202AB") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 2 M202AB") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 2 M202AB") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 2 M202AB") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 2 M202AB") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON INFEED SIDESHIFT 2 M202AB") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("STACK ZONE M202AC NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("JAM ON STACK ZONE CONV M202AC") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON STACK ZONE CONV M202AC") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON STACK ZONE CONV M202AC") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON STACK ZONE CONV M202AC") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON STACK ZONE CONV M202AC") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON STACK ZONE CONV M202AC") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT SIDESHIFT ROLLS M202AD NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("JAM ON EXIT 1 SIDESHIFT ROLLS M202AD") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT 1 SIDESHIFT ROLLS M202AD") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT 1 SIDESHIFT ROLLS M202AD") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT 1 SIDESHIFT ROLLS M202AD") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT 1 SIDESHIFT ROLLS M202AD") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT 1 SIDESHIFT ROLLS M202AD") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT TRANSFER ROLLS M202AE NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER ROLLS M202AE") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER ROLLS M202AE") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER ROLLS M202AE") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER ROLLS M202AE") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER ROLLS M202AE") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER ROLLS M202AE") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT TRANSFER CHAINS M202AF NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER CHAINS M202AF") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER CHAINS M202AF") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER CHAINS M202AF") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER CHAINS M202AF") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER CHAINS M202AF") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("JAM ON EXIT TRANSFER CHAINS M202AF") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0112 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0113 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0114 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0115 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0116 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0117 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0118 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0119 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CHAINS NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 JAM ON CHAINS") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 JAM ON CHAINS") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 JAM ON CHAINS") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 JAM ON CHAINS") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 JAM ON CHAINS") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 JAM ON CHAINS") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CHAINS NOT UP") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CHAINS NOT DOWN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 PALLET STOP NOT UP") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 PALLET STOP NOT DOWN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CONV 2 NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CONV 2 JAM") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CONV 2 JAM") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CONV 2 JAM") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CONV 2 JAM") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CONV 2 JAM") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 1 CONV 2 JAM") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0128 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0129 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CHAINS NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 JAM ON CHAINS") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 JAM ON CHAINS") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 JAM ON CHAINS") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 JAM ON CHAINS") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 JAM ON CHAINS") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 JAM ON CHAINS") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CHAINS NOT UP") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CHAINS NOT DOWN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 PALLET STOP NOT UP") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 PALLET STOP NOT DOWN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CONV 2 NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CONV 2 JAM") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CONV 2 JAM") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CONV 2 JAM") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CONV 2 JAM") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CONV 2 JAM") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("INF SIDESHIFT 2 CONV 2 JAM") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0138 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0139 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT SIDESHIFT CHAINS NOT RUNNING") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT SIDESHIFT JAM ON CHAINS") Then
                    If .Reason2.Contains("Pallet Skewed") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("EXIT SIDESHIFT JAM ON CHAINS") Then
                    If .Reason2.Contains("Load Not Centered") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("EXIT SIDESHIFT JAM ON CHAINS") Then
                    If .Reason2.Contains("Poly Tails") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("EXIT SIDESHIFT JAM ON CHAINS") Then
                    If .Reason2.Contains("Fell Over") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("EXIT SIDESHIFT JAM ON CHAINS") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Other"
                        .Tier2 = "Other"
                        .Tier3 = .Fault
                    End If
                End If

                If .Fault.Contains("EXIT SIDESHIFT JAM ON CHAINS") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT SIDESHIFT CHAINS NOT UP") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT SIDESHIFT CHAINS NOT DOWN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT SIDESHIFT PALLET STOP NOT UP") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("EXIT SIDESHIFT PALLET STOP NOT DOWN") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0146 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0147 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0148 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0149 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0150 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0151 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0152 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0153 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0154 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0155 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0156 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0157 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0158 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If

                If .Fault.Contains("#0159 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = .Fault
                End If



            Else 'Planned
                If .Reason1.Contains("GEN01") Then
                    If .Reason2.Contains("RLS") Then
                        .Tier1 = "RLS/CIL"
                        .Tier2 = .Team
                        .Tier3 = BLANK_INDICATOR
                    End If
                End If

                If .Reason1.Contains("GEN01") Then
                    If .Reason2.Contains("CIL") Then
                        .Tier1 = "RLS/CIL"
                        .Tier2 = .Team
                        .Tier3 = BLANK_INDICATOR
                    End If
                End If

                If .Reason1.Contains("GEN01") Then
                    If .Reason2.Contains("Maintenance") Then
                        .Tier1 = "MAINTENANCE"
                        .Tier2 = .Team
                        .Tier3 = BLANK_INDICATOR
                    End If
                End If

                If .Reason1.Contains("GEN01") Then
                    .Tier1 = "OTHER"
                    .Tier2 = .Team
                    .Tier3 = BLANK_INDICATOR
                End If

                If .Reason1.Contains("GEN02") Then
                    .Tier1 = "CHANGEOVER"
                    .Tier2 = .Team
                    .Tier3 = BLANK_INDICATOR
                End If
            End If
        End With
    End Sub

    Public Sub getFamilyCareUnitOP_Stretchwrapper(ByRef searchevent As DowntimeEvent)
        With searchevent
            .Tier1 = "Other"
            .Tier2 = "Other"
            .Tier3 = BLANK_INDICATOR

            If .isUnplanned Then
                If .Fault.Contains("ACCESS GATE 1 OPEN / UNLOCKED") Then
                    .Tier1 = "Operator Initiated"
                    .Tier2 = "Guard Door"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0002 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0003 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0004 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0005 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0006 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0007 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0008 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0009 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0010 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ACCESS GATE 1 FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Guard Door"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0012 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0013 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0014 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0015 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0016 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0017 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0018 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0019 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0020 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0021 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0022 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0023 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0024 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0025 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0026 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0027 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0028 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0029 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0030 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0031 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0032 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("PANEL E-STOP PB PRESSED") Then
                    .Tier1 = "Operator Initiated"
                    .Tier2 = "E-Stop PB"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0034 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("REMOTE E-STOP PB PRESSED") Then
                    .Tier1 = "Operator Initiated"
                    .Tier2 = "E-Stop PB"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0036 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0037 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0038 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0039 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0040 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0041 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0042 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0043 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0044 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0045 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0046 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0047 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0048 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0049 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0050 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0051 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0052 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0053 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0054 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0055 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0056 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0057 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0058 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0059 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0060 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0061 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0062 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0063 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0064 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("INFEED LIGHT CURTAIN TRIPPED") Then
                    If .Reason2.Contains("Pallet Debris") Then
                        .Tier1 = "Infeed Conveyor"
                        .Tier2 = "Pallet Debris"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("INFEED LIGHT CURTAIN TRIPPED") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Infeed Conveyor"
                        .Tier2 = "Bad Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("INFEED LIGHT CURTAIN TRIPPED") Then
                    If .Reason2.Contains("Unit Fell") Then
                        .Tier1 = "Infeed Conveyor"
                        .Tier2 = "Unit Fell Over_Tipped"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("INFEED LIGHT CURTAIN TRIPPED") Then
                    .Tier1 = "Infeed Conveyor"
                    .Tier2 = "Light Curtain Fault"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("DISCHARGE LIGHT CURTAIN TRIPPED") Then
                    If .Reason2.Contains("Pallet Debris") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Pallet Debris"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE LIGHT CURTAIN TRIPPED") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Bad Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE LIGHT CURTAIN TRIPPED") Then
                    If .Reason2.Contains("Unit Fell") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Unit Fell Over_Tipped"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE LIGHT CURTAIN TRIPPED") Then
                    If .Reason2.Contains("QN Poly Tails") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Poly Tails on Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE LIGHT CURTAIN TRIPPED") Then
                    .Tier1 = "Exit Conveyor"
                    .Tier2 = "Light Curtain Fault"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0067 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0068 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0069 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0070 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0071 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0072 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0073 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0074 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0075 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0076 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0077 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0078 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0079 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0080 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0081 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0082 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0083 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0084 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0085 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0086 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0087 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0088 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0089 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0090 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0091 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0092 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0093 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0094 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0095 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0096 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("NO 24VDC INPUT POWER") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("NO 24VDC OUTPUT POWER") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("NO 24VDC LIGHT CURTAIN POWER") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0164 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0165 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0166 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0167 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0168 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("SAFETY CIRCUIT FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("SAFETY RELAY FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("SAFETY CONTACTOR FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("SAFETY ROSS VALVE FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0173 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0174 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0175 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0176 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("LOW AIR PRESSURE") Then
                    .Tier1 = "Other"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0178 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0179 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0180 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0181 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0182 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0183 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0184 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ELEVATOR FLEX I/O COMM FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ELEVATOR ETHERNET I/O ARMORBLOCK COMMUNICATION FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0187 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0188 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0189 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0190 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0191 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0192 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0193 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0194 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0195 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("INFEED JAM") Then
                    If .Reason2.Contains("Pallet Debris") Then
                        .Tier1 = "Infeed Conveyor"
                        .Tier2 = "Pallet Debris"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("INFEED JAM") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Infeed Conveyor"
                        .Tier2 = "Bad Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("INFEED JAM") Then
                    If .Reason2.Contains("Unit Fell") Then
                        .Tier1 = "Infeed Conveyor"
                        .Tier2 = "Unit Fell Over_Tipped"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("INFEED JAM") Then
                    .Tier1 = "Infeed Conveyor"
                    .Tier2 = "Jam Detected"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("DISCHARGE JAM") Then
                    If .Reason2.Contains("Pallet Debris") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Pallet Debris"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE JAM") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Bad Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE JAM") Then
                    If .Reason2.Contains("Unit Fell") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Unit Fell Over_Tipped"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE JAM") Then
                    If .Reason2.Contains("QN Poly Tails") Then
                        .Tier1 = "Exit Conveyor"
                        .Tier2 = "Poly Tails on Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("DISCHARGE JAM") Then
                    .Tier1 = "Exit Conveyor"
                    .Tier2 = "Jam Detected"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("UPSTREAM CONVEYOR JAM") Then
                    .Tier1 = "Infeed Conveyor"
                    .Tier2 = "Jam Detected"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0199 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("TOP FILM POST NOT LATCHED") Then
                    .Tier1 = "Film Tension Assembly"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0201 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0202 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0203 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0204 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0205 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0206 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0207 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0208 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0209 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0210 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0211 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0212 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0213 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0214 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0215 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0216 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ELEVATOR CHAIN SLACK FAULT") Then
                    .Tier1 = "Elevator"
                    .Tier2 = "Mechanical"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0218 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0219 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0220 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0221 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0222 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0223 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0224 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("FILM BREAK AT CLAMP") Then
                    If .Reason1.Contains("STR11") Then
                        .Tier1 = "Incoming Quality -Poly"
                        .Tier2 = "Film Break"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("FILM BREAK AT CLAMP") Then
                    .Tier1 = "Film Clamp"
                    .Tier2 = "Film Break"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("FILM BREAK AT LOAD") Then
                    If .Reason1.Contains("STR11") Then
                        .Tier1 = "Incoming Quality -Poly"
                        .Tier2 = "Film Break"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("FILM BREAK AT LOAD") Then
                    .Tier1 = "Film Tension Assembly"
                    .Tier2 = "Film Break"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0227 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0228 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0229 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0230 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0231 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0232 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0233 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0234 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0235 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0236 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0237 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0238 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0239 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0240 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0241 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0242 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0243 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0244 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0245 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0246 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0247 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0248 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0249 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0250 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0251 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0252 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0253 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0254 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0255 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0256 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION VFD COMM FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION VFD FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION BRAKING RESISTOR FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION AUTOTUNE REQUIRED") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0389 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION OVERSPEED FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION NOT AT HOME") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION JAM") Then
                    If .Reason2.Contains("Pallet Debris") Then
                        .Tier1 = "Pallet Turner"
                        .Tier2 = "Pallet Debris"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("ROTATION JAM") Then
                    If .Reason2.Contains("Bad Pallet") Then
                        .Tier1 = "Pallet Turner"
                        .Tier2 = "Bad Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("ROTATION JAM") Then
                    If .Reason2.Contains("Unit Fell") Then
                        .Tier1 = "Pallet Turner"
                        .Tier2 = "Unit Fell Over_Tipped"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("ROTATION JAM") Then
                    If .Reason2.Contains("QN Poly Tails") Then
                        .Tier1 = "Pallet Turner"
                        .Tier2 = "Poly Tails on Pallet"
                        .Tier3 = "fault code"
                    End If
                End If

                If .Fault.Contains("ROTATION JAM") Then
                    .Tier1 = "Pallet Turner"
                    .Tier2 = "Jam Detected"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("ROTATION PROX FAULT") Then
                    .Tier1 = "Electrical"
                    .Tier2 = "Other"
                    .Tier3 = "fault code"
                End If

                If .Fault.Contains("#0394 NOT DEFINED") Then
                    .Tier1 = "Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = "fault code"
                End If



            Else 'Planned
                If .Reason1.Contains("GEN01") Then
                    If .Reason2.Contains("RLS") Then
                        .Tier1 = "RLS/CIL"
                        .Tier2 = .Team
                        .Tier3 = "N/A"
                    End If
                End If

                If .Reason1.Contains("GEN01") Then
                    If .Reason2.Contains("CIL") Then
                        .Tier1 = "RLS/CIL"
                        .Tier2 = .Team
                        .Tier3 = "N/A"
                    End If
                End If

                If .Reason1.Contains("GEN01") Then
                    If .Reason2.Contains("Maintenance") Then
                        .Tier1 = "MAINTENANCE"
                        .Tier2 = .Team
                        .Tier3 = "N/A"
                    End If
                End If

                If .Reason1.Contains("GEN01") Then
                    .Tier1 = "OTHER"
                    .Tier2 = .Team
                    .Tier3 = "N/A"
                End If

                If .Reason1.Contains("GEN02") Then
                    .Tier1 = "CHANGEOVER"
                    .Tier2 = .Team
                    .Tier3 = "N/A"
                End If

                If .Reason1.Contains("GEN03 #OR# GEN04") Then
                    .Tier1 = "OTHER"
                    .Tier2 = .Team
                    .Tier3 = "N/A"
                End If

                If .Fault.Contains("Blocked") Then
                    .Tier1 = "COUNT UPTIME, IGNORE DOWNTIME"
                    .Tier2 = .Team
                    .Tier3 = "N/A"
                End If

                If .Fault.Contains("Starved") Then
                    .Tier1 = "COUNT UPTIME, IGNORE DOWNTIME"
                    .Tier2 = .Team
                    .Tier3 = "N/A"
                End If

                If .Fault.Contains("FILM BREAK AT CLAMP") Then
                    If .Reason1.Contains("Poly Roll Change") Then
                        If .Reason2.Contains("Reject") Then
                            .Tier1 = "POLY CHANGE"
                            .Tier2 = .Team
                            .Tier3 = "N/A"
                        End If
                    End If
                End If

                If .Fault.Contains("FILM BREAK AT CLAMP") Then
                    If .Reason1.Contains("Poly Roll Change") Then
                        If .Reason2.Contains("End Of Roll") Then
                            .Tier1 = "POLY CHANGE"
                            .Tier2 = .Team
                            .Tier3 = "N/A"
                        End If
                    End If
                End If

                If .Fault.Contains("FILM BREAK AT LOAD") Then
                    If .Reason1.Contains("Poly Roll Change") Then
                        If .Reason2.Contains("Reject") Then
                            .Tier1 = "POLY CHANGE"
                            .Tier2 = .Team
                            .Tier3 = "N/A"
                        End If
                    End If
                End If

                If .Fault.Contains("FILM BREAK AT LOAD") Then
                    If .Reason1.Contains("Poly Roll Change") Then
                        If .Reason2.Contains("End Of Roll") Then
                            .Tier1 = "POLY CHANGE"
                            .Tier2 = .Team
                            .Tier3 = "N/A"
                        End If
                    End If
                End If
            End If
        End With
    End Sub


    Public Sub getFamilyCareUnitOP_Napkins(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason2.Contains("QP") Or .Reason1.Contains("Web Loss") Then
                    .Tier1 = "Quality Paper"
                    .Tier2 = .Tier1
                ElseIf .Reason1.Contains("ATU17") Then
                    .Tier1 = "ATU"
                    If .Reason2.Contains("loose") Then
                        .Tier2 = "ATU Stack Jam - Lose Napkin"
                    Else
                        .Tier2 = "ATU Stack Jam - Misc"
                    End If
                ElseIf .Reason1.Contains("AT03") Then
                    .Tier1 = "ATU"
                    .Tier2 = "ATU - Electrical"
                ElseIf .Reason1.Contains("AT02") Then
                    .Tier1 = "ATU"
                    .Tier2 = "ATU - Mechanical"
                ElseIf .Reason1.Contains("ATU") And .Reason1.Contains("switch") Then
                    .Tier1 = "ATU"
                    .Tier2 = "ATU - Switch"
                ElseIf .Reason1.Contains("fingers") Then
                    .Tier1 = "ATU"
                    .Tier2 = "ATU Finger Jam"
                ElseIf .Reason1.Contains("slitter") Then
                    .Tier1 = "MISC"
                    .Tier2 = "Slitter Issues"
                ElseIf .Reason1.Contains("ATU") Then
                    .Tier1 = "ATU"
                    .Tier2 = "ATU - Misc"
                ElseIf .Reason1.Contains("folding cylinder") Then
                    .Tier1 = "FOLDING"
                    .Tier2 = "Folding Issues"
                ElseIf .Reason1.Contains("WRP36") Then
                    .Tier1 = "WRP"
                    .Tier2 = "WRP Poly Trouble"
                ElseIf .Reason1.Contains("WRP37") Then
                    .Tier1 = "WRP"
                    .Tier2 = "WRP Path Closing Jam"
                ElseIf .Reason1.Contains("WRP02") Then
                    .Tier1 = "WRP"
                    .Tier2 = "WRP - Mechanical issues"
                ElseIf .Reason1.Contains("WRP") Then
                    .Tier1 = "WRP"
                    .Tier2 = "WRP - Misc"
                ElseIf .Reason1.Contains("Wrapper") Then
                    .Tier1 = "WRP"
                    .Tier2 = "WRP - Misc"
                ElseIf .Reason1.Contains("UWS") Then
                    .Tier1 = "MISC"
                    .Tier2 = "UWS - Misc"
                ElseIf .Reason1.Contains("MCD20") Then
                    .Tier1 = "MCD"
                    .Tier2 = "MCD - Quality Issue"
                ElseIf .Reason1.Contains("MCD04") Then
                    .Tier1 = "MCD"
                    .Tier2 = "MCD - Hygiene"
                ElseIf .Reason1.Contains("MCD96") Then
                    .Tier1 = "MCD"
                    .Tier2 = "MCD - Doctor Blade"
                ElseIf .Reason1.Contains("MCD") Or .Location.Contains("MCD") Then
                    .Tier1 = "MCD"
                    .Tier2 = "MCD - Misc"
                ElseIf .Reason1.Contains("Tresu04") Then
                    .Tier1 = "MCD"
                    .Tier2 = "MCD - Hygiene"

                ElseIf .Reason1.Contains("FCS46") Then
                    .Tier1 = "MISC"
                    .Tier2 = "Band Saw Fault"

                ElseIf .Reason1.Contains("FCS") Then
                    .Tier1 = "MISC"
                    .Tier2 = "FCS - Misc"
                ElseIf .Reason1.Contains("EMB") Then
                    .Tier1 = "MISC"
                    .Tier2 = "EMB - Misc"
                ElseIf .Reason1.Contains("AUX") Then
                    .Tier1 = "MISC"
                    .Tier2 = "AUX - Misc"
                ElseIf .Reason1.Contains("Line Normal") Then
                    .Tier1 = "MISC"
                    .Tier2 = "Line - Misc"
                ElseIf .Reason1.Contains("transfer") Then
                    .Tier1 = "MISC"
                    .Tier2 = "ATU - Misc"


                ElseIf .Location.Contains("fold cut") Then
                    .Tier1 = "FOLDING"
                    .Tier2 = "Folding Issues"
                Else
                    .Tier1 = "MISC"
                    .Tier2 = "Other"
                End If
                .Tier3 = .Fault
            Else
                .Tier2 = .Team
                .Tier3 = .Fault
                If .Reason1.Contains("GEN02") Then
                    .Tier1 = "Product Change"
                ElseIf .Reason1.Contains("UWS05") Then
                    .Tier1 = "PRC"
                ElseIf .Reason1.Contains("WRP05") Then
                    .Tier1 = "WRP Poly Change"
                ElseIf .Reason2.Contains("blowdown") Then
                    .Tier1 = "Blowdown"
                Else
                    .Tier1 = "Planned Intervention"
                End If

            End If
        End With
    End Sub

    Public Sub getFamilyCareUnitOP_ModPACK(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason2.Contains("itrak") Then
                    .Tier1 = "805, 810, 815"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QF") Then
                    .Tier1 = "Quality Film"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason1.Contains("WRP10 Incoming Quality - Rolls") Then
                    .Tier1 = "Quality Rolls"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QR") Then
                    .Tier1 = "Quality Rolls"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("Quality - Rolls") Then
                    .Tier1 = "Quality Rolls"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("Quality - Film") Then
                    .Tier1 = "Quality Film"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QP Package Ears") Then
                    .Tier1 = "830, 835"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QP Poor / No Lap Seal") Then
                    .Tier1 = "820, 825"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QP Packages Stuck Together") Then
                    .Tier1 = "820, 825"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason1.Contains("WRP36 Package Separation Failure") Then
                    .Tier1 = "820, 825"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("805") Then
                    .Tier1 = "805, 810, 815"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("810") Then
                    .Tier1 = "805, 810, 815"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("815") Then
                    .Tier1 = "805, 810, 815"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("820") Then
                    .Tier1 = "820, 825"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("825") Then
                    .Tier1 = "820, 825"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("830") Then
                    .Tier1 = "830, 835"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("835") Then
                    .Tier1 = "830, 835"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("840") Then
                    .Tier1 = "840, 845"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Fault.Contains("845") Then
                    .Tier1 = "840, 845"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QP High Log Compressibility") Then
                    .Tier1 = "Quality Rolls"
                    .Tier2 = "Compressibility"
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QP Low Log Compressibility") Then
                    .Tier1 = "Quality Rolls"
                    .Tier2 = "Compressibility"
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QP Log Compressibility") Then
                    .Tier1 = "Quality Rolls"
                    .Tier2 = "Compressibility"
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("QP Paper Downstream") Then
                    .Tier1 = "Quality Rolls"
                    .Tier2 = "Compressibility"
                    .Tier3 = .Fault
                    Exit Sub
                End If

            Else
                If .Reason2.Contains("Changeover") Then
                    .Tier1 = "Planned Events"
                    .Tier2 = "Changeover"
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    .Tier1 = "Planned Events"
                    .Tier2 = "Planned Intervention"
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("CIL") Then
                    .Tier1 = "Planned Events"
                    .Tier2 = "CIL _ RLS"
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason2.Contains("RLS") Then
                    .Tier1 = "Planned Events"
                    .Tier2 = "CIL _ RLS"
                    .Tier3 = .Fault
                    Exit Sub
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    If .Reason2.Contains("autonomous") Then
                        .Tier1 = "Planned Events"
                        .Tier2 = "AM - PI"
                        .Tier3 = .Fault
                        Exit Sub
                    End If
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    If .Reason2.Contains("CIL") Then
                        .Tier1 = "Planned Events"
                        .Tier2 = "AM - PI"
                        .Tier3 = .Fault
                        Exit Sub
                    End If
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    If .Reason2.Contains("autonomous") Then
                        .Tier1 = "Planned Events"
                        .Tier2 = "CIL"
                        .Tier3 = .Fault
                        Exit Sub
                    End If
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    If .Reason2.Contains("CIL") Then
                        .Tier1 = "Planned Events"
                        .Tier2 = "CIL"
                        .Tier3 = .Fault
                        Exit Sub
                    End If
                End If

                If .Reason1.Contains("Planned Intervention") Then
                    If .Reason2.Contains("maint") Then
                        .Tier1 = "Planned Events"
                        .Tier2 = "AM - PI"
                        .Tier3 = .Fault
                        Exit Sub
                    End If
                End If



            End If
        End With
    End Sub


    Public Sub getFamilyCareUnitOP_WrapperprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then





                If .Reason1.Contains("Mechanical") Then

                    .Tier1 = "Breakdown"

                    .Tier2 = "Mechanical"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Electrical") Then

                    .Tier1 = "Breakdown"

                    .Tier2 = "Electrical"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Guard") Then

                    .Tier1 = "Breakdown"

                    .Tier2 = "Electrical - Guard"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Mechanical") And .Reason2.Contains("Drive") Then

                    .Tier1 = "Breakdown"

                    .Tier2 = "Mechanical - Drive"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Drive") Then

                    .Tier1 = "Breakdown"

                    .Tier2 = "Electrical - Drive"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Set-up") Then

                    .Tier1 = "Centerlines Set up"

                    .Tier2 = "Centerline Setup"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Setup") Then

                    .Tier1 = "Centerlines Set up"

                    .Tier2 = "Centerline Setup"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Height") Then

                    .Tier1 = "Centerlines Set up"

                    .Tier2 = "Centerline Setup - Height"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Timing") Then

                    .Tier1 = "Centerlines Set up"

                    .Tier2 = "Centerline Setup - Timing"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Re-centerlin") Then

                    .Tier1 = "Centerlines Set up"

                    .Tier2 = "Re-Centerline"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Adjust") Then

                    .Tier1 = "Centerlines Set up"

                    .Tier2 = "Centerline Setup Adj"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("QF") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Film Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Incoming Quality - Poly") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Film Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Poly shifted on core") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Film Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Edge Curl") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Film Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Dancer") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Dancer Fault"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Dancer") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Dancer Fault"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Eye Mark Alignment") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Poly-Eye Alignment"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Poly core slid") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Poly-Eye Alignment"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Film") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Film Fault & Web Loss"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Product Registration") Then

                    .Tier1 = "Film Loss & uws"

                    .Tier2 = "Registration Fault"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("QR Tails") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Cut") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Trim & Cut Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("QR Short") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Cut Length"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("QR Long") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Cut Length"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("QR Trim") Or .Reason2.Contains("QR Streamers") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Trim & Cut Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("QR Core") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Core Tip & Length"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Breakout") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Roll Diam & Roll Comp"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Package Separation") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Sep- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Package Separation") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Separation"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Lower Discharge") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Lower Dis Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Lower Dis.") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Lower Dis Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Flex Wipes") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Lower Dis Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Transfer Jam") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Angle OH Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Transfer Fault") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Angle OH Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Transfer oh fault") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Angle OH Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Angle Overhead") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Angle OH Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Product Overhead") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Prod OH Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Former Overhead") Or .Reason1.Contains("FormerOver") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Prod OH Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Infeed Jam") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Infeed Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Wrapper Discharge") And .Reason2.Contains("QR") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Wpr Disch Jam- Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Mushy") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Roll Diam & Roll Comp"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Hard") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Roll Diam & Roll Comp"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Diameter Large") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Roll Diam & Roll Comp"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Diameter Small") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Roll Diam & Roll Comp"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Quality - Rolls") And .Reason2.Contains("QR Large") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Roll Diam & Roll Comp"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Quality - Rolls") And .Reason2.Contains("QR Small") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Roll Diam & Roll Comp"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Incoming Quality") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Incoming Quality Other"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Quality") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Improper Fold") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Sideways/Standing") Or .Reason2.Contains("Sideways / Standing") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Sideways-Standing Roll"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Standing/Sideways") Or .Reason1.Contains("Sideways / Standing") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Sideways-Standing Roll"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Lower Dis.") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lower Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Lower Discharge") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lower Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("LowerDischarge") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lower Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Flex Wipes") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lower Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Flex Wipe") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lower Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Transfer Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Angle Overhead Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Transfer Fault") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Angle Overhead Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Transfer oh fault") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Angle Overhead Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Angle Overhead") Or .Reason1.Contains("AngleOver") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Angle Overhead Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Product Overhead") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Product Overhead Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Former Overhead") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Product Overhead Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Infeed Jam") And .Reason2.Contains("QR Missing") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Infeed Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Infeed Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Infeed Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Flighted Conv") Or .Reason1.Contains("FlightedConv") Or .Reason1.Contains("Flight Con") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Infeed Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Extra") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Infeed Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Missing Roll") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Infeed Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Infeed Lane") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Infeed Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Package Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Wrapper Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Roll Turner") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Wrapper Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Discharge") And .Reason2.Contains("False Discharge") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Wpr Disch Jam- False"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Discharge") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Wrapper Discharge Jam"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Downender") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Wrapper Discharge Jam"

                    .Tier3 = .Fault





                ElseIf .Reason1.Contains("Divert") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "External- Conv"

                    .Tier3 = .Fault



                ElseIf .Location.Contains("Cutroll Conv") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "External- Conv"

                    .Tier3 = .Fault



                ElseIf .Location.Contains("Package Conv") And .Reason1.Contains("Roll Turner") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "External- Conv"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Hygiene") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Hygiene"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Operational") And .Reason2.Contains("Uncoded") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Uncoded"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Operational") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Operational"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("E-Stop") And .Reason2.Contains("Uncoded") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Uncoded"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Interlock") And .Reason2.Contains("Uncoded") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Uncoded"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Uncoded") And .Reason2.Contains("Uncoded") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Uncoded"

                    .Tier3 = .Fault





                ElseIf .Reason2.Contains("QU") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Quality"

                    .Tier3 = .Fault



                ElseIf .Reason2.Contains("Lap Seal") Or .Reason2.Contains("Overlap") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Quality"

                    .Tier3 = .Fault

                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2

                End If




            Else
                If .Reason1.Contains("Product Change") Then
                    .Tier1 = "Changeover"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("Poly Roll") Then
                    .Tier1 = "Poly Change"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("Planned Intervention") Then
                    .Tier1 = "Other Maintenance"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("GEN01") And .Reason2.Contains("Maintenance") Then
                    .Tier1 = "Other Maintenance"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("GEN01") And .Reason2.Contains("Blowdown") Then
                    .Tier1 = "CL/RLS/CIL"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("GEN01") And .Reason2.Contains("CIL") Then
                    .Tier1 = "CL/RLS/CIL"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("GEN01") And .Reason2.Contains("Clean") Then
                    .Tier1 = "CL/RLS/CIL"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("GEN01") And .Reason2.Contains("centerline") Then
                    .Tier1 = "CL/RLS/CIL"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("GEN01") And .Reason2.Contains("AM") Then
                    .Tier1 = "AM"
                    .Tier2 = .Team
                ElseIf .DTGroup.Contains("Holiday/Curtail") Then
                    .Tier1 = "Unscheduled Time"
                    .Tier2 = .Team
                ElseIf .DTGroup.Contains("E.O./Projects") Then
                    .Tier1 = "Unscheduled Time"
                    .Tier2 = .Team
                ElseIf .DTGroup.Contains("Special Causes") Then
                    .Tier1 = "Unscheduled Time"
                    .Tier2 = .Team
                ElseIf .DTGroup.Contains("PR/Poly Change") Then
                    .Tier1 = "Poly Change"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("STARVE") Or .Reason1.Contains("Starv") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "STARVED"
                    .Tier3 = .Fault
                ElseIf .Reason1.Contains("BLOCK") Or .Reason1.Contains("Block") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "BLOCKED"
                    .Tier3 = .Fault
                ElseIf .Reason1.Contains("No Backlog") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "STARVED"
                    .Tier3 = .Fault
                ElseIf .Location.Contains("Starv") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "STARVED"
                    .Tier3 = .Fault
                ElseIf .Location.Contains("Block") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "BLOCKED"
                    .Tier3 = .Fault
                ElseIf .Fault.Contains("Starv") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "STARVED"
                    .Tier3 = .Fault
                ElseIf .Fault.Contains("Block") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "BLOCKED"
                    .Tier3 = .Fault
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                End If




            End If
        End With
    End Sub

    Public Sub getFamilyCareUnitOP_MFprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then


                If .Reason1.Contains("Dropped") Then

                    .Tier1 = "Drop/Tip Rolls"

                    .Tier2 = "Dropped Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Tipped") Then

                    .Tier1 = "Drop/Tip Rolls"

                    .Tier2 = "Tipped roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL30") And .Reason2.Contains("FALS") Then

                    .Tier1 = "Gap Fault"

                    .Tier2 = "Gap Fault False"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL31") And .Reason2.Contains("FALS") Then

                    .Tier1 = "Gap Fault"

                    .Tier2 = "Gap Fault False"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL30") Then

                    .Tier1 = "Gap Fault"

                    .Tier2 = "Gap Fault"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL31") Then

                    .Tier1 = "Gap Fault"

                    .Tier2 = "Gap Fault"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Gap Fault") Then

                    .Tier1 = "Gap Fault"

                    .Tier2 = "Gap Fault"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Gap Fault") And .Reason2.Contains("FALS") Then

                    .Tier1 = "Gap Fault"

                    .Tier2 = "Gap Fault False"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL03") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Elect Breakdown"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Overload") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Electrical Overload"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Overload") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Electrical Overload"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL02") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Mech Breakdown"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL04") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Hygiene"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Hygiene") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Hygiene"

                    .Tier3 = .Fault



                ElseIf .Reason1.Contains("Gap Fault") And .Reason2.Contains("tail") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Quality") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Quality Bundle Other"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Conveyor Jam") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Quality Bundle Other"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("QU") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Incoming Quality"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL10") And .Reason2.Contains("tail") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL11") And .Reason2.Contains("tail") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL10") And .Reason2.Contains("Short Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Short Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL11") And .Reason2.Contains("Short Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Short Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL10") And .Reason2.Contains("Long Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Long Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL11") And .Reason2.Contains("Long Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Long Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL10") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Incoming Quality"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL11") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Incoming Quality"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Dropped") And .Reason2.Contains("Long Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Long Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Dropped") And .Reason2.Contains("Short Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Short Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Tipped") And .Reason2.Contains("Long Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Long Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Tipped") And .Reason2.Contains("Short Roll") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Short Roll"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Dropped") And .Reason2.Contains("Fals") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Dropped") And .Reason2.Contains("tail") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Tipped") And .Reason2.Contains("tail") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Tipped") And .Reason2.Contains("Fals") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Tails"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Lap") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Lap Seal"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Lap") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Lap Seal"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("End Seal") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Quality End Seal"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("End Seal") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Quality End Seal"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Fused") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Quality End Seal"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Fused") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Quality End Seal"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL39") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Die Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL40") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Die Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL41") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Die Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Clamp") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Clamp Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Die") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Die Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Poly") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "PolyFilm related"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Vacuum") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Vacuum"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BDL36") Then

                    .Tier1 = "Sealing"

                    .Tier2 = "Collating Flight"

                    .Tier3 = .Fault
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else

                If .Reason2.Contains("Curtail") Then

                    .Tier1 = "Unscheduled Time"

                    .Tier2 = .Team

                ElseIf .Reason1.Contains("BDL05") Then

                    .Tier1 = "Poly Splice"

                    .Tier2 = .Team

                ElseIf .Reason1.Contains("Change") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Team

                ElseIf .Reason2.Contains("Change") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Team

                ElseIf .Reason1.Contains("Planned") Then

                    .Tier1 = "Planned Intervention"

                    .Tier2 = .Team
                ElseIf .Reason1.Contains("STARVE") Or .Reason1.Contains("Starv") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("BLOCK") Or .Reason1.Contains("Block") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "BLOCKED"

                    .Tier3 = .Fault
                ElseIf .Fault.Contains("BLOCK") Or .Fault.Contains("Block") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "BLOCKED"

                    .Tier3 = .Fault
                ElseIf .Fault.Contains("STARVE") Or .Fault.Contains("STARVE") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault
                ElseIf .Reason1.Contains("Backlog") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Die Jam") And .Reason2.Contains("Turned") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Tipped") And .Reason2.Contains("Turned") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Fault") And .Reason2.Contains("Incomplete") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Fault") And .Reason2.Contains("Tipped") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Fault") And .Reason2.Contains("Turned") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Slug") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Divert") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Disch") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "BLOCKED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("CUSTOMER") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "BLOCKED"

                    .Tier3 = .Fault

                ElseIf .Location.Contains("Starved") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                End If
            End If
        End With
    End Sub

    Public Sub getFamilyCareUnitOP_PalletizerprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason2.Contains("QB") Or .Reason2.Contains("QR") Or .Reason2.Contains("QS") Or .Reason2.Contains("QN") Then
                    .Tier1 = "Quality"
                    .Tier2 = .Reason2
                    .Tier3 = .Fault
                ElseIf .Reason1.Contains("Quality") Then
                    .Tier1 = "Quality"
                    .Tier2 = .Reason1
                    .Tier3 = .Fault


                ElseIf .Fault.Contains("STRIP APRON VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "apron fault"

                ElseIf .Fault.Contains("STRIP APRON NOT FORWARD") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "apron fault"

                ElseIf .Fault.Contains("STRIP APRON NOT BACK") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "apron fault"

                ElseIf .Fault.Contains("STRIP APRON OVER TRAVEL FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "apron fault"

                ElseIf .Fault.Contains("LANE 1 INVALID BARCODE DATA") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Barcode Data Fault"

                ElseIf .Fault.Contains("LANE 2 INVALID BARCODE DATA") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Barcode Data Fault"

                ElseIf .Fault.Contains("BARCODE AT READER DOES NOT MATCH LANE 1 OR LANE 2") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Barcode Data Fault"

                ElseIf .Fault.Contains("NO READ AT BARCODE SCANNER") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Barcode Data Fault"

                ElseIf .Fault.Contains("BARCODE NOT FOUND IN DATABASE") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Barcode Data Fault"

                ElseIf .Fault.Contains("CASE TURN CONVEYOR VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Case Turn Fault"

                ElseIf .Fault.Contains("CASE TURN ZONE 1 VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Case Turn Fault"

                ElseIf .Fault.Contains("CASE TURN ZONE 2 VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Case Turn Fault"

                ElseIf .Fault.Contains("CASE TURN ZONE 3 VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Case Turn Fault"

                ElseIf .Fault.Contains("CASE TURN JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Case Turn Jam"

                ElseIf .Fault.Contains("CENTERING CONVEYOR VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "CENTERING CONVETOR Fault"

                ElseIf .Fault.Contains("CENTERING CONVETOR JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "CENTERING CONVETOR JAM"

                ElseIf .Fault.Contains("CYCLE STOP REQUESTED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Cycle Stop Requested"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("DISCHARGE CONVEYOR 1 VFD FAULT") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Discharge Conv Fault"

                ElseIf .Fault.Contains("DISCHARGE CONVEYOR 2 VFD FAULT") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Discharge Conv Fault"

                ElseIf .Fault.Contains("DISCHARGE LEFT DISCONNECT [LOTO 3]") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Discharge Disconnect"

                ElseIf .Fault.Contains("DISCHARGE RIGHT DISCONNECT [LOTO 4]") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Discharge Disconnect"

                ElseIf .Fault.Contains("DISCHARGE STAGE 1 JAM") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Discharge Jam"

                ElseIf .Fault.Contains("DISCHARGE STAGE 2 JAM") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Discharge Jam"

                ElseIf .Fault.Contains("DUAL SKU FAULT. SAME UPC, DIFFERENT BRANDCODE") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Dual SKU fault"

                ElseIf .Fault.Contains("E-STOP MCR RELAY OFF") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("E-STOP AT SLIP SHEET DISPENSER") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("E-STOP AT PANELVIEW [STATION 1]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("E-STOP AT DISCHARGE RIGHT") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("E-STOP AT DISCHARGE (LEFT)") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("E-STOP AT PALLET DISPENSER") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("E-STOP AT TIE SHEET FEEDER") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("E-STOP AT SLIP SHEET FEEDER") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "E-Stop"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("LANE DIVERTER RIGHT DOOR [DS-1]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("LAYER TABLE RIGHT DOOR [DS-2]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("STRIP APRON DOOR [DS-3]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("SINGLE PALLET CONVEYOR DOOR [DS-4]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("FULL LOAD CONVEYOR DOOR [DS-5]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("LANE DIVERTER LEFT DOOR [DS-6]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("LAYER TABLE LEFT DOOR [DS-7]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("PALLET DISPENSER DOOR [DS-40]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET DISP. LOWER DOOR [DS-60]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET DISPENSER UPPER DOOR [DS-61]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("SLIP SHEET DOOR [DS-80]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Door Faults"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("FRONT RETAINER NOT IN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Front Retainer Fault"

                ElseIf .Fault.Contains("FRONT RETAINER NOT BACK") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Front Retainer Fault"

                ElseIf .Fault.Contains("FULL LOAD CONVEYOR VFD FAULT") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Full Load Conv Fault"

                ElseIf .Fault.Contains("FULL LOAD CONV GUIDES NOT IN") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Full Load Conv Fault"

                ElseIf .Fault.Contains("FULL LOAD CONV GUIDES NOT OUT") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Full Load Conv Fault"

                ElseIf .Fault.Contains("FULL LOAD CONV STOPS NOT BACK") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Full Load Conv Fault"

                ElseIf .Fault.Contains("HOIST APRON SAFETY 1") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Apron Safety"

                ElseIf .Fault.Contains("HOIST APRON SAFETY 2") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Apron Safety"

                ElseIf .Fault.Contains("HOIST DISCHARGE JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Discharge Jam"

                ElseIf .Fault.Contains("APRON COUNT FAULT - RESET HOIST COUNT - CLEAR MACH") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST PINS INSERTED") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST NOT DOWN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST NOT IN POSITION") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST DOWN AND PC14 OR PC15 BLOCKED") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST FALLEN PRODUCT DETECT PC25") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST FALLEN PRODUCT DETECT PC26") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST TOP OF LOAD PC14 OR PC15") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST NOT CLEAR TO LOWER - CHECK PEC'S") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("NO PALLET ON HOIST") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST FAULT PC20 BLOCKED OUT OF SEQUENCE") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST FAULT PC21 BLOCKED OUT OF SEQUENCE") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("HOIST COUNT FAULT - CLEAR MACH OR RESET COUNT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Hoist Fault"

                ElseIf .Fault.Contains("JAM ON PALLET STAGING CONVEYOR 1") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "JAM ON PALLET STAGING CONVEYOR"

                ElseIf .Fault.Contains("JAM ON PALLET STAGING CONVEYOR 2") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "JAM ON PALLET STAGING CONVEYOR"

                ElseIf .Fault.Contains("JAM ON PALLET STAGING CONVEYOR 3") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "JAM ON PALLET STAGING CONVEYOR"

                ElseIf .Fault.Contains("JAM ON PALLET DISPENSER ROLLERS") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "JAM ON PALLET STAGING CONVEYOR"

                ElseIf .Fault.Contains("LANE DIVERTER FVD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "LANE DIVERTERFAULT"

                ElseIf .Fault.Contains("LANER PIN DETECT [PC-6 FAILURE]") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "laner fault"

                ElseIf .Fault.Contains("LANER COUNT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "laner fault"

                ElseIf .Fault.Contains("LANER JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "LANER JAM"

                ElseIf .Fault.Contains("LAYER TABLE 1 VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "layer Table Fault"

                ElseIf .Fault.Contains("LAYER TABLE 2 VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "layer Table Fault"

                ElseIf .Fault.Contains("LAYER TABLE 1 COUNT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "layer Table Fault"

                ElseIf .Fault.Contains("LAYER TABLE 2 COUNT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "layer Table Fault"

                ElseIf .Fault.Contains("LAYER TABLE 1 JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "LAYER TABLE JAM"

                ElseIf .Fault.Contains("LAYER TABLE 2 JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "LAYER TABLE JAM"

                ElseIf .Fault.Contains("LAYER TABLE 1 TO LAYER TABLE 2 TRANSITION JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "LAYER TABLE JAM"

                ElseIf .Fault.Contains("LIGHT CURTAIN @ DISCHARGE [LC-1]") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Light Curtain Discharge"

                ElseIf .Fault.Contains("LIGHT CURTAIN @ PALLET INFEED [LC-40]") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Discharge"
                    .Tier3 = "Light Curtain Fault"

                ElseIf .Fault.Contains("LOW AIR PRESSURE [PS-2]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Low Air Pressure"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("LOW TIE SHEETS") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = "Low Tie Sheets"

                ElseIf .Fault.Contains("MEZZANINE  DISCONNECT [LOTO 1]") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Mezz Disconnect"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("MIXED CASES DETECTED ON METER BELT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Mixed Cased on Meter Belt"

                ElseIf .Fault.Contains("#0000 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0029 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0034 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0042 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0061 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0065 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0068 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0070 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0076 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0077 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0078 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0079 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0081 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0084 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0085 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0086 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0087 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0088 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0089 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0090 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0091 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0092 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0093 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0094 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0095 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0098 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0099 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0102 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0109 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0110 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0111 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0112 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0113 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0116 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0118 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0119 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0120 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0121 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0122 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0123 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0124 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0125 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0130 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0142 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0153 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0154 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0158 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0159 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0163 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0169 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0170 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0171 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0172 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0176 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0187 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0188 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0201 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0202 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0203 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0204 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0205 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0206 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0207 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0212 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0216 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0217 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0218 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0219 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0220 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0221 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0222 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0223 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0225 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0226 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0227 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0228 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0229 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0230 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0231 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0235 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0236 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0237 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0238 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0239 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0249 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0250 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0251 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0252 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0253 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0254 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("#0255 NOT DEFINED") Then
                    .Tier1 = "Cycle Stop / Other"
                    .Tier2 = "Not Defined"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("OVERHEAD GATE VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Overhead gate fault"

                ElseIf .Fault.Contains("OVERHEAD GATE NOT UP") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Overhead gate fault"

                ElseIf .Fault.Contains("OVERHEAD GATE NOT DOWN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Overhead gate fault"

                ElseIf .Fault.Contains("PACER BELT COUNT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Pacer belt count fault"

                ElseIf .Fault.Contains("PACER BELT VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "PACER_METER BELT Fault"

                ElseIf .Fault.Contains("METER BELT VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "PACER_METER BELT Fault"

                ElseIf .Fault.Contains("METER BELT COUNT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "PACER_METER BELT Fault"

                ElseIf .Fault.Contains("PACER/METER BELT JAM") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "PACER_METER BELT JAM"

                ElseIf .Fault.Contains("PALLET DISPENSER DISCONNECT [LOTO 2]") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET DISPENSER ROLLERS NOT RUNNING") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("SINGLE PALLET CONVEYOR VFD FAULT") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET DISPENSER TRANSFER NOT COMPLETE") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET RETAINERS NOT BACK") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET DISPENSER STACK SAFETY (PC44)") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET DISPENSER LIFT NOT RUNNING UP") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET DISPENSER LIFT NOT RUNNING DOWN") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET DISPENSER LIFT SAFETY (PX47 OR PX48)") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Dispenser Fault"

                ElseIf .Fault.Contains("PALLET ENTERING HOIST JAM") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet entering hoist jam"

                ElseIf .Fault.Contains("PALLET TRANSFER JAM IN DISPENSER") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Jam"

                ElseIf .Fault.Contains("PALLET TRANSFER JAM LEAVING DISPENSER") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Jam"

                ElseIf .Fault.Contains("PALLET JAM ON SINGLE PALLET CONV (PC55)") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Jam"

                ElseIf .Fault.Contains("PALLET JAM ON SINGLE PALLET CONV (PC45)") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Jam"

                ElseIf .Fault.Contains("PALLET DISPENSER LIFT JAM") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Jam"

                ElseIf .Fault.Contains("PALLET OVERFEED AT HOIST") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Overfeed at Hoist"

                ElseIf .Fault.Contains("PALLET RETAINERS NOT IN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Retainer Fault"

                ElseIf .Fault.Contains("PALLET RETAINERS NOT OUT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Retainer Fault"

                ElseIf .Fault.Contains("PALLET RETAINER STUCK OUT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Retainer Fault"

                ElseIf .Fault.Contains("PALLET ROTATOR NOT UP") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Rotator Fault"

                ElseIf .Fault.Contains("PALLET ROTATOR NOT DOWN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Rotator Fault"

                ElseIf .Fault.Contains("PALLET ROTATOR NOT IN CW POSITION") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Rotator Fault"

                ElseIf .Fault.Contains("PALLET ROTATOR NOT IN CCW POSITION") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Rotator Fault"

                ElseIf .Fault.Contains("PALLET ROTATOR GUIDES NOT OUT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Rotator Fault"

                ElseIf .Fault.Contains("PALLET ROTATOR GUIDES NOT BACK") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Pallet Rotator Fault"

                ElseIf .Fault.Contains("PALLET STAGING 1 NOT RUNNING") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Staging Not Running"

                ElseIf .Fault.Contains("PALLET STAGING 2 NOT RUNNING") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Staging Not Running"

                ElseIf .Fault.Contains("PALLET STAGING 3 NOT RUNNING") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Pallet infeed"
                    .Tier3 = "Pallet Staging Not Running"

                ElseIf .Fault.Contains("ROBO CENTERING CONVEYOR FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Robo centering conv fault"

                ElseIf .Fault.Contains("ROBO LANE 1 FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Robo Lane Fault"

                ElseIf .Fault.Contains("ROBO LANE 2 FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Robo Lane Fault"

                ElseIf .Fault.Contains("ROBO LANE 4 FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Robo Lane Fault"

                ElseIf .Fault.Contains("ROBO LANE 5 FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Robo Lane Fault"

                ElseIf .Fault.Contains("ROBO LAYER TABLE 1 LEFT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "Robo Layer Table Fault"

                ElseIf .Fault.Contains("ROBO LAYER TABLE 1 RIGHT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "Robo Layer Table Fault"

                ElseIf .Fault.Contains("ROBO LAYER TABLE 2 LEFT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "Robo Layer Table Fault"

                ElseIf .Fault.Contains("ROBO LAYER TABLE 2 RIGHT FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Layer Table"
                    .Tier3 = "Robo Layer Table Fault"

                ElseIf .Fault.Contains("SCANNER PEC FAULT MISALIGNED/PUSHTHRU") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Scanner PEC Fault"

                ElseIf .Fault.Contains("SIDE RETAINER NOT IN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Side Retainer Fault"

                ElseIf .Fault.Contains("SIDE RETAINER NOT OUT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Hoist"
                    .Tier3 = "Side Retainer Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISPENSER CHAIN JAM") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "SLIP SHEET DISPENSER CHAIN JAM"

                ElseIf .Fault.Contains("SLIP SHEET DISPENSER ARM NOT FORWARD TO PALLET") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "SLIP SHEET DISPENSER Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISP CUP DETECT NOT MADE") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISP LIFT NOT RUNNING") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET PINS INSERTED") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET LEFT CHAIN SAFETY (PX89)") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET RIGHT CHAIN SAFETY (PX90)") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISPENSER LIFT NOT UP") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISPENSER LIFT NOT DOWN") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISPENSER ARM NOT BACK OVER BIN") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("LOW SLIP SHEETS") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISP SHEET PICKUP FAULT") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("SLIP SHEET DISP SHEET LOST IN TRANSIT") Then
                    .Tier1 = "Lower Level"
                    .Tier2 = "Slip Sheet"
                    .Tier3 = "Slip Sheet Fault"

                ElseIf .Fault.Contains("TIE SHEET DISPENSER GATE [DS-62]") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET CARRIAGE OVERTRAVEL AT PALLET") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET CARRIAGE OVERTRAVEL AT SHEETS") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET CARRIAGE VFD FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET LIFT NOT RUNNING") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET PINS NOT IN HOLDERS") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET HEAD NOT UP") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET HEAD NOT DOWN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET CARRIAGE NOT FORWARD") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET CARRIAGE NOT BACK") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET TABLE NOT UP") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET TABLE NOT DOWN") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET LIFT GUIDE FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET PICKUP FAULT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TIE SHEET LOST IN TRANSIT") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Tie Sheet"
                    .Tier3 = .Fault

                ElseIf .Fault.Contains("TRACKING FAULT-CHECK FOR MIXED CASES-CLEAR MACHINE") Then
                    .Tier1 = "Upper Level"
                    .Tier2 = "Infeed"
                    .Tier3 = "Tracking Fault"



                ElseIf .Reason2.Contains("Starved") Or .Reason2.Contains("Blocked") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = .Reason1
                    .Tier3 = .Fault


                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                .Tier3 = .Fault
                If .Reason1.Contains("Product Change") Or .Reason2.Contains("Changeover") Then
                    .Tier1 = "Changeover"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("Planned Intervention") Then
                    .Tier1 = "Planned Intervention"
                    .Tier2 = .Team
                ElseIf .Reason2.Contains("CIL") Or .Reason2.Contains("RLS") Then
                    .Tier1 = ("CIL/RLS")
                    .Tier2 = .Team
                ElseIf .Reason2.Contains("Down") Or .Fault.Contains("Blocked") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "Blocked"
                ElseIf .Reason2.Contains("No Product") Or .Fault.Contains("Starved") Then
                    .Tier1 = "Blocked/Starved"
                    .Tier2 = "Starved"
                ElseIf .Reason1.Contains("Roll Change") Or .Reason2.Contains("Roll Change") Then
                    .Tier1 = "Roll Change"
                    .Tier2 = .Team
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Team
                End If
            End If
        End With
    End Sub


    Public Sub getFamilyCareUnitOP_ACPprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                If .Reason2.Contains("Air Pressure") Then

                    .Tier1 = "AIR-Vacuum"

                    .Tier2 = "Air Pressure"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("LOW") And .Reason1.Contains("AIR") Then

                    .Tier1 = "AIR-Vacuum"

                    .Tier2 = "Air Pressure"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Vacuum") Then

                    .Tier1 = "AIR-Vacuum"

                    .Tier2 = "Vacuum"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("SERVO") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Electrical"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Axis Fault") Or .Reason2.Contains("SERVO FAULT") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Electrical"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Axis Faulted") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Electrical"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("PLC") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Programming"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("PLC") Or .Reason1.Contains("CONTROLLER") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Programming"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Unassigned") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Programming"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("DRIVE FAULT") Or .Reason1.Contains("OVERLOAD") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Electrical"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Electrical Equipment") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Electrical"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Logic") Or .Reason2.Contains("Program") Then

                    .Tier1 = "Electrical / Programming"

                    .Tier2 = "Programming"

                    .Tier3 = .Fault

                ElseIf .Location.Contains("LANER") Or .Location.Contains("DIVERTER") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Laner-Diverter Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("CASE Not") And .Reason1.Contains("FORM") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Forming Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("CASE") And .Reason1.Contains("Not OPEN") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Forming Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Index") And .Reason1.Contains("Head Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Index Head Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Indexer") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Index Head Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Seal") And .Reason1.Contains("Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Sealer Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("MAGAZ") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Magazine Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Lift") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lifter-Infeed Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("INFEED") And .Reason1.Contains("Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lifter-Infeed Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("dispens") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Case Dispenser Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("CASE FORM") And .Reason1.Contains("Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Forming Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Case Not Delivered To Packer") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Forming Jam"

                    .Tier3 = .Fault

                ElseIf (.Reason1.Contains("push") Or .Reason1.Contains("LOADER")) And .Reason2.Contains("NO Case") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Forming Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Product jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Lifter-Infeed Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("LOAD") And .Reason1.Contains("Jam") And .Reason2.Contains("Misaligned Stack") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Pusher-Loading Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("push") Or .Reason1.Contains("LOADER") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Pusher-Loading Jam"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("TRANSFER") And .Reason1.Contains("Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "TRANSFER JAM"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Case") And .Reason1.Contains("Chute") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Jam at Loader"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Case") And .Reason1.Contains("Loader") And .Reason1.Contains("Jam") Then

                    .Tier1 = "Jam"

                    .Tier2 = "Jam at Loader"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("AUX") And .Reason1.Contains("MAG") Then

                    .Tier1 = "Loader"

                    .Tier2 = "Low KDF"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("AUX") And .Reason1.Contains("MAG") Then

                    .Tier1 = "Loader"

                    .Tier2 = "Low KDF"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("OUT Of CASES") Then

                    .Tier1 = "Loader"

                    .Tier2 = "Low KDF"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("LOW KDF") Then

                    .Tier1 = "Loader"

                    .Tier2 = "Low KDF"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Hygiene") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Hygiene"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("OPERATIONAL") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Operational"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("DIRT") Or .Reason2.Contains("DUST") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Hygiene"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Not Run") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "ACP Stopped"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Stop") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "ACP Stopped"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Interlock") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Interlock"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("DISC") And .Reason1.Contains("OPEN") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "ACP Stopped"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("UPEND") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Upender"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("REJ") And .Reason1.Contains("CONV") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "REJECT CONVEYOR"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Glue System Failure") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Mechanical"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Mechanical") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Mechanical"

                    .Tier3 = .Fault


                ElseIf .Reason2.Contains("CAs") And .Reason2.Contains("BACKW") Then

                    .Tier1 = OTHERS_STRING

                    .Tier2 = "Operational"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Handling") And .Reason2.Contains("P&G") Then

                    .Tier1 = "Quality"

                    .Tier2 = "KDF Damage"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Wet") Or .Reason2.Contains("Damp") Then

                    .Tier1 = "Quality"

                    .Tier2 = "KDF Damage"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("Clamp Damage") Then

                    .Tier1 = "Quality"

                    .Tier2 = "KDF Damage"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("QR") Or .Reason2.Contains("QU") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Quality"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("QK") Then

                    .Tier1 = "Quality"

                    .Tier2 = "KDF Quality"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Quality") And .Reason1.Contains("KDF") Then

                    .Tier1 = "Quality"

                    .Tier2 = "KDF Quality"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Quality") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Package Quality"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("score") Then

                    .Tier1 = "Quality"

                    .Tier2 = "KDF Quality"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("FLAP") And Not .Reason2.Contains("CLOS") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Open Flap"

                    .Tier3 = .Fault

                ElseIf .Reason2.Contains("FLAP") And .Reason2.Contains("OPEN") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Open Flap"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("FLAP") And .Reason1.Contains("OPEN") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Open Flap"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("FLAP") And Not .Reason1.Contains("CLOS") Then

                    .Tier1 = "Quality"

                    .Tier2 = "Open Flap"

                    .Tier3 = .Fault

                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2

                End If
            Else
                If .Reason2.Contains("Curtail") Then

                    .Tier1 = "Unscheduled Time"

                    .Tier2 = .Team

                ElseIf .Reason1.Contains("Change") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Team

                ElseIf .Reason2.Contains("Change") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Team

                ElseIf .Reason1.Contains("Planned") Then

                    .Tier1 = "Planned Intervention"

                    .Tier2 = .Team
                ElseIf .Reason2.Contains("DISC") And .Reason2.Contains("FULL") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "BLOCKED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("FCC Down") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "BLOCKED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("STARVE") Or .Reason1.Contains("Starv") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Location.Contains("Package Conveyor") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Reason1.Contains("Divert") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                ElseIf .Location.Contains("Starved") Then

                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault
                ElseIf .Fault.Contains("Block") Or .Fault.Contains("BLOCK") Then
                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "BLOCKED"

                    .Tier3 = .Fault
                ElseIf .Fault.Contains("STARVE") Or .Fault.Contains("starve") Then
                    .Tier1 = "Blocked/Starved"

                    .Tier2 = "STARVED"

                    .Tier3 = .Fault

                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1

                End If
            End If
        End With
    End Sub


    Public Sub getFamilyMakingprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                ''''''''''''''''''''''''Zone 1''''''''''''''''''''''
                If .Reason1.Contains("Zone 1") Then
                    .Tier1 = "Zone 1"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 1 SB"
                        .Tier3 = .Reason3

                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If

                    End If

                    'Zone 1 SB Too
                ElseIf .Reason1.Contains("Softener") Then
                    .Tier1 = "Zone 1"
                    .Tier2 = "Zone 1 SB"


                    'Zone 1 Too but DT
                ElseIf .Location.Contains("Stock") Or .Location.Contains("Headbox") Or .Location.Contains("Water") Or .Location.Contains("Steam") Or .Location.Contains("Compressed") Or .Location.Contains("Power Distribution") Or .Location.Contains("Stock") Or .Location.Contains("Operations DT") Or .Location.Contains("Process & C") Then
                    .Tier1 = "Zone 1"
                    .Tier2 = "Zone 1 DT"
                    If .Location.Contains("Stock") Then
                        .Tier3 = "Stock Prep"
                    End If
                    If .Location.Contains("Headbox") Then
                        .Tier3 = "Headbox"
                    End If
                    If .Location.Contains("Steam") Or .Location.Contains("Compressed") Then
                        .Tier3 = "Plant Util"
                    End If
                    If .Location.Contains("Water") Then
                        .Tier3 = "Water System"
                    End If
                    If .Location.Contains("Power") Then
                        .Tier3 = "Power"
                    End If
                    If .Location.Contains("Stock Prep") And .Fault.Contains("Quick Mix/Stock Pump") Then
                        .Tier3 = "Quick Mix"
                    End If
                    If .Location.Contains("Process & C") And .Fault.Contains("Distribution") Then
                        .Tier3 = "Logic Failure"
                    End If
                    If .Location.Contains("Operations DT") And .Fault.Contains("Natural Causes") Then
                        .Tier3 = "Natural Causes"
                    End If


                    'End of Zone 1

                    ''''''''''Zone 2'''''''''''''''''''
                ElseIf .Reason1.Contains("Zone 2") Then
                    .Tier1 = "Zone 2"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 2 SB"
                        .Tier3 = .Reason3
                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If
                    End If
                ElseIf .Location.Contains("Backing") Or .Location.Contains("Forming") Then

                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Roll") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Backing Wire Roll"
                    End If
                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Guide") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Backing Wire Guide"
                    End If
                    If .Location.Contains("Forming Wire Section") And .Fault.Contains("Forming Wire") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Forming Wire"
                    End If
                    If .Location.Contains("Forming Wire Section") And .Fault.Contains("Wire Tension") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Wire Tension System"
                    End If
                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Trim System") And .Reason3.Contains("Nozzle") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "BPD Nozzle Plug"
                    End If
                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Wire Tension") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Backing Wire Tension"
                    End If

                    ''''''''''Zone 3'''''''''''''''''''


                ElseIf .Reason1.Contains("Zone 3") Then
                    .Tier1 = "Zone 3"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 3 SB"
                        .Tier3 = .Reason3
                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If
                    End If
                ElseIf .Location.Contains("Hot Air") Or .Location.Contains("Press Section") Or .Location.Contains("Hydraulic") Or .Location.Contains("Machine Hot Air System") Or .Location.Contains("AC/DC Variable") Or .Location.Contains("Lube System") Then

                    If .Location.Contains("Hot Air") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Hot Air System"
                    End If
                    If .Reason1.Contains("Belt") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Belt Problem Damage"
                    End If
                    If .Location.Contains("Press Section") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Press Section"
                    End If
                    If .Location.Contains("Hydraulic") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Hydraulic System"
                    End If
                    If .Location.Contains("Machine Hot Air System") And .Fault.Contains("Predyer Hot Air System") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Hot Air System"
                    End If
                    If .Location.Contains("AC/DC Variable") And .Fault.Contains("DC Drive System") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Load Share Work"
                    End If
                    If .Location.Contains("Lube System") And .Fault.Contains("Central Lube") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Central Lube"
                    End If




                    ''''''''''Zone 4'''''''''''''''''''
                ElseIf .Reason1.Contains("Zone 4") Then
                    .Tier1 = "Zone 4"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 4 SB"
                        .Tier3 = .Reason3
                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If
                    End If
                ElseIf .Location.Contains("Reel") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Reel Section"
                ElseIf .Reason1.Contains("Reel") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Reel Section"
                ElseIf .Location.Contains("Yankee") And .Fault = "Hood" And .Reason2.Contains("Fire") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Yankee Hood Fire"
                ElseIf .Location.Contains("Yankee") And .Fault = "Hood" And .Reason2.Contains("No/Low Air Flow") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Yankee System SIL"
                ElseIf .Location.Contains("Yankee") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Yankee System"
                ElseIf .Location.Contains("Dust") Or .Location.Contains("Cranes") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Dust-Lube-Crane"
                ElseIf .Location.Contains("Threading") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Threading System"
                ElseIf .Location.Contains("Repulper") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Repulper"
                ElseIf .Location.Contains("Reel Section") And .Fault.Contains("Sheet Break") And .Reason1.Contains("Detector") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Sheetbreak Detector"
                ElseIf .Location.Contains("Dry End Selection") And .Fault.Contains("Shaft Puller") And .Reason1.Contains("Shaft Puller") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Shaftpuller"
                ElseIf .Location.Contains("Variable Drives") And .Fault.Contains("AC Drive System") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Start Up Logic"
                ElseIf .Location.Contains("Calendar System") And .Fault.Contains("Calendar System Loading") And .Reason2.Contains("Loading System Failure") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Calender Loading System Failure"
                ElseIf .Location.Contains("Sheet") And .Reason1.Contains("Calendar System") And .Reason2.Contains("Loading System Failure") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone 4 SB"
                    .Tier3 = "Sheetbreak Other"
                ElseIf .Location.Contains("Forming Wire") Then
                    .Tier1 = "Zone 2"
                    .Tier2 = "Zone 2 DT"
                    .Tier3 = "Forming Wire"
                ElseIf .Location.Contains("Lube") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone 4 DT"
                    .Tier3 = "Central Lube"
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If


                'now make the desired switches
                If .Tier1 <> OTHERS_STRING Then

                    If .Tier1 = "Zone 1" Then
                        If .Tier2.Contains("SB") Then
                            .Tier1 = "Zone SB"
                            .Tier2 = "Zone 1 SB"
                        ElseIf .Tier2.Contains("DT") Then
                            .Tier1 = "Zone DT"
                            .Tier2 = "Zone 1 DT"
                        End If
                    ElseIf .Tier1 = "Zone 2" Then
                        If .Tier2.Contains("SB") Then
                            .Tier1 = "Zone SB"
                            .Tier2 = "Zone 2 SB"
                        ElseIf .Tier2.Contains("DT") Then
                            .Tier1 = "Zone DT"
                            .Tier2 = "Zone 2 DT"
                        End If
                    ElseIf .Tier1 = "Zone 3" Then
                        If .Tier2.Contains("SB") Then
                            .Tier1 = "Zone SB"
                            .Tier2 = "Zone 3 SB"
                        ElseIf .Tier2.Contains("DT") Then
                            .Tier1 = "Zone DT"
                            .Tier2 = "Zone 3 DT"
                        End If
                    ElseIf .Tier1 = "Zone 4" Then
                        If .Tier2.Contains("SB") Then
                            .Tier1 = "Zone SB"
                            .Tier2 = "Zone 4 SB"
                        ElseIf .Tier2.Contains("DT") Then
                            .Tier1 = "Zone DT"
                            .Tier2 = "Zone 4 DT"
                        End If
                    End If

                End If

                If .Tier1 = BLANK_INDICATOR Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If


            Else 'PLANNED
                If .Location.Contains("DT") And .Reason2.Contains("Brand Change") Then
                    .Tier1 = "Brand Change"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Location.Contains("DT") And .Reason2.Contains("Downday") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Location.Contains("SB") And .Reason3.Contains("Brand Change") Then
                    .Tier1 = "Soft Swing"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Fault.Contains("Sheet") And .Reason2.Contains("Creping Blade") Then
                    .Tier1 = "RLS"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Fault.Contains("Sheet") And .Reason3.Contains("Creping Blade") Then
                    .Tier1 = "RLS"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Location.Contains("Operations DT") And .Fault.Contains("Operators") And .Reason1.Contains("Operators") And .Reason4.Contains("Proces/Operational") Then
                    .Tier1 = "DT RLS"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Not Applicable") And .Reason4.Contains("Outage") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Not Applicable") And .Reason4.Contains("Downday") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Not Applicable") And .Reason4.Contains("Brandswing") Then
                    .Tier1 = "Brand Change"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Production Planning") And .Reason2.Contains("Outage") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Production Planning") And .Reason3.Contains("Process/Operational") Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason2.Contains("Outage") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason2.Contains("Downday") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""

                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Team
                End If

            End If
        End With

    End Sub

    Public Sub getFamilyMakingprstoryMapping_OLD(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                ''''''''''''''''''''''''Zone 1''''''''''''''''''''''
                If .Reason1.Contains("Zone 1") Then
                    .Tier1 = "Zone 1"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 1 SB"
                        .Tier3 = .Reason3

                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If

                    End If

                    'Zone 1 SB Too
                ElseIf .Reason1.Contains("Softener") Then
                    .Tier1 = "Zone 1"
                    .Tier2 = "Zone 1 SB"


                    'Zone 1 Too but DT
                ElseIf .Location.Contains("Stock") Or .Location.Contains("Headbox") Or .Location.Contains("Water") Or .Location.Contains("Steam") Or .Location.Contains("Compressed") Or .Location.Contains("Power Distribution") Or .Location.Contains("Stock") Or .Location.Contains("Operations DT") Or .Location.Contains("Process & C") Then
                    .Tier1 = "Zone 1"
                    .Tier2 = "Zone 1 DT"
                    If .Location.Contains("Stock") Then
                        .Tier3 = "Stock Prep"
                    End If
                    If .Location.Contains("Headbox") Then
                        .Tier3 = "Headbox"
                    End If
                    If .Location.Contains("Steam") Or .Location.Contains("Compressed") Then
                        .Tier3 = "Plant Util"
                    End If
                    If .Location.Contains("Water") Then
                        .Tier3 = "Water System"
                    End If
                    If .Location.Contains("Power") Then
                        .Tier3 = "Power"
                    End If
                    If .Location.Contains("Stock Prep") And .Fault.Contains("Quick Mix/Stock Pump") Then
                        .Tier3 = "Quick Mix"
                    End If
                    If .Location.Contains("Process & C") And .Fault.Contains("Distribution") Then
                        .Tier3 = "Logic Failure"
                    End If
                    If .Location.Contains("Operations DT") And .Fault.Contains("Natural Causes") Then
                        .Tier3 = "Natural Causes"
                    End If


                    'End of Zone 1

                    ''''''''''Zone 2'''''''''''''''''''
                ElseIf .Reason1.Contains("Zone 2") Then
                    .Tier1 = "Zone 2"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 2 SB"
                        .Tier3 = .Reason3
                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If
                    End If
                ElseIf .Location.Contains("Backing") Or .Location.Contains("Forming") Then

                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Roll") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Backing Wire Roll"
                    End If
                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Guide") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Backing Wire Guide"
                    End If
                    If .Location.Contains("Forming Wire Section") And .Fault.Contains("Forming Wire") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Forming Wire"
                    End If
                    If .Location.Contains("Forming Wire Section") And .Fault.Contains("Wire Tension") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Wire Tension System"
                    End If
                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Trim System") And .Reason3.Contains("Nozzle") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "BPD Nozzle Plug"
                    End If
                    If .Location.Contains("Backing Wire Section") And .Fault.Contains("Wire Tension") Then
                        .Tier1 = "Zone 2"
                        .Tier2 = "Zone 2 DT"
                        .Tier3 = "Backing Wire Tension"
                    End If

                    ''''''''''Zone 3'''''''''''''''''''


                ElseIf .Reason1.Contains("Zone 3") Then
                    .Tier1 = "Zone 3"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 3 SB"
                        .Tier3 = .Reason3
                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If
                    End If
                ElseIf .Location.Contains("Hot Air") Or .Location.Contains("Press Section") Or .Location.Contains("Hydraulic") Or .Location.Contains("Machine Hot Air System") Or .Location.Contains("AC/DC Variable") Or .Location.Contains("Lube System") Then

                    If .Location.Contains("Hot Air") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Hot Air System"
                    End If
                    If .Reason1.Contains("Belt") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Belt Problem Damage"
                    End If
                    If .Location.Contains("Press Section") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Press Section"
                    End If
                    If .Location.Contains("Hydraulic") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Hydraulic System"
                    End If
                    If .Location.Contains("Machine Hot Air System") And .Fault.Contains("Predyer Hot Air System") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Hot Air System"
                    End If
                    If .Location.Contains("AC/DC Variable") And .Fault.Contains("DC Drive System") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Load Share Work"
                    End If
                    If .Location.Contains("Lube System") And .Fault.Contains("Central Lube") Then
                        .Tier1 = "Zone 3"
                        .Tier2 = "Zone 3 DT"
                        .Tier3 = "Central Lube"
                    End If




                    ''''''''''Zone 4'''''''''''''''''''
                ElseIf .Reason1.Contains("Zone 4") Then
                    .Tier1 = "Zone 4"
                    If .Fault.Contains("Sheet") Then
                        .Tier2 = "Zone 4 SB"
                        .Tier3 = .Reason3
                        ' Exception for R2
                        If .Reason2.Contains("Cleaning Blade") Or .Reason2.Contains("Driver Side") Or .Reason2.Contains("Tending Side") Or .Reason2.Contains("Won't Thread") Or .Reason2.Contains("Center Of Roll") Then
                            .Tier3 = .Reason2
                        End If

                        'Exception for Location
                        If .Location.Contains("Dry End") Or .Location.Contains("Reel Section") Then
                            .Tier3 = "Reel-DryEnd"
                        End If
                    End If
                ElseIf .Location.Contains("Reel") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Reel Section"
                ElseIf .Reason1.Contains("Reel") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Reel Section"
                ElseIf .Location.Contains("Yankee") And .Fault = "Hood" And .Reason2.Contains("Fire") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Yankee Hood Fire"
                ElseIf .Location.Contains("Yankee") And .Fault = "Hood" And .Reason2.Contains("No/Low Air Flow") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Yankee System SIL"
                ElseIf .Location.Contains("Yankee") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Yankee System"
                ElseIf .Location.Contains("Dust") Or .Location.Contains("Cranes") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Dust-Lube-Crane"
                ElseIf .Location.Contains("Threading") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Threading System"
                ElseIf .Location.Contains("Repulper") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Repulper"
                ElseIf .Location.Contains("Reel Section") And .Fault.Contains("Sheet Break") And .Reason1.Contains("Detector") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Sheetbreak Detector"
                ElseIf .Location.Contains("Dry End Selection") And .Fault.Contains("Shaft Puller") And .Reason1.Contains("Shaft Puller") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Shaftpuller"
                ElseIf .Location.Contains("Variable Drives") And .Fault.Contains("AC Drive System") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Start Up Logic"
                ElseIf .Location.Contains("Calendar System") And .Fault.Contains("Calendar System Loading") And .Reason2.Contains("Loading System Failure") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone DT"
                    .Tier3 = "Calender Loading System Failure"
                ElseIf .Location.Contains("Sheet") And .Reason1.Contains("Calendar System") And .Reason2.Contains("Loading System Failure") Then
                    .Tier1 = "Zone 4"
                    .Tier2 = "Zone 4 SB"
                    .Tier3 = "Sheetbreak Other"



                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If


            Else 'PLANNED
                If .Location.Contains("DT") And .Reason2.Contains("Brand Change") Then
                    .Tier1 = "Brand Change"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Location.Contains("DT") And .Reason2.Contains("Downday") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Location.Contains("SB") And .Reason3.Contains("Brand Change") Then
                    .Tier1 = "Soft Swing"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Fault.Contains("Sheet") And .Reason2.Contains("Creping Blade") Then
                    .Tier1 = "RLS"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Fault.Contains("Sheet") And .Reason3.Contains("Creping Blade") Then
                    .Tier1 = "RLS"
                    .Tier2 = .Team
                    .Tier3 = ""

                ElseIf .Location.Contains("Operations DT") And .Fault.Contains("Operators") And .Reason1.Contains("Operators") And .Reason4.Contains("Proces/Operational") Then
                    .Tier1 = "DT RLS"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Not Applicable") And .Reason4.Contains("Outage") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Not Applicable") And .Reason4.Contains("Downday") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Not Applicable") And .Reason4.Contains("Brandswing") Then
                    .Tier1 = "Brand Change"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Production Planning") And .Reason2.Contains("Outage") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason1.Contains("Production Planning") And .Reason3.Contains("Process/Operational") Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason2.Contains("Outage") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""
                ElseIf .Reason2.Contains("Downday") Then
                    .Tier1 = "Downday"
                    .Tier2 = .Team
                    .Tier3 = ""

                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Team
                End If

            End If
        End With

    End Sub

#End Region

#Region "NaucalpanPHC"

    Public Sub getNaucalpanPHC_BprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                If .Reason1.Contains("Cartoneta") Then
                    .Tier1 = "Cartoneta"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Polypack") Then
                    .Tier1 = "Polypack"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Trayformer") Then
                    .Tier1 = "Trayformer"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason1.Contains("Wrap Ade") Then
                    .Tier1 = "Wrap Ade"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("No relacionados con el equipo") Then
                    .Tier1 = "Non-Equip"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Equals("Change Over") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team

                ElseIf .Reason2.Equals("Juntas") Then
                    .Tier1 = "Juntas"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Mantenimiento") Or .Reason2.Equals("Maintenance") Then
                    .Tier1 = "Mantenimiento"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Meeting") Then
                    .Tier1 = "Meeting"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Ini") Then
                    .Tier1 = "Initiatives"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Cost Saving") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Startup") Then
                    .Tier1 = "Startup"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Material") Then
                    .Tier1 = "Material Resupply"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Procedimiento calidad") Then
                    .Tier1 = "Procedimiento calidad"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                    '
                Else 'Planeadas
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                End If
            End If

            .DTGroup = .Tier1 & "-" & .Tier2
        End With
    End Sub
    Public Sub getNaucalpanPHC_JprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                If .Reason1.Contains("Cartoneta") Then
                    .Tier1 = "Cartoneta"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Algodonadora") Then
                    .Tier1 = "Algodonadora"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Etiquetadora") Then
                    .Tier1 = "Etiquetadora"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Llenadora") Then
                    .Tier1 = "Llenadora"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Polypack") Then
                    .Tier1 = "Polypack"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3

                ElseIf .Reason1.Contains("Sorter") Then
                    .Tier1 = "Sorter"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3

                ElseIf .Reason1.Contains("Tapadora") Then
                    .Tier1 = "Tapadora"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3

                ElseIf .Reason1.Contains("No relacionados con el equipo") Then
                    .Tier1 = "Non-Equip"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Equals("Change Over") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team

                ElseIf .Reason2.Equals("Juntas") Then
                    .Tier1 = "Juntas"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Mantenimiento") Or .Reason2.Equals("Maintenance") Then
                    .Tier1 = "Mantenimiento"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Meeting") Then
                    .Tier1 = "Meeting"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Ini") Then
                    .Tier1 = "Initiatives"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Cost Saving") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Startup") Then
                    .Tier1 = "Startup"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Material") Then
                    .Tier1 = "Material Resupply"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Procedimiento calidad") Then
                    .Tier1 = "Procedimiento calidad"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                    '
                Else 'Planeadas
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                End If
            End If

            .DTGroup = .Tier1 & "-" & .Tier2
        End With
    End Sub
    Public Sub getNaucalpanPHC_MexprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                If .Reason1.Contains("Wrap Ade") Then
                    .Tier1 = "Wrap Ade"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3


                ElseIf .Reason1.Contains("No relacionados con el equipo") Then
                    .Tier1 = "Non-Equip"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Equals("Change Over") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team

                ElseIf .Reason2.Equals("Juntas") Then
                    .Tier1 = "Juntas"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Mantenimiento") Or .Reason2.Equals("Maintenance") Then
                    .Tier1 = "Mantenimiento"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Meeting") Then
                    .Tier1 = "Meeting"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Ini") Then
                    .Tier1 = "Initiatives"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Cost Saving") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Startup") Then
                    .Tier1 = "Startup"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Material") Then
                    .Tier1 = "Material Resupply"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Procedimiento calidad") Then
                    .Tier1 = "Procedimiento calidad"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                    '
                Else 'Planeadas
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                End If
            End If

            .DTGroup = .Tier1 & "-" & .Tier2
        End With
    End Sub
    Public Sub getNaucalpanPHC_Vita1prstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                If .Reason1.Contains("Enflex") Then
                    .Tier1 = "Enflex"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3


                ElseIf .Reason1.Contains("No relacionados con el equipo") Then
                    .Tier1 = "Non-Equip"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Equals("Change Over") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team

                ElseIf .Reason2.Equals("Juntas") Then
                    .Tier1 = "Juntas"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Mantenimiento") Or .Reason2.Equals("Maintenance") Then
                    .Tier1 = "Mantenimiento"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("Meeting") Then
                    .Tier1 = "Meeting"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Ini") Then
                    .Tier1 = "Initiatives"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Cost Saving") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Startup") Then
                    .Tier1 = "Startup"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Material") Then
                    .Tier1 = "Material Resupply"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Procedimiento calidad") Then
                    .Tier1 = "Procedimiento calidad"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                    '
                Else 'Planeadas
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                End If
            End If

            .DTGroup = .Tier1 & "-" & .Tier2
        End With
    End Sub
#End Region

#Region "F&HC"
    Public Sub getHyderabadprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason3.Contains("UTILITIES") Then
                    .Tier1 = "Utilities"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                ElseIf .Reason4.Contains("MAKING") Then
                    .Tier1 = "Making"
                    .Tier2 = .Reason3
                    .Tier3 = .Fault
                ElseIf .Reason3.Contains("PACKING MATERIAL") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                ElseIf Left(.DTGroup, 5) = "Equip" Or .Reason1.Contains("ASL") Or .Reason1.Contains("AKASH") Then
                    If .Reason1.Contains("ASL") Then
                        .Tier1 = "Equip-ASL"
                    ElseIf .Reason1.Contains("AKASH") Then
                        .Tier1 = "Equip-AKASH"
                    Else
                        .Tier1 = "Equip-Others"
                    End If


                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("DC") Then
                    .Tier1 = "DC System"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Conveyor") Then
                    .Tier1 = "COMMON CONV"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("CVC") Then
                    .Tier1 = "CVC System"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3

                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason3
                    .Tier3 = .Reason4
                End If

            Else 'PLANNED
                If .Reason3.Contains("CIL") Or .Reason3.Contains("RLS") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Reason4
                ElseIf .Reason3.Contains("CHANGEOVER") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason4
                ElseIf .Reason3.Contains("MAINTENANCE") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason4
                ElseIf .Reason3.Contains("STARTUP") Or .Reason3.Contains("SHUTDOWN") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Reason4
                ElseIf .Reason3.Contains("ROLL End") Then
                    .Tier1 = "Roll Change"
                    .Tier2 = .Reason4
                ElseIf .Reason3.Contains("THREAD") Then
                    .Tier1 = "Thread Change"

                ElseIf .Reason3.Equals("EO") Then
                    .Tier1 = "EO"
                    .Tier2 = .Reason4

                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason3
                    .Tier3 = .Reason4
                End If
            End If

        End With
    End Sub
    Public Sub getRAKONAprstoryMapping(ByRef searchevent As DowntimeEvent)

        With searchevent
            If .isUnplanned Then

                If .Reason1.Contains("LPD") Then

                    .Tier1 = "Constraint"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Plnicka - Ronchi") Then

                    .Tier1 = "Constraint"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Sorter sroubu - cap feeder") Then

                    .Tier1 = "Constraint"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Sroubovacka - Ronchi") Then

                    .Tier1 = "Constraint"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Vaha lahvi") Then

                    .Tier1 = "Constraint"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Plnicka") Then

                    .Tier1 = "Constraint"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Sroubovacka") Then

                    .Tier1 = "Constraint"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Potisk lahvi Linx") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Rozevirac krabic") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Rozevirac krabic CERMEX") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Pakovac") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Horni lepeni") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Potisk a vazeni krabic") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Elevator") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Paletizer - ULF") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Ovinovacka") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravnik palet") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Potisk palet - Eprin") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Depaletizer") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Potisk lahvi - LINX") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Rozevirac krabic OTOR") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Pakovac (Casepacker)") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Horni lepeni (Case sealer)") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Spiralovy dopravnik") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Paletizer (ULF)") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Ovinovacka (Rolo)") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravniky palet") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Rozevirac krabic - Garbo") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Paletizer - Manex") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Divertor - TMG") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Garbo - Ronchi") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Pakovac - Ronchi") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Horni lepeni - Ronchi") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Ovinovacka - Tosa") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravnik palet - TMT") Then

                    .Tier1 = "Downstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("EXTERNI(Logistika,ESA,Manpower,Systemy)") Then

                    .Tier1 = "External"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("Nedostatek materialu") Then

                    .Tier1 = "Material supply"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Nedostatek produktu") Then

                    .Tier1 = "Material supply"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Zasobovani lahvi") Then

                    .Tier1 = "Material supply"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("KVALITA - MATERIALY") Then

                    .Tier1 = "Materials"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KVALITA - PRODUKTU") Then

                    .Tier1 = "Materials"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("KVALITA - MATERIALY") Then

                    .Tier1 = "QA"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KVALITA - PRODUKTU") Then

                    .Tier1 = "QA"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Radic lahvi") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravnik lahvi za radicem") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravnik lahvi za plnickou") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Etiketovacka - PAGO") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravnik lahvi za etiketovackou") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Radic lahvi (Unscrambler)") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Etiketovacka (Labeller)") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("COGNEX Kamera 2D kodu lahvi") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Pago (Sticker labeller)") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravnik lahvi za radicim stolem") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Dopravnik lahvi za sroubovackou") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("Etiketovacka") Then

                    .Tier1 = "Upstream"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2


                End If
            Else
                If .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni (12min)") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni (12min)") And .Reason3.Contains("Kodove etikety-kody-krabice") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni (12min)") And .Reason3.Contains("Standartni plnicka-etik-kod-krabice/lahv") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni (12min)") And .Reason3.Contains("Plne plnicka-etik-kod-krabice a lahve") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni (12min)") And .Reason3.Contains("Hercules sticker a Standartni prejizdeni") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni (12min)") And .Reason3.Contains("Zmena kodu") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni") And .Reason3.Contains("Velka prestavba (40min)") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni") And .Reason3.Contains("Mala prestavba (18min)") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PROCES PLAN") And .Reason2.Contains("PLANOVANE ZASTAVENI") And .Reason3.Contains("Barevne prejizdeni (BCO)") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PROCES PLAN") And .Reason2.Contains("PLANOVANE ZASTAVENI") And .Reason3.Contains("Zmena kodu") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PROCES PLAN") And .Reason2.Contains("PLANOVANE ZASTAVENI") And .Reason3.Contains("Formatova prestavba (SCO)") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni (34min)") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni") And .Reason3.Contains("Jine lahve-etikety-srouby") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni") And .Reason3.Contains("Jine lahve-etikety") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni") And .Reason3.Contains("Jine etikety") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni") And .Reason3.Contains("Prejeti s vymenou pasky") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Barevne prejizdeni") And .Reason3.Contains("Jine etikety a srouby") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Zmena kodu") And .Reason3.Contains("") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni") And .Reason3.Contains("Vcetne barevneho prejizdejni") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni") And .Reason3.Contains("Bez bareveneho prejizdeni") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni") And .Reason3.Contains("Vcetne barevneho prej. - Paletizer GB") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Formatove prejizdeni") And .Reason3.Contains("Bez barevneho prej. - Paletizer GB") Then

                    .Tier1 = "Changeover"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2



                ElseIf .Reason1.Contains("PROCES PLAN") And .Reason2.Contains("PLANOVANE ZASTAVENI") And .Reason3.Contains("Sanitace") Then

                    .Tier1 = "C&S"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Sanitace") Then

                    .Tier1 = "C&S"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("CIL/ Cisteni") Then

                    .Tier1 = "CIL/RLS"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("RLS") Then

                    .Tier1 = "CIL/RLS"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PROCES PLAN") And .Reason2.Contains("PLANOVANE ZASTAVENI") And .Reason3.Contains("RLS") Then

                    .Tier1 = "CIL/RLS"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("Cisteni dopln komentar") Then

                    .Tier1 = "CIL/RLS"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("Ostatni - dopln komentar") Then

                    .Tier1 = "CIL/RLS"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("Planovana udrzba") Then

                    .Tier1 = "Maintenance"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("Prace servisniho technika") Then

                    .Tier1 = "Maintenance"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("Planovane odstraneni abnormality pri CIL") Then

                    .Tier1 = "Maintenance"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("Vymena prolozek") Then

                    .Tier1 = "Raw material change"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Vymena etiket") Then

                    .Tier1 = "Raw material change"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Vymena bile pasky") Then

                    .Tier1 = "Raw material change"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Planovane zastaveni") And .Reason3.Contains("IWS aktivity (Autonomni udrzba)") Then

                    .Tier1 = "Training"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Tymova porada") Then

                    .Tier1 = "Training"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PLANOVANE ZASTAVENI") And .Reason2.Contains("Skoleni") Then

                    .Tier1 = "Training"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                Else

                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    '   .Tier3 = .Reason2
                End If
            End If

        End With

    End Sub
#End Region

#Region "FemCare"
    Public Sub getJijonaUltraprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            .Tier1 = "Otros"
            .Tier3 = .Reason2
            If .isUnplanned Then
                .Tier2 = .Reason1
                If .Location.Contains("Area 1") Then
                    .Tier1 = "Area 1"
                ElseIf .Location.Contains("Area 2") Then
                    .Tier1 = "Area 2"
                ElseIf .Location.Contains("Area 3") Then
                    .Tier1 = "Area 3"
                ElseIf .Location.Contains("Area 4") Then
                    .Tier1 = "Area 4"
                ElseIf .Location.Contains("OFFLINE") Then
                    If .Reason1.Contains("INSTALACIONES") Then
                        .Tier1 = "OffLine"
                    Else
                        .Tier1 = "Otros"
                    End If
                ElseIf .Location.Contains("Area 4") Then
                    .Tier1 = "Area 4"
                ElseIf .Location.Contains("NO AREA") Or .Location.Contains("UNKNOWN") Then
                    .Tier1 = "No Codificado"
                ElseIf .Reason2.Contains("FALLO IMAJIE") Then
                    .Tier1 = "Area 4"
                End If
            Else ' planned
                .Tier2 = .Team
                If .Location.Contains("UNKNOWN") And .Reason1.Contains("ERROR SEPARAR TIEMPO PARO") Then
                    .Tier1 = "Otros"
                ElseIf .Location.Contains("Planned Stops") And .Reason1.Contains("990") Then
                    .Tier1 = "Cambio Talla"
                ElseIf .Location.Contains("Planned Stops") And .Reason1.Contains("991") Then
                    .Tier1 = "Cambio Referencia"
                ElseIf .Location.Contains("Planned Stops") And .Reason1.Contains("992") Then
                    .Tier1 = "Cambio Producto"
                ElseIf .Location.Contains("Planned Stops") And .Reason1.Contains("993") Then
                    .Tier1 = "AM"
                ElseIf .Location.Contains("Planned Stops") And .Reason1.Contains("994") Then
                    .Tier1 = "PM"
                ElseIf .Location.Contains("Planned Stops") And .Reason1.Contains("995") Then
                    .Tier1 = "Proyectos"
                ElseIf .Location.Contains("Planned Stops") And .Reason1.Contains("996") Then
                    .Tier1 = "Organización"
                ElseIf .Location.Contains("Planned Stops") Then
                    .Tier1 = "Otros"

                ElseIf .Reason2.Contains("CAMBIO DE PRODUCTO") Then
                    .Tier1 = "Cambio Producto"
                ElseIf .Reason2.Contains("CAMBIO DE REFERENCIA") Then
                    .Tier1 = "Cambio Referencia"
                ElseIf .Reason2.Contains("CAMBIO DE TALLA") Then
                    .Tier1 = "Cambio Talla"
                ElseIf .Reason2.Contains("MANTENIMIENTO") Then
                    .Tier1 = "PM"
                ElseIf .Reason2.Contains("PARO ADMINISTRATIVO") Then
                    .Tier1 = "Otros"
                ElseIf .Reason2.Contains("PARO IWS") Then
                    .Tier1 = "AM"
                ElseIf .Reason2.Contains("PROYECTOS") Then
                    .Tier1 = "Proyectos"
                End If
            End If
        End With
    End Sub


    Public Sub getBudapestLCCprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                If .Reason1.Contains("MALOMHAZ") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MAGFORMAZO DOB") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STORA") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CBA RAGASZTO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KRIMPER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PFA RAGASZTO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KALANDER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VEGKES") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BACKSHEET LETEKERO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("EMBOSSER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPT LETEKERO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("LOTIONSYSTEM") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FOHORDSZALAG") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("TOPSHEET LETEKERO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PARFUM") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RING ROLLING") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MAG ATADO DOB") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SFA RAGASZTO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPT CUT & SLIP") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MRP CUT & SLIP") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RST LETEKERO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PRE-KALANDER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RST") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WLA / VSA RAGASZTO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("HUTO VEZETEK") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MRP LETEKERO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PRINTER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("REJECT 1") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("POUCH TRI-FOLDING") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("POUCH KRIMPER/KES") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BETETHAJTOGATO") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("POUCH LETEKERO") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SZARNYHAJTOGATO") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("HA") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("POUCH KES") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FORDITOSZALAG") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO FILMLETEKERO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("3M") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO PUSHER") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO PALCAS") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PACKAGING KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO INFEED") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER UCS") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO OLDALSZALAG") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO HOSSZHEGESZTO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO KIMENETI KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO KERESZTHEGESZTO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO FORMAZOVALL") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("M9") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO REPACK FEEDER") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO LIFT") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO OUTFEED") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHRISTO OHP") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("infeed konvejorok") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                End If
            Else


                If .Reason1.Contains("ATALLAS TCO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T5 IP 2.0 - T5 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1S NW - T1 IP3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS CCO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS SCO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 NW FRESH - T3 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T5 IP 2.0 - T5 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS PMC") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3S IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS Single - DUO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS DUO - Single") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 FRESH - T5 IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3S IP 2.0 - T3 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 FRESH - T1 IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 2.0 FRESH - T3 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 FRESH - T3S IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 NW FRESH - T3 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1S IP2.0 - T1 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP2.0 - T1 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1    T1 IP3.0 - T1S IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 FRESH - T3S IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3  T3 IP 3.0 FRESH - T3 IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T3 IP 2.0 NW -T5 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 FRESH - T3 NW") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T3S IP 2.0 - T5 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1S NW FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 NW - T3 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3 NW") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 FRESH - T1 IP3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 - T5 IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS SCO LCC1   T2S IP2.0 - T1 IP3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3  T3 IP 2.0 - T3 IP 3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1 IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3S IP 2.0 - T3 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0  FRESH - T3 IP 2.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T3 IP 2.0 NW -T5 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T5 IP 3.0 FRESH - T3 IP 2.0 NW") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 FRESH - T3 NW FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3  T3 IP 3.0 - T3 IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1S NW FRESH - T1 IP3.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1S NW") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 FRESH - T1S  IP 2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 NW - T3 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3  T3 IP 2.0 - T3 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC4 T3S IP 2.0 - T5 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1S NW - T1 IP3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1 IP3.0 - T1 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC1   T1S IP2.0 - T1 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS SCO LCC1   T2S IP2.0 - T1 IP3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 2.0 FRESH - T3 IP 3.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3 NW FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO LCC3 T3 IP 3.0 - T3 IP 2.0 FRESH") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS SCO LCC1   T1 IP3.0 FRESH - T2S IP2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS SCO LCC1   T1 IP3.0 - T2S IP2.0") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("TEK") Then

                    .Tier1 = "CIL"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("SOR LEALLITAS - INDITAS") Then

                    .Tier1 = "Logistics"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MEETING") Then

                    .Tier1 = "Meeting"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PM02 - KARBANTARTAS") Then

                    .Tier1 = "PM"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("KVALIFIKACIO") Then

                    .Tier1 = "Projects"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("EO") Then

                    .Tier1 = "Projects"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("TRAINING") Then

                    .Tier1 = "Training"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("GEPEGYSEG FELELOS Stop") Then

                    .Tier1 = "Training"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason3
                End If

            End If
        End With
    End Sub
    Public Sub getBudapestFGCprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then

                If .Reason1.Contains("AWA RAGASZTAS") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("AWA RAGASZTO RENDSZER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("AWA-CIA TANK") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CIA RAGASZTAS") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CPM LETEKERES") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CPM LETEKERES & SZALLITO RENDSZER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CPM VALTOEGYSEG") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CREASER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("EX HEX LETEKERES") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FUSION BOND") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FUSION BOND BEVEZETO") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FUSION BOND ERZEKELO SOR") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("GLUE & EMBOSSER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MAG FIFE") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MAG LAMINAT KABIN") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MAGKABIN") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MAGKES") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MAGKES ALATTI KONVEJOR") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OVB FOEGYSEG") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OVERBONDING") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RING ROLLING SELF") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RRS BEVEZETO KONVEJOR") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RRS FOEGYSEG") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SCENT SELENA") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SPREADER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS ATVEVODOB") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS BEVEZETO KONV") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS FIFE") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS KES") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS KIVEZETO FELSO KONVEJOR (SPUNLACE)") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS KIVEZETO KONVEJOR") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS LETEKERES") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS LETEKERES & SZALLITO RENDSZER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STS VALTOEGYSEG") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("TBA RAGASZTO RENDSZER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VIDEO JET PRINTER") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VIDEOJET") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VIDEOJET COGNEX") Then

                    .Tier1 = "Area 1"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("CLEO MAGKES") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BACKSHEET LETEKERES/SZALLITAS") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WLA RAGASZTO RENDSZER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SCENT(NATURELLA)") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RING ROLLING (P77)") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ANYAG HUTES") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VEGKES") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WLA BONDING") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SCENT(ANGEL_LOVE)") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KRIMPER - FOEGYSEG") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WLA BONDING - BEVEZETO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BACKSHEET VALTOEGYSEG") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VEGKES BEVEZETO KONVEJOR") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KRIMPER - BEVEZETO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHILL ROLL") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPT BEHUZO KONVEJOR") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("LOTIONSYSTEM") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KRIMPER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SBAA RAGASZTO RENDSZER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RING ROLLING BEVEZETO KONVEJOR") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VEGKES KIVEZETO KONVEJOR") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("P77 KIVEZETO KONVEJOR") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BACKSHEET LETEKERES & SZALLITAS") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KRIMPER -PNEUMATIKA") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("TRANSITION 2 KONVEJOR") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KRIMPER - ELEKTROMOS") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("LOTION") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SCENT(MARGARETA)") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SCENT(FRUITY_FLORAL)") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FRUIT FLORAL") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WLA TANK") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPT KES") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WLA RAGASZTAS") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("TRANSITION MODUL #2") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BAA RAGASZTO RENDSZER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ONM") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SCENT(GREEN_TEA)") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SFA NYOMOGORGO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SCENT(CALENDULA)") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VEGKES ATVEVODOB") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RING ROLLING HUTES RENDSZER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("VEGKES OLAJZO GORGO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BETETKODOLO") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("CHILL ROLL HUTES RENDSZER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RING ROLLING") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WLA BONDING HUTES RENDSZER") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("LOTION TANK") Then

                    .Tier1 = "Area 2"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2


                ElseIf .Reason1.Contains("DISCHARGE DRUM") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPT EGYSEG") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RPW REJECT #2") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("REJECT2 HIBAS SELEJT LEFUJAS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FORDITO KEREK ATVEVO DOB") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("REJECT2 KIVEZETO ALSO KONVEJOR") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SFA RAGASZTO RENDSZER") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KES/KRIMPER") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("UJRAZARO LETEKERES") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("UJRAZARO EGYSEG") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KES/KRIMPER - FOEGYSEG") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PFA RAGASZTAS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("REJECT2 KONVEJOR") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPT VALTOEGYSEG") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("POUCH LETEKERES & SZALLITAS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SFA RAGASZTAS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SZARNY HAJTOGATO") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FORDITO KEREK BEVEZETO FELSO KONVEJOR") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FORDITO KEREK") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPT LETEKERES") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("RPW REJECT #1") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FORDITO KEREK BEVEZETO") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PFA TANK") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KES/KRIMPER - BEVEZETO") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("REJECT1 BEVEZETO KONVEJOR") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("REJECT2 KIVEZETO FELSO KONVEJOR") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KES/KRIMPER - HAJTAS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KES/KRIMPER -PNEUMATIKA") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SFA TANK") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("REJECT1 LEFUJAS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("POUCH TRI-FOLDING") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("FORDITO KEREK BEVEZETO ALSO KONVEJOR") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("PFA RAGASZTO RENDSZER") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SZITAKONVEJOR") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KES/KRIMPER - ELEKTROMOS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("KES/KRIMPER - KIVEZETO") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SZARNY HAJTOGATAS") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MELTEX SZEKRENY") Then

                    .Tier1 = "Area 3"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 1") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-CSOROS") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-HEGESZTO EGYSEG") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-1") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("DIVERTER") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 2 - OHP") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 1- KAZETTAS") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM2-HEGESZTO EGYSEG") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM-2") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SZAKASZOLO KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("U KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("L KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("WEDGE KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP HIBA") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SZALLITO SZALAG") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP-LANC KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-PALCAS LANC") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("3M") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("MERLEG") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP ELOTTI KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP-BEVEZETO KONVEJOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM2-ZACSI FELVETEL") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-ZACSI FELVETEL") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 1 - OHP") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("BIS") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP-SIAT-DOBOZHAJTOGATO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM2-PALCAS LANC") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 2 - INFEED") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM2-CSOROS") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("METAL DETECTOR") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 2 - OUTFEED") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 1 - INFEED") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP-P&P FEJ HIBA") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 2") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("DIVERTER - NOSE") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-OHP") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-ZACSKOBEHORDO-WICKET") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STICKER 1") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STICKER 2") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM2-ZACSKOBEHORDO-WICKET") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("DIVERTER - GATE4") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP-CSOPORTOSITO ASZTAL") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("DIVERTER-TRANSITION") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP-LIFT") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OS2E-KAZETTA-ZOMITO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ACP-DOBOZLEZARO 3M") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 1 - OUTFEED") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("STACKER 2- KAZETTAS") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM2-KAZETTA-ZOMITO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ZACSI KODOLO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("OPTIMA OM2-OHP") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ELEKTROMOS PROGRAM HIBA") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("DOBOZFELHORDO_1") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("DOBOZFELHORDO_2") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ALAPANYAG HIBA ZACSKO") Then

                    .Tier1 = "Area 4"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2
                Else

                    .Tier1 = "Others"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else 'PLANNED
                If .Reason1.Equals("TEK") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Reason1
                ElseIf .Reason1.Contains("MOI_STOP") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Reason1
                ElseIf .Reason1.Contains("RLS") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Reason1


                ElseIf .Reason1.Contains("ATALLAS SCO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS TCCO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS CCO OS2 DUO - Single") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS  KAZETTAS - ZOMITOS") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS  ZOMITOS - KAZETTAS") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS STCO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS PMC") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS Single - DUO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS DUO - Single") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS CCO OM2 DUO - Single") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS CCO OS2 Single - DUO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("ATALLAS CCO OM2 Single - DUO") Then

                    .Tier1 = "CO"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2



                ElseIf .Reason1.Contains("SOR LEALLITAS - INDITAS") Then

                    .Tier1 = "Logistics"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SOR INDITAS") Then

                    .Tier1 = "Logistics"

                    .Tier2 = .Reason1

                    .Tier3 = .Reason2

                ElseIf .Reason1.Contains("SOR LEALLITAS") Then

                    .Tier1 = "Logistics"

                    .Tier2 = .Reason1



                ElseIf .Reason1.Contains("SEGEDBERENDEZESEK") Then

                    .Tier1 = "Meeting"

                    .Tier2 = .Reason1

                ElseIf .Reason1.Contains("KIURITES") Then

                    .Tier1 = "Meeting"

                    .Tier2 = .Reason1
                ElseIf .Reason1.Contains("PM02") Then
                    .Tier1 = "PM"
                    .Tier2 = .Reason1

                ElseIf .Reason1.Contains("EO") Then

                    .Tier1 = "Projects"

                    .Tier2 = .Reason1

                ElseIf .Reason1.Contains("KVALIFIKACIO") Then

                    .Tier1 = "Projects"

                    .Tier2 = .Reason1

                ElseIf .Reason1.Contains("TRAINING") Then

                    .Tier1 = "Training"

                    .Tier2 = .Reason1

                ElseIf .Reason1.Contains("GEPEGYSEG FELELOS Stop") Then

                    .Tier1 = "Training"

                    .Tier2 = .Reason1

                ElseIf .Reason1.Contains("SEGEDBERENDEZESEK") Then

                    .Tier1 = "Others"

                    .Tier2 = .Reason1

                ElseIf .Reason1.Contains("FLOW To WORK") Then

                    .Tier1 = "Others"

                    .Tier2 = .Reason1
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1

                End If
            End If
        End With
    End Sub

    Public Sub getFemCare_Pads_HPMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Location.Equals("Area 1") Or .Location.Equals("AREA 1") Or .Location.Equals("Area1") Then
                    .Tier1 = "Area 1"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 2") Or .Location.Equals("AREA 2") Or .Location.Equals("Area2") Then
                    .Tier1 = "Area 2"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 3") Or .Location.Equals("AREA 3") Or .Location.Equals("Area3") Then
                    .Tier1 = "Area 3"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 4") Or .Location.Equals("AREA 4") Or .Location.Equals("Area4") Then
                    .Tier1 = "Area 4"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason1.Contains("CIL") Or .Reason2.Contains("RLS") Or .Reason2.Contains("01_CIL") Or .Reason2.Contains("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .Reason1.Equals("Changeover") Or .Reason2.Contains("03_Changeover") Or .Reason2.Contains("ChangeOver") Or .Reason1.Contains("CHANGE") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason2.Contains("AM") Then
                    .Tier1 = "AM"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .Reason2.Contains("PM") Then
                    .Tier1 = "PM"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .Reason2.Contains("ORGANIZATION") Or .Reason1.Contains("ORGANIZATION") Then
                    .Tier1 = "Org"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .Reason2.Contains("Project Work") Or .Reason2.Contains("PROJECT") Then
                    .Tier1 = "Project"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .Reason1.Equals("EO sellable") Then
                    .Tier1 = "EO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason1.Equals("MEETING") Then
                    .Tier1 = "Meeting"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason1.Equals("Logistics") Then
                    .Tier1 = .Reason1
                    .Tier2 = .Team
                    .Tier3 = .Product
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub
    Public Sub getBellevilleprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Location.Equals("Area 1") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 2") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 3") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 4") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else

                If .Reason1.Contains("CIL") Or .Reason2.Contains("RLS") Or .Reason2.Contains("01_CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team
                    .Tier3 = .Product
                ElseIf .Reason1.Equals("Changeover") Or .Reason2.Contains("03_Changeover") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason1.Equals("EO sellable") Then
                    .Tier1 = "EO"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason1.Equals("MEETING") Then
                    .Tier1 = "Meeting"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason1.Equals("Logistics") Then
                    .Tier1 = .Reason1
                    .Tier2 = .Team
                    .Tier3 = .Product
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub
    Public Sub getBoryspilprstoryMapping(ByRef searchevent As DowntimeEvent)
        Dim A1 As String = "Area 1"
        Dim A2 As String = "Area 2"
        Dim A3 As String = "Area 3"
        Dim A4 As String = "Area 4"
        '   Dim A5 As String = "Area/площадь 5"
        '   Dim A6 As String = "Area/площадь 6"

        With searchevent
            If .isUnplanned Then
                If .Location.Equals("Area 1") Then
                    .Tier1 = A1
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 2") Then
                    .Tier1 = A2
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 3") Then
                    .Tier1 = A3
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 4") Then
                    .Tier1 = A4
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Contains("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("ПЕРЕХОД_") Or .Reason1.Contains("changeover") Or .Reason1.Contains("Change") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason2
                ElseIf .Reason3.Contains("Purge") Then
                    .Tier1 = "Purge"
                    .Tier2 = .Team
                ElseIf .Reason2.Equals("AM") Then
                    .Tier1 = "AM"
                    .Tier2 = .Team
                ElseIf .Reason1.Contains("ЗАМЕНА") Then
                    .Tier1 = "ЗАМЕНА МАТРИЦЫ"
                    .Tier2 = .Product
                Else

                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    '   .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub

    Public Sub getTepejiFemprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Location.Equals("Area 1") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 2") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 3") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Location.Equals("Area 4") Then
                    .Tier1 = .Location
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Contains("CHANGEOVER") Or .Reason1.Contains("CHANGE OVER") Or .Reason2.Contains("CHANGE OVER") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("ADMIN") Then
                    .Tier1 = "ADMIN SHUTDOWN"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("IWS") Then
                    .Tier1 = "IWS SHUTDOWN"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("MAINT") Then
                    .Tier1 = "MAINTENANCE"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("PROJECTS") Then
                    .Tier1 = "PROJECTS"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("OPERATIONAL") Then
                    .Tier1 = "OPERATIONAL"
                    .Tier2 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                End If
            End If

        End With
    End Sub

#End Region

#Region "GBO"
    Public Sub getOralCareCruxprstoryMapping(ByRef searchEvent As DowntimeEvent) ', ByRef Tier1 As String, ByRef Tier2 As String, ByRef Tier3 As String)
        With searchEvent
            .Tier1 = "Others"
            .Tier2 = .Reason1
            .Tier3 = .Reason2
            If .isUnplanned Then
                If .Reason1.Equals("Qualidade de Material") Then
                    .Tier1 = "Material"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                    Exit Sub
                End If

                If .Reason1.Equals("Perdas de Fornecimento") Then
                    If .Reason4.Contains("Matéria-prima de Making não entregue; Falha de Making; Quantidade entregue a menos") Then
                        .Tier1 = "Paste"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                        Exit Sub
                    End If
                End If

                If .Reason1.Equals("Perdas de Fornecimento") Then
                    .Tier1 = "Supply Losses"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                    Exit Sub
                End If

                If .Reason1.Equals("Sistemas da Planta & Outros") Then
                    If .Reason2.Equals("Utilidades") Then
                        .Tier1 = "Utilities"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                        Exit Sub
                    End If
                End If

                If .Reason1.Equals("Paste Supply") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "PS"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Envasadeira 1") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 1"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Descarregador de Tubos 1") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 1"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Envasadeira 2") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 2"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Descarregador de Tubos 2") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 2"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Encartuchadeira") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Cartoner"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Balança OCS") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Checkweigher"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Embaladora (S-100)") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Bundler"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Encaixotadora (YZX20U)") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "CasePacker"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Paletizadora (Kuka)") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Palletizer"
                    .Tier3 = .Reason2
                    Exit Sub
                End If

                If .Reason1.Equals("Filmadora de Palete (Lantech)") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "StretchWrapper"
                    .Tier3 = .Reason2
                    Exit Sub
                End If


            Else 'planned
                If .Reason1.Equals("Parada Planejada") Then
                    If .Reason2.Contains("Manutenção") Then
                        If .Reason3.Contains("CIL") Then
                            .Tier1 = "CIL"
                            .Tier2 = .Reason2
                            .Tier3 = .Team
                            Exit Sub
                        End If
                    End If
                End If

                If .Reason1.Equals("Parada Planejada") Then
                    If .Reason2.Equals("Manutenção") Then
                        .Tier1 = "Maintenance"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                        Exit Sub
                    End If
                End If

                If .Reason1.Equals("Parada Planejada") Then
                    If .Reason2.Contains("Reunião") Then
                        .Tier1 = "Meeting"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                        Exit Sub
                    End If
                End If

                If .Reason1.Equals("Parada Planejada") Then
                    If .Reason2.Equals("CO") Then
                        .Tier1 = "CO"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                        Exit Sub
                    End If
                End If

                If .Reason1.Equals("Parada Planejada") Then
                    If .Reason2.Equals("Reabastecimento de Material") Then
                        .Tier1 = "Material"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                        Exit Sub
                    End If
                End If

                If .Reason1.Equals("Parada Planejada") Then
                    If .Reason2.Contains("Desligamento") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                        Exit Sub
                    End If
                End If

                If .Reason1.Equals("Parada Planejada") Then
                    If .Reason2.Contains("Partida de Linha") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                        Exit Sub
                    End If
                End If




            End If
        End With
    End Sub

    Public Sub getOralCareprstoryMapping(ByRef searchEvent As DowntimeEvent) ', ByRef Tier1 As String, ByRef Tier2 As String, ByRef Tier3 As String)
        With searchEvent
            'check if its blank
            If Len(searchEvent.DTGroup) < 2 Then
                .Tier1 = OTHERS_STRING

                Exit Sub
            End If

            'ok not blank

            If searchEvent.isUnplanned Then
                If .DTGroup.Contains("Equip") Then
                    .Tier1 = "Equipment"
                    If .DTGroup.Contains("Equip-Filler") Then
                        .Tier2 = "Filler"
                        If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                            .Tier3 = "Tube Unloading"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                            .Tier3 = "Tube Transport"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                            .Tier3 = "Tube Filling"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                            .Tier3 = "Tube Sealing"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                            .Tier3 = "Tube Trimming"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                            .Tier3 = "Tube Discharge"
                        ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                            .Tier3 = "F_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Cartoner") Then
                        .Tier2 = "Cartoner"
                        If .DTGroup.Equals("Equip-Cartoner-CartonFeeding") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonPicking") Then
                            .Tier3 = "Carton Picking"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeInsertion") Then
                            .Tier3 = "Tube Insertion"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeTransfer") Then
                            .Tier3 = "Tube Transfer"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonClosing") Then
                            .Tier3 = "Carton Closing"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonCoding") Then
                            .Tier3 = "Carton Coding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-VisionSystem") Then
                            .Tier3 = "Vision System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Discharge") Then
                            .Tier3 = "Discharge"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Electrical") Then
                            .Tier3 = "CT_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bundler") Then
                        .Tier2 = "Bundler"
                        If .DTGroup.Equals("Equip-Bundler-CartonInfeed") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Bundler-StackingArea") Then
                            .Tier3 = "Stacking"
                        ElseIf .DTGroup.Equals("Equip-Bundler-FilmFeeding") Then
                            .Tier3 = "Film Transport"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Cut&Wrap") Then
                            .Tier3 = "Wrap Film Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Fold&Seal") Then
                            .Tier3 = "Film Fold End Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Discharge") Then
                            .Tier3 = "Bundle Discharge"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Electrical") Then
                            .Tier3 = "B_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-CP") Then
                        .Tier2 = "Casepacker"
                        If .DTGroup.Equals("Equip-CP-CaseEntrance") Then
                            .Tier3 = "Case Entrance"
                        ElseIf .DTGroup.Equals("Equip-CP-CartonInfeed") Then
                            .Tier3 = "Carton Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-Infeed") Then
                            .Tier3 = "Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-CP-StackingArea") Then
                            .Tier3 = "Stacking Area"
                        ElseIf .DTGroup.Equals("Equip-CP-PushIntoCase") Then
                            .Tier3 = "PushInCase"
                        ElseIf .DTGroup.Equals("Equip-CP-CaseTransport") Then
                            .Tier3 = "Case Transport"
                        ElseIf .DTGroup.Equals("Equip-CP-Electrical") Then
                            .Tier3 = "CP_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bulkcases") Then
                        .Tier2 = "Bulk packer"
                        If .DTGroup.Equals("Equip-BulkCaseErector") Then
                            .Tier3 = "Case Former"
                        ElseIf .DTGroup.Equals("Equip-BulkCasePacker") Then
                            .Tier3 = "Tray Packer"
                        ElseIf .DTGroup.Equals("Equip-BulkCaseSealer") Then
                            .Tier3 = "Case Sealer"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Case") Or .DTGroup.Contains("Equip-Scan") Then
                        .Tier2 = "End Of line"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-Palletizer") Then
                        .Tier2 = "Palletizer"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-BulkCasePacker") Then
                        .Tier2 = "Bulk CP"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Pump") Then
                        .Tier2 = "PS-Pump"
                        .Tier3 = .Reason3
                    Else
                        .Tier2 = OTHERS_STRING
                        .Tier3 = .Reason3
                    End If
                ElseIf .DTGroup.Equals("Quality-Tubes") Or .DTGroup.Equals("Quality-Cartons") Or .DTGroup.Equals("Quality-Film") Or .DTGroup.Equals("Quality-Caseblanks") Then
                    .Tier1 = "Materials"

                    If .DTGroup.Contains("Tube") Then
                        .Tier2 = "Tube Quality"
                    ElseIf .DTGroup.Contains("Carton") Then
                        .Tier2 = "Carton Quality"
                    ElseIf .DTGroup.Contains("Case") Then
                        .Tier2 = "Caseblank Quality"
                    Else
                        .Tier2 = "Other Quality"
                    End If
                    .Tier3 = .Reason4

                ElseIf .DTGroup.Equals("Supply-PackMaterial") Then
                    If .Reason4 = "Material not in warehouse" Or .Reason4 = "Underdelivered quantity" Then
                        .Tier1 = "Materials"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4
                    Else

                        .Tier1 = "Supply Losses"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4

                    End If

                ElseIf .DTGroup.Equals("Supply-Making") Or .DTGroup.Equals("Quality-Paste") Then
                    .Tier1 = "Paste"
                    If .DTGroup.Equals("Quality-Paste") Then
                        .Tier2 = "Quality" '"Paste Quality"
                        .Tier3 = .Product
                    ElseIf .DTGroup.Equals("Supply-Making") Then
                        .Tier2 = "Availability" '"Paste Availability"
                        .Tier3 = .Product
                    Else
                        .Tier2 = .Reason1
                        .Tier3 = .Reason2
                    End If
                ElseIf .DTGroup.Equals("Systems-Utilities") Then
                    .Tier1 = "Utilities"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2




                Else
                    .Tier2 = .DTGroup
                    .Tier3 = .Reason1
                    'here are modifications 
                    If .Reason1.Contains("Material") Then
                        .Tier1 = "Materials"

                    ElseIf .Reason1.Equals("Encartuchadeira") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Cartoner"
                    Else

                        .Tier1 = OTHERS_STRING
                    End If
                End If

                'PLANNED LOSSES
            ElseIf searchEvent.isPlanned Then
                If .DTGroup.Contains("C/O") Then
                    .Tier1 = "CO"
                    .Tier2 = .Team
                    If .DTGroup.Equals("C/O-Size") Then
                        .Tier2 = "Size"
                    ElseIf .DTGroup.Equals("C/O-WO+Size") Then
                        .Tier2 = "WO+Size"
                    ElseIf .DTGroup.Equals("C/O-WO") Then
                        .Tier2 = "WO"
                    ElseIf .DTGroup.Equals("C/O-PO") Then
                        .Tier2 = "PO Change"
                    ElseIf .DTGroup.Equals("C/O-Pigging") Then
                        .Tier2 = "Pigging"
                    ElseIf .DTGroup.Equals("C/O-WO-Sanitization") Then
                        .Tier2 = "Sanitization"
                    ElseIf .DTGroup.Equals("C/O-Platform") Then
                        .Tier2 = "Platform"
                    End If
                ElseIf .DTGroup.Equals("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team '.Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Maintenance") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Training/Meeting") Then
                    .Tier1 = "Train/Meet"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("ProjectWork") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Material resupply") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Contains("SU/SD") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Team '.Reason1 & "-" & .Reason2



                Else
                    .Tier2 = .Reason2
                    If .Reason2.Equals("Changeover") Then
                        .Tier1 = "CO"
                    Else
                        .Tier1 = OTHERS_STRING

                    End If
                End If
                .Tier3 = .Team

            End If
        End With
    End Sub
    Public Sub getOralCareprstoryMappingDF(ByRef searchEvent As DowntimeEvent) ', ByRef Tier1 As String, ByRef Tier2 As String, ByRef Tier3 As String)
        With searchEvent

            If searchEvent.isUnplanned Then
                If .DTGroup.Contains("Equip") Then
                    .Tier1 = "Equipment"
                    If .DTGroup.Contains("Equip-Filler") And .Reason1.Contains("1") Then
                        .Tier2 = "Filler 1"
                        If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                            .Tier3 = "Tube Unloading"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                            .Tier3 = "Tube Transport"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                            .Tier3 = "Tube Filling"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                            .Tier3 = "Tube Sealing"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                            .Tier3 = "Tube Trimming"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                            .Tier3 = "Tube Discharge"
                        ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                            .Tier3 = "F_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Filler") And .Reason1.Contains("2") Then
                        .Tier2 = "Filler 2"
                        If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                            .Tier3 = "Tube Unloading"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                            .Tier3 = "Tube Transport"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                            .Tier3 = "Tube Filling"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                            .Tier3 = "Tube Sealing"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                            .Tier3 = "Tube Trimming"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                            .Tier3 = "Tube Discharge"
                        ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                            .Tier3 = "F_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Cartoner") Then
                        .Tier2 = "Cartoner"
                        If .DTGroup.Equals("Equip-Cartoner-CartonFeeding") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonPicking") Then
                            .Tier3 = "Carton Picking"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeInsertion") Then
                            .Tier3 = "Tube Insertion"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeTransfer") Then
                            .Tier3 = "Tube Transfer"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonClosing") Then
                            .Tier3 = "Carton Closing"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonCoding") Then
                            .Tier3 = "Carton Coding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-VisionSystem") Then
                            .Tier3 = "Vision System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Discharge") Then
                            .Tier3 = "Discharge"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Electrical") Then
                            .Tier3 = "CT_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bundler") Then
                        .Tier2 = "Bundler"
                        If .DTGroup.Equals("Equip-Bundler-CartonInfeed") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Bundler-StackingArea") Then
                            .Tier3 = "Stacking"
                        ElseIf .DTGroup.Equals("Equip-Bundler-FilmFeeding") Then
                            .Tier3 = "Film Transport"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Cut&Wrap") Then
                            .Tier3 = "Wrap Film Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Fold&Seal") Then
                            .Tier3 = "Film Fold End Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Discharge") Then
                            .Tier3 = "Bundle Discharge"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Electrical") Then
                            .Tier3 = "B_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-CP") Then
                        .Tier2 = "Casepacker"
                        If .DTGroup.Equals("Equip-CP-CaseEntrance") Then
                            .Tier3 = "Case Entrance"
                        ElseIf .DTGroup.Equals("Equip-CP-CartonInfeed") Then
                            .Tier3 = "Carton Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-Infeed") Then
                            .Tier3 = "Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-CP-StackingArea") Then
                            .Tier3 = "Stacking Area"
                        ElseIf .DTGroup.Equals("Equip-CP-PushIntoCase") Then
                            .Tier3 = "PushInCase"
                        ElseIf .DTGroup.Equals("Equip-CP-CaseTransport") Then
                            .Tier3 = "Case Transport"
                        ElseIf .DTGroup.Equals("Equip-CP-Electrical") Then
                            .Tier3 = "CP_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bulkcases") Then
                        .Tier2 = "Bulk packer"
                        If .DTGroup.Equals("Equip-BulkCaseErector") Then
                            .Tier3 = "Case Former"
                        ElseIf .DTGroup.Equals("Equip-BulkCasePacker") Then
                            .Tier3 = "Tray Packer"
                        ElseIf .DTGroup.Equals("Equip-BulkCaseSealer") Then
                            .Tier3 = "Case Sealer"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Case") Or .DTGroup.Contains("Equip-Scan") Then
                        .Tier2 = "End Of line"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-Palletizer") Then
                        .Tier2 = "Palletizer"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-BulkCasePacker") Then
                        .Tier2 = "Bulk CP"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Pump") Then
                        .Tier2 = "PS-Pump"
                        .Tier3 = .Reason3
                    Else
                        .Tier2 = OTHERS_STRING
                        .Tier3 = .Reason3
                    End If
                ElseIf .DTGroup.Equals("Quality-Tubes") Or .DTGroup.Equals("Quality-Cartons") Or .DTGroup.Equals("Quality-Film") Or .DTGroup.Equals("Quality-Caseblanks") Then
                    .Tier1 = "Materials"

                    If .DTGroup.Contains("Tube") Then
                        .Tier2 = "Tube Quality"
                    ElseIf .DTGroup.Contains("Carton") Then
                        .Tier2 = "Carton Quality"
                    ElseIf .DTGroup.Contains("Case") Then
                        .Tier2 = "Caseblank Quality"
                    Else
                        .Tier2 = "Other Quality"
                    End If
                    .Tier3 = .Reason4

                ElseIf .DTGroup.Equals("Supply-PackMaterial") Then
                    If .Reason4 = "Material not in warehouse" Or .Reason4 = "Underdelivered quantity" Then
                        .Tier1 = "Materials"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4
                    Else

                        .Tier1 = "Supply Losses"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4

                    End If
                ElseIf .DTGroup.Equals("Supply-Making") Or .DTGroup.Equals("Quality-Paste") Then
                    .Tier1 = "Paste"
                    If .DTGroup.Equals("Quality-Paste") Then
                        .Tier2 = "Quality" '"Paste Quality"
                    ElseIf .DTGroup.Equals("Supply-Making") Then
                        .Tier2 = "Availability" '"Paste Availability"
                    End If
                ElseIf .DTGroup.Equals("Systems-Utilities") Then
                    .Tier1 = "Utilities"
                ElseIf .DTGroup.Equals(BLANK_INDICATOR) And .Location.Contains("1") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 1"
                    If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                        .Tier3 = "Tube Unloading"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                        .Tier3 = "Tube Transport"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                        .Tier3 = "Tube Filling"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                        .Tier3 = "Tube Sealing"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                        .Tier3 = "Tube Trimming"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                        .Tier3 = "Tube Discharge"
                    ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                        .Tier3 = "F_Electrical"
                    End If
                ElseIf .DTGroup.Equals(BLANK_INDICATOR) And .Location.Contains("2") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 2"
                    If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                        .Tier3 = "Tube Unloading"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                        .Tier3 = "Tube Transport"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                        .Tier3 = "Tube Filling"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                        .Tier3 = "Tube Sealing"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                        .Tier3 = "Tube Trimming"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                        .Tier3 = "Tube Discharge"
                    ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                        .Tier3 = "F_Electrical"
                    End If
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            ElseIf searchEvent.isPlanned Then
                If .DTGroup.Contains("C/O") Then
                    .Tier1 = "CO"
                    If .DTGroup.Equals("C/O-Size") Then
                        .Tier2 = "Size"
                    ElseIf .DTGroup.Equals("C/O-WO+Size") Then
                        .Tier2 = "WO+Size"
                    ElseIf .DTGroup.Equals("C/O-WO") Then
                        .Tier2 = "WO"
                    ElseIf .DTGroup.Equals("C/O-PO") Then
                        .Tier2 = "PO Change"
                    ElseIf .DTGroup.Equals("Pigging") Then
                        .Tier2 = "Pigging"
                    ElseIf .DTGroup.Equals("C/O-WO-Sanitization") Then
                        .Tier2 = "Sanitization"
                    ElseIf .DTGroup.Equals("C/O-Platform") Then
                        .Tier2 = "Platform"
                    End If
                ElseIf .DTGroup.Equals("CIL") Then
                    .Tier1 = "CIL"
                ElseIf .DTGroup.Equals("Maintenance") Then
                    .Tier1 = "Maintenance"
                ElseIf .DTGroup.Equals("Training/Meeting") Then
                    .Tier1 = "Train/Meet"
                ElseIf .DTGroup.Equals("ProjectWork") Then
                    .Tier1 = "Projects"
                ElseIf .DTGroup.Equals("Material resupply") Then
                    .Tier1 = "Materials"
                ElseIf .DTGroup.Contains("SU/SD") Then
                    .Tier1 = "SU/SD"
                Else
                    .Tier1 = OTHERS_STRING
                End If


            End If
        End With
        'check if its blank
        If Len(searchEvent.Tier1) = 0 Then
            searchEvent.Tier1 = OTHERS_STRING
        End If
        If searchEvent.DTGroup = "" Then searchEvent.DTGroup = BLANK_INDICATOR

        'ok not blank

    End Sub

    Public Sub getSkinCareprstoryMapping(ByRef searchEvent As DowntimeEvent) ', ByRef Tier1 As String, ByRef Tier2 As String, ByRef Tier3 As String)
        With searchEvent
            .Product = .ProductGroup
            .ProductCode = .ProductGroup
            'check if its blank
            If Len(searchEvent.DTGroup) < 2 Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
            ElseIf .DTGroup.Equals(BLANK_INDICATOR) Then
                If .Reason1 = "Utilities" Or .Reason1 = "Cartoner" Or .Reason1 = "Capper" Or .Reason1 = "Bundler" Or .Reason1 = "Labeler" Or .Reason1 = "End Of Line" Or .Reason1.Contains("case") Then
                    .Tier1 = .Reason1
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                    '  If .Reason1 = "Filler" Then .Tier1 = "Filler"
                    If .Reason1.Contains("case") Then .Tier1 = "End Of Line"

                ElseIf .Reason1.Contains("Filler") Then
                    .Tier1 = "Filler"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                End If

            ElseIf searchEvent.isUnplanned Then
                If .DTGroup.Contains("Equip") Then
                    ' .Tier1 = "Equipment"
                    If .DTGroup.Contains("Equip-Filler") Then
                        ' If .Reason1.Contains("A") Then
                        .Tier1 = "Filler"
                        '  ElseIf .Reason1.Contains("B") Then
                        '      .Tier1 = "Filler B"
                        '   End If

                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .DTGroup.Contains("Capper") Then
                        .Tier1 = "Capper"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .DTGroup.Contains("Equip-Cartoner") Then
                        .Tier1 = "Cartoner"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .DTGroup.Contains("Equip-Bundler") Then
                        .Tier1 = "Bundler"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .DTGroup.Contains("Equip-Labeler") Then
                        .Tier1 = "Labeler"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .DTGroup.Contains("Equip-Case") Or .DTGroup.Contains("Equip-Scanner") Then
                        .Tier1 = "End Of Line"
                        .Tier2 = .Reason1
                        .Tier3 = .Reason2


                    End If
                ElseIf .DTGroup.Contains("Quality") Or .DTGroup.Equals("Supply-PackMaterial") Then
                    .Tier1 = "Materials"
                    If .DTGroup.Equals("Supply-PackMaterial") Then
                        .Tier2 = "Availability" '"Pack Material Available"
                        If .DTGroup.Contains("Bottle") Then
                            .Tier3 = "Bottle/Jar"
                        ElseIf .DTGroup.Contains("Carton") Then
                            .Tier3 = "Cartons"
                        ElseIf .DTGroup.Contains("Case") Then
                            .Tier3 = "Caseblanks"
                        ElseIf .DTGroup.Contains("Caps") Then
                            .Tier3 = "Caps"
                        ElseIf .DTGroup.Contains("Plug") Then
                            .Tier3 = "Plug / Coverdisk"
                        ElseIf .DTGroup.Contains("Lotion") Then
                            .Tier3 = "Lotion / Cream"
                        ElseIf .DTGroup.Contains("Pallet") Then
                            .Tier3 = "Pallets"
                        ElseIf .DTGroup.Contains("Labels") Then
                            .Tier3 = "Labels"
                        ElseIf .DTGroup.Contains("Film") Then
                            .Tier3 = "Film"
                        Else
                            .Tier3 = "Other"
                        End If
                    ElseIf .DTGroup.Contains("Quality") Then
                        .Tier2 = "Quality" '"Pack Material Quality"
                        If .DTGroup.Contains("Bottle") Then
                            .Tier3 = "Bottle/Jar"
                        ElseIf .DTGroup.Contains("Carton") Then
                            .Tier3 = "Cartons"
                        ElseIf .DTGroup.Contains("Case") Then
                            .Tier3 = "Caseblanks"
                        ElseIf .DTGroup.Contains("Caps") Then
                            .Tier3 = "Caps"
                        ElseIf .DTGroup.Contains("Plug") Then
                            .Tier3 = "Plug / Coverdisk"
                        ElseIf .DTGroup.Contains("Lotion") Then
                            .Tier3 = "Lotion / Cream"
                        ElseIf .DTGroup.Contains("Pallet") Then
                            .Tier3 = "Pallets"
                        ElseIf .DTGroup.Contains("Labels") Then
                            .Tier3 = "Labels"
                        ElseIf .DTGroup.Contains("Film") Then
                            .Tier3 = "Film"

                        Else
                            .Tier3 = "Other"
                        End If
                    End If
                    'ElseIf .DTGroup.Equals("Supply-Making") Or .DTGroup.Equals("Quality-Paste") Then
                    '    .Tier1 = "Paste"
                    '    If .DTGroup.Equals("Quality-Paste") Then
                    ' .Tier2 = "Quality" '"Paste Quality"
                    ' .Tier3 = .Product
                    'ElseIf .DTGroup.Equals("Supply-Making") Then
                    '    .Tier2 = "Availability" '"Paste Availability"
                    '    .Tier3 = .Product
                    'End If
                ElseIf .DTGroup.Equals("Systems-Utilities") Then
                    .Tier1 = "Utilities"
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            ElseIf searchEvent.isPlanned Then
                If .DTGroup.Contains("C/O") Then
                    .Tier1 = "CO"
                    If .Reason4.Contains("art") Or .Reason4.Contains("Art") Then
                        .Tier2 = "Artwork"
                    Else
                        .Tier2 = .Reason4
                    End If
                ElseIf .DTGroup.Equals("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team
                ElseIf .DTGroup.Equals("Maintenance") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Team
                ElseIf .DTGroup.Equals("Training/Meeting") Then
                    .Tier1 = "Train/Meet"
                    .Tier2 = .Team
                ElseIf .DTGroup.Equals("ProjectWork") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Team
                ElseIf .DTGroup.Equals("Material resupply") Then
                    .Tier1 = "Materials"
                ElseIf .DTGroup.Contains("SU/SD") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Team
                Else
                    .Tier1 = OTHERS_STRING
                End If
            End If


            If .isPlanned And .Tier1 = "" Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
            ElseIf .isUnplanned And .Tier1 = "" Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
                .Tier3 = .Reason2
            End If
        End With
    End Sub

    Public Sub getAPDOprstoryMapping_I(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                If .Reason1.Contains("Filler") Then
                    .Tier1 = .Reason1
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason3.Contains("Twins") Then
                    .Tier1 = "Twins"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                ElseIf .Reason3.Contains("Labeler") Then
                    .Tier1 = "Labeler"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                ElseIf .Reason3.Contains("Trimmer") Then
                    .Tier1 = "Trimmer"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                ElseIf .Reason3.Contains("Casepacker") Then
                    .Tier1 = "Casepacker"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                ElseIf .Reason3.Contains("Twins") Then
                    .Tier1 = "Twins"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                ElseIf .Reason3.Contains("Sorter") Or .Reason1.Contains("Sorter") Then
                    .Tier1 = "F. Sorter"
                    .Tier2 = .Reason4
                    .Tier3 = .Fault
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                End If
            Else 'If .isPlanned Then
                If .Reason2.Contains("Change") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("Flush") Then
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason3
                End If
            End If
        End With
    End Sub
    Public Sub getAPDOprstoryMapping_J(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                If .Reason1.Contains("Bundler") Then

                ElseIf .Reason1.Contains("Sticker") Then
                    .Tier1 = "Cap Sticker"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Capper") Then
                    .Tier1 = "Capper"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Case Packer") Then
                    .Tier1 = "Case Packer"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Filler") Then
                    .Tier1 = "Filler"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Labeler") Then
                    .Tier1 = "Labeler"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Bulk Material Supply") Then
                    .Tier1 = "Bulk Supply"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Case Checkweigher") Then
                    .Tier1 = "Case CW"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Case Code Dater") Then
                    .Tier1 = "Case Code Dater"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Material Quality") Then
                    .Tier1 = "Material Quality"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3

                ElseIf .Reason1.Contains("Product Supply") Then
                    .Tier1 = "Product Supply"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3

                ElseIf .Reason1.Contains("Supply Losses") Then
                    .Tier1 = "Supply Losses"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3

                ElseIf .Reason1.Contains("EOL") Then
                    .Tier1 = "EOL"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3




                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                End If
            Else 'If .isPlanned Then
                If .Reason2.Contains("Change") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("Flush") Then
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason3
                End If
            End If
        End With
    End Sub

#End Region

#Region "OralCare_Other"
    Public Sub getOralCareNauprstoryMapping(ByRef searchEvent As DowntimeEvent) ', ByRef Tier1 As String, ByRef Tier2 As String, ByRef Tier3 As String)
        With searchEvent
            'check if its blank
            If Len(searchEvent.DTGroup) < 2 Then
                .Tier1 = OTHERS_STRING

                Exit Sub
            End If

            'ok not blank

            If searchEvent.isUnplanned Then
                If .DTGroup.Contains("Equip") Then
                    .Tier1 = "Equipment"
                    If .DTGroup.Contains("Equip-Filler") Then
                        .Tier2 = "Filler"
                        If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                            .Tier3 = "Tube Unloading"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                            .Tier3 = "Tube Transport"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                            .Tier3 = "Tube Filling"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                            .Tier3 = "Tube Sealing"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                            .Tier3 = "Tube Trimming"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                            .Tier3 = "Tube Discharge"
                        ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                            .Tier3 = "F_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Cartoner") Then
                        .Tier2 = "Cartoner"
                        If .DTGroup.Equals("Equip-Cartoner-CartonFeeding") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonPicking") Then
                            .Tier3 = "Carton Picking"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeInsertion") Then
                            .Tier3 = "Tube Insertion"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeTransfer") Then
                            .Tier3 = "Tube Transfer"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonClosing") Then
                            .Tier3 = "Carton Closing"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonCoding") Then
                            .Tier3 = "Carton Coding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-VisionSystem") Then
                            .Tier3 = "Vision System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Discharge") Then
                            .Tier3 = "Discharge"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Electrical") Then
                            .Tier3 = "CT_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bundler") Then
                        .Tier2 = "Bundler"
                        If .DTGroup.Equals("Equip-Bundler-CartonInfeed") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Bundler-StackingArea") Then
                            .Tier3 = "Stacking"
                        ElseIf .DTGroup.Equals("Equip-Bundler-FilmFeeding") Then
                            .Tier3 = "Film Transport"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Cut&Wrap") Then
                            .Tier3 = "Wrap Film Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Fold&Seal") Then
                            .Tier3 = "Film Fold End Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Discharge") Then
                            .Tier3 = "Bundle Discharge"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Electrical") Then
                            .Tier3 = "B_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-CP") Then
                        .Tier2 = "Casepacker"
                        If .DTGroup.Equals("Equip-CP-CaseEntrance") Then
                            .Tier3 = "Case Entrance"
                        ElseIf .DTGroup.Equals("Equip-CP-CartonInfeed") Then
                            .Tier3 = "Carton Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-Infeed") Then
                            .Tier3 = "Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-CP-StackingArea") Then
                            .Tier3 = "Stacking Area"
                        ElseIf .DTGroup.Equals("Equip-CP-PushIntoCase") Then
                            .Tier3 = "PushInCase"
                        ElseIf .DTGroup.Equals("Equip-CP-CaseTransport") Then
                            .Tier3 = "Case Transport"
                        ElseIf .DTGroup.Equals("Equip-CP-Electrical") Then
                            .Tier3 = "CP_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bulkcases") Then
                        .Tier2 = "Bulk packer"
                        If .DTGroup.Equals("Equip-BulkCaseErector") Then
                            .Tier3 = "Case Former"
                        ElseIf .DTGroup.Equals("Equip-BulkCasePacker") Then
                            .Tier3 = "Tray Packer"
                        ElseIf .DTGroup.Equals("Equip-BulkCaseSealer") Then
                            .Tier3 = "Case Sealer"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Case") Or .DTGroup.Contains("Equip-Scan") Then
                        .Tier2 = "End Of line"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-Palletizer") Then
                        .Tier2 = "Palletizer"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-BulkCasePacker") Then
                        .Tier2 = "Bulk CP"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Pump") Then
                        .Tier2 = "PS-Pump"
                        .Tier3 = .Reason3
                    ElseIf .DTGroup.Contains("Tubeunloading") Then
                        .Tier2 = "TubeUnloader"
                        .Tier3 = .Reason3
                    Else
                        .Tier2 = OTHERS_STRING
                        .Tier3 = .Reason3
                    End If
                ElseIf .DTGroup.Equals("Quality-Tubes") Or .DTGroup.Equals("Quality-Cartons") Or .DTGroup.Equals("Quality-Film") Or .DTGroup.Equals("Quality-Caseblanks") Then
                    .Tier1 = "Materials"

                    If .DTGroup.Contains("Tube") Then
                        .Tier2 = "Tube Quality"
                    ElseIf .DTGroup.Contains("Carton") Then
                        .Tier2 = "Carton Quality"
                    ElseIf .DTGroup.Contains("Case") Then
                        .Tier2 = "Caseblank Quality"
                    Else
                        .Tier2 = "Other Quality"
                    End If
                    .Tier3 = .Reason4

                ElseIf .DTGroup.Equals("Supply-PackMaterial") Then
                    If .Reason4 = "Material not in warehouse" Or .Reason4 = "Underdelivered quantity" Then
                        .Tier1 = "Materials"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4
                    Else

                        .Tier1 = "Supply Losses"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4

                    End If
                ElseIf .DTGroup.Equals("Supply-Making") Or .DTGroup.Equals("Quality-Paste") Then
                    .Tier1 = "Paste"
                    If .DTGroup.Equals("Quality-Paste") Then
                        .Tier2 = "Quality" '"Paste Quality"
                        .Tier3 = .Product
                    ElseIf .DTGroup.Equals("Supply-Making") Then
                        .Tier2 = "Availability" '"Paste Availability"
                        .Tier3 = .Product
                    Else
                        .Tier2 = .Reason1
                        .Tier3 = .Reason2
                    End If
                ElseIf .DTGroup.Equals("Systems-Utilities") Then
                    .Tier1 = "Utilities"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .DTGroup
                    .Tier3 = .Reason1
                End If
            ElseIf searchEvent.isPlanned Then
                If .DTGroup.Contains("C/O") Then
                    .Tier1 = "CO"
                    .Tier2 = .Team
                    If .DTGroup.Equals("C/O-Size") Then
                        .Tier2 = "Size"
                    ElseIf .DTGroup.Equals("C/O-WO+Size") Then
                        .Tier2 = "WO+Size"
                    ElseIf .DTGroup.Equals("C/O-WO") Then
                        .Tier2 = "WO"
                    ElseIf .DTGroup.Equals("C/O-PO") Then
                        .Tier2 = "PO Change"
                    ElseIf .DTGroup.Equals("C/O-Pigging") Then
                        .Tier2 = "Pigging"
                    ElseIf .DTGroup.Equals("C/O-WO-Sanitization") Then
                        .Tier2 = "Sanitization"
                    ElseIf .DTGroup.Equals("C/O-Platform") Then
                        .Tier2 = "Platform"
                    End If
                ElseIf .DTGroup.Equals("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team '.Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Maintenance") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Training/Meeting") Then
                    .Tier1 = "Train/Meet"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("ProjectWork") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Material resupply") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Contains("SU/SD") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Team '.Reason1 & "-" & .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                End If
                .Tier3 = .Team

            End If
        End With
    End Sub

    Public Sub getOralCareNauprstoryMappingDF(ByRef searchEvent As DowntimeEvent) ', ByRef Tier1 As String, ByRef Tier2 As String, ByRef Tier3 As String)
        With searchEvent

            If searchEvent.isUnplanned Then
                If .DTGroup.Contains("Equip") Then
                    .Tier1 = "Equipment"
                    If (.DTGroup.Contains("Equip-Filler") Or .DTGroup.Contains("TubeUnloading")) And .Reason1.Contains("1") Then
                        .Tier2 = "Filler 1"
                        If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("TubeUnloading") Then
                            .Tier3 = "Tube Unloading"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                            .Tier3 = "Tube Transport"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                            .Tier3 = "Tube Filling"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                            .Tier3 = "Tube Sealing"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                            .Tier3 = "Tube Trimming"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                            .Tier3 = "Tube Discharge"
                        ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                            .Tier3 = "F_Electrical"
                        End If
                    ElseIf (.DTGroup.Contains("Equip-Filler") Or .DTGroup.Contains("TubeUnloading")) And .Reason1.Contains("2") Then
                        .Tier2 = "Filler 2"
                        If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                            .Tier3 = "Tube Unloading"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                            .Tier3 = "Tube Transport"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                            .Tier3 = "Tube Filling"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                            .Tier3 = "Tube Sealing"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                            .Tier3 = "Tube Trimming"
                        ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                            .Tier3 = "Tube Discharge"
                        ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                            .Tier3 = "F_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Cartoner") Then
                        .Tier2 = "Cartoner"
                        If .DTGroup.Equals("Equip-Cartoner-CartonFeeding") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonPicking") Then
                            .Tier3 = "Carton Picking"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeInsertion") Then
                            .Tier3 = "Tube Insertion"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-TubeTransfer") Then
                            .Tier3 = "Tube Transfer"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonClosing") Then
                            .Tier3 = "Carton Closing"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-CartonCoding") Then
                            .Tier3 = "Carton Coding"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-VisionSystem") Then
                            .Tier3 = "Vision System"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Discharge") Then
                            .Tier3 = "Discharge"
                        ElseIf .DTGroup.Equals("Equip-Cartoner-Electrical") Then
                            .Tier3 = "CT_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bundler") Then
                        .Tier2 = "Bundler"
                        If .DTGroup.Equals("Equip-Bundler-CartonInfeed") Then
                            .Tier3 = "Carton Feeding"
                        ElseIf .DTGroup.Equals("Equip-Bundler-StackingArea") Then
                            .Tier3 = "Stacking"
                        ElseIf .DTGroup.Equals("Equip-Bundler-FilmFeeding") Then
                            .Tier3 = "Film Transport"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Cut&Wrap") Then
                            .Tier3 = "Wrap Film Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Fold&Seal") Then
                            .Tier3 = "Film Fold End Sealing"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Discharge") Then
                            .Tier3 = "Bundle Discharge"
                        ElseIf .DTGroup.Equals("Equip-Bundler-Electrical") Then
                            .Tier3 = "B_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-CP") Then
                        .Tier2 = "Casepacker"
                        If .DTGroup.Equals("Equip-CP-CaseEntrance") Then
                            .Tier3 = "Case Entrance"
                        ElseIf .DTGroup.Equals("Equip-CP-CartonInfeed") Then
                            .Tier3 = "Carton Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-Infeed") Then
                            .Tier3 = "Infeed"
                        ElseIf .DTGroup.Equals("Equip-CP-GlueSystem") Then
                            .Tier3 = "Glue System"
                        ElseIf .DTGroup.Equals("Equip-CP-StackingArea") Then
                            .Tier3 = "Stacking Area"
                        ElseIf .DTGroup.Equals("Equip-CP-PushIntoCase") Then
                            .Tier3 = "PushInCase"
                        ElseIf .DTGroup.Equals("Equip-CP-CaseTransport") Then
                            .Tier3 = "Case Transport"
                        ElseIf .DTGroup.Equals("Equip-CP-Electrical") Then
                            .Tier3 = "CP_Electrical"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Bulkcases") Then
                        .Tier2 = "Bulk packer"
                        If .DTGroup.Equals("Equip-BulkCaseErector") Then
                            .Tier3 = "Case Former"
                        ElseIf .DTGroup.Equals("Equip-BulkCasePacker") Then
                            .Tier3 = "Tray Packer"
                        ElseIf .DTGroup.Equals("Equip-BulkCaseSealer") Then
                            .Tier3 = "Case Sealer"
                        End If
                    ElseIf .DTGroup.Contains("Equip-Case") Or .DTGroup.Contains("Equip-Scan") Then
                        .Tier2 = "End Of line"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-Palletizer") Then
                        .Tier2 = "Palletizer"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Equip-BulkCasePacker") Then
                        .Tier2 = "Bulk CP"
                        .Tier3 = .Tier2
                    ElseIf .DTGroup.Contains("Pump") Then
                        .Tier2 = "PS-Pump"
                        .Tier3 = .Reason3
                    Else
                        .Tier2 = OTHERS_STRING
                        .Tier3 = .Reason3
                    End If
                ElseIf .DTGroup.Equals("Quality-Tubes") Or .DTGroup.Equals("Quality-Cartons") Or .DTGroup.Equals("Quality-Film") Or .DTGroup.Equals("Quality-Caseblanks") Then
                    .Tier1 = "Materials"

                    If .DTGroup.Contains("Tube") Then
                        .Tier2 = "Tube Quality"
                    ElseIf .DTGroup.Contains("Carton") Then
                        .Tier2 = "Carton Quality"
                    ElseIf .DTGroup.Contains("Case") Then
                        .Tier2 = "Caseblank Quality"
                    Else
                        .Tier2 = "Other Quality"
                    End If
                    .Tier3 = .Reason4

                ElseIf .DTGroup.Equals("Supply-PackMaterial") Then
                    If .Reason4 = "Material not in warehouse" Or .Reason4 = "Underdelivered quantity" Then
                        .Tier1 = "Materials"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4
                    Else

                        .Tier1 = "Supply Losses"
                        .Tier2 = .Reason3
                        .Tier3 = .Reason4

                    End If
                ElseIf .DTGroup.Equals("Supply-Making") Or .DTGroup.Equals("Quality-Paste") Then
                    .Tier1 = "Paste"
                    If .DTGroup.Equals("Quality-Paste") Then
                        .Tier2 = "Quality" '"Paste Quality"
                    ElseIf .DTGroup.Equals("Supply-Making") Then
                        .Tier2 = "Availability" '"Paste Availability"
                    End If
                ElseIf .DTGroup.Equals("Systems-Utilities") Then
                    .Tier1 = "Utilities"
                ElseIf .DTGroup.Equals(BLANK_INDICATOR) And .Location.Contains("NAUP OASIS Perdida de velocidad") And .Reason1.Contains("Llenadora 1") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 1"
                    If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                        .Tier3 = "Tube Unloading"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                        .Tier3 = "Tube Transport"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                        .Tier3 = "Tube Filling"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                        .Tier3 = "Tube Sealing"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                        .Tier3 = "Tube Trimming"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                        .Tier3 = "Tube Discharge"
                    ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                        .Tier3 = "F_Electrical"
                    Else
                        .Tier3 = .Fault
                    End If
                ElseIf .DTGroup.Equals(BLANK_INDICATOR) And .Location.Contains("NAUP OASIS Perdida de velocidad") And .Reason1.Contains("Llenadora 2") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler 2"
                    If .DTGroup.Equals("Equip-Filler-TubeUnloading") Or .Equals("-Filler-TubeUnloading") Then
                        .Tier3 = "Tube Unloading"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTransport") Then
                        .Tier3 = "Tube Transport"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeFilling") Then
                        .Tier3 = "Tube Filling"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeSealing") Then
                        .Tier3 = "Tube Sealing"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeTrimming") Then
                        .Tier3 = "Tube Trimming"
                    ElseIf .DTGroup.Equals("Equip-Filler-TubeDischarge") Then
                        .Tier3 = "Tube Discharge"
                    ElseIf .DTGroup.Equals("Equip-Filler-Electrical") Then
                        .Tier3 = "F_Electrical"
                    Else
                        .Tier3 = .Fault
                    End If
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            ElseIf searchEvent.isPlanned Then
                If .DTGroup.Contains("C/O") Then
                    .Tier1 = "CO"
                    If .DTGroup.Equals("C/O-Size") Then
                        .Tier2 = "Size"
                    ElseIf .DTGroup.Equals("C/O-WO+Size") Then
                        .Tier2 = "WO+Size"
                    ElseIf .DTGroup.Equals("C/O-WO") Then
                        .Tier2 = "WO"
                    ElseIf .DTGroup.Equals("C/O-PO") Then
                        .Tier2 = "PO Change"
                    ElseIf .DTGroup.Equals("Pigging") Then
                        .Tier2 = "Pigging"
                    ElseIf .DTGroup.Equals("C/O-WO-Sanitization") Then
                        .Tier2 = "Sanitization"
                    ElseIf .DTGroup.Equals("C/O-Platform") Then
                        .Tier2 = "Platform"
                    End If
                ElseIf .DTGroup.Equals("CIL") Then
                    .Tier1 = "CIL"
                ElseIf .DTGroup.Equals("Maintenance") Then
                    .Tier1 = "Maintenance"
                ElseIf .DTGroup.Equals("Training/Meeting") Then
                    .Tier1 = "Train/Meet"
                ElseIf .DTGroup.Equals("ProjectWork") Then
                    .Tier1 = "Projects"
                ElseIf .DTGroup.Equals("Material resupply") Then
                    .Tier1 = "Materials"
                ElseIf .DTGroup.Contains("SU/SD") Then
                    .Tier1 = "SU/SD"
                Else
                    .Tier1 = OTHERS_STRING
                End If


            End If
        End With
        'check if its blank
        If Len(searchEvent.Tier1) = 0 Then
            searchEvent.Tier1 = OTHERS_STRING
        End If
        If searchEvent.DTGroup = "" Then searchEvent.DTGroup = BLANK_INDICATOR
    End Sub

    Public Sub getOralCareGrossprstoryMapping(ByRef searchEvent As DowntimeEvent) ', ByRef Tier1 As String, ByRef Tier2 As String, ByRef Tier3 As String)
        With searchEvent
            'check if its blank
            If Len(searchEvent.DTGroup) < 2 Then
                .Tier1 = OTHERS_STRING
                Exit Sub
            End If

            If searchEvent.isUnplanned Then
                If .DTGroup.Contains("Equip") Then
                    .Tier1 = "Equipment"
                    .Tier2 = Strings.Mid(.DTGroup, InStr(.DTGroup, "-") + 1)
                    If InStr(.Tier2, "-") > 0 Then
                        .Tier3 = Strings.Mid(.Tier2, InStr(.Tier2, "-") + 1)
                        .Tier2 = Strings.Left(.Tier2, InStr(.Tier2, "-") - 1)
                    Else
                        .Tier3 = .Reason3
                    End If
                ElseIf .DTGroup.Contains("Quality") And Not .DTGroup.Contains("Making") Then
                    .Tier1 = "Materials"
                    .Tier3 = Strings.Mid(.DTGroup, InStr(.DTGroup, "-") + 1)
                    .Tier2 = "Quality"



                ElseIf .DTGroup.Contains("Supply-PackMaterial") Then
                    .Tier1 = "Supply Losses"
                    .Tier2 = .Reason3
                    .Tier3 = .Reason4
                ElseIf .DTGroup.Contains("Making") Then
                    .Tier1 = "Paste"
                    If .DTGroup.Contains("Supply") Then
                        .Tier2 = "Availability"
                        .Tier3 = .Reason3
                    Else
                        .Tier2 = "Quality"
                        .Tier3 = .Reason3
                    End If
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason3
                End If
            ElseIf searchEvent.isPlanned Then
                If .DTGroup.Contains("C/O") Then
                    .Tier1 = "CO"
                    .Tier2 = .Team
                    If .DTGroup.Equals("C/O-Size") Then
                        .Tier2 = "Size"
                    ElseIf .DTGroup.Equals("C/O-WO+Size") Then
                        .Tier2 = "WO+Size"
                    ElseIf .DTGroup.Equals("C/O-WO") Then
                        .Tier2 = "WO"
                    ElseIf .DTGroup.Equals("C/O-PO") Then
                        .Tier2 = "PO Change"
                    ElseIf .DTGroup.Equals("C/O-Pigging") Then
                        .Tier2 = "Pigging"
                    ElseIf .DTGroup.Equals("C/O-WO-Sanitization") Then
                        .Tier2 = "Sanitization"
                    ElseIf .DTGroup.Equals("C/O-Platform") Then
                        .Tier2 = "Platform"
                    End If
                ElseIf .DTGroup.Equals("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team '.Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Maintenance") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Training/Meeting") Then
                    .Tier1 = "Train/Meet"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("ProjectWork") Then
                    .Tier1 = "Projects"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Equals("Material resupply") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason1 & "-" & .Reason2
                ElseIf .DTGroup.Contains("SU/SD") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Team '.Reason1 & "-" & .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason2
                End If
                .Tier3 = .Team

            End If
        End With
    End Sub

    Public Sub getFem_LuisCustomMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            .Tier1 = OTHERS_STRING
            .Tier2 = .Reason1
            .Tier3 = .Reason2
            If .isUnplanned Then
                If .Location.Contains("Area 1") Then
                    .Tier1 = "Area 1"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("Area 2") Then
                    .Tier1 = "Area 2"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("Area 3") Then
                    .Tier1 = "Area 3"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("Area 4") Then
                    .Tier1 = "Area 4"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("EA1") Then
                    .Tier1 = "Area 1"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("EA2") Then
                    .Tier1 = "Area 2"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("EA3") Then
                    .Tier1 = "Area 3"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("EA4") Then
                    .Tier1 = "Area 4"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("MATERIA PRIMA") Then
                        .Tier1 = "Matéria Prima"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("RAZÃO EXTERNA") Then
                        .Tier1 = "Razão Externa"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("OUTROS") Then
                        .Tier1 = "Outros"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("Utilidades") Then
                        .Tier1 = "Razão Externa"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("OSPREY") Then
                        .Tier1 = "Osprey"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("EXTERNAL REASONS") Then
                        If .Reason2.Contains("SPECIAL CAUSES") Then
                            .Tier1 = "Outros"
                            .Tier2 = .Reason3
                            .Tier3 = .Team
                        End If
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("CONTROLLERS") Then
                        .Tier1 = "Controladores"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("MAIN DRIVE AND MOTOR") Then
                        .Tier1 = "Main Drive"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("Programmed") Then
                        If .Reason2.Contains("RAW MATERIAL FAILURE") Then
                            .Tier1 = "Matéria Prima"
                            .Tier2 = .Reason3
                            .Tier3 = .Team
                        End If
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("Programmed") Then
                        If .Reason2.Contains("SPECIAL CAUSES") Then
                            .Tier1 = "Outros"
                            .Tier2 = .Reason3
                            .Tier3 = .Team
                        End If
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("RTCIS") Then
                        .Tier1 = "RTCIS"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("SERVICES AND FACILITIES") Then
                        .Tier1 = "Razão Externa"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("STARTUP") Then
                        .Tier1 = "Start-up"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("CONTROL-NET") Then
                        .Tier1 = "Controladores"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("OFFLINE") Then
                    If .Reason1.Contains("CENTRAL FACILITIES") Then
                        .Tier1 = "Razão Externa"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("OFFLINE") Then
                    If .Reason2.Contains("FAILURE - DATA NOT ENTERED") Then
                        .Tier1 = "Outros"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("OffLine") Then
                    .Tier1 = "Razão Externa"
                    .Tier2 = .Reason1
                    .Tier3 = .Team
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("CHILLER") Then
                        .Tier1 = "Chiller"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("EVENTS") Then
                        .Tier1 = "Outros"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("MATERIAL") Then
                        .Tier1 = "Matéria Prima"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If
            Else 'if .isplanned
                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("CHANGEOVER PARTS HANGING") Then
                        .Tier1 = "Change Over"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("EXTERNAL REASONS") Then
                        If .Reason2.Contains("Programmed") Then
                            .Tier1 = "Outros"
                            .Tier2 = .Reason3
                            .Tier3 = .Team
                        End If
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("IWS") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("MPSa") Then
                        .Tier1 = "MPSa"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("PREVENTIVE MAINTENACE (PM)") Then
                        .Tier1 = "Manutenção Programada"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("Programmed") Then
                        If .Reason2.Contains("Programmed") Then
                            .Tier1 = "Outros"
                            .Tier2 = .Reason3
                            .Tier3 = .Team
                        End If
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("SAFETY") Then
                        .Tier1 = "Segurança"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("CHANGE OVER") Then
                        .Tier1 = "Change Over"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("DOWNTIME PLANEJADO") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("EO") Then
                        .Tier1 = "EO"
                        .Tier2 = .Reason4
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("PARADA INTENCIONAL") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("SEGURANÇA") Then
                        .Tier1 = "Segurança"
                        .Tier2 = .Reason4
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("SHUTDOWN") Then
                        .Tier1 = "Shutdown"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("CIL") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("EO Não Vendável") Then
                        .Tier1 = "EO"
                        .Tier2 = .Reason4
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("EO Vendável") Then
                        .Tier1 = "EO"
                        .Tier2 = .Reason4
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("TREINAMENTO") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("Startup Shutdown") Then
                        .Tier1 = "Shutdown"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("REUNIÃO") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("Projeto") Then
                        .Tier1 = "EO"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("PtD") Then
                        .Tier1 = "MPSa"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("Manutenção Planejada") Then
                        .Tier1 = "Manutenção Programada"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("INTENTIONAL STOP") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("AUTONOMOUS MAINTENANCE") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("4HR SURVIVAL") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason3
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("OFFLINE") Then
                    If .Reason2.Contains("SHUT DOWN") Then
                        .Tier1 = "Shutdown"
                        .Tier2 = .Team
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("CHANGEOVER") Then
                        .Tier1 = "Change Over"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("COMISSIONING") Then
                        .Tier1 = "EO"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("E/O") Then
                        .Tier1 = "EO"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("PROGRESSIVE MAINTENANCE") Then
                        .Tier1 = "Manutenção Programada"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("PROJECTS") Then
                        .Tier1 = "EO"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If

                If .Location.Contains("No Area") Then
                    If .Reason1.Contains("SURVIVAL") Then
                        .Tier1 = "Parada Programada"
                        .Tier2 = .Reason2
                        .Tier3 = .Team
                    End If
                End If



            End If

        End With
    End Sub

    Public Sub getIowaCityprstoryMapping(ByRef searchEvent As DowntimeEvent)

        With searchEvent
            .Tier1 = .Reason1
            .Tier2 = .Reason2
            .Tier3 = .Reason3
            .DTGroup = .Reason1 + "-" + .Reason2

            If .Reason1.Equals("Scheduled Downtime") Then
                If .Reason2.Equals("Format Change") Then
                    .Tier1 = "Format Change"
                    .Tier2 = .Reason3
                    .Tier3 = .Reason4
                Else
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                    .Tier3 = .Reason4
                End If
            End If



            If .Tier3 = "Comment" Then
                .Tier3 = ""
            End If

            If .isPlanned Then
                If .Tier1 = "SAP" Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = "SAP"
                ElseIf .Tier1 = "Cleaning" Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = "Cleaning"
                ElseIf .Tier1 = "OQ" Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = "OQ"
                ElseIf .Tier1 = "Sanitization" Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = "Sanitization"
                ElseIf .Tier1 = "Tank Change" Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = "Tank Change"
                ElseIf .Tier1 = "SU/SD" Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = "SU/SD"
                End If
            End If

            If .isPlanned And .Tier1 = "" Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
            End If
            If .isUnplanned And .Tier1 = "" Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
                .Tier3 = .Reason2
            End If
        End With
    End Sub


#End Region

#Region "BabyCare"

    Public Sub getMandideepprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                If Not .Location.Contains("/") Then
                    If .Reason1.Contains("010") Then
                        .Tier1 = "Utilities"
                        .Tier2 = .Reason1 & "-" & .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Location.Contains("OFFLINE") Then
                        .Tier1 = "Offline"
                        .Tier2 = .Reason1 & "-" & .Reason2
                        .Tier3 = .Reason3
                    Else
                        .Tier1 = .Location
                        .Tier2 = .Reason1
                        .Tier3 = Strings.Left(.Reason1, 4) & .Reason2   'LG Code (change made per Aditi's request to provide context for reason 2s)
                    End If
                ElseIf .Location.Contains("Unknown") Or .Location.Contains("UNKNOWN") Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1 & "-" & .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1 & "-" & .Reason2
                    .Tier3 = .Reason3
                End If

            ElseIf .isPlanned Then
                If .Reason1.Contains("990") Then
                    .Tier1 = "CO"
                    .Tier2 = "SIZE CHANGE"
                ElseIf .Reason1.Contains("991") Then
                    .Tier1 = "CO"
                    .Tier2 = "PACK CHANGE"
                ElseIf .Reason1.Contains("992") Then
                    .Tier1 = "CO"
                    .Tier2 = "BRAND/PRODUCT CHANGE"
                ElseIf .Reason1.Contains("993") Then
                    .Tier1 = "AM/CIL/RLS" 'mm
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("994") Then
                    .Tier1 = "PM"
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("995") Then
                    .Tier1 = "Projects / Construction"
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("996") Then
                    .Tier1 = "Organization"
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("998") Then
                    .Tier1 = "Induced Stops"
                    .Tier2 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1 & "-" & .Reason2
                End If
            End If
        End With
    End Sub

#End Region


    Public Sub getMandideepFemprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                If Not .Location.Contains("/") Then
                    If .Reason1.Contains("010") Then
                        .Tier1 = "Utilities"
                        .Tier2 = .Reason1 & "-" & .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Location.Contains("OFFLINE") Then
                        .Tier1 = "Offline"
                        .Tier2 = .Reason1 & "-" & .Reason2
                        .Tier3 = .Reason3
                    Else
                        .Tier1 = .Location
                        .Tier2 = .Reason1
                        .Tier3 = Strings.Left(.Reason1, 4) & .Reason2   'LG Code (change made per Aditi's request to provide context for reason 2s)
                    End If
                ElseIf .Location.Contains("Unknown") Or .Location.Contains("UNKNOWN") Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1 & "-" & .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1 & "-" & .Reason2
                    .Tier3 = .Reason3
                End If

            ElseIf .isPlanned Then
                If .Reason1.Contains("990") Then
                    .Tier1 = "CO"
                    .Tier2 = "SIZE CHANGE"
                ElseIf .Reason1.Contains("991") Then
                    .Tier1 = "CO"
                    .Tier2 = "PACK CHANGE"
                ElseIf .Reason1.Contains("992") Then
                    .Tier1 = "CO"
                    .Tier2 = "BRAND/PRODUCT CHANGE"
                ElseIf .Reason2.Contains("09_RLS") Or .Reason2.Contains("01_CIL") Or .Reason3.Contains("RLS") Then
                    .Tier1 = "AM/CIL/RLS" 'mm
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("994") Then
                    .Tier1 = "PM"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("07_Project") Then
                    .Tier1 = "Projects / Construction"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("Meeting") Then
                    .Tier1 = "Organization"
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("998") Then
                    .Tier1 = "Induced Stops"
                    .Tier2 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1 & "-" & .Reason2
                End If
            End If
        End With
    End Sub



#Region "PHC"
    Public Sub getPhoenixprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent


            If .Reason1.Equals("Material Availability") Then
                If .Reason2.Equals("Bottles") Then
                    .Tier1 = "Material"
                    .Tier2 = "Availibility"
                    .Tier3 = "Bottles"
                End If
            End If

            If .Reason1.Equals("Material Availability") Then
                If .Reason2.Equals("Sleeves") Then
                    .Tier1 = "Material"
                    .Tier2 = "Availibility"
                    .Tier3 = "Sleeves"
                End If
            End If

            If .Reason1.Equals("Material Availability") Then
                If .Reason2.Equals("Caps") Then
                    .Tier1 = "Material"
                    .Tier2 = "Availibility"
                    .Tier3 = "Caps"
                End If
            End If

            If .Reason1.Equals("Material Availability") Then
                If .Reason2.Equals("Stickers") Then
                    .Tier1 = "Material"
                    .Tier2 = "Availibility"
                    .Tier3 = "Stickers"
                End If
            End If

            If .Reason1.Equals("Material Availability") Then
                If .Reason2.Equals("Shippers") Then
                    .Tier1 = "Material"
                    .Tier2 = "Availibility"
                    .Tier3 = "Shippers"
                End If
            End If

            If .Reason1.Equals("Material Availability") Then
                If .Reason2.Equals("Powder Available") Then
                    .Tier1 = "Powder"
                    .Tier2 = "Availibility"
                    .Tier3 = ""
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Application") Then
                    If .Reason3.Equals("Bad Sleeves") Then
                        If .Reason4.Equals("Excessive Crease") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Sleever"
                            .Tier3 = "Application"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Application") Then
                    If .Reason3.Equals("Bad Sleeves") Then
                        If .Reason4.Equals("Comment") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Sleever"
                            .Tier3 = "Application"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Application") Then
                    If .Reason3.Equals("Bad Sleeves") Then
                        If .Reason4.Equals("Brittle Sleeves") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Sleever"
                            .Tier3 = "Application"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Application") Then
                    If .Reason3.Equals("Bad Sleeves") Then
                        If .Reason4.Equals("Coming Apart At Seam") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Sleever"
                            .Tier3 = "Application"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Application") Then
                    If .Reason3.Equals("Registration Off") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Application"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Application") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Application"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Application") Then
                    If .Reason3.Equals("Jam at Shot Wheel") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Application"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Bottle discharge") Then
                    If .Reason3.Equals("Down Bottle") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Bottle discharge"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Bottle discharge") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Bottle discharge"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Bottle Infeed") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Bottle Infeed"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Bottle Infeed") Then
                    If .Reason3.Equals("Down Bottle") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Bottle Infeed"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Bottle Infeed") Then
                    If .Reason3.Equals("Feedscrew Not Synchronized") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Bottle Infeed"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Heat & Shrink") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Heat & Shrink"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Heat & Shrink") Then
                    If .Reason3.Equals("Down Bottle") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Heat & Shrink"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Heat & Shrink") Then
                    If .Reason3.Equals("No Heat") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Heat & Shrink"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Heat & Shrink") Then
                    If .Reason3.Equals("Upper Conveyor Not Synchronized") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Heat & Shrink"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Sleeve Wipe Down") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Sleever"
                        .Tier3 = "Sleeve Wipe Down"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Operator Error") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Sleever"
                    .Tier3 = "Operator Error"
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Comments") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Sleever"
                    .Tier3 = "Comments"
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Sleeves") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Sleeves Quality in Sleever"
                    End If
                End If
            End If

            If .Reason1.Equals("Sleever") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Bottles") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Bottles Quality in Sleever"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Bottle discharge") Then
                    If .Reason3.Equals("Down Bottle") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Bottle discharge"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Bottle discharge") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Bottle discharge"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Layer Removal") Then
                    If .Reason3.Equals("Layer Not Found") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Layer Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Layer Removal") Then
                    If .Reason3.Equals("Layer Not Swept") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Layer Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Layer Removal") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Layer Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Pallet Indexing") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Pallet Indexing"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Pallet Indexing") Then
                    If .Reason3.Equals("Bad Pallet") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Pallet Indexing"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Tier Sheet Removal") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Tier Sheet Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Tier Sheet Removal") Then
                    If .Reason3.Equals("Pic Frame not Picked") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Tier Sheet Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Pallet Removal") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Pallet Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Pallet Removal") Then
                    If .Reason3.Equals("Pallet Not Discharged") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Pallet Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Pallet Removal") Then
                    If .Reason3.Equals("Bad Pallet") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Depal"
                        .Tier3 = "Pallet Removal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Operator Error") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Depal"
                    .Tier3 = "Operator Error"
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Pallet") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Pallet Quality in Depal"
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Tier Sheet") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Tier Sheet Quality "
                    End If
                End If
            End If

            If .Reason1.Equals("Depal") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Bottle") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Bottle Quality in Depal"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Discharge Pallet") Then
                    If .Reason3.Equals("Cases fell off") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Discharge Pallet"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Discharge Pallet") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Discharge Pallet"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Index Pallet") Then
                    If .Reason3.Equals("Cases obstructing pallet") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Index Pallet"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Index Pallet") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Index Pallet"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Orient Cases") Then
                    If .Reason3.Equals("Cases not stable") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Orient Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Orient Cases") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Orient Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Transfer Layer") Then
                    If .Reason3.Equals("Carriage not at correct height") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Transfer Layer"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Transfer Layer") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Transfer Layer"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Operator Error") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "ULF"
                    .Tier3 = "Operator Error"
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Brake") Then
                    If .Reason3.Equals("Brake staying up") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Brake"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Brake") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Brake"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Open Flap from WACP") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "ULF"
                        .Tier3 = "Open Flap from WACP"
                    End If
                End If
            End If

            If .Reason1.Equals("ULF") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Pallet") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Pallet Quality in ULF"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Discharge Case") Then
                    If .Reason3.Equals("Case Coder Not Printing") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Discharge Case"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Discharge Case") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Discharge Case"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Discharge Case") Then
                    If .Reason3.Equals("Conveyor Rails Not Set-up") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Discharge Case"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Forming Cases") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Forming Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Forming Cases") Then
                    If .Reason3.Equals("North side flaps not tucked") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Forming Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Forming Cases") Then
                    If .Reason3.Equals("South side flaps not tucked") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Forming Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Sealing Cases") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Sealing Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Sealing Cases") Then
                    If .Reason3.Equals("Mfg flap not sealed") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Sealing Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Sealing Cases") Then
                    If .Reason3.Equals("North side not sealed") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Sealing Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Sealing Cases") Then
                    If .Reason3.Equals("South side not sealed") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Sealing Cases"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Operator Error") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "WACP"
                    .Tier3 = "Operator Error"
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Bottles") Then
                    If .Reason3.Equals("Adjust Loader Section") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Bottles"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Bottles") Then
                    If .Reason3.Equals("Case Not Placed Correctly") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Bottles"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Bottles") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Bottles"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Bottles") Then
                    If .Reason3.Equals("Down Bottle") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Bottles"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Shippers") Then
                    If .Reason3.Equals("Adjust Loader Section") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Shippers"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Shippers") Then
                    If .Reason3.Equals("Case Not Placed Correctly") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Shippers"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Shippers") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Shippers"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Shippers") Then
                    If .Reason3.Equals("No Shippers") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Shippers"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Infeed Shippers") Then
                    If .Reason3.Equals("Shippers not picked") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "WACP"
                        .Tier3 = "Infeed Shippers"
                    End If
                End If
            End If

            If .Reason1.Equals("WACP") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Shippers") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Shipper Quality"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Discharged") Then
                    If .Reason3.Equals("Bottle Jam") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Bottle Jam"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Discharged") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Comment"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Discharged") Then
                    If .Reason3.Equals("Rail Adjustment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Rail Adjustment"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Grip & Ionize") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Comment"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Grip & Ionize") Then
                    If .Reason3.Equals("Gearbox") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Gearbox"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Grip & Ionize") Then
                    If .Reason3.Equals("Grippers Worn") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Grippers Worn"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Grip & Ionize") Then
                    If .Reason3.Equals("Motor") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Motor"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Grip & Ionize") Then
                    If .Reason3.Equals("Wear Strip Worn") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Wear Strip Worn"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Infeed/Spaced") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Comment"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Infeed/Spaced") Then
                    If .Reason3.Equals("Down Bottle") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Down Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Bottle Infeed/Spaced") Then
                    If .Reason3.Equals("Starwheel not turning") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Starwheel not turning"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Speed") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Rinser"
                    .Tier3 = "Speed"
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Vision System Not Ready") Then
                    If .Reason3.Equals("Five Consecutive Rejects") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Vision System Not Ready"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Vision System Not Ready") Then
                    If .Reason3.Equals("Ejector Failed to advance") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Vision System Not Ready"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Vision System Not Ready") Then
                    If .Reason3.Equals("Ejector Failed to return") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Vision System Not Ready"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Vision System Not Ready") Then
                    If .Reason3.Equals("No Response from Vision System") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Vision System Not Ready"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Vision System Not Ready") Then
                    If .Reason3.Equals("Low Air Pressure (Reject Air Cylinder)") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Vision System Not Ready"
                    End If
                End If
            End If

            If .Reason1.Equals("Rinser") Then
                If .Reason2.Equals("Vision System Not Ready") Then
                    If .Reason3.Equals("Vision System Not Ready") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Rinser"
                        .Tier3 = "Vision System Not Ready"
                    End If
                End If
            End If

            If .Reason1.Equals("Sticker") Then
                If .Reason2.Equals("Comment") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Sticker"
                    .Tier3 = "Comment"
                End If
            End If

            If .Reason1.Equals("Sticker") Then
                If .Reason2.Equals("Sticker Placement") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Sticker"
                    .Tier3 = "Sticker Placement"
                End If
            End If

            If .Reason1.Equals("Sticker") Then
                If .Reason2.Equals("Web Tracking") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Sticker"
                    .Tier3 = "Web Tracking"
                End If
            End If

            If .Reason1.Equals("Sticker") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Stickers") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Sticker Quality"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Bottle discharge") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Bottle Discharge"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Bottle discharge") Then
                    If .Reason3.Equals("Starwheel out of Time") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Bottle Discharge "
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 1") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 1"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 2") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 2"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 10") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 10"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 11") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 11"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 12") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 12"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 13") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 13"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 14") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 14"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 15") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 15"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 16") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 16"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 17") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 17"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 18") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 18"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 19") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 19"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 20") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 20"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 3") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 3"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 4") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 4"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 5") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 5"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 6") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 6"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 7") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 7"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 8") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 8"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Cap stuck in airveyor") Then
                        If .Reason4.Equals("Zone 9") Then
                            .Tier1 = "Equipment"
                            .Tier2 = "Cap Delivery"
                            .Tier3 = "Zone 9"
                        End If
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Cap Delivery"
                        .Tier3 = "Comments"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Delivered") Then
                    If .Reason3.Equals("Loose Liners") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Cap Delivery"
                        .Tier3 = "Loose Liners"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Fasten to Bottle") Then
                    If .Reason3.Equals("Chuck not picking cap") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap Fasten to Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Fasten to Bottle") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap Fasten to Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Fasten to Bottle") Then
                    If .Reason3.Equals("Gripper Star not Align") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap Fasten to Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Pick-Off") Then
                    If .Reason3.Equals("Adjust Finger") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap Pick-Off"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Pick-Off") Then
                    If .Reason3.Equals("Arm is bent") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap Pick-Off"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Pick-Off") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap Pick-Off"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap Pick-Off") Then
                    If .Reason3.Equals("Upside down caps") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap Pick-Off"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap-Innerseal Check") Then
                    If .Reason3.Equals("Cocked Cap Detector not adjusted") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap-Innerseal Check"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap-Innerseal Check") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap-Innerseal Check"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Cap-Innerseal Check") Then
                    If .Reason3.Equals("No Foil") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Capper"
                        .Tier3 = "Cap-Innerseal Check"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Caps") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Caps Quality"
                    End If
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Operator Error") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Capper"
                    .Tier3 = "Operator Error"
                End If
            End If

            If .Reason1.Equals("Capper") Then
                If .Reason2.Equals("Material Quality") Then
                    If .Reason3.Equals("Bottles") Then
                        .Tier1 = "Material"
                        .Tier2 = "Quality"
                        .Tier3 = "Bottles Quality in Capper"
                    End If
                End If
            End If

            If .Reason1.Equals("Checkwiegher") Then
                If .Reason2.Equals("Transfer Bottle") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Checkwiegher"
                        .Tier3 = "Transfer Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Checkwiegher") Then
                If .Reason2.Equals("Transfer Bottle") Then
                    If .Reason3.Equals("Infeed Rails Adjusted") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Checkwiegher"
                        .Tier3 = "Transfer Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Checkwiegher") Then
                If .Reason2.Equals("Transfer Bottle") Then
                    If .Reason3.Equals("Outfeed Rails Adjusted") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Checkwiegher"
                        .Tier3 = "Transfer Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Checkwiegher") Then
                If .Reason2.Equals("Weigh Bottle") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Checkwiegher"
                        .Tier3 = "Weigh Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Checkwiegher") Then
                If .Reason2.Equals("Weigh Bottle") Then
                    If .Reason3.Equals("Scale not Calibrated") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Checkwiegher"
                        .Tier3 = "Weigh Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Infeed") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Infeed"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Infeed") Then
                    If .Reason3.Equals("Down Bottle") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Infeed"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Infeed") Then
                    If .Reason3.Equals("Feedscrew out of Time") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Infeed"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Infeed") Then
                    If .Reason3.Equals("Starwheel out of Time") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Infeed"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Transferred") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Transferred"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Transferred") Then
                    If .Reason3.Equals("Dog House Sticking") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Transferred"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Transferred") Then
                    If .Reason3.Equals("Starwheel out of Time") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Transferred"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Material Quality/Availability") Then
                    If .Reason3.Equals("Bottles") Then
                        .Tier1 = "Material"
                        .Tier2 = "Filler "
                        .Tier3 = "Bottles Quality/Availibility"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("FP Dust Collector") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler"
                    .Tier3 = "FP Dust Collector"
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Mezzanine") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Mezannine"
                    .Tier3 = "Filler"
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Dust Collector") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler"
                    .Tier3 = "Dust Collector"
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Operator Error") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler"
                    .Tier3 = "Operator Error"
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #1") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #10") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #2") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #3") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #4") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #5") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #6") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #7") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #8") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #9") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 1-10 Bottle Filled Properly") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 1-10 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #11") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #12") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #13") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #14") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #15") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #16") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #17") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #18") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #19") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #20") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 11-20 Bottle Filled Properly") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 11-20 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #21") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #22") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #23") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #24") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #25") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #26") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #27") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #28") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #29") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #30") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 21-30 Bottle Filled Properly") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 21-30 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #31") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #32") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #33") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #34") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #35") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #36") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #37") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #38") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #39") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #40") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Head 31-40 Bottle Filled Properly") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Head 31-40 Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #1, 11, 21, 31") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #10, 20, 30, 40") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #2, 12, 22, 32") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #3, 13, 23, 33") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #4, 14, 24, 34") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #5, 15, 25, 35") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #6, 16, 26, 36") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #7, 17, 27, 37") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #8, 18, 28, 38") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("LCI Card #9, 19, 29, 39") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly LCI Card") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly LCI Card"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #1") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #10") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #11") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #12") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #13") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #14") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #15") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #16") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #17") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #18") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #19") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #2") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #20") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #21") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #22") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #23") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #24") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #25") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #26") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #27") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #28") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #29") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #3") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #30") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #31") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #32") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #33") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #34") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #35") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #36") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #37") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #38") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #39") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #4") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #40") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #5") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #6") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #7") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #8") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Head #9") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Bottle Filled Properly") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Filler"
                        .Tier3 = "Bottle Filled Properly"
                    End If
                End If
            End If

            If .Reason1.Equals("Filler") Then
                If .Reason2.Equals("Room Not at Pressure") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Filler"
                    .Tier3 = "Room Not at Pressure"
                End If
            End If

            If .Reason1.Equals("Lepel") Then
                If .Reason2.Equals("Seal Bottle") Then
                    If .Reason3.Equals("Comment") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Lepel"
                        .Tier3 = "Seal Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Lepel") Then
                If .Reason2.Equals("Seal Bottle") Then
                    If .Reason3.Equals("No Heat Induction") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Lepel"
                        .Tier3 = "Seal Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Lepel") Then
                If .Reason2.Equals("Seal Bottle") Then
                    If .Reason3.Equals("Water not circulating") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Lepel"
                        .Tier3 = "Seal Bottle"
                    End If
                End If
            End If

            If .Reason1.Equals("Video Jet Coder") Then
                If .Reason2.Equals("Comment") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Video Jet Coder"
                    .Tier3 = "Comments"
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("PROFICY/RTCIS/SAP") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Desert Surge"
                    .Tier3 = "PROFICY/RTCIS/SAP"
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("Mezzanine") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Mezannine"
                    .Tier3 = "Desert Surge"
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("Tote Stands") Then
                    If .Reason3.Equals("A1 & A2-STR/OTR") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Tote Stands"
                        .Tier3 = "A1 & A2-STR/OTR"
                    End If
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("Tote Stands") Then
                    If .Reason3.Equals("A3 & A4-STOS/OTO") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Tote Stands"
                        .Tier3 = "A3 & A4-STOS/OTO"
                    End If
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("Tote Stands") Then
                    If .Reason3.Equals("A5 &A6-STOF/STCF/STBF/STLF") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Tote Stands"
                        .Tier3 = "A5 &A6-STOF/STCF/STBF/STLF"
                    End If
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("Transporters") Then
                    If .Reason3.Equals("Trans #1-STOF/STCF/STBF/STLF") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Transporters"
                        .Tier3 = "Trans #1-STOF/STCF/STBF/STLF"
                    End If
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("Transporters") Then
                    If .Reason3.Equals("Trans #2-STOS/OTO") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Transporters"
                        .Tier3 = "Trans #2-STOS/OTO"
                    End If
                End If
            End If

            If .Reason1.Equals("Desert Surge") Then
                If .Reason2.Equals("Transporters") Then
                    If .Reason3.Equals("Trans #3-STR/OTR") Then
                        .Tier1 = "Equipment"
                        .Tier2 = "Transporters"
                        .Tier3 = "Trans #3-STR/OTR"
                    End If
                End If
            End If

            If .Reason1.Equals("Accumulation Table") Then
                If .Reason2.Equals("Comment") Then
                    .Tier1 = "Equipment"
                    .Tier2 = "Accumulation Table"
                    .Tier3 = ""
                End If
            End If

            If .Reason1.Equals("Quality Related Shutdown") Then
                .Tier1 = "Others"
                .Tier2 = "Quality Related Shutdown"
                .Tier3 = ""
            End If

            If .Reason1.Equals("Changeover Over Target") Then
                .Tier1 = "PDT Over Target"
                .Tier2 = "Changeover "
                .Tier3 = ""
            End If

            If .Reason1.Equals("CIL Over Goal") Then
                .Tier1 = "PDT Over Target"
                .Tier2 = "CIL"
                .Tier3 = ""
            End If

            If .Reason1.Equals("Shift Exchange Over Goal") Then
                .Tier1 = "PDT Over Target"
                .Tier2 = "Shift Exchange"
                .Tier3 = ""
            End If

            If .Reason1.Equals("Bottle Rejector") Then
                .Tier1 = "Equipment"
                .Tier2 = "Bottle Rejector"
                .Tier3 = ""
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Changeover") Then
                    If .Reason3.Equals("Size&Flavor") Then
                        .Tier1 = "Changeover"
                        .Tier2 = "Size&Flavor"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Changeover") Then
                    If .Reason3.Equals("Component") Then
                        .Tier1 = "Changeover"
                        .Tier2 = "Component"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Changeover") Then
                    If .Reason3.Equals("Formula") Then
                        .Tier1 = "Changeover"
                        .Tier2 = "Formula"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Changeover") Then
                    If .Reason3.Equals("Size") Then
                        .Tier1 = "Changeover"
                        .Tier2 = "Size"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Maintenance") Then
                    If .Reason3.Equals("AM") Then
                        .Tier1 = "AM"
                        .Tier2 = ""
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Maintenance") Then
                    If .Reason3.Equals("CIL") Then
                        .Tier1 = "CIL"
                        .Tier2 = ""
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Start-Up & Shut-Down") Then
                    If .Reason3.Equals("BOW") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = "BOW"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Start-Up & Shut-Down") Then
                    If .Reason3.Equals("EOW") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = "EOW"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Start-Up & Shut-Down") Then
                    If .Reason3.Equals("BOS") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = "BOS"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Start-Up & Shut-Down") Then
                    If .Reason3.Equals("EOS") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = "EOS"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Start-Up & Shut-Down") Then
                    If .Reason3.Equals("BOL") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = "BOL"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Start-Up & Shut-Down") Then
                    If .Reason3.Equals("EOL") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = "EOL"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Start-Up & Shut-Down") Then
                    If .Reason3.Equals("Weekend") Then
                        .Tier1 = "SU/SD"
                        .Tier2 = "Weekend"
                        .Tier3 = ""
                    End If
                End If
            End If

            If .Reason1.Equals("Not Equipment") Then
                If .Reason2.Equals("Meeting/Training") Then
                    .Tier1 = "Meeting/Training"
                    .Tier2 = ""
                    .Tier3 = ""
                End If
            End If


            If .Tier3 = "Comments" Or .Tier1 = "Comment" Then
                .Tier3 = ""
            End If


            If .Tier1 = "CIL" Then .Tier2 = .Team
            If .Tier1 = "AM" Then .Tier2 = .Team


            'DTGROUP WRITE TO
            If Not .Tier1.Equals("Equipment") Then
                .DTGroup = .Tier1 & "-" & .Tier2 & "-" & .Tier3
            Else
                If Not .Reason3.Equals("Comment") Or Not .Reason3.Equals("Comments") Then
                    .DTGroup = .Tier2 & "-" & .Tier3 & "-" & .Reason3
                Else
                    .DTGroup = .Tier2 & "-" & .Tier3
                End If
            End If


        End With
    End Sub

    Public Sub getPheonix_DprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                If .Reason1.Contains("Material Q") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Equals("Filler") Then
                    .Tier1 = "Equipment"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason1.Equals("Splicer") Then
                    .Tier1 = "Equipment"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason1.Equals("Cartoner") Then
                    .Tier1 = "Equipment"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason1.Contains("Blocked") Or .Reason1.Contains("Starved") Then
                    If .Reason2.Equals("Cartoner") Then
                        .Tier1 = "Equipment"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Reason2.Equals("Knife") Then
                        .Tier1 = "Equipment"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Reason2.Equals("Shrink Bundler") Then
                        .Tier1 = "Equipment"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Reason2.Equals("Labeler") Then
                        .Tier1 = "Equipment"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Reason2.Equals("Capper") Then
                        .Tier1 = "Equipment"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Reason2.Equals("Polypack") Then
                        .Tier1 = "Equipment"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    ElseIf .Reason2.Equals("Unscrambler") Then
                        .Tier1 = "Equipment"
                        .Tier2 = .Reason2
                        .Tier3 = .Reason3
                    End If
                End If
            Else
                If .Reason2.Contains("Change") Then
                    .Tier1 = "Changeover"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("Start") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Reason3 & "-" & .Team
                ElseIf .Reason2.Contains("Start") Then
                    .Tier1 = "Meeting/Training"
                    .Tier2 = .Reason3 & "-" & .Team
                End If
            End If
            If .Tier1 = BLANK_INDICATOR Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
                .Tier3 = .Fault
            End If
        End With
    End Sub

    Public Sub getSwingRoadprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason1.Contains("Material") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason2
                    .Tier3 = .ProductCode
                Else
                    .Tier1 = "Equipment"
                    .Tier2 = .Location
                    .Tier3 = .Reason1 & "-" & .Reason2
                End If
                '  .Tier1 = OTHERS_STRING
                '  .Tier2 = .Reason1
                '  .Tier3 = .Reason2
            Else
                If .Reason1.Contains("Changeover") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("CIL") Or .Reason2.Contains("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("AM") Then
                    .Tier1 = "AM Work"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("Cleaning") Then
                    .Tier1 = "Cleaning"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("Startup") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("Change") Then
                    If .Reason2.Contains("Batch") Then
                        .Tier1 = "CO"
                        .Tier2 = "Batch"
                    ElseIf .Reason2.Contains("Flavor") Then
                        .Tier1 = "CO"
                        .Tier2 = "Flavor"
                    ElseIf .Reason2.Contains("Process") Then
                        .Tier1 = "CO"
                        .Tier2 = "Process"
                    ElseIf .Reason2.Contains("Label") Then
                        .Tier1 = "CO"
                        .Tier2 = "Label"
                    End If
                End If

            End If

            If .Tier1.Equals(BLANK_INDICATOR) Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
                .Tier3 = .Reason2
            End If



        End With
    End Sub

    Public Sub getSwingRoadprstoryMapping_67(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason1.Contains("Material") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason2
                    .Tier3 = .ProductCode
                ElseIf .Reason3.Equals("Utilities") Then
                    .Tier1 = .Reason3
                    .Tier2 = .Reason4
                    .Tier3 = .ProductCode

                ElseIf .Reason1.Equals("Thermoformer") Then
                    .Tier1 = "Thermoformer"
                    If Left(.Reason2, 2).Equals("L6") And Len(.Reason2) > 3 Then
                        .Tier2 = Right(.Reason2, Len(.Reason2) - 2)
                    Else
                        .Tier2 = .Reason2
                    End If
                    .Tier3 = .Reason3

                ElseIf .Reason1.Equals("Blisterformer") Then
                    .Tier1 = "Blisterformer"
                    If Left(.Reason2, 2).Equals("L7") And Len(.Reason2) > 3 Then
                        .Tier2 = Right(.Reason2, Len(.Reason2) - 2)
                    Else
                        .Tier2 = .Reason2
                    End If
                    .Tier3 = .Reason3

                ElseIf .Reason2.Contains("Cartoner") Then
                    .Tier1 = "Equip-Other"
                    .Tier2 = "Cartoner"
                    .Tier3 = .Reason3
                ElseIf .Reason2.Contains("Bander") Then
                    .Tier1 = "Equip-Other"
                    .Tier2 = "Bander"
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Not Equipment Related") And .Reason3.Equals(BLANK_INDICATOR) Then
                    .Tier1 = "Equip-Other"
                    .Tier2 = .Fault
                    .Tier3 = .Product
                ElseIf .Reason1.Contains("Blocked") And .Reason2.Equals(BLANK_INDICATOR) Then
                    .Tier1 = "Equip-Other"
                    .Tier2 = .Fault
                    .Tier3 = .Product
                End If




            Else 'PLANNED!!!
                If .Reason1.Contains("Changeover") Or .Reason1.Contains("Product Change") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason2
                ElseIf .Reason1.Contains("CIL") Or .Reason2.Contains("CIL") Or .Reason3.Contains("CILs") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("AM") Then
                    .Tier1 = "AM Work"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("Cleaning") Or .Reason3.Contains("Cleaning") Then
                    .Tier1 = "Cleaning"
                    .Tier2 = .Product
                ElseIf .Reason2.Contains("Startup") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Reason2
                ElseIf .Reason2.Contains("Change") Then
                    If .Reason2.Contains("Batch") Then
                        .Tier1 = "CO"
                        .Tier2 = "Batch"
                    ElseIf .Reason2.Contains("Flavor") Then
                        .Tier1 = "CO"
                        .Tier2 = "Flavor"
                    ElseIf .Reason2.Contains("Process") Then
                        .Tier1 = "CO"
                        .Tier2 = "Process"
                    ElseIf .Reason2.Contains("Label") Then
                        .Tier1 = "CO"
                        .Tier2 = "Label"
                    ElseIf .Reason2.Contains("Product") Then
                        .Tier1 = "CO"
                        .Tier2 = "Product"
                    ElseIf .Reason2.Contains("PO") Then
                        .Tier1 = "CO"
                        .Tier2 = "PO"
                    End If
                ElseIf .Reason4.Contains("Splice") Then
                    .Tier1 = "Splice"
                    .Tier2 = .Reason3
                ElseIf .Reason2.Contains("Maintenance") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason3
                End If
            End If

            If .Tier1.Equals(BLANK_INDICATOR) Then
                .Tier1 = OTHERS_STRING
                .Tier2 = .Reason1
                .Tier3 = .Reason2
            End If



        End With
    End Sub


#End Region

#Region "Beauty Care"

    Public Sub getSingaporePioneerMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If Left(.Reason2, 6).Equals("Filler") Then
                    .Tier1 = "Filler"
                    .Tier2 = .Reason3
                    .Tier3 = .Fault
                ElseIf .Reason1.Contains("End of line") Or .Reason1.Contains("End Of Line") Then
                    .Tier1 = "EOL"
                    .Tier2 = .Reason3
                    .Tier3 = .Fault
                ElseIf .Reason2.Equals("Utilities") Then
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                    .Tier3 = .Fault
                ElseIf .Reason1.Contains("SHUBHAM") Then
                    .Tier1 = "Shubham"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                '#Region "Planned"
                If .Reason2.Equals("Induced Stop") Then
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                    .Tier3 = .Fault
                ElseIf .Reason2.Equals("Changeover") Then
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                    .Tier3 = .Fault
                ElseIf .Reason3.Equals("CIL") Then
                    .Tier1 = .Reason3
                    .Tier2 = .Fault
                    .Tier3 = .Product
                ElseIf .Reason3.Equals("DDS") Or .Reason4.Equals("DDS") Then
                    .Tier1 = "DDS"
                    .Tier2 = .Fault
                    .Tier3 = .Product
                ElseIf .Reason3.Contains("SPLICE") Or .Reason3.Contains("Splicing") Then
                    .Tier1 = "Splice"
                    .Tier2 = .Fault
                    .Tier3 = .Product
                ElseIf .Reason3.Equals("Passivation") Then
                    .Tier1 = .Reason3
                    .Tier2 = .Fault
                    .Tier3 = .Product
                ElseIf .Reason3.Equals("Startup/Shutdown") Then
                    .Tier1 = .Reason3
                    .Tier2 = .Fault
                    .Tier3 = .Product
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
                '#End Region
            End If
        End With
    End Sub



    Public Sub getHuangpuprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .DTGroup.Contains("Equip") And Not .Reason1.Contains("Product Delivery") And Not .DTGroup.Contains("Bad") Then
                    .Tier1 = "Equipment"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .DTGroup.Contains("Mater") Or .Reason2.Contains("Material") Or .Reason1.Contains("material") Or .Reason3.Contains("Material") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason2.Contains("Utilit") Then
                    .Tier1 = "Utilities"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason1.Equals("Filler") Or .Reason1.Equals("Capper") Or .Reason1.Equals("Labeler") Or .Reason1.Equals("PSL") Or .Reason1.Equals("Filler/Capper") Then
                    .Tier1 = "Equipment"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Tier1.Equals(BLANK_INDICATOR) Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .DTGroup.Contains("CIL") Or .Reason1.Contains("CIL") Or .Reason2.Contains("CIL") Or .Reason3.Contains("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team
                ElseIf .DTGroup.Contains("CO") Or .DTGroup.Contains("Change") Or .Reason1.Contains("Change") Or .Reason2.Contains("Changeover") Or .Reason3.Contains("Changeover") Or .Reason2.Equals("Product Change") Then
                    .Tier1 = "CO"
                    .Tier2 = .Reason3
                ElseIf .DTGroup.Contains("Maint") Or .DTGroup.Contains("Maint") Or .Reason1.Contains("Maint") Or .Reason2.Contains("Maint") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason3
                    '  ElseIf .Reason3.Contains("Pause") Then
                    '      .Tier1 = "Pause"
                    '      .Tier2 = .Reason2
                    '  ElseIf .Reason3.Contains("Split PO") Then
                    '      .Tier1 = "Split PO"
                    '      .Tier2 = .Reason4
                Else

                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    '   .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub

    Public Sub getMariscalaprstoryMapping2(ByRef searchevent As DowntimeEvent)
        '"AJUSTES", "EMERGENCIAS", "EQUIPO", "INTERVENCIONES", "MAKING", "MATERIALES", "RCO", "SERVICIOS", "TERMINO DE CEDULA"

        With searchevent
            .Tier2 = .Reason2
            .Tier3 = .Reason3


            If .Reason1.Contains("Planeadas") Or .Reason1.Contains("Vendible") Or .Reason1.Contains("PSUs") Or .Reason1.Contains("Ingenieria") Then
                .Tier1 = "INTERVENCIONES"
            ElseIf .Reason1.Contains("Emergencias") Then
                .Tier1 = "EMERGENCIAS"
            ElseIf .Reason1.Contains("Personal") Or .Reason1.Contains("Servicios") Then
                .Tier1 = "SERVICIOS"
            ElseIf .Reason1.Contains("Materiales") Then
                .Tier1 = "MATERIALES"
            ElseIf .Reason1.Contains("Ajustes") Then
                .Tier1 = "AJUSTES"
            ElseIf .Reason1.Contains("Making") Then
                .Tier1 = "MAKING"
            ElseIf .Reason1.Contains("cedula") Then
                .Tier1 = "TERMINO DE CEDULA"
            ElseIf .Reason1.Contains("Encintadora") Or .Reason1.Contains("Laminado") Or .Reason1.Contains("Sellado") Then
                .Tier1 = "EQUIPO"
            ElseIf .Reason1.Contains("Corte") Or .Reason1.Contains("Antiretorno") Or .Reason1.Contains("Codificado") Then
                .Tier1 = "EQUIPO"
            ElseIf .Reason1.Contains("Manifold") Or .Reason1.Contains("Dosification") Or .Reason1.Contains("Llenadora") Then
                .Tier1 = "EQUIPO"
            ElseIf .Reason1.Contains("Gabinete") Or .Reason1.Contains("Corrugados") Or .Reason1.Contains("Sachets") Or .Reason1.Contains("IJ3000") Then
                .Tier1 = "EQUIPO"
            ElseIf .Reason1.Contains("Case Former") Or .Reason1.Contains("Case Packer") Or .Reason1.Contains("Checadora de peso") Or .Reason1.Contains("Hanger Feeder") Then
                .Tier1 = "EQUIPO"
            ElseIf .Reason1.Contains("Falla de Sistema") Then
                .Tier1 = "SERVICIOS"
            ElseIf .Reason1.Contains("Changeovers") Then
                .Tier1 = "RCO"
            Else
                .Tier1 = OTHERS_STRING
            End If

        End With
    End Sub

    Public Sub getMariscalaprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                ' If .Reason1.Equals("Personal") Then
                '   .Tier1 = .Reason1
                '  .Tier2 = .Reason2
                ' .Tier3 = .Reason3
                If .Reason1.Equals("Encartonadora") Then
                    .Tier1 = "Equipo"
                    .Tier2 = "Encartonadora"
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Materiales") Then
                    .Tier1 = "Materiales"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Intellifeed") Then
                    .Tier1 = "Intellifeed"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason1.Contains("Encelofanadora") Then
                    .Tier1 = "Equipo"
                    .Tier2 = "Encelofanadora"
                    .Tier3 = .Reason2
                ElseIf .Reason1.Contains("Llenadora") Then
                    .Tier1 = "Equipo"
                    .Tier2 = "Llenadora"
                    .Tier3 = .Reason2
                ElseIf .Reason1.Contains("Case Packer") Then
                    .Tier1 = "Equipo"
                    .Tier2 = "Case Packer"
                    .Tier3 = .Reason2
                ElseIf .Reason1.Contains("Sistemas") Then
                    .Tier1 = "Sistemas"
                    .Tier2 = .Reason2
                    .Tier3 = .Product
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Equals("Mantenimiento Autonomo") Then
                    .Tier1 = "AM"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                ElseIf .Reason1.Equals("Cambios") Or .Reason1.Contains("Changeovers") Then
                    .Tier1 = "Cambios"
                    .Tier2 = .Reason2
                    .Tier3 = .Team
                ElseIf .Reason2.Equals("Planeadas") Then
                    .Tier1 = "Planeadas"
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                Else 'Planeadas
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub


#End Region

    Public Sub getICOCprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                If .Reason1.Contains("Blister") Then
                    .Tier1 = "Blister Machine"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Cartoner") Then
                    .Tier1 = "Cartoner"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Labeler") Then
                    .Tier1 = "Labeler"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Case") Then
                    .Tier1 = "Case Packer"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason2.Contains("Material") Then
                    .Tier1 = "Material"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Brush") Then
                    .Tier1 = "Brush Transfer"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Tray") Then
                    .Tier1 = "Tray"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("ATS") Then
                    .Tier1 = "ATS"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Rollover") Then
                    .Tier1 = "Rollover"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason1.Contains("Pick") Then
                    .Tier1 = "Pick & Place"
                    If .Reason1.Contains("2") Then
                        .Tier2 = "Pick & Place 2"
                    Else
                        .Tier2 = "Pick & Place 1"
                    End If

                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason3.Contains("Changeover") Then
                    .Tier1 = "Changeover"
                    .Tier2 = .DTGroup
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("CILs") Then
                    .Tier1 = "CILs"
                    .Tier2 = .Team
                    .Tier3 = .Fault
                ElseIf .Reason3.Contains("Material") Then
                    .Tier1 = "Material"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Training") Then
                    .Tier1 = "Training"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Maintenance") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("AM") Then
                    .Tier1 = "AM"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                ElseIf .Reason3.Contains("Startup/Shutdown") Then
                    .Tier1 = "SU/SD"
                    .Tier2 = .Reason4
                    .Tier3 = .Team
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason3
                    .Tier3 = .Team
                End If
            End If
        End With
    End Sub

    Public Sub getICOCMakingprstoryMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .isUnplanned Then
                If Left(.Reason1, 9) = "ICOC 304 " Or Left(.Reason2, 9) = "ICOC 305 " Then
                    .Tier1 = Right(.Reason1, .Reason1.Length - 9)
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                Else
                    .Tier1 = .Reason1
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                End If
            Else

                .Tier1 = .Reason3
                .Tier2 = .Team
                .Tier3 = .Fault
            End If
        End With
    End Sub


    Public Sub getGENERICprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .DTGroup.Contains("Equip") And Not .Reason1.Contains("Product Delivery") And Not .DTGroup.Contains("Bad") Then
                    .Tier1 = "Equipment"
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                ElseIf .Reason2.Contains("Material") Or .Reason1.Contains("Material") Then
                    .Tier1 = "Materials"
                    .Tier2 = .Reason2
                    .Tier3 = .Reason3
                ElseIf .Reason2.Contains("Utilit") Then
                    .Tier1 = "Utilities"
                    .Tier2 = .Reason1
                    .Tier3 = .Product
                ElseIf .Reason1.Contains("Product Delivery") Then
                    .Tier1 = "Product Delivery"
                    .Tier2 = .Reason3
                    .Tier3 = .ProductGroup
                ElseIf .Tier1.Equals(BLANK_INDICATOR) Then
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Contains("CIL") Then
                    .Tier1 = "CIL"
                    .Tier2 = .Team
                ElseIf .Reason3.Contains("Washout") Then
                    .Tier1 = "Washout"
                    .Tier2 = .Team
                ElseIf .Reason3.Contains("Purge") Then
                    .Tier1 = "Purge"
                    .Tier2 = .Team
                ElseIf .Reason2.Equals("AM") Then
                    .Tier1 = "AM"
                    .Tier2 = .Team
                ElseIf .Reason2.Contains("Maint") Then
                    .Tier1 = "Maintenance"
                    .Tier2 = .Reason3
                ElseIf .Reason3.Contains("Pause") Then
                    .Tier1 = "Pause"
                    .Tier2 = .Reason2
                ElseIf .Reason3.Contains("Split PO") Then 'ЗАМЕНА МАТРИЦЫ_
                    .Tier1 = "Split PO"
                    .Tier2 = .Reason4
                Else

                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    '   .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub
    '{"3M", "CERMEX",
    '  "NGS", "Bundles", "Frascos", OTHERS_STRING}
    Public Sub getRIOprstoryMapping(ByRef searchevent As DowntimeEvent)
        With searchevent
            If .isUnplanned Then
                If .Reason2.Contains("3M") Or .Reason2.Contains("CERMEX") Or .Reason2.Contains("NGS") Or .Reason2.Contains("Bundles") Or .Reason2.Contains("Frascos") Then
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                    .Tier3 = .Reason4
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            Else
                If .Reason2.Contains("3M") Or .Reason2.Contains("CERMEX") Or .Reason2.Contains("NGS") Or .Reason2.Contains("Bundles") Or .Reason2.Contains("Frascos") Then
                    .Tier1 = .Reason2
                    .Tier2 = .Reason3
                    .Tier3 = .Reason4
                Else
                    .Tier1 = OTHERS_STRING
                    .Tier2 = .Reason1
                    .Tier3 = .Reason2
                End If
            End If

        End With
    End Sub



End Module

Module Mapping_prstory_FIXEDFIELDS

    Public Function getprStoryCardField(mappingIndicator As Integer, prStoryCardNum As Integer, fieldNum As Integer) As String
        Select Case mappingIndicator
            Case prStoryMapping.BudapestFGC
                Return getBudapestFGCprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.BudapestLCC
                Return getBudapestLCCprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Hyderabad
                Return getHyderabadprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Rakona
                Return getRakonaprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.SwingRoad
                Return getSwingRoadprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.SwingRoad_6
                Return getSwingRoadprStoryCardFixedFieldName_6(prStoryCardNum, fieldNum)
            Case prStoryMapping.SwingRoad_7
                Return getSwingRoadprStoryCardFixedFieldName_7(prStoryCardNum, fieldNum)
            Case prStoryMapping.OralCare
                Return getOralCareprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.OralCareNau
                Return getOralCareprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.OralCareGross
                Return getOralCareprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.OralCare_DF
                Return getOralCareprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.OralCareNau_DF
                Return getOralCareprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Mandideep_Fem
                Return getMandideepprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Mandideep
                Return getMandideepprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.SkinCare
                Return getSkinCareprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.IowaCity
                Return getIowaCityprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.IowaCityBeauty
                Return getIowaCityBeautyprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Phoenix
                Return getPhoenixprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Pheonix_D
                Return getPhoenixprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.HuangPu
                Return getHuangpuprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.APDO_I
                Return getAPDO_I_prStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.APDO_J
                Return getAPDO_J_prStoryCardFixedFieldName(prStoryCardNum, fieldNum)
          '  Case prStoryMapping.GENERIC
               ' Return getIowaCityBeautyprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Boryspil
                Return getBoryspilprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.FemCare_Pads
                Return getBellevilleprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.TepejiFem
                Return getTepejiFemprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Mariscala
                Return getMariscalaprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.Albany
                Return getAlbanyprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.FamilyCareUnitOP_Wrapper
                Return getFamilyCareUnitOP_wrapperprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.FamilyCareUnitOP_mf
                Return getFamilyCareUnitOP_mfprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.FamilyCareUnitOP_ACP
                Return getFamilyCareUnitOP_ACPprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.FamilyMaking
                Return getFamilyMakingprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.NaucalpanPHC_B
                Return getNaucalpanPHC_BprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.NaucalpanPHC_J
                Return getNaucalpanPHC_JprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.NaucalpanPHC_Mex
                Return getNaucalpanPHC_MexprStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.NaucalpanPHC_Vita1
                Return getNaucalpanPHC_Vita1prStoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case prStoryMapping.SingaporePioneer
                Return getSingaporePioneer_prstoryCardFixedFieldName(prStoryCardNum, fieldNum)
            Case Else
                Return BLANK_INDICATOR
        End Select
    End Function

#Region "F&HC"
    Private Function getRakonaprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Constraint"
                    Case 2
                        Return "Downstream"
                    Case 3
                        Return "Upstream"
                    Case 4
                        Return "External"
                    Case 5
                        Return "Material supply"
                    Case 6
                        Return "Materials"
                    Case 7
                        Return "QA"
                    Case 8
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Changeover"
                    Case 2
                        Return "CIL/RLS"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "C&S"
                    Case 5
                        Return "Training"
                    Case 6
                        Return "Raw material change"
                    Case 7
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function

    Private Function getHyderabadprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equip-AKASH"
                    Case 2
                        Return "Equip-ASL"
                    Case 3
                        Return "Making"
                    Case 4
                        Return "Materials"
                    Case 5
                        Return "DC System"
                    Case 6
                        Return "Utilities"
                    Case 7
                        Return "CVC System"
                    Case 8
                        Return "COMMON CONV"
                    Case 9
                        Return "Equip-Others"
                    Case 10
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Thread Change"
                    Case 5
                        Return "Roll Change"
                    Case 6
                        Return "SU/SD"

                    Case 7
                        Return "EO"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function


#End Region

#Region "Family"
    Private Function getAlbanyprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Converter"
                    Case 2
                        Return "Block/Starved"
                    Case 3
                        Return "Materials"
                    Case 4
                        Return "ELP"
                    Case 5
                        Return "Utilities"
                    Case 6
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "PR Changes"
                    Case 2
                        Return "CO"
                    Case 3
                        Return "CIL"
                    Case 4
                        Return "AM"
                    Case 5
                        Return "Blowdown"
                    Case 6
                        Return "Maintenance"
                    Case 7
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getFamilyCareUnitOP_ACPprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Jam"
                    Case 2
                        Return "AIR-Vacuum"
                    Case 3
                        Return "Electrical / Programming"
                    Case 4
                        Return "Loader"
                    Case 5
                        Return "Quality"
                    Case 6
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Changeover"
                    Case 2
                        Return "Unscheduled Time"
                    Case 3
                        Return "Planned Intervention"
                    Case 4
                        Return "Blocked/Starved"
                    Case 5
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getFamilyCareUnitOP_mfprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"

                    Case 1
                        Return "Drop/Tip Rolls"
                    Case 2
                        Return "Gap Fault"
                    Case 3
                        Return "Quality"
                    Case 4
                        Return "Sealing"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Changeover"
                    Case 2
                        Return "Poly Splice"
                    Case 3
                        Return "Planned Intervention"
                    Case 4
                        Return "Blocked/Starved"
                    Case 5
                        Return "Unscheduled Time"
                    Case 6
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getFamilyCareUnitOP_wrapperprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"

                    Case 1
                        Return "Breakdown"
                    Case 2
                        Return "Jam"
                    Case 3
                        Return "Centerlines Set up"
                    Case 4
                        Return "Film Loss & uws"
                    Case 5
                        Return "Quality"
                    Case 6
                        Return "Film Loss & uws"
                    Case 7
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Changeover"
                    Case 2
                        Return "AM"
                    Case 3
                        Return "CL/RLS/CIL"
                    Case 4
                        Return "Other Maintenance"
                    Case 5
                        Return "Poly Change"
                    Case 6
                        Return "Blocked/Starved"
                    Case 7
                        Return "Unscheduled Time"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function

    Private Function getFamilyMakingprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Zone 1"
                    Case 2
                        Return "Zone 2"
                    Case 3
                        Return "Zone 3"
                    Case 4
                        Return "Zone 4"

                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Brand Change"
                    Case 2
                        Return "Downday"
                    Case 3
                        Return "RLS"
                    Case 4
                        Return "DT RLS"
                    Case 5
                        Return "Soft Swing"

                    Case 6
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function


#End Region

#Region "Fem"
    Private Function getBoryspilprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Area 1"
                    Case 2
                        Return "Area 2"
                    Case 3
                        Return "Area 3"
                    Case 4
                        Return "Area 4"
                    '  Case 5
                    '      Return "Area/площадь 5"
                    '   Case 5
                    '       Return "Material"
                    '  Case 6
                    '      Return "Utilities"
                    ' Case 8
                    '     Return "Offline"
                    Case 5
                        Return OTHERS_STRING

                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "PM"
                    Case 3
                        Return "Organization"
                    Case 4
                        Return "Induced Stops"
                    Case 5
                        Return "ЗАМЕНА МАТРИЦЫ"
                    Case 6
                        Return "CIL"
                    Case 7
                        Return OTHERS_STRING
                    Case 8
                        Return ""
                End Select
        End Select

        Return ""
    End Function
    Private Function getBellevilleprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Area 1"
                    Case 2
                        Return "Area 2"
                    Case 3
                        Return "Area 3"
                    Case 4
                        Return "Area 4"
                    '  Case 5
                    '      Return "Area/площадь 5"
                    '   Case 5
                    '       Return "Material"
                    '  Case 6
                    '      Return "Utilities"
                    ' Case 8
                    '     Return "Offline"
                    Case 5
                        Return OTHERS_STRING

                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Meeting"
                    Case 4
                        Return "Logistics"
                    Case 5
                        Return "PM"
                    Case 6
                        Return "Training"
                    Case 7
                        Return OTHERS_STRING
                    Case 8
                        Return ""
                End Select
        End Select

        Return ""
    End Function
    Private Function getTepejiFemprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Area 1"
                    Case 2
                        Return "Area 2"
                    Case 3
                        Return "Area 3"
                    Case 4
                        Return "Area 4"

                    Case 5
                        Return OTHERS_STRING

                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "MAINTENANCE"
                    Case 3
                        Return "IWS SHUTDOWN"
                    Case 4
                        Return "ADMIN SHUTDOWN"
                    Case 5
                        Return "PROJECTS"
                    Case 6
                        Return "OPERATIONAL"
                    Case 7
                        Return OTHERS_STRING
                    Case 8
                        Return ""
                End Select
        End Select

        Return ""
    End Function
    Private Function getBudapestFGCprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Area 1"
                    Case 2
                        Return "Area 2"
                    Case 3
                        Return "Area 3"
                    Case 4
                        Return "Area 4"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "PM"
                    Case 4
                        Return "Logistics"
                    Case 5
                        Return "Training"
                    Case 6
                        Return "Meeting"
                    Case 7
                        Return "Projects"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getBudapestLCCprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Area 1"
                    Case 2
                        Return "Area 2"
                    Case 3
                        Return "Area 3"
                    Case 4
                        Return "Area 4"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "PM"
                    Case 4
                        Return "Logistics"
                    Case 5
                        Return "Training"
                    Case 6
                        Return "Meeting"
                    Case 7
                        Return "Projects"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function



#End Region

#Region "Baby"

    Private Function getMandideepprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "EA1"
                    Case 2
                        Return "EA2"
                    Case 3
                        Return "EA3"
                    Case 4
                        Return "EA4"
                    Case 5
                        Return "EA6"
                    Case 6
                        Return "Utilities"
                    Case 7
                        Return "Offline"
                    Case 8
                        Return OTHERS_STRING

                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "PM"
                    Case 3
                        Return "Organization"
                    Case 4
                        Return "Induced Stops"
                    Case 5
                        Return "Projects / Construction"
                    Case 6
                        Return "AM/CIL/RLS"
                    Case 7
                        Return OTHERS_STRING
                    Case 8
                        Return ""
                End Select
            Case prStoryCard.Materials
                Select Case fieldNumber
                    Case 0
                        Return "Availability"
                    Case 1
                        Return "Quality"
                End Select
            Case prStoryCard.Bulk
                Select Case fieldNumber
                    Case 0
                        Return "Availability"
                    Case 1
                        Return "Quality"
                End Select
        End Select

        Return ""
    End Function


#End Region

#Region "Oral"
    Private Function getIowaCityprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equip_Minor"
                    Case 2
                        Return "Equip_PF"
                    Case 3
                        Return "Equip_BD"
                    Case 4
                        Return "Equip_Others"
                    Case 5
                        Return "Material"
                    Case 6
                        Return "Product"
                    Case 7
                        Return "Utilities"
                    Case 8
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Format Change"
                    Case 2
                        Return "LE work"
                    Case 3
                        Return "CIL"
                    Case 4
                        Return "AM"
                    Case 5
                        Return "Training-Meeting"
                    Case 6
                        Return "Lunch/Break"
                    Case 7
                        Return "PM" 'THEY DONT WANT THIS ONE
                    Case 8
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Materials
                Select Case fieldNumber
                    Case 0
                        Return "Supply"
                    Case 1
                        Return "Quality"
                End Select
            Case prStoryCard.Bulk
                Select Case fieldNumber
                    Case 0
                        Return "Supply"
                    Case 1
                        Return "Quality"
                End Select
        End Select

        Return ""
    End Function
    Private Function getOralCareprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equipment"
                    Case 2
                        Return "Materials"
                    Case 3
                        Return "Paste"
                    Case 4
                        Return "Supply Losses"
                    Case 5
                        Return "Utilities"
                    Case 6
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Train/Meet"
                    Case 5
                        Return "Materials"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "Projects"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function

    Private Function getICOCprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
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
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Cartoner"
                    Case 2
                        Return "Blister Machine"
                    Case 3
                        Return "EOL"
                    Case 4
                        Return "Tray"
                    Case 5
                        Return "Rollover"
                    Case 6
                        Return "ATS"
                    Case 7
                        Return "Pick & Place"
                    Case 8
                        Return "Case Packer"
                    Case 9
                        Return "Brush Transfer"
                    Case 10
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Changeover"
                    Case 2
                        Return "CILs"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Training"
                    Case 5
                        Return "Material"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "AM"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function



#End Region

#Region "PHC"

    Private Function getNaucalpanPHC_BprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Cartoneta"
                    Case 2
                        Return "Trayformer"
                    Case 3
                        Return "Polypack"
                    Case 4
                        Return "Wrap Ade"
                    Case 5
                        Return "Non-Equip"
                    Case 6
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "Mantenimiento"
                    Case 3
                        Return "Meeting"
                    Case 4
                        Return "Initiatives"
                    Case 5
                        Return "Projects"
                    Case 6
                        Return "Material Resupply"
                    Case 7
                        Return "Procedimiento calidad"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getNaucalpanPHC_JprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Cartoneta"
                    Case 2
                        Return "Algodonadora"
                    Case 3
                        Return "Polypack"
                    Case 4
                        Return "Etiquetadora"
                    Case 5
                        Return "Llenadora"
                    Case 6
                        Return "Sorter"
                    Case 7
                        Return "Tapadora"
                    Case 8
                        Return "Non-Equip"
                    Case 9
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "Mantenimiento"
                    Case 3
                        Return "Meeting"
                    Case 4
                        Return "Initiatives"
                    Case 5
                        Return "Projects"
                    Case 6
                        Return "Material Resupply"
                    Case 7
                        Return "Procedimiento calidad"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getNaucalpanPHC_MexprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Wrap Ade"
                    Case 2
                        Return "Non-Equip"
                    Case 3
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "Mantenimiento"
                    Case 3
                        Return "Meeting"
                    Case 4
                        Return "Initiatives"
                    Case 5
                        Return "Projects"
                    Case 6
                        Return "Material Resupply"
                    Case 7
                        Return "Procedimiento calidad"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getNaucalpanPHC_Vita1prStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Enflex"
                    Case 2
                        Return "Non-Equip"
                    Case 3
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "Mantenimiento"
                    Case 3
                        Return "Meeting"
                    Case 4
                        Return "Initiatives"
                    Case 5
                        Return "Projects"
                    Case 6
                        Return "Material Resupply"
                    Case 7
                        Return "Procedimiento calidad"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function

    Private Function getSwingRoadprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equipment"
                    Case 2
                        Return "Materials"
                    Case 3
                        Return "Bulk"
                    Case 4
                        Return "Utilities"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Cleaning"
                    Case 5
                        Return "Materials"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "AM Work"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getSwingRoadprStoryCardFixedFieldName_6(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Thermoformer"
                    Case 2
                        Return "Materials"
                    Case 3
                        Return "Equip-Other"
                    Case 4
                        Return "Utilities"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Cleaning"
                    Case 5
                        Return "Splice"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "AM Work"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getSwingRoadprStoryCardFixedFieldName_7(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Blisterformer"
                    Case 2
                        Return "Materials"
                    Case 3
                        Return "Equip-Other"
                    Case 4
                        Return "Utilities"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Cleaning"
                    Case 5
                        Return "Splice"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "AM Work"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function

    Private Function getPhoenixprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equipment"
                    Case 2
                        Return "Material"
                    Case 3
                        Return "Utilities"
                    Case 4
                        Return "Powder"
                    Case 5
                        Return OTHERS_STRING

                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Changeover"
                    Case 2
                        Return "Maintenance"
                    Case 3
                        Return "CIL"
                    Case 4
                        Return "AM"
                    Case 5
                        Return "SU/SD"
                    Case 6
                        Return "Meeting/Training"
                    Case 7
                        Return OTHERS_STRING
                    Case 8
                        Return ""

                End Select
            Case prStoryCard.Materials
                Select Case fieldNumber
                    Case 0
                        Return "Availability"
                    Case 1
                        Return "Quality"
                End Select
            Case prStoryCard.Bulk
                Select Case fieldNumber
                    Case 0
                        Return "Availability"
                    Case 1
                        Return "Quality"
                End Select
        End Select

        Return ""
    End Function


#End Region

#Region "Beauty"
    Private Function getSingaporePioneer_prstoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "EOL"
                    Case 2
                        Return "Resources"
                    Case 3
                        Return "Shubham"
                    Case 4
                        Return "Filler"
                    Case 5
                        Return "Utilities"
                    Case 6
                        Return "Materials"
                    Case 7
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Changeover"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Startup/Shutdown"
                    Case 4
                        Return "DDS"
                    Case 5
                        Return "PM"
                    Case 6
                        Return "Induced Stop"
                    Case 7
                        Return "Splice"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        Return ""
    End Function

    Private Function getAPDO_I_prStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Filler"
                    Case 2
                        Return "Twins"
                    Case 3
                        Return "F. Sorter"
                    Case 4
                        Return "Labeler"
                    Case 5
                        Return "Trimmer"
                    Case 6
                        Return "Casepacker"
                    Case 7
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "Flush"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Train/Meet"
                    Case 5
                        Return "Materials"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "Projects"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function
    Private Function getAPDO_J_prStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Filler"
                    Case 2
                        Return "Labeler"
                    Case 3
                        Return "Bundler"
                    Case 4
                        Return "Cap Sticker"
                    Case 5
                        Return "Case Packer"
                    Case 6
                        Return "Case Code Dater"
                    Case 7
                        Return "Case CW"
                    Case 8
                        Return "Bulk Supply"
                    Case 9
                        Return "EOL"
                    Case 10
                        Return "Material Quality"
                    Case 11
                        Return "Supply Losses"
                    Case 12
                        Return "Product Supply"
                    Case 13
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "Flush"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Train/Meet"
                    Case 5
                        Return "Materials"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "Projects"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function

    Private Function getSkinCareprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Capper"
                    Case 2
                        Return "Cartoner"
                    Case 3
                        Return "Bundler"
                    Case 4
                        Return "Labeler"
                    Case 5
                        Return "Filler"
                    Case 6
                        Return "End Of Line"
                    Case 7
                        Return "Materials"
                    Case 8
                        Return "Utilities"
                    Case 9
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Train/Meet"
                    Case 5
                        Return "Materials"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "Projects"
                    Case 8
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Materials
                Select Case fieldNumber
                    Case 0
                        Return "Availability"
                    Case 1
                        Return "Quality"
                End Select
            Case prStoryCard.Bulk
                Select Case fieldNumber
                    Case 0
                        Return "Availability"
                    Case 1
                        Return "Quality"
                End Select
        End Select
        '  
        Return ""
    End Function
    Private Function getIowaCityBeautyprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equipment"
                    Case 2
                        Return "Materials"
                    Case 3
                        Return "Product Delivery"
                    Case 4
                        Return "Utilities"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Split PO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Washout"
                    Case 5
                        Return "Purge"
                    Case 6
                        Return "Pause"
                    Case 7
                        Return "AM"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select

        Return ""
    End Function

    Private Function getMariscalaprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equipo"
                    Case 2
                        Return "Materiales"
                    Case 3
                        Return "Intellifeed"
                    Case 4
                        Return "Sistemas"
                    Case 5
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Cambios"
                    Case 2
                        Return "AM"
                    Case 3
                        Return "Planeadas"
                    Case 4
                        Return "Train/Meet"
                    Case 5
                        Return "Materials"
                    Case 6
                        Return "SU/SD"
                    Case 7
                        Return "Projects"
                    Case 8
                        Return OTHERS_STRING
                End Select
        End Select
        ' 
        Return ""
    End Function

    Private Function getHuangpuprStoryCardFixedFieldName(prStoryCardNum As Integer, fieldNumber As Integer) As String
        Select Case prStoryCardNum
            Case prStoryCard.Unplanned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "Equipment"
                    Case 2
                        Return "Materials"
                    Case 3
                        Return "Utilities"
                    Case 4
                        Return OTHERS_STRING
                End Select
            Case prStoryCard.Planned
                Select Case fieldNumber
                    Case 0
                        Return "Total"
                    Case 1
                        Return "CO"
                    Case 2
                        Return "CIL"
                    Case 3
                        Return "Maintenance"
                    Case 4
                        Return "Train/Meet"
                    Case 5
                        Return "Materials"
                    Case 6
                        Return OTHERS_STRING
                End Select
        End Select
        '  
        Return ""
    End Function

#End Region

End Module
