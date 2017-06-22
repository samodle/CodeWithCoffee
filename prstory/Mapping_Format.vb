Module Mapping_Format
    Public Sub getSkinCareFormatMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .ProductCode.Equals("80245098") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80245096") Then
                .Format = "21"
            ElseIf .ProductCode.Equals("80245099") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80245101") Then
                .Format = "17"
            ElseIf .ProductCode.Equals("80245102") Then
                .Format = "18"
            ElseIf .ProductCode.Equals("80245104") Then
                .Format = "39"
            ElseIf .ProductCode.Equals("80245105") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80245106") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245108") Then
                .Format = "21"
            ElseIf .ProductCode.Equals("80245112") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80245115") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245116") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80245118") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245120") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80245123") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245127") Then
                .Format = "3"
            ElseIf .ProductCode.Equals("80245129") Then
                .Format = "17"
            ElseIf .ProductCode.Equals("80245136") Then
                .Format = "16"
            ElseIf .ProductCode.Equals("80245141") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245144") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80245145") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80245148") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245152") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80245154") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80245163") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245164") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80245166") Then
                .Format = "17"
            ElseIf .ProductCode.Equals("80245167") Then
                .Format = "21"
            ElseIf .ProductCode.Equals("80245168") Then
                .Format = "17"
            ElseIf .ProductCode.Equals("80251678") Then
                .Format = "76"
            ElseIf .ProductCode.Equals("80252959") Then
                .Format = "18"
            ElseIf .ProductCode.Equals("80252982") Then
                .Format = "18"
            ElseIf .ProductCode.Equals("80253061") Then
                .Format = "39"
            ElseIf .ProductCode.Equals("80253069") Then
                .Format = "39"
            ElseIf .ProductCode.Equals("80253391") Then
                .Format = "17"
            ElseIf .ProductCode.Equals("80255023") Then
                .Format = "4"
            ElseIf .ProductCode.Equals("80255030") Then
                .Format = "4"
            ElseIf .ProductCode.Equals("80255139") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80255140") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80255199") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80255200") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80255201") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80255202") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80255203") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80255211") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80255212") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80255213") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80255214") Then
                .Format = "34"
            ElseIf .ProductCode.Equals("80255216") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80255217") Then
                .Format = "13"
            ElseIf .ProductCode.Equals("80255219") Then
                .Format = "22"
            ElseIf .ProductCode.Equals("80255226") Then
                .Format = "4"
            ElseIf .ProductCode.Equals("80255227") Then
                .Format = "4"
            Else
                .Format = OTHERS_STRING
            End If
        End With
    End Sub

    Public Function getSkinCareFormatFromSku(ByVal SKUnumber As String) As String
        If SKUnumber.Equals("80245098") Then
            Return "34"
        ElseIf SKUnumber.Equals("80245096") Then
            Return "21"
        ElseIf SKUnumber.Equals("80245099") Then
            Return "13"
        ElseIf SKUnumber.Equals("80245101") Then
            Return "17"
        ElseIf SKUnumber.Equals("80245102") Then
            Return "18"
        ElseIf SKUnumber.Equals("80245104") Then
            Return "39"
        ElseIf SKUnumber.Equals("80245105") Then
            Return "34"
        ElseIf SKUnumber.Equals("80245106") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245108") Then
            Return "21"
        ElseIf SKUnumber.Equals("80245112") Then
            Return "13"
        ElseIf SKUnumber.Equals("80245115") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245116") Then
            Return "34"
        ElseIf SKUnumber.Equals("80245118") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245120") Then
            Return "13"
        ElseIf SKUnumber.Equals("80245123") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245127") Then
            Return "3"
        ElseIf SKUnumber.Equals("80245129") Then
            Return "17"
        ElseIf SKUnumber.Equals("80245136") Then
            Return "16"
        ElseIf SKUnumber.Equals("80245141") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245144") Then
            Return "13"
        ElseIf SKUnumber.Equals("80245145") Then
            Return "34"
        ElseIf SKUnumber.Equals("80245148") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245152") Then
            Return "13"
        ElseIf SKUnumber.Equals("80245154") Then
            Return "13"
        ElseIf SKUnumber.Equals("80245163") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245164") Then
            Return "22"
        ElseIf SKUnumber.Equals("80245166") Then
            Return "17"
        ElseIf SKUnumber.Equals("80245167") Then
            Return "21"
        ElseIf SKUnumber.Equals("80245168") Then
            Return "17"
        ElseIf SKUnumber.Equals("80251678") Then
            Return "76"
        ElseIf SKUnumber.Equals("80252959") Then
            Return "18"
        ElseIf SKUnumber.Equals("80252982") Then
            Return "18"
        ElseIf SKUnumber.Equals("80253061") Then
            Return "39"
        ElseIf SKUnumber.Equals("80253069") Then
            Return "39"
        ElseIf SKUnumber.Equals("80253391") Then
            Return "17"
        ElseIf SKUnumber.Equals("80255023") Then
            Return "4"
        ElseIf SKUnumber.Equals("80255030") Then
            Return "4"
        ElseIf SKUnumber.Equals("80255139") Then
            Return "34"
        ElseIf SKUnumber.Equals("80255140") Then
            Return "13"
        ElseIf SKUnumber.Equals("80255199") Then
            Return "13"
        ElseIf SKUnumber.Equals("80255200") Then
            Return "34"
        ElseIf SKUnumber.Equals("80255201") Then
            Return "22"
        ElseIf SKUnumber.Equals("80255202") Then
            Return "22"
        ElseIf SKUnumber.Equals("80255203") Then
            Return "34"
        ElseIf SKUnumber.Equals("80255211") Then
            Return "34"
        ElseIf SKUnumber.Equals("80255212") Then
            Return "22"
        ElseIf SKUnumber.Equals("80255213") Then
            Return "13"
        ElseIf SKUnumber.Equals("80255214") Then
            Return "34"
        ElseIf SKUnumber.Equals("80255216") Then
            Return "22"
        ElseIf SKUnumber.Equals("80255217") Then
            Return "13"
        ElseIf SKUnumber.Equals("80255219") Then
            Return "22"
        ElseIf SKUnumber.Equals("80255226") Then
            Return "4"
        ElseIf SKUnumber.Equals("80255227") Then
            Return "4"
        Else
            Return OTHERS_STRING
        End If
    End Function
End Module

Module Mapping_Shape
    'SKIN CARE
    Public Function getSkinCareShapeFromSku(ByVal SKUnumber As String) As String
        If SKUnumber.Equals("80245098") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245096") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245099") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80245101") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245102") Then
            Return "Pump"
        ElseIf SKUnumber.Equals("80245104") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80245105") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245106") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245108") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245112") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80245115") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245116") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245118") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245120") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80245123") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245127") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245129") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245136") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245141") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245144") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80245145") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245148") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245152") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80245154") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80245163") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245164") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245166") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245167") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80245168") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80251678") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80252959") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80252982") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80253061") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80253069") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80253391") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255023") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80255030") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80255139") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255140") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80255199") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80255200") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255201") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255202") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255203") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255211") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255212") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255213") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80255214") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255216") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255217") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80255219") Then
            Return "Bottle"
        ElseIf SKUnumber.Equals("80255226") Then
            Return "Jar"
        ElseIf SKUnumber.Equals("80255227") Then
            Return "Jar"
        Else
            Return OTHERS_STRING
        End If
    End Function

    Public Sub getSkinCareShapeMapping(ByRef searchEvent As DowntimeEvent)
        With searchEvent
            If .ProductCode.Equals("80245098") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245096") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245099") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80245101") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245102") Then
                .Shape = "Pump"
            ElseIf .ProductCode.Equals("80245104") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80245105") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245106") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245108") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245112") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80245115") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245116") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245118") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245120") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80245123") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245127") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245129") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245136") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245141") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245144") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80245145") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245148") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245152") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80245154") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80245163") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245164") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245166") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245167") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80245168") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80251678") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80252959") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80252982") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80253061") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80253069") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80253391") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255023") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80255030") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80255139") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255140") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80255199") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80255200") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255201") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255202") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255203") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255211") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255212") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255213") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80255214") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255216") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255217") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80255219") Then
                .Shape = "Bottle"
            ElseIf .ProductCode.Equals("80255226") Then
                .Shape = "Jar"
            ElseIf .ProductCode.Equals("80255227") Then
                .Shape = "Jar"
            Else
                .Shape = OTHERS_STRING
            End If
        End With
    End Sub

    'IOWA CITY
    ' Public Sub getIC()
    '  Public Function getICShapeMapping(ByVal SKUnumber As String) As String

    '  End Sub
End Module
