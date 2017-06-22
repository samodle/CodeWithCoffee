Module RE_Dependency
   

    Function executeDependencyAnalysis(mappedData() As String)
        Dim x_expect, pois, lambda, pois1, y, LogY, rank As Double
        Dim x_total_r, x_rows_r, Check, min_take, x_total_i_r, x_total_j_r As Double
        Dim alpha As Single
        Dim eLog As Double
        Dim x_rows, i, x_max, n_chance, j, i_row, j_col, m, k, i_list As Long
        Dim yExtend As Long
        Dim neg_i_list As Long
        Dim pos_i_list As Long
        Dim x_list_old, x_list, x_actual, x_total, x_total_i, x_total_j, r, LastRow As Long
        Dim text_field, text_field_test As String
        Dim i_code, i_flag, j_code As Long
        Dim sum_actual, sum_expected As Double
        Dim neg_sum_actual, neg_sum_expected As Double
        Dim pos_sum_actual, pos_sum_expected As Double

        'Dim RawData(0 To 3005, 0 To 1) As String
        Dim RawData(,) As String
        Dim Codes() As Long
        Dim Texts() As String
        Dim Counts(,) As Long
        Dim Totals(,) As Double

        Const DATA_COL As Integer = 2
        Const TABLE_COL As Integer = DATA_COL + 5
        Const RESULT_COL As Integer = 5
        Const RESULT_ROW As Integer = 66
        Const MAP_COL As Integer = DATA_COL + 1
        Const DEP_TABLE_ROW As Integer = 33
        Const DEP_TABLE_COL As Integer = 4

        Const Z As Integer = 3
        Const LAG As Integer = 1
        Const MINIM = 0

        'Dim excelSheet(0 To 4100, 0 To 200, 0 To 2)
        Dim excelSheet(,,)
        Const DEPENDENCY_ANALYSIS As Integer = 0
        Const RAW_DATA As Integer = 1
        Dim currentSheet As Integer
        Dim returnArray(,) 'As Array 'SRO
        currentSheet = RAW_DATA
        LastRow = mappedData.Length 'excelsheet(1, 1).End(xlDown).Row
        ReDim excelSheet(0 To LastRow + 1, 0 To 200, 0 To 2)
        ReDim RawData(0 To LastRow, 0 To 2)
        ReDim Codes(0 To LastRow)
        ReDim Texts(0 To LastRow)

        'copy the data into the "sheet" -sro
        For i = 1 To LastRow
            excelSheet(i, 1, RAW_DATA) = mappedData(i - 1)
        Next
        'end sro

        sum_actual = 0
        sum_expected = 0
        neg_sum_actual = 0
        neg_sum_expected = 0
        neg_i_list = 0
        pos_sum_actual = 0
        pos_sum_expected = 0
        pos_i_list = 0
        i = 1
        i_code = 0
        text_field = excelSheet(i, DATA_COL - 1, currentSheet)
        Do While text_field <> ""
            RawData(i, 1) = text_field
            i_flag = 1
            If i_code > 0 Then
                For j = 2 To i_code + 1
                    text_field_test = Texts(j - 1)
                    If text_field_test = text_field Then
                        i_flag = 0
                        j_code = j
                        j = i_code + 1
                    End If
                Next j
            End If
            If i_flag = 1 Then
                i_code = i_code + 1
                Codes(i_code) = i_code
                Texts(i_code) = text_field
                j_code = i_code + 1
            End If
            RawData(i, 2) = Codes(j_code - 1)
            i = i + 1
            text_field = excelSheet(i, DATA_COL - 1, currentSheet)
        Loop
        i_list = 1
        x_rows = i - 1
        x_max = i_code

        'Sheets("Dependency_Analysis").Activate()
        currentSheet = DEPENDENCY_ANALYSIS
        excelSheet(23, 7, currentSheet) = x_rows
        ' excelsheet(14, 7).Value = MAX_FAILURES
        alpha = NORMSDIST(-Z)

        excelSheet(24, 7, currentSheet) = x_max
        excelSheet(25, 7, currentSheet) = x_rows / x_max

        ReDim Counts(0 To x_max + 1, 0 To x_max)
        ReDim Totals(0 To x_max, 0 To 2)

        ' Build the table of data of actual counts
        'Sheets("Raw_Data").Activate()
        currentSheet = RAW_DATA
        n_chance = 0
        For i = 1 To x_max
            i_row = i
            For j = 1 To x_max
                j_col = TABLE_COL + j - 1
                Counts(i, j) = 0
            Next j
        Next i

        For i = LAG + 1 To x_rows
            excelSheet(x_list_old, TABLE_COL + x_list - 1, currentSheet) = x_actual + 1

            x_list_old = RawData(i - LAG, 2)
            x_list = RawData(i, 2)
            x_actual = Counts(x_list_old, x_list)
            Counts(x_list_old, x_list) = x_actual + 1
        Next i

        For j = 1 To x_max

            x_total = 0
            For i = 1 To x_max
                x_total = x_total + Counts(i, j)
            Next i
            Counts(x_max + 1, j) = x_total
        Next j

        eLog = Math.Log(10)

        'First Pass, just look for Same-Same Dependencies
        For i = 1 To x_max
            j = i

            x_total_i = Counts(x_max + 1, i)
            x_total_j = Counts(x_max + 1, j)
            x_expect = (x_rows - LAG) * x_total_i * x_total_j / x_rows / x_rows
            x_actual = Counts(i, j)

            If ((x_expect >= 5) Or (x_actual >= 5)) And (x_total_i >= MINIM) And (x_total_j >= MINIM) Then
                pois = 0
                lambda = x_expect
                r = x_actual
                If lambda < 1 Then
                    pois1 = 0
                    For m = 0 To r
                        y = 1
                        yExtend = 0
                        If m > 0 Then
                            For k = 1 To m
                                LogY = Math.Log10(y)
                                If (Math.Abs(LogY) > 300) Then
                                    yExtend = yExtend + CLng(LogY)
                                    y = (10) ^ (LogY - CLng(LogY))
                                End If
                                y = y * lambda / k
                            Next k
                        End If
                        If (yExtend = 0) Then
                            pois1 = pois1 + y * Math.Exp(-lambda)
                        Else
                            LogY = Math.Log(y)
                            pois1 = pois1 + Math.Exp(LogY - lambda + yExtend * eLog)
                        End If
                    Next m
                    If pois1 > 1 Then
                        pois1 = 1
                    End If
                    n_chance = n_chance + (1 - pois1)
                    lambda = 1
                Else
                    n_chance = n_chance + alpha
                End If
                For m = 0 To r
                    y = 1
                    yExtend = 0
                    If m > 0 Then
                        For k = 1 To m
                            LogY = Math.Log10(y)
                            If (Math.Abs(LogY) > 300) Then
                                yExtend = yExtend + CLng(LogY)
                                y = (10) ^ (LogY - CLng(LogY))
                            End If
                            y = y * lambda / k
                        Next k
                    End If
                    If (yExtend = 0) Then
                        pois = pois + y * Math.Exp(-lambda)
                    Else
                        LogY = Math.Log(y)
                        pois = pois + Math.Exp(LogY - lambda + yExtend * eLog)
                    End If
                Next m
                If pois > 1 Then
                    pois = 1
                End If
                rank = Math.Abs(pois - 0.5) + 0.5
                If rank >= 1 - alpha Then
                    'Sheets("Dependency_Analysis").Activate()
                    currentSheet = DEPENDENCY_ANALYSIS
                    If excelSheet(i + 1, MAP_COL + 1, RAW_DATA) = excelSheet(j + 1, MAP_COL + 1, RAW_DATA) Then
                        excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL, currentSheet) = "Same-Same"
                    Else
                        excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL, currentSheet) = "Not-Same"
                    End If

                    excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 1, currentSheet) = Texts(i)

                    excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 2, currentSheet) = Texts(j)
                    excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 3, currentSheet) = x_actual - x_expect
                    excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 4, currentSheet) = 100.0# * (x_actual - x_expect) / x_rows
                    excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 5, currentSheet) = x_actual
                    excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 6, currentSheet) = 100.0# * x_actual / x_rows
                    excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 7, currentSheet) = x_expect

                    'Sheets("Raw_Data").Activate()
                    currentSheet = RAW_DATA
                    sum_actual = sum_actual + x_actual
                    sum_expected = sum_expected + x_expect
                    i_list = i_list + 1
                    If x_actual < x_expect Then
                        neg_sum_actual = neg_sum_actual + x_actual
                        neg_sum_expected = neg_sum_expected + x_expect
                        neg_i_list = neg_i_list + 1
                    Else
                        pos_sum_actual = pos_sum_actual + x_actual
                        pos_sum_expected = pos_sum_expected + x_expect
                        pos_i_list = pos_i_list + 1
                    End If
                End If
            End If
        Next i

        'Sheets("Dependency_Analysis").Activate()
        currentSheet = DEPENDENCY_ANALYSIS

        excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL, currentSheet) = pos_i_list
        excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 2, currentSheet) = x_max
        excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 4, currentSheet) = pos_sum_actual
        excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 5, currentSheet) = pos_sum_expected
        excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 6, currentSheet) = pos_sum_actual - pos_sum_expected
        excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 8, currentSheet) = 100 * excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 6, currentSheet) / x_rows

        excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL, currentSheet) = neg_i_list
        excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 2, currentSheet) = x_max
        excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 4, currentSheet) = neg_sum_actual
        excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 5, currentSheet) = neg_sum_expected
        excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 6, currentSheet) = neg_sum_actual - neg_sum_expected
        excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 8, currentSheet) = 100 * Math.Abs(excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 6, currentSheet)) / x_rows


        ' Sheets("Raw_Data").Activate()
        currentSheet = RAW_DATA
        'Replace all Same-Same pairs in the table with counts needed to make the Actual=Expected.
        'Use a convergence algorithm to do this
        For i = 1 To x_max
            Totals(i, 2) = 0

        Next i

        x_rows_r = 0
        For i = 1 To x_max
            i_row = i
            x_total_r = 0
            For j = 1 To x_max
                j_col = TABLE_COL + j - 1
                If i <> j Then

                    x_total_r = x_total_r + Counts(i, j)
                Else

                    x_total_r = x_total_r + Totals(i, 2)
                End If
            Next j

            Totals(i, 1) = x_total_r
            x_rows_r = x_rows_r + x_total_r
        Next i

        x_rows_r = x_rows_r + LAG
        Check = x_rows_r
        For k = 1 To 100
            For i = 1 To x_max

                x_total_i = Totals(i, 1)
                x_actual = (x_rows_r - LAG) * x_total_i * x_total_i / x_rows_r / x_rows_r

                Totals(i, 2) = x_actual
            Next i
            x_rows_r = 0
            For i = 1 To x_max
                i_row = i
                x_total_r = 0
                For j = 1 To x_max
                    j_col = TABLE_COL + j - 1
                    If i <> j Then

                        x_total_r = x_total_r + Counts(i, j)
                    Else

                        x_total_r = x_total_r + Totals(i, 2)
                    End If
                Next j

                Totals(i, 1) = x_total_r
                x_rows_r = x_rows_r + x_total_r
            Next i
            x_rows_r = x_rows_r + LAG
            If (Math.Abs(Check - x_rows_r) <= 0.0001) Then
                k = 100
            Else
                Check = x_rows_r
            End If
        Next k

        ' Now look for Not Same Dependencies
        min_take = x_rows_r
        For i = 1 To x_max
            For j = 1 To x_max
                If i <> j Then

                    x_total_i_r = Totals(i, 1)

                    x_total_j_r = Totals(j, 1)
                    x_expect = (x_rows_r - LAG) * x_total_i_r * x_total_j_r / x_rows_r / x_rows_r

                    x_actual = Counts(i, j)
                    If ((x_expect >= 5) Or (x_actual >= 5)) And (x_total_i_r >= MINIM) And (x_total_j_r >= MINIM) Then
                        pois = 0
                        lambda = x_expect
                        r = x_actual
                        If lambda < 1 Then
                            pois1 = 0
                            For m = 0 To r
                                y = 1
                                yExtend = 0
                                If m > 0 Then
                                    For k = 1 To m
                                        LogY = Math.Log10(y)
                                        If (Math.Abs(LogY) > 300) Then
                                            yExtend = yExtend + CLng(LogY)
                                            y = (10) ^ (LogY - CLng(LogY))
                                        End If
                                        y = y * lambda / k
                                    Next k
                                End If
                                If (yExtend = 0) Then
                                    pois1 = pois1 + y * Math.Exp(-lambda)
                                Else
                                    LogY = Math.Log(y)
                                    pois1 = pois1 + Math.Exp(LogY - lambda + yExtend * eLog)
                                End If
                            Next m
                            If pois1 > 1 Then
                                pois1 = 1
                            End If
                            n_chance = n_chance + (1 - pois1)
                            lambda = 1
                        Else
                            n_chance = n_chance + alpha
                        End If
                        For m = 0 To r
                            y = 1
                            yExtend = 0
                            If m > 0 Then
                                For k = 1 To m
                                    LogY = Math.Log10(y)
                                    If (Math.Abs(LogY) > 300) Then
                                        yExtend = yExtend + CLng(LogY)
                                        y = (10) ^ (LogY - CLng(LogY))
                                    End If
                                    y = y * lambda / k
                                Next k
                            End If
                            If (yExtend = 0) Then
                                pois = pois + y * Math.Exp(-lambda)
                            Else
                                LogY = Math.Log(y)
                                pois = pois + Math.Exp(LogY - lambda + yExtend * eLog)
                            End If
                        Next m
                        If pois > 1 Then
                            pois = 1
                        End If
                        rank = Math.Abs(pois - 0.5) + 0.5
                        If rank >= 1 - alpha Then
                            'Sheets("Dependency_Analysis").Activate()
                            currentSheet = DEPENDENCY_ANALYSIS
                            ' If Sheets("Raw_Data").Cells(i + 1, MAP_COL + 1).Value = Sheets("Raw_Data").Cells(j + 1, MAP_COL + 1).Value Then
                            If Texts(i) = Texts(j) Then
                                excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL, currentSheet) = "Same-Same"
                            Else
                                excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL, currentSheet) = "Not-Same"
                            End If

                            excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 1, currentSheet) = Texts(i)

                            excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 2, currentSheet) = Texts(j)
                            excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 3, currentSheet) = x_actual - x_expect
                            excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 4, currentSheet) = 100.0# * (x_actual - x_expect) / x_rows
                            excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 5, currentSheet) = x_actual
                            excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 6, currentSheet) = 100.0# * x_actual / x_rows
                            excelSheet(RESULT_ROW - 1 + i_list, RESULT_COL + 7, currentSheet) = x_expect
                            'Sheets("Raw_Data").Activate()
                            currentSheet = RAW_DATA
                            sum_actual = sum_actual + x_actual
                            sum_expected = sum_expected + x_expect
                            i_list = i_list + 1
                            If x_actual < x_expect Then
                                neg_sum_actual = neg_sum_actual + x_actual
                                neg_sum_expected = neg_sum_expected + x_expect
                                neg_i_list = neg_i_list + 1
                            Else
                                pos_sum_actual = pos_sum_actual + x_actual
                                pos_sum_expected = pos_sum_expected + x_expect
                                pos_i_list = pos_i_list + 1
                            End If
                            If x_total_i_r < min_take Then
                                min_take = x_total_i_r
                            End If
                            If x_total_j_r < min_take Then
                                min_take = x_total_j_r
                            End If
                        End If
                    End If
                End If
            Next j
        Next i
        'Sheets("Dependency_Analysis").Activate()
        currentSheet = DEPENDENCY_ANALYSIS
        excelSheet(27, 7, currentSheet) = n_chance

        excelSheet(DEP_TABLE_ROW + 2, DEP_TABLE_COL, currentSheet) = pos_i_list
        excelSheet(DEP_TABLE_ROW + 2, DEP_TABLE_COL + 2, currentSheet) = x_max * x_max
        excelSheet(DEP_TABLE_ROW + 2, DEP_TABLE_COL + 4, currentSheet) = pos_sum_actual
        excelSheet(DEP_TABLE_ROW + 2, DEP_TABLE_COL + 5, currentSheet) = pos_sum_expected
        excelSheet(DEP_TABLE_ROW + 2, DEP_TABLE_COL + 6, currentSheet) = pos_sum_actual - pos_sum_expected
        excelSheet(DEP_TABLE_ROW + 2, DEP_TABLE_COL + 8, currentSheet) = 100 * excelSheet(DEP_TABLE_ROW + 2, DEP_TABLE_COL + 6, currentSheet) / x_rows
        excelSheet(DEP_TABLE_ROW + 1, DEP_TABLE_COL, currentSheet) = pos_i_list - excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL, currentSheet)
        excelSheet(DEP_TABLE_ROW + 1, DEP_TABLE_COL + 2, currentSheet) = x_max * x_max - x_max
        excelSheet(DEP_TABLE_ROW + 1, DEP_TABLE_COL + 4, currentSheet) = pos_sum_actual - excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 4, currentSheet)
        excelSheet(DEP_TABLE_ROW + 1, DEP_TABLE_COL + 5, currentSheet) = pos_sum_expected - excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 5, currentSheet)
        excelSheet(DEP_TABLE_ROW + 1, DEP_TABLE_COL + 6, currentSheet) = pos_sum_actual - pos_sum_expected - excelSheet(DEP_TABLE_ROW, DEP_TABLE_COL + 6, currentSheet)
        excelSheet(DEP_TABLE_ROW + 1, DEP_TABLE_COL + 8, currentSheet) = 100 * excelSheet(DEP_TABLE_ROW + 1, DEP_TABLE_COL + 6, currentSheet) / x_rows

        excelSheet(DEP_TABLE_ROW + 8, DEP_TABLE_COL, currentSheet) = neg_i_list
        excelSheet(DEP_TABLE_ROW + 8, DEP_TABLE_COL + 2, currentSheet) = x_max * x_max
        excelSheet(DEP_TABLE_ROW + 8, DEP_TABLE_COL + 4, currentSheet) = neg_sum_actual
        excelSheet(DEP_TABLE_ROW + 8, DEP_TABLE_COL + 5, currentSheet) = neg_sum_expected
        excelSheet(DEP_TABLE_ROW + 8, DEP_TABLE_COL + 6, currentSheet) = neg_sum_actual - neg_sum_expected
        excelSheet(DEP_TABLE_ROW + 8, DEP_TABLE_COL + 8, currentSheet) = 100 * Math.Abs(excelSheet(DEP_TABLE_ROW + 8, DEP_TABLE_COL + 6, currentSheet)) / x_rows
        excelSheet(DEP_TABLE_ROW + 7, DEP_TABLE_COL, currentSheet) = neg_i_list - excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL, currentSheet)
        excelSheet(DEP_TABLE_ROW + 7, DEP_TABLE_COL + 2, currentSheet) = x_max * x_max - x_max
        excelSheet(DEP_TABLE_ROW + 7, DEP_TABLE_COL + 4, currentSheet) = neg_sum_actual - excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 4, currentSheet)
        excelSheet(DEP_TABLE_ROW + 7, DEP_TABLE_COL + 5, currentSheet) = neg_sum_expected - excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 5, currentSheet)
        excelSheet(DEP_TABLE_ROW + 7, DEP_TABLE_COL + 6, currentSheet) = neg_sum_actual - neg_sum_expected - excelSheet(DEP_TABLE_ROW + 6, DEP_TABLE_COL + 6, currentSheet)
        excelSheet(DEP_TABLE_ROW + 7, DEP_TABLE_COL + 8, currentSheet) = 100 * Math.Abs(excelSheet(DEP_TABLE_ROW + 7, DEP_TABLE_COL + 6, currentSheet)) / x_rows

        'LETS TRY TO GET SOME OUTPUT
        currentSheet = DEPENDENCY_ANALYSIS
        i = 66 'To 106
        j = 0
        While excelSheet(i, 5, currentSheet) = "Same-Same" Or excelSheet(i, 5, currentSheet) = "Not-Same"
            j += 1
            i += 1
        End While
        ReDim returnArray(0 To j, 0 To 8)
        j = 1
        returnArray(0, 0) = "Dependency Type"
        returnArray(0, 1) = "Pre Stop Failure Mode"
        returnArray(0, 2) = "Post Stop Failure Mode"
        returnArray(0, 3) = "Act-Exp #"
        returnArray(0, 4) = "Act-Exp %"
        returnArray(0, 5) = "Act #"
        returnArray(0, 6) = "Act %"
        returnArray(0, 7) = "Exp #"
        returnArray(0, 8) = "Exp %"
        i = 66
        While excelSheet(i, 5, currentSheet) = "Same-Same" Or excelSheet(i, 5, currentSheet) = "Not-Same"
            returnArray(j, 0) = excelSheet(i, 5, currentSheet)
            returnArray(j, 1) = excelSheet(i, 6, currentSheet)
            returnArray(j, 2) = excelSheet(i, 7, currentSheet)
            returnArray(j, 3) = excelSheet(i, 8, currentSheet)
            returnArray(j, 4) = excelSheet(i, 9, currentSheet)
            returnArray(j, 5) = excelSheet(i, 10, currentSheet)
            returnArray(j, 6) = excelSheet(i, 11, currentSheet)
            returnArray(j, 7) = excelSheet(i, 12, currentSheet)
            returnArray(j, 8) = (excelSheet(i, 11, currentSheet) - excelSheet(i, 9, currentSheet))
            j += 1
            i += 1
        End While
        Return returnArray
    End Function

End Module


Public Class dependencyEvent2
#Region "Variables & Properties"
    Private _Name As String
    Private _Stops As Integer
    Private _SPD As Double


#End Region


End Class