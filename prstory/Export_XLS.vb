Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Collections.ObjectModel

Module Export_XLS
    Public Sub ExcelExportLossTree(ByVal savePath As String, prReport As prStoryMainPageReport)
        Dim rowR As Long, colC As Long ', i As Long
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim Tier1Incrementer As Integer, Tier2Incrementer As Integer, Tier3Incrementer As Integer, tmpTier2List As List(Of DTevent), tmpTier3list As List(Of DTevent)
        Dim netSchedTime As Double, netUptimeDT As Double

        Dim saveName As String = "LossTree"
        Dim firstRow As Integer, firstCol As Integer

        If xlApp Is Nothing Then
            Throw New genericMessageException("Excel is not properly installed") 'MessageBox.Show("Excel is not properly installed!!")
            Return
        End If

        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet, xlLossTreeSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim chartRange As Excel.Range

        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        xlLossTreeSheet = xlWorkBook.Sheets("sheet2")

        xlLossTreeSheet.Name = "LossTree"
        rowR = 2
        colC = 3

        With prReport
            netSchedTime = .MainLEDSReport.schedTime
            netUptimeDT = .MainLEDSReport.UT_DT

            'header row
            xlLossTreeSheet.Cells(rowR, colC + 1) = "Stops"
            xlLossTreeSheet.Cells(rowR, colC + 2) = "Stops/Day"
            xlLossTreeSheet.Cells(rowR, colC + 3) = "Total DT"
            If My.Settings.AdvancedSettings_isAvailabilityMode Then
                xlLossTreeSheet.Cells(rowR, colC + 4) = "Aval. Loss"
            Else
                xlLossTreeSheet.Cells(rowR, colC + 4) = "PR Loss"
            End If

            xlLossTreeSheet.Cells(rowR, colC + 5) = "MTTR"
            xlLossTreeSheet.Cells(rowR, colC + 6) = "MTBF"
            xlLossTreeSheet.Range(xlLossTreeSheet.Cells(rowR, colC), xlLossTreeSheet.Cells(rowR, colC + 6)).Font.Bold = True
            rowR += 1

            'header row
            firstRow = rowR
            firstCol = colC
            xlLossTreeSheet.Cells(rowR, colC) = "UNPLANNED LOSSES"
            xlLossTreeSheet.Range(xlLossTreeSheet.Cells(rowR, colC), xlLossTreeSheet.Cells(rowR, colC)).Font.Bold = True
            xlLossTreeSheet.Cells(rowR, colC + 1) = .ActualStops
            xlLossTreeSheet.Cells(rowR, colC + 2) = .StopsPerDay
            xlLossTreeSheet.Cells(rowR, colC + 3) = 0
            xlLossTreeSheet.Cells(rowR, colC + 4) = .UPDT
            xlLossTreeSheet.Cells(rowR, colC + 5) = 0
            xlLossTreeSheet.Cells(rowR, colC + 6) = .MTBF
            rowR += 1

            'raw data
            For Tier1Incrementer = 0 To .UnplannedList.Count - 1
                xlLossTreeSheet.Cells(rowR, colC) = .UnplannedList(Tier1Incrementer).Name
                xlLossTreeSheet.Cells(rowR, colC + 1) = .UnplannedList(Tier1Incrementer).Stops
                xlLossTreeSheet.Cells(rowR, colC + 2) = .UnplannedList(Tier1Incrementer).SPD
                xlLossTreeSheet.Cells(rowR, colC + 3) = .UnplannedList(Tier1Incrementer).DT
                xlLossTreeSheet.Cells(rowR, colC + 4) = .UnplannedList(Tier1Incrementer).DTpct
                xlLossTreeSheet.Cells(rowR, colC + 5) = .UnplannedList(Tier1Incrementer).MTTR
                xlLossTreeSheet.Cells(rowR, colC + 6) = netUptimeDT / .UnplannedList(Tier1Incrementer).Stops
                rowR += 1

                tmpTier2List = .MainLEDSReport.DT_Report.getTier2Directory(.UnplannedList(Tier1Incrementer).Name)
                For Tier2Incrementer = 0 To tmpTier2List.Count - 1
                    xlLossTreeSheet.Cells(rowR, colC) = tmpTier2List(Tier2Incrementer).Name
                    xlLossTreeSheet.Cells(rowR, colC + 1) = tmpTier2List(Tier2Incrementer).Stops
                    xlLossTreeSheet.Cells(rowR, colC + 2) = tmpTier2List(Tier2Incrementer).SPD
                    xlLossTreeSheet.Cells(rowR, colC + 3) = tmpTier2List(Tier2Incrementer).DT
                    xlLossTreeSheet.Cells(rowR, colC + 4) = tmpTier2List(Tier2Incrementer).DT / netSchedTime
                    xlLossTreeSheet.Cells(rowR, colC + 5) = tmpTier2List(Tier2Incrementer).MTTR
                    xlLossTreeSheet.Cells(rowR, colC + 6) = netUptimeDT / tmpTier2List(Tier2Incrementer).Stops
                    rowR += 1

                    tmpTier3list = .MainLEDSReport.DT_Report.getTier3Directory(.UnplannedList(Tier1Incrementer).Name, tmpTier2List(Tier2Incrementer).Name)
                    For Tier3Incrementer = 0 To tmpTier3list.Count - 1
                        xlLossTreeSheet.Cells(rowR, colC) = tmpTier3list(Tier3Incrementer).Name
                        xlLossTreeSheet.Cells(rowR, colC + 1) = tmpTier3list(Tier3Incrementer).Stops
                        xlLossTreeSheet.Cells(rowR, colC + 2) = tmpTier3list(Tier3Incrementer).SPD
                        xlLossTreeSheet.Cells(rowR, colC + 3) = tmpTier3list(Tier3Incrementer).DT
                        xlLossTreeSheet.Cells(rowR, colC + 4) = tmpTier3list(Tier3Incrementer).DT / netSchedTime
                        xlLossTreeSheet.Cells(rowR, colC + 5) = tmpTier3list(Tier3Incrementer).MTTR
                        xlLossTreeSheet.Cells(rowR, colC + 6) = netUptimeDT / tmpTier3list(Tier3Incrementer).Stops
                        rowR += 1
                    Next


                Next
            Next
        End With

        xlLossTreeSheet.Range(xlLossTreeSheet.Cells(firstRow, firstCol), xlLossTreeSheet.Cells(rowR, colC + 6)).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)

        'add data 
        xlWorkSheet.Cells(4, 2) = ""
        xlWorkSheet.Cells(4, 3) = "Student1"
        xlWorkSheet.Cells(4, 4) = "Student2"
        xlWorkSheet.Cells(4, 5) = "Student3"

        xlWorkSheet.Cells(9, 2) = "Total"
        xlWorkSheet.Cells(9, 3) = "315"
        xlWorkSheet.Cells(9, 4) = "299"
        xlWorkSheet.Cells(9, 5) = "238"

        chartRange = xlWorkSheet.Range("b2", "e3")
        chartRange.Merge()

        chartRange = xlWorkSheet.Range("b2", "e3")
        chartRange.FormulaR1C1 = "MARK LIST"
        chartRange.HorizontalAlignment = 3
        chartRange.VerticalAlignment = 3
        chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)
        chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        chartRange.Font.Size = 20

        'lets save this guy
        Try
            xlWorkBook.SaveAs(savePath & saveName & ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
             Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        Catch ex As System.Runtime.InteropServices.COMException
            xlWorkBook.SaveAs(savePath & saveName & "_" & Second(Now()) & ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        End Try
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Module
