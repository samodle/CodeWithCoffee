Imports System.Windows.Forms
Imports System.Collections.ObjectModel
Imports System.Threading
Module Export_CSV

   public sub CSV_exportString(output as string, fileName as string)
        Dim appPath As String
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop 'lg
        dialog.SelectedPath = "C:\"
        dialog.Description = "Select Path To Save prStory Raw Data Files"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            appPath = dialog.SelectedPath
            fileName = appPath & "\" & fileName & ".csv"

            dim fsT = CreateObject("ADODB.Stream")
            fsT.Type = 2 'Specify stream type - we want To save text/string data.
            fsT.Charset = "utf-8" 'Specify charset For the source text data.
            fsT.Open() 'Open the stream And write binary data To the object

            fst.writetext(output)
            Try
                fsT.SaveToFile(fileName, 2) 'Save binary data To disk
            Catch ex As Exception
                MsgBox("A file with same name is open. Please close that file and export again.")
                fsT = Nothing
                Exit Sub

            End Try

            fsT = Nothing
            'show the folder
            Process.Start(appPath)
        End If
   End sub

#Region "Export Raw Data From prstory Objects"
    'Exporting Lists Of Raw Data
    'Exporting Raw Data From Main Menu
    Public Sub CSV_exportRawLEDsDataFromList(parentLine As ProdLine, dataList As ObservableCollection(Of DowntimeEvent), startTime As Date, endTime As Date, fieldName As String)
        Dim appPath As String
        Dim dialog As New FolderBrowserDialog()

        Dim fsT As Object
        Dim fileName As String, row As Integer
        Dim tempcomment As String
        Dim commaposition As Integer
        dialog.RootFolder = Environment.SpecialFolder.Desktop 'lg
        dialog.SelectedPath = "C:\"
        dialog.Description = "Select Path To Save prStory Raw Data Files"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            appPath = dialog.SelectedPath
            fileName = appPath & "\" & parentLine.Name & "_RawDowntime_" & startTime.Day & startTime.Month & startTime.Year & "_" & endTime.Day & endTime.Month & endTime.Year & "_" & fieldName & ".csv"


            fsT = CreateObject("ADODB.Stream")
            fsT.Type = 2 'Specify stream type - we want To save text/string data.
            fsT.Charset = "utf-8" 'Specify charset For the source text data.
            fsT.Open() 'Open the stream And write binary data To the object

            fsT.WriteText(parentLine.Name & ",Start Time, End Time, Downtime, Uptime, Location, Fault, " & parentLine.Reason1Name & ", " & parentLine.Reason2Name & ", " & parentLine.Reason3Name & ", " & parentLine.Reason4Name & ", Tier 1, Tier 2, Tier 3, DTsched, DTplanned, DTgroup, Description, GCAS, Comments" & vbCrLf)
            'actually export the data
            For row = 0 To dataList.Count - 1
                With dataList(row)
                    tempcomment = .Comment


                    'remove commas
                    While InStr(tempcomment, ",") > 0
                        commaposition = InStr(tempcomment, ",")
                        tempcomment = tempcomment.Remove(commaposition - 1, 1)
                    End While

                    While InStr(tempcomment, vbCrLf) > 0
                        commaposition = InStr(tempcomment, vbCrLf)
                        tempcomment = tempcomment.Remove(commaposition - 1, 2)
                    End While


                    fsT.WriteText("," & .startTime & "," & .endTime & "," & .DT & "," & .UT & "," & .Location & "," & .Fault & "," & .Reason1 & "," & .Reason2 & "," & .Reason3 & "," & .Reason4 & "," & .Tier1 & "," & .Tier2 & "," & .Tier3 & "," & .PR_inout & "," & .PlannedUnplanned & "," & .DTGroup & "," & .Product & "," & .ProductCode & "," & tempcomment)
                    fsT.WriteText(vbCrLf)
                End With
            Next row
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

    Public Sub CSV_exportRawProdDataFromList(parentLine As ProdLine, dataList As ObservableCollection(Of ProductionEvent), startTime As Date, endTime As Date, fieldName As String)
        Dim appPath As String
        Dim dialog As New FolderBrowserDialog()

        Dim fsT As Object
        Dim fileName As String, row As Integer


        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Select Path To Save prStory Raw Data Files"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            appPath = dialog.SelectedPath
            fileName = appPath & "\" & parentLine.Name & "_RawProduction_" & startTime.Day & startTime.Month & startTime.Year & "_" & endTime.Day & endTime.Month & endTime.Year & "_" & fieldName & ".csv"


            fsT = CreateObject("ADODB.Stream")
            fsT.Type = 2 'Specify stream type - we want To save text/string data.
            fsT.Charset = "utf-8" 'Specify charset For the source text data.
            fsT.Open() 'Open the stream And write binary data To the object

            fsT.WriteText(parentLine.Name & ",Start Time, End Time, PR, Uptime,  Line, Actual Rate, Target Rate, Description, SKU, PR In/Out , Actual Cases, Stat Units, Adjusted Cases, Team,, Adjusted Units, Units/Case" & vbCrLf)
            'actually export the data
            For row = 0 To dataList.Count - 1
                With dataList(row)
                    fsT.WriteText("," & .startTime & "," & .endTime & "," & .PR & "," & .UT & "," & .MasterProductionUnit & "," & .ActualRate & "," & .TargetRate & "," & .Product & "," & .ProductCode & "," & .PR_inout & "," & .ActCases & "," & .StatUnits & "," & .AdjCases & "," & .Team & ", ," & .AdjUnits & "," & .UnitsPerCase)
                    fsT.WriteText(vbCrLf)
                End With
            Next row
            'fin
            Try
                fsT.SaveToFile(fileName, 2) 'Save binary data To disk
            Catch ex As Exception
                MsgBox("A file with same name is open. Please close that file and export again.")

                fsT = Nothing
                Exit Sub
            End Try
            fsT = Nothing
            'show the folder
            Process.Start(appPath)
        End If
    End Sub
#End Region

#Region "Export Raw Data From array"
    'Exporting subsets of raw data
    Public Sub CSV_exportRawLEDsData(parentLine As ProdLine, prReport As prStoryMainPageReport, exportDowntimeData As Boolean, exportProductionData As Boolean, exportReDEPdata As Boolean, exportLossTree As Boolean)
        Dim appPath As String, startTime As String, endTime As String, reArr As Array
        ' exportPath = My.Settings.CSV_ExportPath
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Select Path To Save prStory Raw Data Files"
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            appPath = dialog.SelectedPath
            startTime = parentLine.rawProfStartTime.Day & Month(parentLine.rawProfStartTime) & Year(parentLine.rawProfStartTime)
            endTime = parentLine.rawProfEndTime.Day & Month(parentLine.rawProfEndTime) & Year(parentLine.rawProfEndTime)
            If exportDowntimeData Then exportDowntimeAsCSV(parentLine.rawProficyData, appPath & "\" & parentLine.Name & "_" & "DOWNTIME_" & startTime & "_" & endTime & ".csv", parentLine)
            If exportProductionData Then exportProductionAsCSV(parentLine.rawProficyProductionData, appPath & "\" & parentLine.Name & "_" & "PRODUCTION_" & startTime & "_" & endTime & ".csv", parentLine.Name)
            If exportReDEPdata Then
                reArr = RE_getDependencyAnalysisForLine(selectedindexofLine_temp)
                exportREDEPAsCSV(reArr, appPath & "\" & parentLine.Name & "_" & "REDependencyAnalysis_" & startTime & "_" & endTime & ".csv", parentLine.Name)
            End If
            If exportLossTree Then ExcelExportLossTree(appPath & "\" & parentLine.Name & "_", prReport)
            Process.Start(appPath)
        End If
    End Sub

    Private Sub exportDowntimeAsCSV(arr As Array, targetPath As String, parentLine As ProdLine) 'myArray As Array)
        Dim fsT As Object
        Dim fileName As String
        Dim row As Integer, col As Integer
        Dim colArr(0 To 15) As Integer
        Dim lineName As String
        lineName = parentLine.Name
        fileName = targetPath 'error hider
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        If parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.OneClick  Or parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.Maple Then
            fsT.WriteText(lineName & ",Start Time, End Time, Downtime, Uptime, Location, Fault, " & parentLine.Reason1Name & ", " & parentLine.Reason2Name & ", " & parentLine.Reason3Name & ", " & parentLine.Reason4Name & ", Team, Shift, Product Description, GCAS, DTPlanned, DTGroup, DTSched, Comment" & vbCrLf)
            colArr = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 16, 17, 19, 20}
        ElseIf parentLine.SQLdowntimeProcedure = DefaultProficyDowntimeProcedure.GLEDS Then
            fsT.WriteText(lineName & ",Start Time, End Time, Downtime, Uptime, Location, Fault, " & parentLine.Reason1Name & ", " & parentLine.Reason2Name & ", " & parentLine.Reason3Name & ", " & parentLine.Reason4Name & ", Team, Shift, DTGroup, DTPlanned, DTSched, Product, Comment" & vbCrLf)
            colArr = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 15, 16, 17, 18, 20}
        Else
            fsT.WriteText(lineName & ",Start Time, End Time, Downtime, Uptime, Location, Fault, " & parentLine.Reason1Name & ", " & parentLine.Reason2Name & ", " & parentLine.Reason3Name & ", " & parentLine.Reason4Name & ", DTsched, DTplanned, DTgroup, Description, GCAS, Comment" & vbCrLf)
            colArr = {0, 1, 2, 3, 5, 6, 7, 8, 9, 10, 12, 15, 16, 19, 20, 22}

        End If



        For row = 0 To arr.GetLength(1) - 1
            For col = 0 To colArr.Length - 1
                fsT.WriteText("," & arr(colArr(col), row))
            Next col
            fsT.WriteText(vbCrLf)
        Next row

        'fin
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing
    End Sub

    Private Sub exportProductionAsCSV(arr As Array, targetPath As String, lineName As String) 'myArray As Array)
        Dim fsT As Object
        Dim fileName As String
        Dim row As Integer, col As Integer
        Dim colArr(0 To 15) As Integer
        fileName = targetPath 'error hider
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        fsT.WriteText(lineName & ",Start Time, End Time" & vbCrLf) ', a, b, c, d, Reason 1, Reason 2, Reason 3, Reason 4, DTsched") ', DTplanned, DTgroup" & vbCrLf)
        ' colArr = {0, 1, 2, 3, 5, 6, 7, 8, 9, 10, 12} ', 15, 16, 17, 19, 20}
        colArr = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}
        For row = 0 To arr.GetLength(1) - 1
            For col = 0 To arr.GetLength(0) - 1 'colArr.Length - 1 '15
                fsT.WriteText("," & arr(col, row)) 'arr(colArr(col), row))
            Next col
            fsT.WriteText(vbCrLf)
        Next row

        'fin
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing
    End Sub

    'export all raw data
    Public Sub CSV_exportAllRawDataFromArrayObject(ByVal parentLine As ProdLine)
        Dim exportThread As Thread
        Dim objectForThread(4) As Object
        exportThread = New Thread(AddressOf exportAllRaw_Threaded)

        objectForThread(0) = parentLine.Name
        objectForThread(1) = parentLine.rawProficyData
        objectForThread(2) = parentLine.rawProficyProductionData
        objectForThread(3) = parentLine.SiteName

        exportThread.Start(objectForThread)
    End Sub
    Private Sub exportAllRaw_Threaded(ByVal paramObj As Object)

        CleanTempFOlder()  'clear the raw data folder to save space!


        Dim fsT As Object
        Dim row As Integer, col As Integer
        Dim lineName As String = paramObj(0)
        Dim DTarr As Array = paramObj(1)
        Dim PRODarr As Array = paramObj(2)
        Dim siteName As String = paramObj(3)

        '''''''''''''''''''''''''''''''''''''''
        'First, Lets get the Downtime
        '''''''''''''''''''''''''''''''''''''''

        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        For row = 0 To DTarr.GetLength(1) - 1
            For col = 0 To DTarr.GetLength(0) - 1
                fsT.WriteText("," & DTarr(col, row))
            Next col
            fsT.WriteText(vbCrLf)
        Next row

        'fin
        Try
            fsT.SaveToFile(PATH_PRSTORY_RAWDATA & "DT_" & siteName & "_" & lineName & ".csv", 2) 'Save binary data To disk
        Catch e As Exception
            MessageBox.Show("Error Writing Downtime To Local Disk")
        End Try
        ' fsT = Nothing

        '''''''''''''''''''''''''''''''''''''''
        'Now, Lets get the Production
        '''''''''''''''''''''''''''''''''''''''
        If Not My.Settings.AdvancedSettings_isAvailabilityMode Then
            fsT = CreateObject("ADODB.Stream")
            fsT.Type = 2 'Specify stream type - we want To save text/string data.
            fsT.Charset = "utf-8" 'Specify charset For the source text data.
            fsT.Open() 'Open the stream And write binary data To the object

            For row = 0 To PRODarr.GetLength(1) - 1
                For col = 0 To PRODarr.GetLength(0) - 1
                    fsT.WriteText("," & PRODarr(col, row))
                Next col
                fsT.WriteText(vbCrLf)
            Next row

            'fin
            fsT.SaveToFile(PATH_PRSTORY_RAWDATA & "PROD_" & siteName & "_" & lineName & ".csv", 2) 'Save binary data To disk
        End If
        fsT = Nothing
    End Sub
    Private Sub CleanTempFOlder()
        Try


            For Each foundFile As String In My.Computer.FileSystem.GetFiles(
    PATH_PRSTORY_RAWDATA, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.*")

                My.Computer.FileSystem.DeleteFile(foundFile,
                    Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                    Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently)
            Next
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
#End Region

#Region "Export Dependency Analysis"
    'Exporting RE Dependency Analysis
    Public Function RE_getDependencyAnalysisForLine(lineIndex As Integer)
        Dim depArray As Array
        '
        'Dim mappingCol As Integer = My.Settings.defaultMappingLevel
        Dim tmpList As New List(Of String)
        With AllProdLines(lineIndex).rawDowntimeData
            For i = 0 To .UnplannedData.Count - 1
                tmpList.Add(.UnplannedData(i).MappedField)
            Next
            depArray = executeDependencyAnalysis(tmpList.ToArray)
        End With
        Return depArray
    End Function
    Private Sub exportREDEPAsCSV(arr As Array, targetPath As String, lineName As String) 'myArray As Array)
        Dim fsT As Object
        Dim fileName As String
        Dim row As Integer, col As Integer
        Dim colArr(0 To 15) As Integer
        fileName = targetPath 'error hider
        fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2 'Specify stream type - we want To save text/string data.
        fsT.Charset = "utf-8" 'Specify charset For the source text data.
        fsT.Open() 'Open the stream And write binary data To the object

        'fileName = "path.csv"

        fsT.WriteText(lineName & vbCrLf)
        For col = 0 To arr.GetLength(0) - 1
            For row = 0 To arr.GetLength(1) - 1
                fsT.WriteText("," & arr(col, row))
            Next row
            fsT.WriteText(vbCrLf)
        Next col

        'fin
        fsT.SaveToFile(fileName, 2) 'Save binary data To disk
        fsT = Nothing
    End Sub
#End Region

End Module
