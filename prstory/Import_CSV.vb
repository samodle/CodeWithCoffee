Imports System.Data
Imports System.IO
Imports System.Reflection

Module Import_CSV
    ' Private Const NEW_LINE_INDICATOR As String = "ABC"

    Public Sub CSV_readTargetsFile()
        Dim i As Integer
        Dim tmpCard As Integer, tmpValue As Double, tmpFieldName As String = "", tmpLineSite As String = "", tmpLine As String = "", tmpSite As String = ""
        Dim tmpCharIndex As Integer, lineIndex As Integer
        Dim isFirstLine As Boolean = True

        '   Dim _assembly As [Assembly]
        '   _assembly = [Assembly].GetExecutingAssembly()
        Try
            '  Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(PATH_PRSTORY_TARGETS & FILE_RAWTARGETS_CSV)
            ' Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(System.IO.Path.Combine(Environment.CurrentDirectory, "Resources\" & FILE_RAWTARGETS_CSV))
            ' Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser( _assembly.GetManifestResourceStream("DigitalFactory." & FILE_RAWTARGETS_CSV))

            Dim fileContent As String = My.Resources.target 'My.Resources.prstory_dtpct_targets
            Dim stringStream As New System.IO.StringReader(fileContent)

            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(stringStream)

                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        Dim currentField As String
                        i = 0
                        For Each currentField In currentRow
                            Select Case i
                                Case 0
                                    tmpLineSite = currentField
                                    tmpCharIndex = tmpLineSite.IndexOf(":")
                                    If tmpCharIndex > -1 Then
                                        tmpLine = tmpLineSite.Substring(0, tmpCharIndex)
                                        tmpSite = tmpLineSite.Substring(tmpCharIndex + 2, Len(tmpLineSite) - Len(tmpLine) - 2)
                                    End If
                                Case 1
                                    tmpCard = CInt(currentField)
                                Case 2
                                    tmpFieldName = currentField
                                Case 3
                                    tmpValue = CDbl(currentField)
                            End Select
                            i += 1
                        Next
                        If tmpValue < 1 Then

                            lineIndex = -1
                            For ix As Integer = 0 To AllProdLines.Count - 1
                                With AllProdLines(ix)
                                    If .Name = tmpLine Then
                                        If .SiteName = tmpSite Or .parentSite.ThreeLetterID = tmpSite Then
                                            lineIndex = ix
                                            GoTo found
                                        End If
                                    End If
                                    'this is in case the order in the template gets reversed upon manual entry (as of 5/1/17 many of the entries are like this)
                                    If .Name = tmpSite Then

                                        If .SiteName = tmpLine Or .parentSite.ThreeLetterID = tmpLine Then
                                            Dim tmpZ As String = tmpLine
                                            tmpLine = tmpSite
                                            tmpSite = tmpZ
                                            lineIndex = ix
                                            GoTo found
                                        End If
                                    End If
                                End With
                            Next
found:
                            If lineIndex > -1 Then
                                With AllProdLines(lineIndex)
                                    If Not .doIhaveTargets Then .DowntimePercentTargets = New DTPct_Targets(tmpLine, tmpSite)
                                    .DowntimePercentTargets.addNewTarget(tmpFieldName, tmpCard, tmpValue)
                                    If tmpCard = 2 Then
                                        .DowntimePercentTargets.addNewTarget(tmpFieldName, 3, tmpValue) 'this is a hack to make amanda's targets work. could result in planned t1/upt2 targets mixing.
                                    End If
                                End With
                            End If
                        End If
                    Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While

            End Using

        Catch ex As System.IO.FileNotFoundException
            MsgBox("Error Importing Default Settings. Please Contact LG/Sam If Problem Persists." & ex.Message)
        Catch ex As Exception
            MsgBox("Error Importing Default Settings. Please Contact LG/Sam If Problem Persists." & ex.Message)
        End Try
    End Sub






End Module

