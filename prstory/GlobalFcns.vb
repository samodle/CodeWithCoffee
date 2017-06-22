Imports System.Net
Module GlobalFcns

    Public Sub sortEventList_ByStops(ByRef tgtList As List(Of DTevent))
        Dim i As Integer
        For i = 0 To tgtList.Count - 1
            tgtList(i).sortBy_Stops()
        Next
        tgtList.Sort()
    End Sub
    Public Sub sortEventList_ByDT(ByRef tgtList As List(Of DTevent))
        Dim i As Integer
        For i = 0 To tgtList.Count - 1
            tgtList(i).sortBy_DT()
        Next
        tgtList.Sort()
    End Sub

 
    Public Function getTimeDifference_Minutes(earlierTime As Object, laterTime As Object) As Long
        Return timeToInteger(laterTime.subtract(earlierTime))
    End Function
    Public Function timeToInteger(time As TimeSpan) As Long
        Return time.Minutes + time.Hours * 60 + time.Seconds / 60 + time.Days * 1440 '24 * 60
    End Function

    Function NORMSDIST(ByVal x As Single) As Single
        Dim result As Single
        Dim y As Single = 1 / (1 + (0.2316419 * Math.Abs(x)))
        Dim z As Single = 0.3989423 * (Math.Exp((-x ^ 2) / 2))

        result = 1 - z * ((1.33027 * (y ^ 5)) - (1.821256 * (y ^ 4)) + (1.781478 * (y ^ 3)) - (0.356538 * (y ^ 2)) + (0.3193815 * y))

        If x > 0 Then
            Return result
        Else
            Return 1 - result
        End If
    End Function




    Public Function onlyDigits(s As String) As String
        ' Variables needed (remember to use "option explicit").   '
        Dim retval As String    ' This is the return string.      '
        Dim i As Integer        ' Counter for character position. '

        ' Initialise return string to empty                       '
        retval = ""

        ' For every character in input string, copy digits to     '
        '   return string.                                        '
        For i = 1 To Len(s)

            If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
                retval = retval + Mid(s, i, 1)
            End If
        Next

        ' Then return the return string.                          '
        onlyDigits = retval
    End Function


    Public Function decidecolor(score As Integer) As SolidColorBrush

        Dim whichcolor As New SolidColorBrush
        Select Case score
            Case 1
                whichcolor = bubblecolorGreen
            Case 2
                whichcolor = bubblecolorYellow
            Case 3
                whichcolor = bubblecolorOrange
            Case 4
                whichcolor = bubblecolorRed
            Case Else
                whichcolor = bubblecolorGreen
        End Select
        Return whichcolor
    End Function

    Public Function CheckIfFtpFileExists(ByVal fileUri As String, username As String, password As String) As Boolean


        Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create(fileUri), System.Net.FtpWebRequest)
        request.Credentials = New System.Net.NetworkCredential(username, password)
        'request.Method = System.Net.WebRequestMethods.Ftp.UploadFile


        'Dim request As FtpWebRequest = WebRequest.Create(fileUri)
        'request.Credentials = New NetworkCredential(username, password)
        request.Method = WebRequestMethods.Ftp.GetFileSize
        Try
            Dim response As FtpWebResponse = request.GetResponse()
            ' THE FILE EXISTS
        Catch ex As WebException
            Dim response As FtpWebResponse = ex.Response
            If FtpStatusCode.ActionNotTakenFileUnavailable = response.StatusCode Then
                ' THE FILE DOES NOT EXIST
                Return False
            End If
        End Try
        Return True
    End Function

    Public Function ReturnNullifzero(number_to_analyze As Double) As Object
        If number_to_analyze <= 0 Then
            Return "null"
        Else
            Return number_to_analyze
        End If

    End Function


    Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        ' by making Generator static, we preserve the same instance '
        ' (i.e., do not create new instances with the same seed over and over) '
        ' between calls '
        Static Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function
End Module
