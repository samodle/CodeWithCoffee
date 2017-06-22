Public Class DTPct_Targets
    Private _parentLineName As String
    Private _parentSiteName As String

    Private _targetList As New List(Of PRTarget)
    Public ReadOnly Property RawTargetList As List(Of PRTarget)
        Get
            Return _targetList
        End Get
    End Property

    Public Sub New(lineName As String, siteName As String)
        _parentLineName = lineName
        _parentSiteName = siteName
    End Sub

    'returns -1 if target does not exist
    Public Function getTargetValue(fieldName As String, CardNumber As Integer) As Double
        Dim tmpIndex As Integer = _targetList.IndexOf(New PRTarget(0, fieldName, CardNumber))
        If tmpIndex > -1 Then
            Return _targetList(tmpIndex).targetValue
        Else
            Return -1
        End If
    End Function

    Public Sub addNewTarget(fieldName As String, CardNumber As Integer, fieldValue As Double)
        Dim tmpIndex As Integer = _targetList.IndexOf(New PRTarget(0, fieldName, CardNumber))
        If tmpIndex > -1 Then 'already exists!
            _targetList(tmpIndex).targetValue = fieldValue
        Else
            _targetList.Add(New PRTarget(fieldValue, fieldName, CardNumber))
        End If
    End Sub

    Public Function getMaxDTpct(CardNumber As Integer) As Double
        Dim currentMaxVal As Double = 0
        For i As Integer = 0 To _targetList.Count - 1
            With _targetList(i)
                If .Card = CardNumber Then
                    If .targetValue > currentMaxVal Then currentMaxVal = .targetValue
                End If
            End With
        Next
        Return currentMaxVal
    End Function

End Class

Public Class PRTarget
    Implements IEquatable(Of PRTarget)

#Region "Variables & Properties"
    Private _cardNumber As Integer
    Private _targetValue As Double
    Private _targetName As String

    Public Property targetValue As Double
        Get
            Return _targetValue
        End Get
        Set(value As Double)
            _targetValue = value
        End Set
    End Property
    Public ReadOnly Property targetName As String
        Get
            Return _targetName
        End Get
    End Property
    Public ReadOnly Property Card As Integer
        Get
            Return _cardNumber
        End Get
    End Property
#End Region

#Region "Construction"
    Public Sub New(dataVal As Double, Name As String, CardNum As Integer)
        _targetName = Name
        _targetValue = dataVal
        _cardNumber = CardNum
    End Sub
#End Region


    Public Overrides Function ToString() As String
        Return _targetName & " " & _targetValue
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Return False
        End If
        Dim objAsPart As prtarget = TryCast(obj, prtarget)
        If objAsPart Is Nothing Then
            Return False
        Else
            Return Equals(objAsPart)
        End If
    End Function
    Public Overloads Function Equals(other As PRTarget) As Boolean _
        Implements IEquatable(Of PRTarget).Equals
        If other Is Nothing Then
            Return False
        End If
        If (Me.targetName.Equals(other.targetName)) Then
            Return Me.Card = other.Card
        Else
            Return False
        End If
    End Function
End Class
