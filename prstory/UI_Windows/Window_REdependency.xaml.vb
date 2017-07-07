Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Xml 'not sure if i need thsi
Imports System.Threading

Public Class Window_REdependency
#Region "Variables & Properties"
    Dim prstoryReport As prStoryMainPageReport
    Dim _ActiveDataCollection As New ObservableCollection(Of dependencyEvent)()
    Public ReadOnly Property ActiveDataCollection() As ObservableCollection(Of dependencyEvent)
        Get
            Return _ActiveDataCollection
        End Get
    End Property


    'sorting our listview
    Private _lastHeaderClicked_ActiveData As GridViewColumnHeader = Nothing
    Private _lastDirection_ActiveData As ListSortDirection = ListSortDirection.Ascending

#End Region

    Public Sub New(ByVal depArr As Array, ByVal prstoryReportIn As prStoryMainPageReport)
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        prstoryReport = prstoryReportIn

        For i As Integer = 1 To depArr.GetLength(0) - 1
            _ActiveDataCollection.Add(New dependencyEvent(depArr(i, 1), depArr(i, 1), depArr(i, 0), depArr(i, 3), depArr(i, 4), depArr(i, 5), depArr(i, 6), depArr(i, 7), depArr(i, 8)))
        Next

        ' Dim DependencyHTMLthread As New Thread(AddressOf HTML_exportBubbleREDEP)
        ' DependencyHTMLthread.Start()
    End Sub

#Region "Column Sorting"
    Private Sub GridViewColumnHeaderClickedHandler_activedata(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim headerClicked As GridViewColumnHeader = TryCast(e.OriginalSource, GridViewColumnHeader)
        Dim direction As ListSortDirection

        If headerClicked IsNot Nothing Then
            If headerClicked.Role <> GridViewColumnHeaderRole.Padding Then
                If headerClicked IsNot _lastHeaderClicked_ActiveData Then
                    direction = ListSortDirection.Ascending
                Else
                    If _lastDirection_ActiveData = ListSortDirection.Ascending Then
                        direction = ListSortDirection.Descending
                    Else
                        direction = ListSortDirection.Ascending
                    End If
                End If

                Dim header As String = TryCast(headerClicked.Column.Header, String)
                Sort_ActiveData(header, direction)

                If direction = ListSortDirection.Ascending Then
                    headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowUp"), DataTemplate)
                Else
                    headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowDown"), DataTemplate)
                End If

                ' Remove arrow from previously sorted header
                If _lastHeaderClicked_ActiveData IsNot Nothing AndAlso _lastHeaderClicked_ActiveData IsNot headerClicked Then
                    _lastHeaderClicked_ActiveData.Column.HeaderTemplate = Nothing
                End If


                _lastHeaderClicked_ActiveData = headerClicked
                _lastDirection_ActiveData = direction
            End If
        End If
    End Sub
    Private Sub Sort_ActiveData(ByVal sortBy As String, ByVal direction As ListSortDirection)
        Dim dataView As ICollectionView = CollectionViewSource.GetDefaultView(ActiveDataList.ItemsSource)

        dataView.SortDescriptions.Clear()
        Dim sd As New SortDescription(sortBy, direction)
        dataView.SortDescriptions.Add(sd)
        dataView.Refresh()
    End Sub
#End Region
    Public Sub SelectionChangedEventHandler()

    End Sub

End Class

Public Class dependencyEvent
    Private _preStopName As String
    Private _postStopName As String
    Private _depType As String
    Private _ActExpNum As Double
    Private _ActExpPct As Double
    Private _ActNum As Double
    Private _ActPct As Double
    Private _ExpNum As Double
    Private _ExpPct As Double

    Private _Stops As Double

    Friend PostStops

    Public Property Stops As Double
        Get
            Return _Stops
        End Get
        Set(value As Double)
            _Stops = value
        End Set
    End Property

    Public ReadOnly Property preStop As String
        Get
            Return _preStopName
        End Get
    End Property
    Public ReadOnly Property postStop As String
        Get
            Return _postStopName
        End Get
    End Property
    Public ReadOnly Property TypeX As String
        Get
            Return _depType
        End Get
    End Property
    Public ReadOnly Property ActExpNum As String
        Get
            Return Math.Round(_ActExpNum, 1)
        End Get
    End Property
    Public ReadOnly Property ActExpPct As String
        Get
            Return Math.Round(_ActExpPct, 1) & "%"
        End Get
    End Property
    Public ReadOnly Property ActNum As String
        Get
            Return Math.Round(_ActNum, 1)
        End Get
    End Property
    Public ReadOnly Property ActPct As String
        Get
            Return Math.Round(_ActPct, 1) & "%"
        End Get
    End Property
    Public ReadOnly Property ExpNum As String
        Get
            Return Math.Round(_ExpNum, 1)
        End Get
    End Property
    Public ReadOnly Property ExpPct As String
        Get
            Return Math.Round(_ExpPct, 1) & "%"
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return _preStopName & "-" & _postStopName & ": " & ActExpPct
    End Function


    Public Sub New(preN As String, postN As String, Dtype As String, AEn As Double, AEp As Double, An As Double, Ap As Double, En As Double, Ep As Double)
        _preStopName = preN
        _postStopName = postN
        _depType = Dtype
        _ActExpNum = AEn
        _ActExpPct = AEp
        _ActNum = An
        _ActPct = Ap
        _ExpNum = En
        _ExpPct = Ep
    End Sub

End Class
