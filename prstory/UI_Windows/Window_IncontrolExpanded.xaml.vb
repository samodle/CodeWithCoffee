Imports System.IO
Imports Awesomium.Core
Imports System.ComponentModel
Imports Awesomium.Windows.Controls
Public Class Window_IncontrolExpanded
    Public tempbubblesender_expanded As Object
    Sub incontrol_loaded()
        SPCchart.Visibility = Windows.Visibility.Visible

        assignvalues()

    End Sub
    Sub hideSPC()
        SPCchart.Visibility = Windows.Visibility.Hidden
        incontrolstopcount_circle.Visibility = Windows.Visibility.Hidden
        incontrolstopcountLabel.Visibility = Windows.Visibility.Hidden
        incontroltextstopsperdaylabel.Visibility = Windows.Visibility.Hidden
        incontrolvstextlabel.Visibility = Windows.Visibility.Hidden
        incontrol90dayaveragestopslabe.Visibility = Windows.Visibility.Hidden
    End Sub
    Sub showSPC()
        SPCchart.Visibility = Windows.Visibility.Visible
        incontrolstopcount_circle.Visibility = Windows.Visibility.Visible
        incontrolstopcountLabel.Visibility = Windows.Visibility.Visible
        incontroltextstopsperdaylabel.Visibility = Windows.Visibility.Visible
        incontrolvstextlabel.Visibility = Windows.Visibility.Visible
        incontrol90dayaveragestopslabe.Visibility = Windows.Visibility.Visible
    End Sub
    Sub assignvalues()

        incontrolstopcountLabel.Content = stopbubblestops(bubblenumberpublic)
        incontrol90dayaveragestopslabe.Content = stopbubble90daystopsperday(bubblenumberpublic) & " AVG STOPS PER DAY LAST 90 DAYS"
        incontroltextstopsperdaylabel_heading.Content = stopbubblenames(bubblenumberpublic) & " - -  CURRENT ANALYSIS PERIOD STOPS VS LAST 90 DAYS"
    End Sub


    Private Sub RefreshChart()
        RefreshAlertLabel.Visibility = Windows.Visibility.Hidden
    End Sub

    Private Sub Generalmousemove(sender As Object, e As MouseEventArgs)
        sender.Opacity = 0.7
    End Sub

    Private Sub Generalmouseleave(sender As Object, e As MouseEventArgs)
        sender.Opacity = 1.0
    End Sub
End Class
