Public Class Window_TextInput
    Public Sub Window_TextInput()
        InitializeComponent()
    End Sub

    Public Property ResponseText As String
        Get
            Return ResponseTextBox.Text
        End Get
        Set
            ResponseTextBox.Text = Value
        End Set
    End Property

    Private Sub OKButton_Click(sender As Object, e As System.Windows.RoutedEventArgs)

        DialogResult = True
    End Sub
End Class
