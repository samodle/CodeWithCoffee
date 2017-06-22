Imports System.Windows.Forms
Imports System.Windows
Imports System.DirectoryServices.AccountManagement

Partial Public Class WinLogin
    Inherits Window

    ' public ParentWindow as WindowMAIN_prstory

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        txtDomain.Text = Environment.UserDomainName.ToLower()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        DialogResult = False
    End Sub

    Private Sub btnLogin_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Try
            Dim pc As PrincipalContext = New PrincipalContext(ContextType.Domain, System.Environment.UserDomainName, ContextOptions.SecureSocketLayer)
            With pc
                If Not pc.ValidateCredentials(txtUserName.Text, txtPassword.Password, ContextOptions.Negotiate) Then
                    Windows.MessageBox.Show("It seems you might not have permission to use prstory or you might have entered an incorrect username/password. Please try again.", "Authentication Failure", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    DialogResult = True
                End If
            End With
        Catch ex As Exception
            Windows.MessageBox.Show("prstory experienced the following exception while accessing the Active Directory: " & ex.Message & ". prstory will now restart. If this error continues, please contact your SPOC.", "LADP Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            DialogResult = False
        End Try
    End Sub
End Class