Option Strict On
Option Explicit On

Public Class LoginForm
    Public ReadOnly Property PasswordValue As String
        Get
            Return txtPassword.Text
        End Get
    End Property

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

End Class
