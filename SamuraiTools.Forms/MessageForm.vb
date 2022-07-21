Imports System.Windows.Forms

Namespace Forms
    Public Class MessageForm
        Inherits Form

        Public ReadOnly Property lblMessage As Label

        Public Sub New(formMessage As String)
            lblMessage = New Label()
            lblMessage.Text = formMessage
            Me.Controls.Add(lblMessage)

            Width = 200
            Height = 100

            MinimizeBox = False
            MaximizeBox = False
        End Sub

        Private Sub MessageForm_Load(sender As Object, e As EventArgs) Handles Me.Load
            CenterToScreen()
        End Sub
    End Class
End Namespace