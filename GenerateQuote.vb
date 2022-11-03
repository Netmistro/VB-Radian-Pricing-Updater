Public Class frmGenerateQuote
    Dim MainClass As New MainClass

    Private Sub frmGenerateQuote_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MainClass.CentreForm(Me)
        MainClass.LoadQuoteForm()
    End Sub

    Private Sub btnQuote_Click(sender As Object, e As EventArgs) Handles btnQuote.Click
        MainClass.GenerateQuote()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

End Class