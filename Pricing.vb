Public Class frmPricing
    Dim MainClass As New MainClass

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        MainClass.Close()
    End Sub

    Private Sub Pricing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MainClass.LoadPricingForm()
        MainClass.CentreForm(Me)
    End Sub

    Private Sub BtnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        MainClass.SendMail()
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        MainClass.BrowseFiles()
    End Sub

    Private Sub BtnPrice_Click(sender As Object, e As EventArgs) Handles btnPrice.Click
        MainClass.PriceItems()
    End Sub

    Private Sub BtnAttach_Click(sender As Object, e As EventArgs)
        MainClass.AttachFiles()
    End Sub

    Private Sub btnQuote_Click(sender As Object, e As EventArgs) Handles btnQuote.Click
        frmGenerateQuote.ShowDialog()
    End Sub

    Private Sub btnNewFile_Click(sender As Object, e As EventArgs) Handles btnNewFile.Click
        MainClass.OpenPricedFile()
    End Sub

End Class
