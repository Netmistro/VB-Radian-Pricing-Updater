Imports System.Net.Mail
Imports System.Configuration
Imports System.IO
Imports System.Net

Public Class frmPricing
    Dim MainClass As New MainClass

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        MainClass.Close()
    End Sub

    Private Sub Pricing_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MainClass.CentreForm()
        MainClass.LoadForm()
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

    Private Sub BtnAttach_Click(sender As Object, e As EventArgs) Handles btnAttach.Click
        MainClass.AttachFiles()
    End Sub

    Private Sub btnQuote_Click(sender As Object, e As EventArgs) Handles btnQuote.Click
        frmGenerateQuote.ShowDialog()
    End Sub
End Class
