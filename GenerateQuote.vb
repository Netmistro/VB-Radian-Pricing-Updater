Imports System.Configuration
Imports System.IO

Public Class frmGenerateQuote
    Dim MainClass As New MainClass

    Private Sub frmGenerateQuote_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MainClass.CentreForm()
        'MainClass.GenerateQuote()

    End Sub
End Class