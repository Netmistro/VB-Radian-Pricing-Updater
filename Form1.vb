Imports System.Net.Mail
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmMain
    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Me.Close()

    End Sub

    Private Sub TxtTo_TextChanged(sender As Object, e As EventArgs) Handles txtTo.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        txtTo.Text = "arnold.bradshaw@rhatt.com"

    End Sub

    Private Sub BtnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click

        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("arnold.bradshaw@rhatt.com", "Radian1234*")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("arnold.bradshaw@rhatt.com")
            e_mail.To.Add(txtTo.Text)
            e_mail.Subject = "Scaffold Material Pricing"
            e_mail.IsBodyHtml = False
            e_mail.Body = txtMessage.Text
            Smtp_Server.Send(e_mail)
            MsgBox("Quotation email Sent!")


        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Open File Dialogue"
        fd.InitialDirectory = "C:\"

        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            txtFileName.Text = strFileName

        End If

    End Sub

    Private Sub BtnPrice_Click(sender As Object, e As EventArgs) Handles btnPrice.Click

        Dim xlFileName As String = txtFileName.Text
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWB As Excel.Workbook = xlApp.Workbooks.Open(xlFileName)
        Dim Sheet1 As Excel.Worksheet = xlWB.Sheets("Sheet")
        Dim currentDate As Date = Date.Today


        ' Make Excel Visible
        xlApp.Visible = True

        ' Add a new Excel Worksheet
        Dim Sheet2 As Excel.Worksheet = xlWB.Worksheets.Add
        Sheet2 = xlWB.ActiveSheet

        ' Add table headers
        Dim Headers() As String = {"Item No.", "Description", "Quantity", "Sale Unit Price", "Sale Price", "Sale Used Price", "Used Price", "Rental Unit Price", "Rental Price", "Weight (lbs.)"}
        For i As Integer = 1 To 10
            Sheet2.Cells(1, i) = Headers(i - 1)
        Next

                
        'Get the last row of the spreadsheet
        Dim lastRow As String = Sheet1.UsedRange.Rows.Count

        For j As Integer = 12 To lastRow - 4
            ' Create Item Nos.
            Sheet2.Cells(j - 10, 1) = (j - 11)
            ' Copy Descriptions
            Sheet2.Cells(j - 10, 2) = Sheet1.Cells(j, 4)
            ' Copy Quantities
            Sheet2.Cells(j - 10, 3) = Sheet1.Cells(j, 10)
            ' Copy weight in lbs.
            Sheet2.Cells(j - 10, 10) = Sheet1.Cells(j, 13)
        Next

        Dim newFileName As String = xlFileName.Substring(0, xlFileName.Length - 5)
        'Save new workbook
        xlWB.SaveAs(newFileName & " - " & Format(currentDate, "yyyy-MM-dd") & ".xlsx")

        ' Release object references
        Sheet1 = Nothing
        xlWB = Nothing
        xlApp.Quit()
        xlApp = Nothing
        xlFileName = Nothing
        newFileName = Nothing

    End Sub
End Class
