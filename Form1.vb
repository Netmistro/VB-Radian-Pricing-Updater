Imports System.Net.Mail
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmMain
    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Me.Close()

    End Sub

    Private Sub TxtTo_TextChanged(sender As Object, e As EventArgs) Handles txtTo.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Centre window on load
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        Dim formNewLeft As Integer = Val(Val(screenWidth) / Val(2)) - Val(Val(Me.Width) / Val(2))
        Dim formNewTop As Integer = Val(Val(screenHeight) / Val(2)) - Val(Val(Me.Height) / Val(2))
        Me.Left = formNewLeft
        Me.Top = formNewTop

        'Load Default email address
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

        ' Materials Pricelist
        Dim materialPricelist = New List(Of KeyValuePair(Of String, Decimal()))
        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok combined jack and base plate", {150.0, 125.0, 1.5}))
        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok deck adaptor", {200.0, 150.0, 1.5}))

        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 0.8m", {150.0, 130.0, 1.3}))
        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 0.9m", {175.0, 150.0, 1.5}))
        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 1.3m", {200.0, 175.0, 1.75}))
        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 1.8m", {300.0, 260.0, 2.6}))
        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 2.5m", {380.0, 330.0, 3.3}))
        materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 3.0m", {450.0, 400.0, 4.0}))




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
            Dim weight As String = Sheet2.Cells(j - 10, 10).value().ToString
            weight = weight.Substring(0, weight.Length - 4)
            Sheet2.Cells(j - 10, 10) = weight

            ' Perform some cell formatting
            Sheet2.Range("A1:J1").EntireColumn.AutoFit()
            Sheet2.Range("A1:J1").Font.Bold = True
            Sheet2.Range("A1:J1").Interior.ColorIndex = 6
            ' Format cells as Accounting
            Sheet2.Cells(j - 10, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Sheet2.Cells(j - 10, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Sheet2.Cells(j - 10, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Sheet2.Cells(j - 10, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Sheet2.Cells(j - 10, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Sheet2.Cells(j - 10, 9).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Sheet2.Cells(j - 10, 10).NumberFormat = "#,##0.00"

            Dim value1 As Decimal = Sheet2.Cells(j - 10, 4).value()
            Dim value2 As Decimal = Sheet2.Cells(j - 10, 3).value()

            ' Calculate sale values and rentals
            Sheet2.Range("E" & (j - 10)).Formula = "=C" & (j - 10) & "*" & "D" & (j - 10)
            Sheet2.Range("G" & (j - 10)).Formula = "=F" & (j - 10) & "*" & "D" & (j - 10)
            Sheet2.Range("I" & (j - 10)).Formula = "=H" & (j - 10) & "*" & "D" & (j - 10)

            Dim materialDescription As String = Sheet2.Cells(j - 10, 2).value().ToString
            For Each pair In materialPricelist
                If pair.Key = materialDescription Then
                    Sheet2.Cells(j - 10, 4) = pair.Value(0)
                    Sheet2.Cells(j - 10, 6) = pair.Value(1)
                    Sheet2.Cells(j - 10, 8) = pair.Value(2)

                End If
            Next
        Next

        ' Totals of each column for sale and weight
        Sheet2.Range("E" & (lastRow - 12)).Formula = "=SUM(E2:" & "E" & (lastRow - 13)
        Sheet2.Range("G" & (lastRow - 12)).Formula = "=SUM(G2:" & "G" & (lastRow - 13)
        Sheet2.Range("I" & (lastRow - 12)).Formula = "=SUM(I2:" & "I" & (lastRow - 13)
        Sheet2.Range("J" & (lastRow - 12)).Formula = "=SUM(J2:" & "J" & (lastRow - 13)


        Dim newFileName As String = xlFileName.Substring(0, xlFileName.Length - 5)
        'Save new workbook
        xlWB.SaveAs(newFileName & " - " & Format(currentDate, "yyyy-MM-dd") & ".xlsx")

        ' Release object references
        Sheet1 = Nothing
        Sheet2 = Nothing
        xlWB = Nothing
        xlApp.Quit()
        xlApp = Nothing
        xlFileName = Nothing
        newFileName = Nothing

    End Sub

    Private Sub BtnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click

        'Open priced file

    End Sub
End Class
