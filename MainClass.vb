Imports System.Net.Mail
Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word

Public Class MainClass
    Sub Close()
        Application.Exit()
    End Sub
    Sub LoadPricingForm()
        'Set focus on the first textbox
        frmPricing.txtFileName.Select()
        'Load Default email address
        frmPricing.txtTo.Text = "training@rhatt.com"
    End Sub
    Sub LoadQuoteForm()
        frmGenerateQuote.rbAlicia.Checked = True
        frmGenerateQuote.rbCurrentDate.Checked = True
        frmGenerateQuote.rbSale.Checked = True
        CompanyDetails()
    End Sub
    Sub SendMail()
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
            e_mail.To.Add(frmPricing.txtTo.Text)
            e_mail.Bcc.Add("arnold.bradshaw@rhatt.com")
            e_mail.Subject = "Scaffold Material Pricing"
            e_mail.Attachments.Add(New Attachment(frmPricing.txtAttachment.Text))
            e_mail.IsBodyHtml = False
            e_mail.Body = frmPricing.txtMessage.Text & vbCr & vbCr & "RADIAN H.A. Limited"

            If frmPricing.txtTo.Text = "" Or frmPricing.txtMessage.Text = "" Or frmPricing.txtAttachment.Text = "" Then

                MessageBox.Show("Please fill out all fields!", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Else
                Smtp_Server.Send(e_mail)
                MessageBox.Show("Your Spreadsheet has been Sent!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message + "Error sending email!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Sub
    Sub PriceItems()
        If frmPricing.txtFileName.Text = "" Then

            MessageBox.Show("Please Select a File to Price!", "Error Selecting File", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else

            Dim xlFileName As String = frmPricing.txtFileName.Text
            Dim xlApp As Excel.Application = New Excel.Application()
            Dim xlWB As Excel.Workbook = xlApp.Workbooks.Open(xlFileName)
            Dim Sheet1 As Excel.Worksheet = xlWB.Sheets("Sheet")
            Dim currentDate As Date = Date.Today

            ' Materials Pricelist and Pricing
            Dim materialPricelist = New List(Of KeyValuePair(Of String, Decimal())) From {
                New KeyValuePair(Of String, Decimal())("Cuplok combined jack and base plate", {245.0, 210.0, 2.45}),
                New KeyValuePair(Of String, Decimal())("Cuplok deck adaptor", {280.0, 210.0, 2.8}),
                New KeyValuePair(Of String, Decimal())("Ladder hatch", {1400.0, 1100.0, 21.0}),
                New KeyValuePair(Of String, Decimal())("Cuplok ladder safety gate: 0.8m", {1400.0, 1260.0, 21.0}),
                New KeyValuePair(Of String, Decimal())("Swivel coupler", {91, 63.0, 0.46}),
                New KeyValuePair(Of String, Decimal())("Base plate", {77.0, 49.0, 0.39}),
                New KeyValuePair(Of String, Decimal())("Double coupler", {84.0, 56.0, 0.42}),
                New KeyValuePair(Of String, Decimal())("Putlog coupler", {84.0, 56.0, 0.42}),
                New KeyValuePair(Of String, Decimal())("Sleeve coupler", {84.0, 56.0, 0.42}),
                New KeyValuePair(Of String, Decimal())("Sole pads", {105.0, 70.0, 1.05}),
                New KeyValuePair(Of String, Decimal())("Toe board clip", {60.0, 40.0, 0.3}),
                New KeyValuePair(Of String, Decimal())("Ring bolt", {0.0, 0.0, 0.0}),
                New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 0.6m", {175.0, 154.0, 1.5}),
                New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 0.8m", {210.0, 182.0, 1.8}),
                New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 0.9m", {245.0, 210.0, 2.1}),
                New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 1.3m", {280.0, 245.0, 2.4}),
                New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 1.8m", {420.0, 364.0, 3.6}),
                New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 2.5m", {475.0, 320.0, 5.7}),
                New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 3.0m", {650.0, 560.0, 7.6}),
                New KeyValuePair(Of String, Decimal())("Cuplok intermediate transom: 0.9m", {266.0, 182.0, 2.11}),
                New KeyValuePair(Of String, Decimal())("Cuplok intermediate transom: 1.3m", {420.0, 350.0, 2.8}),
                New KeyValuePair(Of String, Decimal())("Cuplok intermediate transom: 1.8m", {665.0, 0.0, 4.2}),
                New KeyValuePair(Of String, Decimal())("Cuplok intermediate transom: 2.5m", {910.0, 0.0, 5.6}),
                New KeyValuePair(Of String, Decimal())("Cuplok standard: 0.9m", {350.0, 0.0, 5.5}),
                New KeyValuePair(Of String, Decimal())("Cuplok standard: 1.0m", {350.0, 0.0, 5.5}),
                New KeyValuePair(Of String, Decimal())("Cuplok standard: 1.52m", {560.0, 455.0, 4.2}),
                New KeyValuePair(Of String, Decimal())("Cuplok standard: 1.50m", {560.0, 455.0, 4.2}),
                New KeyValuePair(Of String, Decimal())("Cuplok standard: 2.0m", {630.0, 504.0, 5.0}),
                New KeyValuePair(Of String, Decimal())("Cuplok standard: 2.5m", {784.0, 630.0, 7.84}),
                New KeyValuePair(Of String, Decimal())("Cuplok standard: 3.0m", {917.0, 770.0, 9.17}),
                New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 1.8 x 2.0m", {490.0, 420.0, 4.26}),
                New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 2.5 x 1.5m", {500.0, 450.0, 4.0}),
                New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 2.5 x 2.0m", {560.0, 490.0, 4.87}),
                New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 3.0 x 2.0m", {630.0, 560.0, 5.4}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 2.0m", {980.0, 840.0, 9.8}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 2.5m", {1225.0, 1050.0, 12.25}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 3.0m", {1470.0, 1344.0, 14.7}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 3.5m", {1715.0, 1568.0, 17.15}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 4.0m", {1960.0, 1792.0, 19.6}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 5.0m", {2450.0, 2240.0, 24.5}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 6.0m", {2940.0, 2688.0, 29.4}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 7.0m", {3430.0, 3136.0, 34.3}),
                New KeyValuePair(Of String, Decimal())("Ladder: steel 8.0m", {3920.0, 3584.0, 39.2}),
                New KeyValuePair(Of String, Decimal())("Steel board: 2.5m", {1260.0, 1050.0, 9.8}),
                New KeyValuePair(Of String, Decimal())("Steel board: 1.8m", {945.0, 0.0, 7.7}),
                New KeyValuePair(Of String, Decimal())("Steel board: 1.2m", {840.0, 630.0, 5.6}),
                New KeyValuePair(Of String, Decimal())("Steel tube:  1'", {35.0, 30.1, 0.35}),
                New KeyValuePair(Of String, Decimal())("Steel tube:  2'", {70.0, 60.2, 0.7}),
                New KeyValuePair(Of String, Decimal())("Steel tube:  4'", {140.0, 120.4, 1.4}),
                New KeyValuePair(Of String, Decimal())("Steel tube:  5'", {175.0, 150.5, 1.75}),
                New KeyValuePair(Of String, Decimal())("Steel tube:  6'", {210.0, 180.6, 2.1}),
                New KeyValuePair(Of String, Decimal())("Steel tube:  8'", {280.0, 240.8, 2.8}),
                New KeyValuePair(Of String, Decimal())("Steel tube: 10'", {350.0, 301.0, 3.5}),
                New KeyValuePair(Of String, Decimal())("Steel tube: 12'", {497.0, 361.2, 4.97}),
                New KeyValuePair(Of String, Decimal())("Steel tube: 14'", {490.0, 421.4, 4.9}),
                New KeyValuePair(Of String, Decimal())("Steel tube: 16'", {560.0, 481.6, 5.6}),
                New KeyValuePair(Of String, Decimal())("Steel tube: 18'", {630.0, 541.8, 6.3}),
                New KeyValuePair(Of String, Decimal())("Steel tube: 20'", {700.0, 666.6, 7.0}),
                New KeyValuePair(Of String, Decimal())("Steel tube: 21'", {735.0, 632.1, 7.35}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 2.0m", {980.0, 840.0, 9.8}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 2.5m", {1225.0, 1050.0, 12.25}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 3.0m", {1470.0, 1344.0, 14.7}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 3.5m", {1715.0, 1568.0, 17.15}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 4.0m", {1960.0, 1792.0, 19.6}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 5.0m", {2450.0, 2240.0, 24.5}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 6.0m", {2940.0, 2688.0, 29.4}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 7.0m", {3430.0, 3136.0, 34.3}),
                New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 8.0m", {3920.0, 3584.0, 39.2}),
                New KeyValuePair(Of String, Decimal())("Timber board:  1'", {59.06, 49.0, 0.59}),
                New KeyValuePair(Of String, Decimal())("Timber board:  2'", {118.13, 98.0, 1.18}),
                New KeyValuePair(Of String, Decimal())("Timber board:  4'", {236.25, 175.0, 2.36}),
                New KeyValuePair(Of String, Decimal())("Timber board:  5'", {262.5, 218.75, 2.63}),
                New KeyValuePair(Of String, Decimal())("Timber board:  6'", {350.0, 262.5, 3.5}),
                New KeyValuePair(Of String, Decimal())("Timber board:  8'", {437.5, 350.0, 4.38}),
                New KeyValuePair(Of String, Decimal())("Timber board: 10'", {525.0, 437.5, 5.25}),
                New KeyValuePair(Of String, Decimal())("Timber board: 13'", {568.75, 481.25, 5.69}),
                New KeyValuePair(Of String, Decimal())("Timber board:  2' - 225mm x 38mm", {118.13, 98.0, 1.18}),
                New KeyValuePair(Of String, Decimal())("Timber board:  4' - 225mm x 38mm", {236.25, 175.0, 2.36}),
                New KeyValuePair(Of String, Decimal())("Timber board:  6' - 225mm x 38mm", {350.0, 262.5, 3.5}),
                New KeyValuePair(Of String, Decimal())("Timber board:  8' - 225mm x 38mm", {437.5, 350.0, 4.38}),
                New KeyValuePair(Of String, Decimal())("Timber board: 10' - 225mm x 38mm", {525.0, 437.5, 5.25}),
                New KeyValuePair(Of String, Decimal())("Timber board: 12' - 225mm x 38mm", {500.0, 460.0, 5.0}),
                New KeyValuePair(Of String, Decimal())("Timber board: 13' - 225mm x 38mm", {568.75, 481.25, 5.69})
            }

            ' Make Excel Visible
            xlApp.Visible = False

            ' Add a new Excel Worksheet
            Dim Sheet2 As Excel.Worksheet = xlWB.Worksheets.Add
            Sheet2 = xlWB.ActiveSheet

            ' Add table headers
            Dim Headers() As String = {"Item No.", "Description", "Quantity", "Sale Unit Price", "Sale Price", "Sale Used Price", "Used Price", "Rental Unit Price", "Rental Price", "Weight (lbs.)"}
            'Covers only the table headers for the new sheet
            For i As Integer = 1 To Headers.Length
                Sheet2.Cells(2, i) = Headers(i - 1)
            Next

            'Get the last row of the spreadsheet
            Dim lastRow As String = Sheet1.UsedRange.Rows.Count

            ' The number 17 was chosen because it is the start of the original sheet where the data begins
            ' The modification of the code for other versions of CADS Smart Scaffolder starts here for cells alignment. Modify the number 14.
            For j As Integer = 17 To lastRow - 4
                ' Create Item Nos.
                Sheet2.Cells(j - 14, 1) = (j - 16)
                ' Copy Descriptions
                Sheet2.Cells(j - 14, 2) = Sheet1.Cells(j, 4)
                ' Copy Quantities
                Sheet2.Cells(j - 14, 3) = Sheet1.Cells(j, 10)
                ' Copy weight in lbs.
                Sheet2.Cells(j - 14, 10) = Sheet1.Cells(j, 13)
                Dim weight As String = Sheet2.Cells(j - 14, 10).value().ToString
                weight = weight.Substring(0, weight.Length - 4)
                Sheet2.Cells(j - 14, 10) = weight

                ' Perform some cell formatting
                Sheet2.Range("A1:J1").EntireColumn.AutoFit()
                Sheet2.Range("A2:J2").Font.Bold = True
                Sheet2.Range("A2:J2").Interior.ColorIndex = 6

                ' Align Cells Centre
                Sheet2.Range("A1:J1").EntireColumn.VerticalAlignment = Excel.Constants.xlCenter
                Sheet2.Range("A1").EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter
                Sheet2.Range("C1").EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter
                ' Increase indent level
                Sheet2.Range("B1").EntireColumn.IndentLevel = 1

                ' Format cells as Accounting
                Sheet2.Cells(j - 14, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 14, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 14, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 14, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 14, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 14, 9).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 14, 10).NumberFormat = "#,##0.00"

                ' Declare variables to perform some calculations
                Dim value1 As Decimal = Sheet2.Cells(j - 14, 4).value()
                Dim value2 As Decimal = Sheet2.Cells(j - 14, 3).value()

                ' Calculate sale Values and Rentals
                Sheet2.Range("E" & (j - 14)).Formula = "=C" & (j - 14) & "*" & "D" & (j - 14)
                Sheet2.Range("G" & (j - 14)).Formula = "=F" & (j - 14) & "*" & "C" & (j - 14)
                Sheet2.Range("I" & (j - 14)).Formula = "=H" & (j - 14) & "*" & "C" & (j - 14)

                Dim materialDescription As String = Sheet2.Cells(j - 14, 2).value().ToString
                For Each pair In materialPricelist
                    If pair.Key = materialDescription Then
                        Sheet2.Cells(j - 14, 4) = pair.Value(0)
                        Sheet2.Cells(j - 14, 6) = pair.Value(1)
                        Sheet2.Cells(j - 14, 8) = pair.Value(2)
                    End If
                Next
            Next

            ' Format cells as Accounting and last Column as a Decimal
            For j As Integer = 13 To lastRow - 5
                Sheet2.Cells(j - 4, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 4, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 4, 9).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 4, 10).NumberFormat = "#,##0.00"
            Next

            ' Totals of each column for sale, used, rental and weight
            Sheet2.Range("E" & (lastRow - 16)).Formula = "=SUM(E3:" & "E" & (lastRow - 17)
            Sheet2.Range("G" & (lastRow - 16)).Formula = "=SUM(G3:" & "G" & (lastRow - 17)
            Sheet2.Range("I" & (lastRow - 16)).Formula = "=SUM(I3:" & "I" & (lastRow - 17)
            Sheet2.Range("J" & (lastRow - 16)).Formula = "=SUM(J3:" & "J" & (lastRow - 17)

            ' Convert weight to kg
            Dim weightInPounds As Decimal = Sheet2.Range("J" & (lastRow - 16)).Value()
            Sheet2.Range("J" & (lastRow - 14)).Formula = weightInPounds / 2.20462

            ' Convert weight to tons
            Sheet2.Range("J" & (lastRow - 12)).Formula = weightInPounds / 2204.62

            ' Format these cells bold
            Sheet2.Range("E" & (lastRow - 16)).Font.Bold = True
            Sheet2.Range("G" & (lastRow - 16)).Font.Bold = True
            Sheet2.Range("I" & (lastRow - 16)).Font.Bold = True
            Sheet2.Range("J" & (lastRow - 16)).Font.Bold = True
            Sheet2.Range("J" & (lastRow - 14)).Font.Bold = True
            Sheet2.Range("J" & (lastRow - 12)).Font.Bold = True

            Sheet2.Range("I" & (lastRow - 14)).Font.Bold = True
            Sheet2.Range("I" & (lastRow - 12)).Font.Bold = True

            ' Make Background color Yellow
            Sheet2.Range("E" & (lastRow - 16)).Interior.ColorIndex = 6
            Sheet2.Range("G" & (lastRow - 16)).Interior.ColorIndex = 6
            Sheet2.Range("I" & (lastRow - 16)).Interior.ColorIndex = 6
            Sheet2.Range("J" & (lastRow - 16)).Interior.ColorIndex = 6
            Sheet2.Range("J" & (lastRow - 14)).Interior.ColorIndex = 6
            Sheet2.Range("J" & (lastRow - 12)).Interior.ColorIndex = 6

            ' Format Weight cells to 2 Decimal places
            Sheet2.Range("J" & (lastRow - 14)).NumberFormat = "#,##0.00"
            Sheet2.Range("J" & (lastRow - 12)).NumberFormat = "#,##0.00"

            ' Include some description for weight cells
            Sheet2.Range("I" & (lastRow - 14)).Value() = "Weight(Kg)"
            Sheet2.Range("I" & (lastRow - 12)).Value() = "Weight(Ton)"

            ' Autofit columns again
            Sheet2.Range("A1:J1").EntireColumn.AutoFit()

            Try
                Dim newFileName As String = xlFileName.Substring(0, xlFileName.Length - 5)
                'Save new workbook
                Dim finalFileName As String = newFileName & " - " & Format(currentDate, "yyyy-MM-dd") & ".xlsx"
                xlWB.SaveAs(finalFileName)
                frmPricing.txtAttachment.Text = finalFileName
                Dim emailMessage As String
                emailMessage =
                "Good day," + vbNewLine +
                "Please find attached quotation breakdown." + vbNewLine + vbNewLine +
                "With Kind Regards," + vbNewLine +
                "RADIAN H.A. Limited" + vbNewLine +
                "187 Helen Street, Marabella," + vbNewLine +
                "Trinidad and Tobago W.I." + vbCrLf + vbNewLine +
                "T:- +1-868-658-0293/8294" + vbNewLine +
                "F: - +1-868-658-5946" + vbNewLine +
                "Email:- operations@rhatt.com" + vbNewLine +
                "Web:- www.radianhaltd.com" + vbNewLine

                frmPricing.txtMessage.Text = emailMessage
                MessageBox.Show("Excel Sheet Created Successfully!", "File Created", MessageBoxButtons.OK, MessageBoxIcon.Information)
                newFileName = Nothing
            Catch ex As Exception
                MessageBox.Show("There was an error saving your file!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            ' Release object references
            Sheet1 = Nothing
            Sheet2 = Nothing
            xlWB = Nothing
            xlApp.Quit()
            xlApp = Nothing
            xlFileName = Nothing

        End If
    End Sub
    Sub OpenPricedFile()
        If frmPricing.txtAttachment.Text = "" Then
            MsgBox("Please input a file to price first!")
        Else
            Process.Start(frmPricing.txtAttachment.Text)
        End If

    End Sub
    Sub BrowseFiles()
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "RADIAN Material Pricing"
        fd.InitialDirectory = "C:\"

        fd.Filter = "Excel Files|*.xlsx"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            frmPricing.txtFileName.Text = strFileName

        End If
        ' Dispose of the instance
        fd.Dispose()
    End Sub
    Sub AttachFiles()
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Select Attachment"
        fd.InitialDirectory = "C:\"

        fd.Filter = "Excel Files|*.xlsx"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            frmPricing.txtAttachment.Text = strFileName

        End If
        ' Dispose of the instance
        fd.Dispose()
    End Sub
    Sub GenerateQuote()
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        'Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara3, oPara4, oPara5, oPara6 As Word.Paragraph

        Dim xlFileName As String = frmPricing.txtAttachment.Text
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWB As Excel.Workbook = xlApp.Workbooks.Open(xlFileName)
        Dim Sheet1 As Excel.Worksheet = xlWB.Sheets("Sheet1")

        ' Make Excel Visible
        xlApp.Visible = True

        'Get the last row of the spreadsheet
        Dim lastRow As String = Sheet1.UsedRange.Rows.Count

        'Select cells
        Sheet1.Range("A1:J" & lastRow).Select()

        'Copy text to clipboard
        Sheet1.Range("A1:J" & lastRow).Copy()

        'Start Word
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document
        oPara1 = oDoc.Content.Paragraphs.Add
        With oPara1.Range.Font
            .Name = "Times New Roman"
            .Size = 9
        End With
        With oPara1.Range
            .Text = "Attorney: A.J. Bhajan & Company" & vbCrLf & "Banker: Scotia Bank Limited" & vbCrLf
        End With
        oPara1.Range.InsertParagraphAfter()

        'Insert another paragraph
        oPara2 = oDoc.Content.Paragraphs.Add
        With oPara2.Range.Font
            .Name = "Times New Roman"
            .Size = 12
        End With
        With oPara2.Range
            .Text = "REF#: ####/" & QuotedInitials() & vbCrLf & GenerateDate() + vbCrLf
        End With
        '24 pt spacing after paragraph
        oPara2.Range.InsertParagraphAfter()

        'Insert another paragraph
        oPara3 = oDoc.Content.Paragraphs.Add
        With oPara3.Range.Font
            .Name = "Times New Roman"
            .Size = 12
            .Bold = True
        End With
        With oPara3.Range
            .Text = frmGenerateQuote.cmbCompany.Text + vbCrLf + vbCrLf +
            QuoteType() + vbCrLf +
            "ATTENTION: " + frmGenerateQuote.txtTo.Text.ToUpper + vbCrLf
        End With
        oPara3.Range.InsertParagraphAfter()

        'Insert another paragraph
        oPara4 = oDoc.Content.Paragraphs.Add
        With oPara4.Range.Font
            .Name = "Times New Roman"
            .Size = 12
        End With
        With oPara4.Range
            .Text = "Dear Sir/Madam," & vbCrLf +
            "Thank you for the opportunity to provide a quote on the SALE/RENTAL of the following Scaffolding items." + " " +
            "Please see pricing below:" & vbCrLf +
            "NOTE:" & vbCrLf +
            "1) Quotation is valid for 30 days." & vbCrLf +
            "2) All scaffold equipment and access products are in accordance with British Standards and OSHA." & vbCrLf +
            "3) Should the above quote meet your approval, please follow with a Purchase Order and payment." +
            vbCrLf + QuoteNotes() +
            "We thank you for your interest and look forward to your continued business." & vbCrLf
        End With
        oPara4.Range.InsertParagraphAfter()

        'Insert another paragraph
        oPara5 = oDoc.Content.Paragraphs.Add
        With oPara5.Range.Font
            .Name = "Times New Roman"
            .Size = 12
        End With
        With oPara5.Range
            .Text = "Yours faithfully," & vbCrLf +
            "RADIAN H.A. LTD" & vbCrLf & vbCrLf +
            "........................................." + vbCrLf
        End With
        '24 pt spacing after paragraph
        oPara5.Range.InsertParagraphAfter()

        'Insert another paragraph
        oPara6 = oDoc.Content.Paragraphs.Add
        With oPara6.Range.Font
            .Name = "Times New Roman"
            .Size = 12
        End With
        With oPara6.Range
            .bold = True
            .Text = QuotedBy()
        End With
        '24 pt spacing after paragraph
        oPara6.Range.InsertParagraphAfter()

        'Move the selection point to end of document
        Dim what As Object = Word.WdGoToItem.wdGoToLine
        Dim which As Object = Word.WdGoToDirection.wdGoToLast
        oWord.Selection.GoTo(what, which, Nothing, Nothing)

        'Insert a new page in word document
        'oWord.Selection.MoveEnd()
        oWord.Selection.InsertNewPage()

        'Make page landscape
        oPara6.Range.InsertBreak()
        oDoc.Range.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape
        oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape



        'Paste clipboard contents

        ' Release Excel object references
        Sheet1 = Nothing
        xlWB = Nothing
        xlApp.Quit()
        xlApp = Nothing
        xlFileName = Nothing

        'Release Word object references
        'oword = nothing
        'odoc=nothing
        'opara1=nothing

        'Clear clipboard
        Clipboard.Clear()

        MsgBox("Document Created!")

    End Sub
    Sub CentreForm(ByVal frm As Form)
        'Centre window on load
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        Dim formNewLeft As Integer = Val(Val(screenWidth) / Val(2)) - Val(Val(frm.Width) / Val(2))
        Dim formNewTop As Integer = Val(Val(screenHeight) / Val(2)) - Val(Val(frm.Height) / Val(2))
        frm.Left = formNewLeft
        frm.Top = formNewTop
    End Sub
    Function GenerateDate()
        If frmGenerateQuote.rbCurrentDate.Checked = True Then
            Return Format(Date.Now, "dd MMMM yyyy")
        Else
            Return Format(frmGenerateQuote.dtpDate.Value, "dd MMMM yyyy")
        End If
    End Function
    Function QuotedBy()
        If frmGenerateQuote.rbAlicia.Checked = True Then
            Return "Mrs. Alicia Singh-Biran" & vbCrLf + "Quotation/Training Coordinator" & vbCrLf
        ElseIf frmGenerateQuote.rbArnold.Checked = True Then
            Return "Mr. Arnold Bradshaw" & vbCrLf + "Engineer/Software Developer" & vbCrLf
        Else
            Return "Ms. Cheryl Khan" & vbCrLf + "Accounts Receivables" & vbCrLf
        End If
    End Function
    Function QuotedInitials()
        If frmGenerateQuote.rbAlicia.Checked = True Then
            Return "ASB"
        ElseIf frmGenerateQuote.rbArnold.Checked = True Then
            Return "AB"
        Else
            Return "CCM"
        End If
    End Function
    Sub CompanyDetails()
        'Company Names
        Dim CompanyNames = New List(Of KeyValuePair(Of String, String))
        CompanyNames.Add(New KeyValuePair(Of String, String)("Woodgroup PLC",
                                                             "Digi Lane, CHAGUANAS"))
        CompanyNames.Add(New KeyValuePair(Of String, String)("Stork Technical Servives Limited",
                                                             "403 Pacific Avenue"))

        For i As Integer = 0 To CompanyNames.Count - 1
            frmGenerateQuote.cmbCompany.Items.Add(CompanyNames(i).Key.ToUpper & vbCrLf & CompanyNames(i).Value)
        Next

    End Sub
    Function QuoteNotes()
        Return frmGenerateQuote.txtQuoteNotes.Text & vbCrLf
    End Function
    Function QuoteType()
        If frmGenerateQuote.rbSale.Checked = True Then
            Return "RE: SALE OF SCAFFOLD MATERIALS"
        ElseIf frmGenerateQuote.rbSaleRental.Checked = True Then
            Return "RE: SALE/RENTAL OF SCAFFOLD MATERIAL"
        ElseIf frmGenerateQuote.rbRental.Checked = True Then
            Return "RE: RENTAL OF SCAFFOLD MATERIAL"
        ElseIf frmGenerateQuote.rbConsumables.Checked = True Then
            Return "RE: SALE OF CONSUMABLEs"
        Else
            Return "RE: ENGINEERING DESIGN"
        End If
    End Function
End Class