Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.Configuration
Imports System.IO
Imports System.Net
Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word

Public Class MainClass

    Sub Close()
        frmPricing.Close()
    End Sub
    Sub LoadForm()
        'Set focus on the first textbox
        frmPricing.txtFileName.Select()
        'Centre Form
        CentreForm()
        'Load Default email address
        frmPricing.txtTo.Text = "training@rhatt.com"
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
            e_mail.From = New MailAddress("training@rhatt.com")
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

            MessageBox.Show("Please select a file!", "Error selecting file", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else

            Dim xlFileName As String = frmPricing.txtFileName.Text
            Dim xlApp As Excel.Application = New Excel.Application()
            Dim xlWB As Excel.Workbook = xlApp.Workbooks.Open(xlFileName)
            Dim Sheet1 As Excel.Worksheet = xlWB.Sheets("Sheet")
            Dim currentDate As Date = Date.Today

            ' Materials Pricelist
            Dim materialPricelist = New List(Of KeyValuePair(Of String, Decimal()))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok combined jack and base plate", {150.0, 125.0, 1.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok deck adaptor", {200.0, 150.0, 1.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder hatch", {1000.0, 900.0, 15.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok ladder safety gate: 0.8m", {1000.0, 900.0, 15.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Swivel coupler", {65, 45.0, 0.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Base plate", {55.0, 35.0, 0.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Double coupler", {60.0, 40.0, 0.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Putlog coupler", {60.0, 40.0, 0.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Sleeve coupler", {60.0, 40.0, 0.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Sole pads", {60.0, 40.0, 0.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Toe board clip", {60.0, 40.0, 0.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ring bolt", {0.0, 0.0, 0.0}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 0.8m", {150.0, 130.0, 1.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 0.9m", {175.0, 150.0, 1.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 1.3m", {200.0, 175.0, 1.75}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 1.8m", {300.0, 260.0, 2.6}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 2.5m", {380.0, 330.0, 3.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok horizontal: 3.0m", {450.0, 400.0, 4.0}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok intermediate transom: 1.3m", {300.0, 0, 2.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok intermediate transom: 1.8m", {475.0, 0, 3.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok intermediate transom: 2.5m", {650.0, 0, 4.0}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok standard: 1.0m", {400.0, 325.0, 3.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok standard: 1.52m", {425.0, 345.0, 3.3}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok standard: 2.0m", {450.0, 360.0, 3.6}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok standard: 2.5m", {560.0, 450.0, 4.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok standard: 3.0m", {655.0, 550.0, 5.5}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 1.8 x 2.0m", {415.0, 300.0, 3.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 2.5 x 1.5m", {450.0, 350.0, 3.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 2.5 x 2.0m", {450.0, 350.0, 3.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Cuplok swivel face brace: 3.0 x 2.0m", {450.0, 400.0, 4.0}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: steel 2.5m", {875.0, 750.0, 12.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: steel 3.0m", {1050.0, 960.0, 15.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: steel 3.5m", {1225.0, 1120.0, 18.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: steel 5.0m", {1750.0, 1600.0, 24.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: steel 6.0m", {2100.0, 1920.0, 30.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: steel 7.0m", {2450.0, 2240.0, 34.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: steel 8.0m", {2800.0, 2560.0, 39.0}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel board: 2.5m", {350.0, 0.0, 4.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel board: 1.8m", {275.0, 0.0, 3.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel board: 1.2m", {200.0, 0.0, 2.5}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube:  2'", {50.0, 43.0, 0.6}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube:  4'", {100.0, 86.0, 1.2}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube:  5'", {125.0, 107.5, 1.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube:  6'", {150.0, 129.0, 1.8}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube:  8'", {200.0, 172.0, 2.4}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube: 10'", {250.0, 215.0, 3.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube: 12'", {300.0, 258.0, 3.6}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube: 14'", {350.0, 301.0, 4.2}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube: 16'", {400.0, 344.0, 4.8}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube: 18'", {450.0, 387.0, 5.4}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube: 20'", {500.0, 430.0, 6.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Steel tube: 21'", {525.0, 451.5, 6.3}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 2.0m", {700.0, 600.0, 9.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 2.5m", {875.0, 750.0, 12.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 3.0m", {1050.0, 960.0, 15.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 3.5m", {1225.0, 1120.0, 18}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 4.0m", {1400.0, 1280.0, 19.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 5.0m", {1750.0, 1600.0, 24.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 6.0m", {2100.0, 1920.0, 30.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 7.0m", {2450.0, 2240.0, 34.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Ladder: orange painted steel 8.0m", {2800.0, 2560.0, 39.0}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  1'", {30.0, 20.0, 0.25}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  2'", {60.0, 40.0, 0.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  4'", {135.0, 100.0, 1.25}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  6'", {200.0, 150.0, 2.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  8'", {250.0, 200.0, 2.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board: 10'", {300.0, 250.0, 3.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board: 13'", {325.0, 275.0, 3.25}))

            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  2' - 225mm x 38mm", {60.0, 40.0, 0.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  4' - 225mm x 38mm", {135.0, 100.0, 1.25}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  6' - 225mm x 38mm", {200.0, 150.0, 2.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board:  8' - 225mm x 38mm", {250.0, 200.0, 2.5}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board: 10' - 225mm x 38mm", {300.0, 250.0, 3.0}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board: 12' - 225mm x 38mm", {325.0, 275.0, 3.25}))
            materialPricelist.Add(New KeyValuePair(Of String, Decimal())("Timber board: 13' - 225mm x 38mm", {325.0, 275.0, 3.25}))

            ' Make Excel Visible
            xlApp.Visible = False

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
                ' Align cells centre
                Sheet2.Range("A1:J1").EntireColumn.VerticalAlignment = Excel.Constants.xlCenter
                Sheet2.Range("A1").EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter
                Sheet2.Range("C1").EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter
                ' Increase indent level
                Sheet2.Range("B1").EntireColumn.IndentLevel = 1

                ' Format cells as Accounting
                Sheet2.Cells(j - 10, 4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 10, 5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 10, 6).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 10, 7).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 10, 8).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 10, 9).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
                Sheet2.Cells(j - 10, 10).NumberFormat = "#,##0.00"

                ' Declare variables to perform some calculations
                Dim value1 As Decimal = Sheet2.Cells(j - 10, 4).value()
                Dim value2 As Decimal = Sheet2.Cells(j - 10, 3).value()

                ' Calculate sale values and rentals
                Sheet2.Range("E" & (j - 10)).Formula = "=C" & (j - 10) & "*" & "D" & (j - 10)
                Sheet2.Range("G" & (j - 10)).Formula = "=F" & (j - 10) & "*" & "C" & (j - 10)
                Sheet2.Range("I" & (j - 10)).Formula = "=H" & (j - 10) & "*" & "C" & (j - 10)

                Dim materialDescription As String = Sheet2.Cells(j - 10, 2).value().ToString
                For Each pair In materialPricelist
                    If pair.Key = materialDescription Then
                        Sheet2.Cells(j - 10, 4) = pair.Value(0)
                        Sheet2.Cells(j - 10, 6) = pair.Value(1)
                        Sheet2.Cells(j - 10, 8) = pair.Value(2)
                    End If
                Next
            Next

            ' Totals of each column for sale, used, rental and weight
            Sheet2.Range("E" & (lastRow - 12)).Formula = "=SUM(E2:" & "E" & (lastRow - 13)
            Sheet2.Range("G" & (lastRow - 12)).Formula = "=SUM(G2:" & "G" & (lastRow - 13)
            Sheet2.Range("I" & (lastRow - 12)).Formula = "=SUM(I2:" & "I" & (lastRow - 13)
            Sheet2.Range("J" & (lastRow - 12)).Formula = "=SUM(J2:" & "J" & (lastRow - 13)

            ' Convert weight to kg
            Dim weightInPounds As Decimal = Sheet2.Range("J" & (lastRow - 12)).Value()
            Sheet2.Range("J" & (lastRow - 10)).Formula = weightInPounds / 2.20462

            ' Convert weight to tons
            Sheet2.Range("J" & (lastRow - 8)).Formula = weightInPounds / 2204.62

            ' Format these cells bold
            Sheet2.Range("E" & (lastRow - 12)).Font.Bold = True
            Sheet2.Range("G" & (lastRow - 12)).Font.Bold = True
            Sheet2.Range("I" & (lastRow - 12)).Font.Bold = True
            Sheet2.Range("J" & (lastRow - 12)).Font.Bold = True
            Sheet2.Range("J" & (lastRow - 10)).Font.Bold = True
            Sheet2.Range("J" & (lastRow - 8)).Font.Bold = True

            Sheet2.Range("I" & (lastRow - 10)).Font.Bold = True
            Sheet2.Range("I" & (lastRow - 8)).Font.Bold = True

            ' Make Background color Yellow
            Sheet2.Range("E" & (lastRow - 12)).Interior.ColorIndex = 6
            Sheet2.Range("G" & (lastRow - 12)).Interior.ColorIndex = 6
            Sheet2.Range("I" & (lastRow - 12)).Interior.ColorIndex = 6
            Sheet2.Range("J" & (lastRow - 12)).Interior.ColorIndex = 6
            Sheet2.Range("J" & (lastRow - 10)).Interior.ColorIndex = 6
            Sheet2.Range("J" & (lastRow - 8)).Interior.ColorIndex = 6

            ' Format weight cells to 2 Decimal places
            Sheet2.Range("J" & (lastRow - 10)).NumberFormat = "#,##0.00"
            Sheet2.Range("J" & (lastRow - 8)).NumberFormat = "#,##0.00"

            ' Include some description for weight cells
            Sheet2.Range("I" & (lastRow - 10)).Value() = "Weight(Kg)"
            Sheet2.Range("I" & (lastRow - 8)).Value() = "Weight(Ton)"

            ' Autofit columns again
            Sheet2.Range("A1:J1").EntireColumn.AutoFit()

            Try
                Dim newFileName As String = xlFileName.Substring(0, xlFileName.Length - 5)
                'Save new workbook
                xlWB.SaveAs(newFileName & " - " & Format(currentDate, "yyyy-MM-dd") & ".xlsx")
                MessageBox.Show("Complete!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
        Dim oPara1, oPara2, oPara3 As Word.Paragraph
        Dim oPara4 As Word.Paragraph
        'Dim oRng As Word.Range

        'Start Word
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document
        oPara1 = oDoc.Content.Paragraphs.Add
        With oPara1.Range.Font
            .Name = "Times New Roman"
            .Size = 12
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
            .Text = "REF#: ####/ASB" & vbCrLf & Format(Date.Now(), "MMMM dd, yyyy") & vbCrLf
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
            .Text = "COMPANY NAME" + vbCrLf + "Company Address," + vbCrLf + "PORT-OF-SPAIN" + vbCr + vbCr +
            "RE: SALE/RENTAL OF SCAFFOLD MATERIALS" + vbCrLf +
            "ATTENTION: " + vbCrLf + vbCrLf
        End With
        oPara3.Range.InsertParagraphAfter()

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
            "3) Should the above quote meet your approval, please follow with a Purchase Order and payment." & vbCrLf & vbCrLf +
            "We thank you for your interest and look forward to your continued business." & vbCrLf + vbCrLf +
            "Yours faithfully," & vbCrLf +
            "RADIAN H.A. LTD" & vbCrLf & vbCrLf +
            "........................................." + vbCrLf +
            "Mrs. Alicia Singh-Biran" & vbCrLf +
            "Quotation/ Training Coordinator" & vbCrLf
        End With
        oPara4.Range.InsertParagraphAfter()
        MsgBox("Document Created!")

    End Sub
    Sub CentreForm()
        'Centre window on load
        Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height
        Dim formNewLeft As Integer = Val(Val(screenWidth) / Val(2)) - Val(Val(frmPricing.Width) / Val(2))
        Dim formNewTop As Integer = Val(Val(screenHeight) / Val(2)) - Val(Val(frmPricing.Height) / Val(2))
        frmPricing.Left = formNewLeft
        frmPricing.Top = formNewTop
    End Sub
End Class
