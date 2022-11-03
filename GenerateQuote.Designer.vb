<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGenerateQuote
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lblHeader = New System.Windows.Forms.Label()
        Me.btnQuote = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.rbArnold = New System.Windows.Forms.RadioButton()
        Me.rbAlicia = New System.Windows.Forms.RadioButton()
        Me.rbCheryl = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbCurrentDate = New System.Windows.Forms.RadioButton()
        Me.rbCustomDate = New System.Windows.Forms.RadioButton()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.cmbCompany = New System.Windows.Forms.ComboBox()
        Me.txtTo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtQuoteNotes = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.rbConsumables = New System.Windows.Forms.RadioButton()
        Me.rbSaleRental = New System.Windows.Forms.RadioButton()
        Me.rbSale = New System.Windows.Forms.RadioButton()
        Me.rbEngDesign = New System.Windows.Forms.RadioButton()
        Me.rbRental = New System.Windows.Forms.RadioButton()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Scaffold_Pricing.My.Resources.Resources.radian_logo
        Me.PictureBox1.Location = New System.Drawing.Point(24, 19)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(73, 71)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 18
        Me.PictureBox1.TabStop = False
        '
        'lblHeader
        '
        Me.lblHeader.AutoSize = True
        Me.lblHeader.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeader.Location = New System.Drawing.Point(237, 38)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(236, 22)
        Me.lblHeader.TabIndex = 17
        Me.lblHeader.Text = "Scaffold Material Pricing"
        '
        'btnQuote
        '
        Me.btnQuote.BackColor = System.Drawing.Color.DarkOrange
        Me.btnQuote.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnQuote.Location = New System.Drawing.Point(530, 414)
        Me.btnQuote.Name = "btnQuote"
        Me.btnQuote.Size = New System.Drawing.Size(90, 35)
        Me.btnQuote.TabIndex = 19
        Me.btnQuote.Text = "&Quote"
        Me.btnQuote.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.Tomato
        Me.btnClose.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(642, 414)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(90, 35)
        Me.btnClose.TabIndex = 20
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'rbArnold
        '
        Me.rbArnold.AutoSize = True
        Me.rbArnold.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbArnold.Location = New System.Drawing.Point(15, 24)
        Me.rbArnold.Name = "rbArnold"
        Me.rbArnold.Size = New System.Drawing.Size(125, 21)
        Me.rbArnold.TabIndex = 21
        Me.rbArnold.TabStop = True
        Me.rbArnold.Text = "Arnold Bradshaw"
        Me.rbArnold.UseVisualStyleBackColor = True
        '
        'rbAlicia
        '
        Me.rbAlicia.AutoSize = True
        Me.rbAlicia.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbAlicia.Location = New System.Drawing.Point(15, 47)
        Me.rbAlicia.Name = "rbAlicia"
        Me.rbAlicia.Size = New System.Drawing.Size(126, 21)
        Me.rbAlicia.TabIndex = 22
        Me.rbAlicia.TabStop = True
        Me.rbAlicia.Text = "Alicia Singh-Biran"
        Me.rbAlicia.UseVisualStyleBackColor = True
        '
        'rbCheryl
        '
        Me.rbCheryl.AutoSize = True
        Me.rbCheryl.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbCheryl.Location = New System.Drawing.Point(15, 70)
        Me.rbCheryl.Name = "rbCheryl"
        Me.rbCheryl.Size = New System.Drawing.Size(125, 21)
        Me.rbCheryl.TabIndex = 23
        Me.rbCheryl.TabStop = True
        Me.rbCheryl.Text = "Cheryl Chin Ming"
        Me.rbCheryl.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.PeachPuff
        Me.GroupBox1.Controls.Add(Me.rbArnold)
        Me.GroupBox1.Controls.Add(Me.rbCheryl)
        Me.GroupBox1.Controls.Add(Me.rbAlicia)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(24, 106)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(169, 112)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Quoted By"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.PeachPuff
        Me.GroupBox2.Controls.Add(Me.rbCurrentDate)
        Me.GroupBox2.Controls.Add(Me.rbCustomDate)
        Me.GroupBox2.Controls.Add(Me.dtpDate)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(207, 106)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(157, 112)
        Me.GroupBox2.TabIndex = 25
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Quotation Date"
        '
        'rbCurrentDate
        '
        Me.rbCurrentDate.AutoSize = True
        Me.rbCurrentDate.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbCurrentDate.Location = New System.Drawing.Point(15, 24)
        Me.rbCurrentDate.Name = "rbCurrentDate"
        Me.rbCurrentDate.Size = New System.Drawing.Size(95, 21)
        Me.rbCurrentDate.TabIndex = 24
        Me.rbCurrentDate.TabStop = True
        Me.rbCurrentDate.Text = "Curent Date"
        Me.rbCurrentDate.UseVisualStyleBackColor = True
        '
        'rbCustomDate
        '
        Me.rbCustomDate.AutoSize = True
        Me.rbCustomDate.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbCustomDate.Location = New System.Drawing.Point(15, 47)
        Me.rbCustomDate.Name = "rbCustomDate"
        Me.rbCustomDate.Size = New System.Drawing.Size(101, 21)
        Me.rbCustomDate.TabIndex = 25
        Me.rbCustomDate.TabStop = True
        Me.rbCustomDate.Text = "Custom Date"
        Me.rbCustomDate.UseVisualStyleBackColor = True
        '
        'dtpDate
        '
        Me.dtpDate.CalendarFont = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpDate.CustomFormat = "dd-MMM-yyyy"
        Me.dtpDate.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(15, 76)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(121, 25)
        Me.dtpDate.TabIndex = 28
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.PeachPuff
        Me.GroupBox3.Controls.Add(Me.cmbCompany)
        Me.GroupBox3.Controls.Add(Me.txtTo)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(379, 106)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(353, 112)
        Me.GroupBox3.TabIndex = 28
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Company Information"
        '
        'cmbCompany
        '
        Me.cmbCompany.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCompany.FormattingEnabled = True
        Me.cmbCompany.Location = New System.Drawing.Point(58, 36)
        Me.cmbCompany.Name = "cmbCompany"
        Me.cmbCompany.Size = New System.Drawing.Size(234, 25)
        Me.cmbCompany.TabIndex = 31
        '
        'txtTo
        '
        Me.txtTo.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTo.Location = New System.Drawing.Point(58, 70)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(234, 25)
        Me.txtTo.TabIndex = 31
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(27, 73)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 17)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "To:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 17)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Name:"
        '
        'txtQuoteNotes
        '
        Me.txtQuoteNotes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuoteNotes.Location = New System.Drawing.Point(207, 245)
        Me.txtQuoteNotes.Multiline = True
        Me.txtQuoteNotes.Name = "txtQuoteNotes"
        Me.txtQuoteNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtQuoteNotes.Size = New System.Drawing.Size(525, 159)
        Me.txtQuoteNotes.TabIndex = 29
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(204, 225)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 17)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "Quotation Notes"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.PeachPuff
        Me.GroupBox4.Controls.Add(Me.rbConsumables)
        Me.GroupBox4.Controls.Add(Me.rbSaleRental)
        Me.GroupBox4.Controls.Add(Me.rbSale)
        Me.GroupBox4.Controls.Add(Me.rbEngDesign)
        Me.GroupBox4.Controls.Add(Me.rbRental)
        Me.GroupBox4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(24, 245)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(169, 159)
        Me.GroupBox4.TabIndex = 25
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Quote Type"
        '
        'rbConsumables
        '
        Me.rbConsumables.AutoSize = True
        Me.rbConsumables.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbConsumables.Location = New System.Drawing.Point(15, 124)
        Me.rbConsumables.Name = "rbConsumables"
        Me.rbConsumables.Size = New System.Drawing.Size(104, 21)
        Me.rbConsumables.TabIndex = 25
        Me.rbConsumables.TabStop = True
        Me.rbConsumables.Text = "Consumables"
        Me.rbConsumables.UseVisualStyleBackColor = True
        '
        'rbSaleRental
        '
        Me.rbSaleRental.AutoSize = True
        Me.rbSaleRental.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbSaleRental.Location = New System.Drawing.Point(15, 97)
        Me.rbSaleRental.Name = "rbSaleRental"
        Me.rbSaleRental.Size = New System.Drawing.Size(91, 21)
        Me.rbSaleRental.TabIndex = 24
        Me.rbSaleRental.TabStop = True
        Me.rbSaleRental.Text = "Sale/Rental"
        Me.rbSaleRental.UseVisualStyleBackColor = True
        '
        'rbSale
        '
        Me.rbSale.AutoSize = True
        Me.rbSale.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbSale.Location = New System.Drawing.Point(15, 24)
        Me.rbSale.Name = "rbSale"
        Me.rbSale.Size = New System.Drawing.Size(50, 21)
        Me.rbSale.TabIndex = 21
        Me.rbSale.TabStop = True
        Me.rbSale.Text = "Sale"
        Me.rbSale.UseVisualStyleBackColor = True
        '
        'rbEngDesign
        '
        Me.rbEngDesign.AutoSize = True
        Me.rbEngDesign.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbEngDesign.Location = New System.Drawing.Point(15, 70)
        Me.rbEngDesign.Name = "rbEngDesign"
        Me.rbEngDesign.Size = New System.Drawing.Size(95, 21)
        Me.rbEngDesign.TabIndex = 23
        Me.rbEngDesign.TabStop = True
        Me.rbEngDesign.Text = "Eng. Design"
        Me.rbEngDesign.UseVisualStyleBackColor = True
        '
        'rbRental
        '
        Me.rbRental.AutoSize = True
        Me.rbRental.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbRental.Location = New System.Drawing.Point(15, 47)
        Me.rbRental.Name = "rbRental"
        Me.rbRental.Size = New System.Drawing.Size(62, 21)
        Me.rbRental.TabIndex = 22
        Me.rbRental.TabStop = True
        Me.rbRental.Text = "Rental"
        Me.rbRental.UseVisualStyleBackColor = True
        '
        'frmGenerateQuote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(758, 461)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtQuoteNotes)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnQuote)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblHeader)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "frmGenerateQuote"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Generate Quote"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Public WithEvents lblHeader As Label
    Friend WithEvents btnQuote As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents rbArnold As RadioButton
    Friend WithEvents rbAlicia As RadioButton
    Friend WithEvents rbCheryl As RadioButton
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents dtpDate As DateTimePicker
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents cmbCompany As ComboBox
    Friend WithEvents txtTo As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtQuoteNotes As TextBox
    Protected Friend WithEvents Label1 As Label
    Friend WithEvents rbCurrentDate As RadioButton
    Friend WithEvents rbCustomDate As RadioButton
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents rbConsumables As RadioButton
    Friend WithEvents rbSaleRental As RadioButton
    Friend WithEvents rbSale As RadioButton
    Friend WithEvents rbEngDesign As RadioButton
    Friend WithEvents rbRental As RadioButton
End Class
