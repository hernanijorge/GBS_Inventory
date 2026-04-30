Partial Class frmAddEquipamento
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private Sub InitializeComponent()

        Me.pnlFooter       = New System.Windows.Forms.Panel()
        Me.btnSave         = New System.Windows.Forms.Button()
        Me.btnCancel       = New System.Windows.Forms.Button()
        Me.pnlBody         = New System.Windows.Forms.Panel()
        Me.lblNote         = New System.Windows.Forms.Label()
        Me.lblUID          = New System.Windows.Forms.Label()
        Me.txtUID          = New System.Windows.Forms.TextBox()
        Me.lblManuf        = New System.Windows.Forms.Label()
        Me.cboManufacturer = New System.Windows.Forms.ComboBox()
        Me.lblModel        = New System.Windows.Forms.Label()
        Me.txtModel        = New System.Windows.Forms.TextBox()
        Me.lblSerial       = New System.Windows.Forms.Label()
        Me.txtSerial       = New System.Windows.Forms.TextBox()
        Me.lblProcessor    = New System.Windows.Forms.Label()
        Me.txtProcessor    = New System.Windows.Forms.TextBox()
        Me.lblRam          = New System.Windows.Forms.Label()
        Me.txtRam          = New System.Windows.Forms.TextBox()
        Me.lblStorage      = New System.Windows.Forms.Label()
        Me.txtStorage      = New System.Windows.Forms.TextBox()
        Me.lblCondition    = New System.Windows.Forms.Label()
        Me.cboCondition    = New System.Windows.Forms.ComboBox()
        Me.lblStatus       = New System.Windows.Forms.Label()
        Me.cboStatus       = New System.Windows.Forms.ComboBox()
        Me.lblBatch        = New System.Windows.Forms.Label()
        Me.txtBatch        = New System.Windows.Forms.TextBox()
        Me.lblNotes        = New System.Windows.Forms.Label()
        Me.txtNotes        = New System.Windows.Forms.TextBox()

        Me.pnlFooter.SuspendLayout()
        Me.SuspendLayout()

        ' ── Form ──────────────────────────────────────────────────────
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        Me.AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize          = New System.Drawing.Size(500, 490)
        Me.FormBorderStyle     = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox         = False
        Me.MinimizeBox         = False
        Me.StartPosition       = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text                = "Add Equipment"
        Me.Font                = New System.Drawing.Font("Segoe UI", 9.0F)

        ' ── pnlFooter (Dock=Bottom) ────────────────────────────────────
        Me.pnlFooter.Dock    = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFooter.Height  = 52

        Me.btnSave.Text      = "Save"
        Me.btnSave.Size      = New System.Drawing.Size(110, 32)
        Me.btnSave.Location  = New System.Drawing.Point(170, 10)
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat

        Me.btnCancel.Text      = "Cancel"
        Me.btnCancel.Size      = New System.Drawing.Size(110, 32)
        Me.btnCancel.Location  = New System.Drawing.Point(290, 10)
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat

        Me.pnlFooter.Controls.Add(Me.btnSave)
        Me.pnlFooter.Controls.Add(Me.btnCancel)

        ' ── pnlBody (Dock=Fill) ───────────────────────────────────────
        Me.pnlBody.Dock = System.Windows.Forms.DockStyle.Fill

        ' ── Layout: label col X=14, control col X=155, control W=310 ──
        ' Row gap = 34px.  All Y coords relative to pnlBody.

        ' Note
        Me.lblNote.AutoSize  = True
        Me.lblNote.Location  = New System.Drawing.Point(14, 10)
        Me.lblNote.Text      = "* = required field"
        Me.lblNote.ForeColor = System.Drawing.Color.FromArgb(120, 120, 120)
        Me.lblNote.Font      = New System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Italic)

        ' Internal UID *  — Y=38
        Me.lblUID.AutoSize = True
        Me.lblUID.Location = New System.Drawing.Point(14, 40)
        Me.lblUID.Text     = "Internal UID *"
        Me.txtUID.Location = New System.Drawing.Point(155, 38)
        Me.txtUID.Size     = New System.Drawing.Size(310, 24)

        ' Manufacturer *  — Y=72
        Me.lblManuf.AutoSize = True
        Me.lblManuf.Location = New System.Drawing.Point(14, 74)
        Me.lblManuf.Text     = "Manufacturer *"
        Me.cboManufacturer.Location          = New System.Drawing.Point(155, 72)
        Me.cboManufacturer.Size              = New System.Drawing.Size(310, 24)
        Me.cboManufacturer.DropDownStyle     = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.cboManufacturer.AutoCompleteMode  = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cboManufacturer.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems

        ' Model *  — Y=106
        Me.lblModel.AutoSize = True
        Me.lblModel.Location = New System.Drawing.Point(14, 108)
        Me.lblModel.Text     = "Model *"
        Me.txtModel.Location = New System.Drawing.Point(155, 106)
        Me.txtModel.Size     = New System.Drawing.Size(310, 24)

        ' Serial Number  — Y=140
        Me.lblSerial.AutoSize = True
        Me.lblSerial.Location = New System.Drawing.Point(14, 142)
        Me.lblSerial.Text     = "Serial Number"
        Me.txtSerial.Location = New System.Drawing.Point(155, 140)
        Me.txtSerial.Size     = New System.Drawing.Size(310, 24)

        ' Processor  — Y=174
        Me.lblProcessor.AutoSize = True
        Me.lblProcessor.Location = New System.Drawing.Point(14, 176)
        Me.lblProcessor.Text     = "Processor"
        Me.txtProcessor.Location = New System.Drawing.Point(155, 174)
        Me.txtProcessor.Size     = New System.Drawing.Size(310, 24)

        ' RAM (GB) + Storage (GB) on same row  — Y=208
        Me.lblRam.AutoSize = True
        Me.lblRam.Location = New System.Drawing.Point(14, 210)
        Me.lblRam.Text     = "RAM (GB)"
        Me.txtRam.Location = New System.Drawing.Point(155, 208)
        Me.txtRam.Size     = New System.Drawing.Size(70, 24)

        Me.lblStorage.AutoSize = True
        Me.lblStorage.Location = New System.Drawing.Point(238, 210)
        Me.lblStorage.Text     = "Storage (GB)"
        Me.txtStorage.Location = New System.Drawing.Point(330, 208)
        Me.txtStorage.Size     = New System.Drawing.Size(70, 24)

        ' Condition Status *  — Y=242
        Me.lblCondition.AutoSize = True
        Me.lblCondition.Location = New System.Drawing.Point(14, 244)
        Me.lblCondition.Text     = "Condition Status *"
        Me.cboCondition.Location     = New System.Drawing.Point(155, 242)
        Me.cboCondition.Size         = New System.Drawing.Size(200, 24)
        Me.cboCondition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCondition.Items.AddRange({"EXCELLENT", "GOOD", "FAIR", "POOR"})

        ' Status *  — Y=276
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(14, 278)
        Me.lblStatus.Text     = "Status *"
        Me.cboStatus.Location     = New System.Drawing.Point(155, 276)
        Me.cboStatus.Size         = New System.Drawing.Size(200, 24)
        Me.cboStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatus.Items.AddRange({"IN_STOCK", "IN_REPAIR", "SCRAPPED"})

        ' Source Batch  — Y=310
        Me.lblBatch.AutoSize = True
        Me.lblBatch.Location = New System.Drawing.Point(14, 312)
        Me.lblBatch.Text     = "Source Batch"
        Me.txtBatch.Location = New System.Drawing.Point(155, 310)
        Me.txtBatch.Size     = New System.Drawing.Size(310, 24)

        ' Notes  — Y=344, multiline H=72
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Location = New System.Drawing.Point(14, 346)
        Me.lblNotes.Text     = "Notes"
        Me.txtNotes.Location   = New System.Drawing.Point(155, 344)
        Me.txtNotes.Size       = New System.Drawing.Size(310, 72)
        Me.txtNotes.Multiline  = True
        Me.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical

        Me.pnlBody.Controls.AddRange(New System.Windows.Forms.Control() {
            Me.lblNote,
            Me.lblUID,       Me.txtUID,
            Me.lblManuf,     Me.cboManufacturer,
            Me.lblModel,     Me.txtModel,
            Me.lblSerial,    Me.txtSerial,
            Me.lblProcessor, Me.txtProcessor,
            Me.lblRam,       Me.txtRam,
            Me.lblStorage,   Me.txtStorage,
            Me.lblCondition, Me.cboCondition,
            Me.lblStatus,    Me.cboStatus,
            Me.lblBatch,     Me.txtBatch,
            Me.lblNotes,     Me.txtNotes
        })

        Me.Controls.Add(Me.pnlBody)
        Me.Controls.Add(Me.pnlFooter)

        Me.pnlFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlFooter       As System.Windows.Forms.Panel
    Friend WithEvents btnSave         As System.Windows.Forms.Button
    Friend WithEvents btnCancel       As System.Windows.Forms.Button
    Friend WithEvents pnlBody         As System.Windows.Forms.Panel
    Friend WithEvents lblNote         As System.Windows.Forms.Label
    Friend WithEvents lblUID          As System.Windows.Forms.Label
    Friend WithEvents txtUID          As System.Windows.Forms.TextBox
    Friend WithEvents lblManuf        As System.Windows.Forms.Label
    Friend WithEvents cboManufacturer As System.Windows.Forms.ComboBox
    Friend WithEvents lblModel        As System.Windows.Forms.Label
    Friend WithEvents txtModel        As System.Windows.Forms.TextBox
    Friend WithEvents lblSerial       As System.Windows.Forms.Label
    Friend WithEvents txtSerial       As System.Windows.Forms.TextBox
    Friend WithEvents lblProcessor    As System.Windows.Forms.Label
    Friend WithEvents txtProcessor    As System.Windows.Forms.TextBox
    Friend WithEvents lblRam          As System.Windows.Forms.Label
    Friend WithEvents txtRam          As System.Windows.Forms.TextBox
    Friend WithEvents lblStorage      As System.Windows.Forms.Label
    Friend WithEvents txtStorage      As System.Windows.Forms.TextBox
    Friend WithEvents lblCondition    As System.Windows.Forms.Label
    Friend WithEvents cboCondition    As System.Windows.Forms.ComboBox
    Friend WithEvents lblStatus       As System.Windows.Forms.Label
    Friend WithEvents cboStatus       As System.Windows.Forms.ComboBox
    Friend WithEvents lblBatch        As System.Windows.Forms.Label
    Friend WithEvents txtBatch        As System.Windows.Forms.TextBox
    Friend WithEvents lblNotes        As System.Windows.Forms.Label
    Friend WithEvents txtNotes        As System.Windows.Forms.TextBox

End Class
