Imports System.Windows.Forms
Imports System.Drawing

Partial Class frmShipmentReportOptions
    Inherits Form

#Region "Designer"

    Friend WithEvents dtpFrom            As DateTimePicker
    Friend WithEvents dtpTo              As DateTimePicker
    Friend WithEvents chkDelivered       As CheckBox
    Friend WithEvents chkInTransit       As CheckBox
    Friend WithEvents chkPending         As CheckBox
    Friend WithEvents chkIssues          As CheckBox
    Friend WithEvents cboRecipient       As ComboBox
    Friend WithEvents rbGroupNone        As RadioButton
    Friend WithEvents rbGroupByPrefix    As RadioButton
    Friend WithEvents rbGroupManualAlias As RadioButton
    Friend WithEvents pnlPrefix          As Panel
    Friend WithEvents nudPrefixLen       As NumericUpDown
    Friend WithEvents lblPrefixHint      As Label
    Friend WithEvents pnlAlias           As Panel
    Friend WithEvents btnEditMapping     As Button
    Friend WithEvents lblMappingStatus   As Label
    Friend WithEvents rbSummary          As RadioButton
    Friend WithEvents rbDetailed         As RadioButton
    Friend WithEvents lblPreview         As Label
    Friend WithEvents clbShipments       As CheckedListBox
    Friend WithEvents btnCheckAll        As Button
    Friend WithEvents btnUncheckAll      As Button
    Friend WithEvents lblSelecao         As Label
    Friend WithEvents btnGenerate        As Button
    Friend WithEvents btnCancel          As Button
    Friend WithEvents pnlFooter          As Panel

    Private Sub InitializeComponent()

        Me.Text            = "Shipment Report by Recipient"
        Me.Size            = New Size(930, 580)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.StartPosition   = FormStartPosition.CenterParent

        ' ════ Left column: filters and options ════════════════════════

        ' ── Date range (by shipment creation date) ───────────────────
        Dim grpDatas As New GroupBox() With { .Text = "Date range (created)", .Location = New Point(15, 10), .Size = New Size(425, 58) }
        Dim lblFrom As New Label() With { .Text = "From:", .Location = New Point(15, 26), .AutoSize = True }
        Me.dtpFrom = New DateTimePicker() With { .Location = New Point(60, 22), .Size = New Size(130, 24), .Format = DateTimePickerFormat.Custom, .CustomFormat = "yyyy-MM-dd" }
        Dim lblTo As New Label() With { .Text = "To:", .Location = New Point(220, 26), .AutoSize = True }
        Me.dtpTo   = New DateTimePicker() With { .Location = New Point(250, 22), .Size = New Size(130, 24), .Format = DateTimePickerFormat.Custom, .CustomFormat = "yyyy-MM-dd" }
        grpDatas.Controls.AddRange({lblFrom, dtpFrom, lblTo, dtpTo})

        ' ── Status groups ────────────────────────────────────────────
        Dim grpStatus As New GroupBox() With { .Text = "Include shipments with status", .Location = New Point(15, 74), .Size = New Size(425, 72) }
        Me.chkDelivered = New CheckBox() With { .Text = "Delivered",  .Location = New Point(15,  22), .AutoSize = True, .Checked = True }
        Me.chkInTransit = New CheckBox() With { .Text = "In Transit", .Location = New Point(215, 22), .AutoSize = True, .Checked = True }
        Me.chkPending   = New CheckBox() With { .Text = "Pending",    .Location = New Point(15,  46), .AutoSize = True, .Checked = True }
        Me.chkIssues    = New CheckBox() With { .Text = "Issues",     .Location = New Point(215, 46), .AutoSize = True, .Checked = True }
        grpStatus.Controls.AddRange({chkDelivered, chkInTransit, chkPending, chkIssues})

        Dim tip As New ToolTip()
        tip.SetToolTip(chkDelivered, "DELIVERED")
        tip.SetToolTip(chkInTransit, "PICKED_UP, IN_TRANSIT, OUT_FOR_DELIVERY")
        tip.SetToolTip(chkPending,   "LABEL_CREATED, OPEN")
        tip.SetToolTip(chkIssues,    "DELAYED, EXCEPTION, RETURNED")

        ' ── Recipient filter ─────────────────────────────────────────
        Dim lblRecipient As New Label() With { .Text = "Recipient (optional filter):", .Location = New Point(18, 154), .AutoSize = True }
        Me.cboRecipient = New ComboBox() With { .Location = New Point(18, 173), .Size = New Size(419, 24), .DropDownStyle = ComboBoxStyle.DropDownList }
        tip.SetToolTip(cboRecipient, "With client grouping on, picking any spelling brings the whole group")

        ' ── Client grouping ──────────────────────────────────────────
        Dim grpGrouping As New GroupBox() With { .Text = "Client grouping (merge similar names)", .Location = New Point(15, 205), .Size = New Size(425, 138) }
        Me.rbGroupNone        = New RadioButton() With { .Text = "No grouping (each variation separate)", .Location = New Point(15, 20), .AutoSize = True }
        Me.rbGroupByPrefix    = New RadioButton() With { .Text = "Auto-group by name prefix (COMPUGANGAS = COMPUGANGAS 15"" = ...)", .Location = New Point(15, 42), .AutoSize = True, .Checked = True }
        Me.rbGroupManualAlias = New RadioButton() With { .Text = "Custom mapping (I'll define which names belong together)", .Location = New Point(15, 64), .AutoSize = True }

        Me.pnlPrefix = New Panel() With { .Location = New Point(12, 88), .Size = New Size(405, 44) }
        Dim lblPrefix As New Label() With { .Text = "Minimum prefix chars to match:", .Location = New Point(3, 4), .AutoSize = True }
        Me.nudPrefixLen  = New NumericUpDown() With { .Location = New Point(200, 1), .Size = New Size(55, 24), .Minimum = 3, .Maximum = 15, .Value = 6 }
        Me.lblPrefixHint = New Label() With { .Text = "Names sharing the first N characters will be grouped", .Location = New Point(3, 26), .AutoSize = True,
                                              .ForeColor = TemaEscuro.TextoMutado, .Font = New Font("Segoe UI", 8) }
        Me.pnlPrefix.Controls.AddRange({lblPrefix, nudPrefixLen, lblPrefixHint})

        Me.pnlAlias = New Panel() With { .Location = New Point(12, 88), .Size = New Size(405, 44), .Visible = False }
        Dim lblMapping As New Label() With { .Text = "Client mapping:", .Location = New Point(3, 8), .AutoSize = True }
        Me.btnEditMapping   = New Button() With { .Text = "Edit Mapping...", .Location = New Point(105, 2), .Size = New Size(120, 28) }
        Me.lblMappingStatus = New Label() With { .Text = "0 mappings defined", .Location = New Point(235, 8), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.pnlAlias.Controls.AddRange({lblMapping, btnEditMapping, lblMappingStatus})

        grpGrouping.Controls.AddRange({rbGroupNone, rbGroupByPrefix, rbGroupManualAlias, pnlPrefix, pnlAlias})

        ' ── Detail level ─────────────────────────────────────────────
        Dim grpDetalhe As New GroupBox() With { .Text = "Detail level", .Location = New Point(15, 349), .Size = New Size(425, 70) }
        Me.rbSummary  = New RadioButton() With { .Text = "Summary only (1 line per recipient)", .Location = New Point(15, 20), .AutoSize = True, .Checked = True }
        Me.rbDetailed = New RadioButton() With { .Text = "Detailed (1 line per recipient + shipment breakdown below)", .Location = New Point(15, 43), .AutoSize = True }
        grpDetalhe.Controls.AddRange({rbSummary, rbDetailed})

        Me.lblPreview = New Label() With { .Text = "", .Location = New Point(18, 430), .Size = New Size(425, 36),
                                           .ForeColor = TemaEscuro.Accent, .Font = New Font("Segoe UI", 9, FontStyle.Bold) }

        ' ════ Right column: pick shipments ═══════════════════════════
        Dim grpRemessas As New GroupBox() With { .Text = "Shipments to include", .Location = New Point(455, 10), .Size = New Size(450, 455) }
        Me.btnCheckAll   = New Button() With { .Text = "Check all",   .Location = New Point(12, 20), .Size = New Size(90, 26) }
        Me.btnUncheckAll = New Button() With { .Text = "Uncheck all", .Location = New Point(108, 20), .Size = New Size(95, 26) }
        Me.lblSelecao    = New Label() With { .Text = "", .Location = New Point(215, 25), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.clbShipments  = New CheckedListBox() With { .Location = New Point(12, 52), .Size = New Size(426, 392), .CheckOnClick = True,
                                                       .IntegralHeight = False, .HorizontalScrollbar = True, .BorderStyle = BorderStyle.FixedSingle,
                                                       .BackColor = TemaEscuro.SurfaceClaro, .ForeColor = TemaEscuro.Texto,
                                                       .Font = New Font("Consolas", 9) }
        grpRemessas.Controls.AddRange({btnCheckAll, btnUncheckAll, lblSelecao, clbShipments})

        ' ── Footer ───────────────────────────────────────────────────
        Me.pnlFooter   = New Panel() With { .Dock = DockStyle.Bottom, .Height = 55, .BackColor = TemaEscuro.Surface }
        Me.btnGenerate = New Button() With { .Text = "Generate Report", .Size = New Size(150, 34), .Location = New Point(645, 10),
                                              .BackColor = TemaEscuro.Accent, .ForeColor = TemaEscuro.Fundo, .FlatStyle = FlatStyle.Flat }
        Me.btnCancel   = New Button() With { .Text = "Cancel", .Size = New Size(100, 34), .Location = New Point(805, 10) }
        Me.pnlFooter.Controls.AddRange({btnGenerate, btnCancel})

        Me.AcceptButton = btnGenerate
        Me.CancelButton = btnCancel

        Me.Controls.AddRange({grpDatas, grpStatus, lblRecipient, cboRecipient, grpGrouping, grpDetalhe, lblPreview, grpRemessas, pnlFooter})

    End Sub

#End Region

End Class
