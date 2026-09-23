Imports System.Windows.Forms
Imports System.Drawing

Partial Class frmAddComponent
    Inherits Form

#Region "Designer"

    Friend WithEvents cboType        As ComboBox
    Friend WithEvents cboCapacity    As ComboBox
    Friend WithEvents cboGeneration  As ComboBox
    Friend WithEvents cboSpeed       As ComboBox
    Friend WithEvents cboBrand       As ComboBox
    Friend WithEvents txtModel       As TextBox
    Friend WithEvents txtPartNumber  As TextBox
    Friend WithEvents cboCondition   As ComboBox
    Friend WithEvents cboStatus      As ComboBox
    Friend WithEvents cboSourceBatch As ComboBox
    Friend WithEvents txtNotes       As TextBox
    Friend WithEvents btnSave        As Button
    Friend WithEvents btnCancel      As Button
    Friend WithEvents pnlFooter      As Panel
    Friend WithEvents lblSpeedLabel  As Label
    Friend WithEvents nudQuantity    As NumericUpDown
    Friend WithEvents lblCap         As Label
    Friend WithEvents lblGen         As Label
    Friend WithEvents lblCpu         As Label
    Friend WithEvents txtCpu         As TextBox
    Friend WithEvents lblStorage     As Label
    Friend WithEvents txtStorage     As TextBox

    Private Sub InitializeComponent()

        Me.Text            = "Add Component"
        Me.Size            = New Size(620, 520)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.StartPosition   = FormStartPosition.CenterParent

        Dim lblTitle As New Label() With {
            .Text      = "New Component",
            .Font      = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .Location  = New Point(20, 15),
            .AutoSize  = True
        }

        ' ── Row 1: Type | Capacity | Generation ──────────────────────
        Dim lblType As New Label() With { .Text = "Type *",         .Location = New Point(20,  65), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboType       = New ComboBox() With { .Location = New Point(20,  85), .Size = New Size(140, 24), .DropDownStyle = ComboBoxStyle.DropDownList }

        Me.lblCap        = New Label() With  { .Text = "Capacity (GB) *",.Location = New Point(180, 65), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboCapacity   = New ComboBox() With { .Location = New Point(180, 85), .Size = New Size(140, 24), .DropDownStyle = ComboBoxStyle.DropDown }

        Me.lblGen        = New Label() With  { .Text = "Generation",     .Location = New Point(340, 65), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboGeneration = New ComboBox() With { .Location = New Point(340, 85), .Size = New Size(140, 24), .DropDownStyle = ComboBoxStyle.DropDownList }

        Me.lblCpu        = New Label() With  { .Text = "CPU *",         .Location = New Point(340, 65), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado, .Visible = False }
        Me.txtCpu        = New TextBox() With { .Location = New Point(340, 85), .Size = New Size(140, 24), .Visible = False }

        ' ── Row 2: Speed | Brand | Model ─────────────────────────────
        Me.lblSpeedLabel = New Label()   With { .Text = "Speed (MHz)",  .Location = New Point(20,  130), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboSpeed      = New ComboBox() With { .Location = New Point(20,  150), .Size = New Size(140, 24), .DropDownStyle = ComboBoxStyle.DropDownList }

        Me.lblStorage    = New Label() With  { .Text = "Storage (GB)",.Location = New Point(20, 130), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado, .Visible = False }
        Me.txtStorage    = New TextBox() With { .Location = New Point(20, 150), .Size = New Size(140, 24), .Visible = False }

        Dim lblBrand As New Label() With { .Text = "Brand",         .Location = New Point(180, 130), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboBrand      = New ComboBox() With { .Location = New Point(180, 150), .Size = New Size(140, 24), .DropDownStyle = ComboBoxStyle.DropDown }

        Dim lblModel As New Label() With { .Text = "Model",         .Location = New Point(340, 130), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.txtModel      = New TextBox()  With { .Location = New Point(340, 150), .Size = New Size(230, 24), .MaxLength = 100 }

        ' ── Row 3: Part Number | Quantity ────────────────────────────
        Dim lblPart As New Label() With  { .Text = "Part Number",   .Location = New Point(20,  195), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.txtPartNumber = New TextBox()  With { .Location = New Point(20,  215), .Size = New Size(300, 24) }

        Dim lblQty As New Label() With   { .Text = "Quantity",      .Location = New Point(340, 195), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.nudQuantity   = New NumericUpDown() With { .Location = New Point(340, 215), .Size = New Size(100, 24),
                                                       .Minimum = 1, .Maximum = 9999, .Value = 1, .DecimalPlaces = 0 }

        ' ── Row 4: Condition | Status | Source Batch ─────────────────
        Dim lblCond As New Label() With  { .Text = "Condition *",   .Location = New Point(20,  260), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboCondition  = New ComboBox() With { .Location = New Point(20,  280), .Size = New Size(140, 24), .DropDownStyle = ComboBoxStyle.DropDownList }

        Dim lblStat As New Label() With  { .Text = "Status *",      .Location = New Point(180, 260), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboStatus     = New ComboBox() With { .Location = New Point(180, 280), .Size = New Size(140, 24), .DropDownStyle = ComboBoxStyle.DropDownList }

        Dim lblBatch As New Label() With { .Text = "Source Batch",  .Location = New Point(340, 260), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboSourceBatch= New ComboBox() With { .Location = New Point(340, 280), .Size = New Size(230, 24), .DropDownStyle = ComboBoxStyle.DropDown }

        ' ── Row 5: Notes ─────────────────────────────────────────────
        Dim lblNotes As New Label() With { .Text = "Notes",         .Location = New Point(20,  320), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.txtNotes      = New TextBox()  With { .Location = New Point(20, 340), .Size = New Size(550, 50), .Multiline = True }

        ' ── Footer ───────────────────────────────────────────────────
        Me.pnlFooter = New Panel() With { .Dock = DockStyle.Bottom, .Height = 55, .BackColor = TemaEscuro.Surface }
        Me.btnSave   = New Button() With { .Text = "Save",   .Size = New Size(120, 34), .Location = New Point(320, 10),
                                            .BackColor = TemaEscuro.Accent, .ForeColor = TemaEscuro.Fundo, .FlatStyle = FlatStyle.Flat }
        Me.btnCancel = New Button() With { .Text = "Cancel", .Size = New Size(100, 34), .Location = New Point(450, 10) }
        Me.pnlFooter.Controls.AddRange({btnSave, btnCancel})

        Me.Controls.AddRange({lblTitle,
                               lblType, cboType, lblCap, cboCapacity, lblGen, cboGeneration,
                               lblCpu, txtCpu, lblStorage, txtStorage,
                               lblSpeedLabel, cboSpeed, lblBrand, cboBrand, lblModel, txtModel,
                               lblPart, txtPartNumber, lblQty, nudQuantity,
                               lblCond, cboCondition, lblStat, cboStatus, lblBatch, cboSourceBatch,
                               lblNotes, txtNotes,
                               pnlFooter})

    End Sub

#End Region

End Class
