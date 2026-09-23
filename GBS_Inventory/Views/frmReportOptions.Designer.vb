Imports System.Windows.Forms
Imports System.Drawing

Partial Class frmReportOptions
    Inherits Form

#Region "Designer"

    Friend WithEvents rbGroupWithoutModel As RadioButton
    Friend WithEvents rbGroupWithModel    As RadioButton
    Friend WithEvents chkDesktop          As CheckBox
    Friend WithEvents chkMiniDesktop      As CheckBox
    Friend WithEvents chkRam              As CheckBox
    Friend WithEvents chkSsd              As CheckBox
    Friend WithEvents chkHdd              As CheckBox
    Friend WithEvents chkOthers           As CheckBox
    Friend WithEvents btnSelectAll        As Button
    Friend WithEvents btnClearAll         As Button
    Friend WithEvents btnOnlyMachines     As Button
    Friend WithEvents btnOnlyParts        As Button
    Friend WithEvents chkColInStock       As CheckBox
    Friend WithEvents chkColInstalled     As CheckBox
    Friend WithEvents chkColSold          As CheckBox
    Friend WithEvents chkColScrapped      As CheckBox
    Friend WithEvents btnGenerate         As Button
    Friend WithEvents btnCancel           As Button
    Friend WithEvents pnlFooter           As Panel

    Private Sub InitializeComponent()

        Me.Text            = "Report Options"
        Me.Size            = New Size(470, 450)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.StartPosition   = FormStartPosition.CenterParent

        ' ── Section 1: Grouping ──────────────────────────────────────
        Dim grpGrouping As New GroupBox() With { .Text = "Grouping", .Location = New Point(15, 12), .Size = New Size(425, 75) }
        Me.rbGroupWithoutModel = New RadioButton() With { .Text = "Consolidated (original — Type + CPU + Specs)", .Location = New Point(15, 22), .AutoSize = True }
        Me.rbGroupWithModel    = New RadioButton() With { .Text = "Detailed (include Model in grouping)",        .Location = New Point(15, 46), .AutoSize = True, .Checked = True }
        grpGrouping.Controls.AddRange({rbGroupWithoutModel, rbGroupWithModel})

        ' ── Section 2: Include component types ───────────────────────
        Dim grpTypes As New GroupBox() With { .Text = "Include component types", .Location = New Point(15, 97), .Size = New Size(425, 125) }
        Me.chkDesktop     = New CheckBox() With { .Text = "Desktop",      .Location = New Point(15,  24), .AutoSize = True, .Checked = True }
        Me.chkMiniDesktop = New CheckBox() With { .Text = "Mini Desktop", .Location = New Point(150, 24), .AutoSize = True, .Checked = True }
        Me.chkOthers      = New CheckBox() With { .Text = "Other types",  .Location = New Point(285, 24), .AutoSize = True, .Checked = True }
        Me.chkRam         = New CheckBox() With { .Text = "RAM",          .Location = New Point(15,  50), .AutoSize = True, .Checked = True }
        Me.chkSsd         = New CheckBox() With { .Text = "SSD",          .Location = New Point(150, 50), .AutoSize = True, .Checked = True }
        Me.chkHdd         = New CheckBox() With { .Text = "HDD",          .Location = New Point(285, 50), .AutoSize = True, .Checked = True }

        Me.btnSelectAll    = New Button() With { .Text = "Select All",    .Location = New Point(15,  82), .Size = New Size(80, 28) }
        Me.btnClearAll     = New Button() With { .Text = "Clear All",     .Location = New Point(100, 82), .Size = New Size(80, 28) }
        Me.btnOnlyMachines = New Button() With { .Text = "Only Machines", .Location = New Point(185, 82), .Size = New Size(110, 28) }
        Me.btnOnlyParts    = New Button() With { .Text = "Only Parts",    .Location = New Point(300, 82), .Size = New Size(110, 28) }
        grpTypes.Controls.AddRange({chkDesktop, chkMiniDesktop, chkOthers, chkRam, chkSsd, chkHdd,
                                    btnSelectAll, btnClearAll, btnOnlyMachines, btnOnlyParts})

        Dim tip As New ToolTip()
        tip.SetToolTip(btnOnlyMachines, "Desktop + Mini Desktop")
        tip.SetToolTip(btnOnlyParts,    "RAM + SSD + HDD")

        ' ── Section 3: Status columns ────────────────────────────────
        Dim grpStatus As New GroupBox() With { .Text = "Status columns", .Location = New Point(15, 232), .Size = New Size(425, 80) }
        Me.chkColInStock   = New CheckBox() With { .Text = "Include 'In Stock' column",  .Location = New Point(15,  24), .AutoSize = True, .Checked = True }
        Me.chkColInstalled = New CheckBox() With { .Text = "Include 'Installed' column", .Location = New Point(215, 24), .AutoSize = True, .Checked = True }
        Me.chkColSold      = New CheckBox() With { .Text = "Include 'Sold' column",      .Location = New Point(15,  50), .AutoSize = True, .Checked = True }
        Me.chkColScrapped  = New CheckBox() With { .Text = "Include 'Scrapped' column",  .Location = New Point(215, 50), .AutoSize = True, .Checked = True }
        grpStatus.Controls.AddRange({chkColInStock, chkColInstalled, chkColSold, chkColScrapped})

        ' ── Footer ───────────────────────────────────────────────────
        Me.pnlFooter   = New Panel() With { .Dock = DockStyle.Bottom, .Height = 55, .BackColor = TemaEscuro.Surface }
        Me.btnGenerate = New Button() With { .Text = "Generate Report", .Size = New Size(150, 34), .Location = New Point(180, 10),
                                              .BackColor = TemaEscuro.Accent, .ForeColor = TemaEscuro.Fundo, .FlatStyle = FlatStyle.Flat }
        Me.btnCancel   = New Button() With { .Text = "Cancel", .Size = New Size(100, 34), .Location = New Point(340, 10) }
        Me.pnlFooter.Controls.AddRange({btnGenerate, btnCancel})

        Me.AcceptButton = btnGenerate
        Me.CancelButton = btnCancel

        Me.Controls.AddRange({grpGrouping, grpTypes, grpStatus, pnlFooter})

    End Sub

#End Region

End Class
