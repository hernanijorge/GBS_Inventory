Imports System.Windows.Forms
Imports System.Drawing

Partial Class frmClientAliasMapping
    Inherits Form

#Region "Designer"

    Friend WithEvents dgvClienteReal  As DataGridView
    Friend WithEvents dgvAliases      As DataGridView
    Friend WithEvents btnAddClient    As Button
    Friend WithEvents btnRemoveClient As Button
    Friend WithEvents cboNovoAlias    As ComboBox
    Friend WithEvents btnAddAlias     As Button
    Friend WithEvents btnRemoveAlias  As Button
    Friend WithEvents btnAutoDetect   As Button
    Friend WithEvents lblInfo         As Label
    Friend WithEvents btnSave         As Button
    Friend WithEvents btnCancel       As Button
    Friend WithEvents pnlFooter       As Panel

    Private Sub InitializeComponent()

        Me.Text            = "Client Alias Mapping"
        Me.Size            = New Size(740, 540)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.StartPosition   = FormStartPosition.CenterParent

        ' ── Column A: canonical names ────────────────────────────────
        Dim lblA As New Label() With { .Text = "Canonical name (shown in the report)", .Location = New Point(15, 12), .AutoSize = True }
        Me.dgvClienteReal = New DataGridView() With { .Location = New Point(15, 34), .Size = New Size(320, 330) }
        Me.dgvClienteReal.Columns.Add("CANONICAL", "Canonical name")

        Me.btnAddClient    = New Button() With { .Text = "+ Add Client",  .Location = New Point(15, 372),  .Size = New Size(110, 28) }
        Me.btnRemoveClient = New Button() With { .Text = "Remove Client", .Location = New Point(131, 372), .Size = New Size(110, 28) }

        ' ── Column B: aliases of the selected canonical name ─────────
        Dim lblB As New Label() With { .Text = "Aliases (variations merged into it)", .Location = New Point(355, 12), .AutoSize = True }
        Me.dgvAliases = New DataGridView() With { .Location = New Point(355, 34), .Size = New Size(355, 330) }
        Me.dgvAliases.Columns.Add("ALIAS", "Aliases (variations)")

        Me.cboNovoAlias   = New ComboBox() With { .Location = New Point(355, 374), .Size = New Size(225, 24), .DropDownStyle = ComboBoxStyle.DropDown,
                                                  .AutoCompleteMode = AutoCompleteMode.SuggestAppend, .AutoCompleteSource = AutoCompleteSource.ListItems }
        Me.btnAddAlias    = New Button() With { .Text = "+ Add Alias",  .Location = New Point(586, 372), .Size = New Size(124, 28) }
        Me.btnRemoveAlias = New Button() With { .Text = "Remove Alias", .Location = New Point(586, 406), .Size = New Size(124, 28) }

        Me.btnAutoDetect = New Button() With { .Text = "Auto-detect from Prefix", .Location = New Point(15, 412), .Size = New Size(180, 28) }
        Me.lblInfo       = New Label() With { .Text = "", .Location = New Point(15, 446), .Size = New Size(560, 18), .ForeColor = TemaEscuro.TextoMutado }

        ' ── Footer ───────────────────────────────────────────────────
        Me.pnlFooter = New Panel() With { .Dock = DockStyle.Bottom, .Height = 55, .BackColor = TemaEscuro.Surface }
        Me.btnSave   = New Button() With { .Text = "Save Mapping", .Size = New Size(140, 34), .Location = New Point(465, 10),
                                            .BackColor = TemaEscuro.Accent, .ForeColor = TemaEscuro.Fundo, .FlatStyle = FlatStyle.Flat }
        Me.btnCancel = New Button() With { .Text = "Cancel", .Size = New Size(100, 34), .Location = New Point(612, 10) }
        Me.pnlFooter.Controls.AddRange({btnSave, btnCancel})

        Me.CancelButton = btnCancel

        Me.Controls.AddRange({lblA, dgvClienteReal, btnAddClient, btnRemoveClient,
                              lblB, dgvAliases, cboNovoAlias, btnAddAlias, btnRemoveAlias,
                              btnAutoDetect, lblInfo, pnlFooter})

    End Sub

#End Region

End Class
