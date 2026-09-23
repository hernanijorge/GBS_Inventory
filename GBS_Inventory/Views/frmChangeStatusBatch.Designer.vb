Imports System.Windows.Forms
Imports System.Drawing

Partial Class frmChangeStatusBatch
    Inherits Form

#Region "Designer"

    Friend WithEvents lblInfo       As Label
    Friend WithEvents cboNovoStatus As ComboBox
    Friend WithEvents lblAviso      As Label
    Friend WithEvents txtNota       As TextBox
    Friend WithEvents btnOK         As Button
    Friend WithEvents btnCancel     As Button
    Friend WithEvents pnlFooter     As Panel

    Private Sub InitializeComponent()

        Me.Text            = "Change Status"
        Me.Size            = New Size(440, 330)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.StartPosition   = FormStartPosition.CenterParent

        Dim lblTitle As New Label() With {
            .Text      = "Change Status",
            .Font      = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .Location  = New Point(20, 15),
            .AutoSize  = True
        }

        Me.lblInfo       = New Label() With { .Text = "Change status of 0 selected items to:", .Location = New Point(20, 60), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboNovoStatus = New ComboBox() With { .Location = New Point(20, 82), .Size = New Size(180, 24), .DropDownStyle = ComboBoxStyle.DropDownList }

        Me.lblAviso      = New Label() With { .Text = "These items will leave available stock", .Location = New Point(20, 112), .AutoSize = True,
                                              .ForeColor = TemaEscuro.Destaque, .Font = New Font("Segoe UI", 9, FontStyle.Bold), .Visible = False }

        Dim lblNota As New Label() With { .Text = "Note (added to all items)", .Location = New Point(20, 140), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.txtNota       = New TextBox() With { .Location = New Point(20, 162), .Size = New Size(380, 50), .Multiline = True, .MaxLength = 500 }

        ' ── Footer ───────────────────────────────────────────────────
        Me.pnlFooter = New Panel() With { .Dock = DockStyle.Bottom, .Height = 55, .BackColor = TemaEscuro.Surface }
        Me.btnOK     = New Button() With { .Text = "OK",     .Size = New Size(100, 34), .Location = New Point(190, 10),
                                            .BackColor = TemaEscuro.Accent, .ForeColor = TemaEscuro.Fundo, .FlatStyle = FlatStyle.Flat }
        Me.btnCancel = New Button() With { .Text = "Cancel", .Size = New Size(100, 34), .Location = New Point(300, 10) }
        Me.pnlFooter.Controls.AddRange({btnOK, btnCancel})

        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel

        Me.Controls.AddRange({lblTitle, lblInfo, cboNovoStatus, lblAviso, lblNota, txtNota, pnlFooter})

    End Sub

#End Region

End Class
