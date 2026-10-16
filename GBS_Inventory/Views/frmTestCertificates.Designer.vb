Imports System.Windows.Forms
Imports System.Drawing

Partial Class frmTestCertificates
    Inherits Form

#Region "Designer"

    Friend WithEvents lblTitulo      As Label
    Friend WithEvents lblCliente     As Label
    Friend WithEvents lblContador    As Label
    Friend WithEvents dgvMaquinas    As DataGridView
    Friend WithEvents rbSoAprovados  As RadioButton
    Friend WithEvents rbTodos        As RadioButton
    Friend WithEvents lblIncluidos   As Label
    Friend WithEvents btnGenerate    As Button
    Friend WithEvents lblArquivo     As Label
    Friend WithEvents btnOpenPdf     As Button
    Friend WithEvents btnSaveAs      As Button
    Friend WithEvents btnEmail       As Button
    Friend WithEvents btnClose       As Button

    Private Sub InitializeComponent()

        Me.Text            = "Test Certificates"
        Me.Size            = New Size(980, 640)
        Me.MinimumSize     = New Size(820, 480)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.MaximizeBox     = True
        Me.MinimizeBox     = False

        ' ── Topo: remessa, cliente e contador ────────────────────────
        Dim pnlTopo As New Panel() With { .Dock = DockStyle.Top, .Height = 78, .Padding = New Padding(12, 8, 12, 4) }

        Me.lblTitulo = New Label() With {
            .Text     = "Test Certificates",
            .Font     = New Font("Segoe UI", 12, FontStyle.Bold),
            .Location = New Point(12, 8),
            .AutoSize = True
        }
        Me.lblCliente = New Label() With {
            .Text     = "",
            .Font     = New Font("Segoe UI", 9),
            .Location = New Point(14, 34),
            .AutoSize = True
        }
        Me.lblContador = New Label() With {
            .Text     = "",
            .Font     = New Font("Segoe UI", 10, FontStyle.Bold),
            .Location = New Point(14, 54),
            .AutoSize = True
        }
        pnlTopo.Controls.AddRange({lblTitulo, lblCliente, lblContador})

        ' ── Grid de pre-visualizacao ─────────────────────────────────
        Me.dgvMaquinas = New DataGridView() With {
            .Dock          = DockStyle.Fill,
            .ReadOnly      = True,
            .MultiSelect   = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AllowUserToAddRows    = False,
            .AllowUserToDeleteRows = False
        }

        ' ── Rodape: opcoes, gerar, acoes ─────────────────────────────
        Dim pnlRodape As New Panel() With { .Dock = DockStyle.Bottom, .Height = 96, .Padding = New Padding(12, 6, 12, 6) }

        Me.rbSoAprovados = New RadioButton() With {
            .Text = "Include only passed machines (PASSED / PASSED WITH NOTES)",
            .Location = New Point(14, 8), .AutoSize = True, .Checked = True
        }
        Me.rbTodos = New RadioButton() With {
            .Text = "Include all machines (failed / incomplete / not tested listed on the cover)",
            .Location = New Point(14, 30), .AutoSize = True
        }
        Me.lblIncluidos = New Label() With { .Text = "", .Location = New Point(16, 56), .AutoSize = True, .Font = New Font("Segoe UI", 8.5F, FontStyle.Italic) }

        Me.btnGenerate = New Button() With { .Text = "Generate", .Location = New Point(560, 8), .Size = New Size(120, 34) }

        Me.lblArquivo = New Label() With { .Text = "", .Location = New Point(718, 71), .Size = New Size(232, 18), .Font = New Font("Segoe UI", 8.5F),
                                           .AutoEllipsis = True }

        Me.btnOpenPdf = New Button() With { .Text = "Open PDF", .Location = New Point(690, 8),  .Size = New Size(85, 34), .Enabled = False }
        Me.btnSaveAs  = New Button() With { .Text = "Save As",  .Location = New Point(780, 8),  .Size = New Size(80, 34), .Enabled = False }
        Me.btnEmail   = New Button() With { .Text = "Email to Customer", .Location = New Point(560, 66), .Size = New Size(150, 26), .Enabled = False }
        Me.btnClose   = New Button() With { .Text = "Close",    .Location = New Point(865, 8),  .Size = New Size(80, 34) }

        pnlRodape.Controls.AddRange({rbSoAprovados, rbTodos, lblIncluidos, btnGenerate, lblArquivo,
                                     btnOpenPdf, btnSaveAs, btnEmail, btnClose})

        ' Fill primeiro, depois Top/Bottom (ordem de docking)
        Me.Controls.Add(dgvMaquinas)
        Me.Controls.Add(pnlTopo)
        Me.Controls.Add(pnlRodape)

    End Sub

#End Region

End Class
