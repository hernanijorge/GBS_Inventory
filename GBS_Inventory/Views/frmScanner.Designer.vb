Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmScanner
    Inherits Form

    Friend WithEvents txtUID As TextBox
    Friend WithEvents btnBuscar As Button
    Friend WithEvents btnAddUpgrade As Button
    Friend WithEvents lblResultado As Label
    Friend WithEvents lblManufacturer As Label
    Friend WithEvents lblSerial As Label
    Friend WithEvents lblCPU As Label
    Friend WithEvents lblRAM As Label
    Friend WithEvents lblStorage As Label
    Friend WithEvents lblCondition As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents lblUpgrades As Label
    Friend WithEvents dgvUpgrades As DataGridView
    Friend WithEvents lblTitulo As Label

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text = "Scanner — Internal UID"
        Me.Size = New Size(900, 650)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = TemaEscuro.Fundo
        Me.Font = New Font("Segoe UI", 9)

        ' ── Título ────────────────────────────────────────────────
        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "Scanner de Equipamentos"
        Me.lblTitulo.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent
        Me.lblTitulo.Location = New Point(20, 15)
        Me.lblTitulo.AutoSize = True

        ' ── Campo UID grande ──────────────────────────────────────
        Dim lblUID As New Label() With {
            .Text = "Internal UID (escaneie ou digite):",
            .Location = New Point(20, 55),
            .AutoSize = True,
            .ForeColor = TemaEscuro.TextoMutado,
            .Font = New Font("Segoe UI", 10)
        }

        Me.txtUID = New TextBox()
        Me.txtUID.Location = New Point(20, 80)
        Me.txtUID.Size = New Size(540, 36)
        Me.txtUID.Font = New Font("Consolas", 18, FontStyle.Bold)
        Me.txtUID.BorderStyle = BorderStyle.FixedSingle
        Me.txtUID.BackColor = TemaEscuro.SurfaceClaro
        Me.txtUID.ForeColor = TemaEscuro.Accent

        Me.btnBuscar = New Button()
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.Location = New Point(575, 78)
        Me.btnBuscar.Size = New Size(120, 40)
        Me.btnBuscar.Font = New Font("Segoe UI", 10, FontStyle.Bold)

        Me.btnAddUpgrade = New Button()
        Me.btnAddUpgrade.Text = "+ Upgrade"
        Me.btnAddUpgrade.Location = New Point(705, 78)
        Me.btnAddUpgrade.Size = New Size(150, 40)
        Me.btnAddUpgrade.Enabled = False

        ' ── Label resultado ───────────────────────────────────────
        Me.lblResultado = New Label()
        Me.lblResultado.Location = New Point(20, 130)
        Me.lblResultado.Size = New Size(840, 24)
        Me.lblResultado.Font = New Font("Segoe UI", 11, FontStyle.Bold)
        Me.lblResultado.Text = "Aguardando leitura..."
        Me.lblResultado.ForeColor = TemaEscuro.TextoMutado

        ' ── Card de detalhes ──────────────────────────────────────
        Dim pnlDetalhes As New Panel() With {
            .Location = New Point(20, 165),
            .Size = New Size(840, 180),
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlDetalhes.Paint, Sub(sender, e)
                                          Using p As New Pen(TemaEscuro.Accent, 2)
                                              e.Graphics.DrawRectangle(p, 0, 0, pnlDetalhes.Width - 1, pnlDetalhes.Height - 1)
                                          End Using
                                      End Sub

        Me.lblManufacturer = criarLabelDetalhe(20, 15, 800, 26, 14, True)
        Me.lblSerial = criarLabelDetalhe(20, 50, 390, 22, 10, False)
        Me.lblCPU = criarLabelDetalhe(410, 50, 410, 22, 10, False)
        Me.lblRAM = criarLabelDetalhe(20, 80, 390, 22, 10, False)
        Me.lblStorage = criarLabelDetalhe(410, 80, 410, 22, 10, False)
        Me.lblCondition = criarLabelDetalhe(20, 110, 390, 22, 10, False)
        Me.lblStatus = criarLabelDetalhe(410, 110, 410, 22, 10, False)
        Me.lblUpgrades = criarLabelDetalhe(20, 140, 800, 22, 10, False)

        pnlDetalhes.Controls.AddRange({lblManufacturer, lblSerial, lblCPU, lblRAM,
                                        lblStorage, lblCondition, lblStatus, lblUpgrades})

        ' ── Grid de upgrades ──────────────────────────────────────
        Dim lblHistUpg As New Label() With {
            .Text = "Histórico de Upgrades:",
            .Location = New Point(20, 360),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        Me.dgvUpgrades = New DataGridView()
        Me.dgvUpgrades.Location = New Point(20, 385)
        Me.dgvUpgrades.Size = New Size(840, 200)
        Me.dgvUpgrades.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        ' ── Montagem ──────────────────────────────────────────────
        Me.Controls.AddRange({lblTitulo, lblUID, txtUID, btnBuscar, btnAddUpgrade,
                               lblResultado, pnlDetalhes, lblHistUpg, dgvUpgrades})

        Me.ResumeLayout(False)

    End Sub

    Private Function criarLabelDetalhe(pX As Integer, pY As Integer,
                                         pW As Integer, pH As Integer,
                                         pSize As Integer, pBold As Boolean) As Label
        Return New Label() With {
            .Location = New Point(pX, pY),
            .Size = New Size(pW, pH),
            .ForeColor = If(pBold, TemaEscuro.Accent, TemaEscuro.Texto),
            .Font = New Font("Segoe UI", pSize, If(pBold, FontStyle.Bold, FontStyle.Regular)),
            .Text = ""
        }
    End Function

End Class
