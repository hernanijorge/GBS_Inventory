Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmPrincipal
    Inherits Form

    Private components As System.ComponentModel.IContainer

    Friend WithEvents pnlTopo As Panel
    Friend WithEvents lblTitulo As Label
    Friend WithEvents lblSubTitulo As Label
    Friend WithEvents dtpData As DateTimePicker
    Friend WithEvents pnlPower As Panel
    Friend WithEvents pnlContainer As Panel
    Friend WithEvents tabPrincipal As TabControl
    Friend WithEvents tabDashboard As TabPage
    Friend WithEvents tabEstoque As TabPage
    Friend WithEvents tabRemessas As TabPage
    Friend WithEvents tabUpgrades As TabPage
    Friend WithEvents tabImportacao As TabPage
    Friend WithEvents pnlAcoes As Panel
    Friend WithEvents btnBuscarUID As Button
    Friend WithEvents btnScanner As Button
    Friend WithEvents btnImportarPlanilha As Button
    Friend WithEvents btnAtualizar As Button
    Friend WithEvents txtPesquisa As TextBox
    Friend WithEvents lblPesquisa As Label
    Friend WithEvents dgvEstoque As DataGridView
    Friend WithEvents pnlDashboard As Panel
    Friend WithEvents lblTotalUnidades As Label
    Friend WithEvents lblEmEstoque As Label
    Friend WithEvents lblCondicaoBoa As Label
    Friend WithEvents lblUpgrades30d As Label
    Friend WithEvents lblRemessasAtivas As Label
    Friend WithEvents lblVersao As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents dgvRemessas As DataGridView
    Friend WithEvents dgvUpgrades As DataGridView

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        ' ═════════════════════════════════════════════════════════════
        ' FORM PRINCIPAL
        ' ═════════════════════════════════════════════════════════════
        Me.Text = "GBS Inventory Manager — Global Business Solution, Boston"
        Me.Size = New Size(1280, 780)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(1100, 650)
        Me.BackColor = TemaEscuro.Fundo
        Me.ForeColor = TemaEscuro.Texto
        Me.Font = New Font("Segoe UI", 9)

        ' ═════════════════════════════════════════════════════════════
        ' PAINEL TOPO — Date picker + Power button
        ' ═════════════════════════════════════════════════════════════
        Me.pnlTopo = New Panel()
        Me.pnlTopo.Dock = DockStyle.Top
        Me.pnlTopo.Height = 60
        Me.pnlTopo.BackColor = TemaEscuro.Fundo

        Me.dtpData = New DateTimePicker()
        Me.dtpData.Format = DateTimePickerFormat.Long
        Me.dtpData.Location = New Point(15, 18)
        Me.dtpData.Size = New Size(240, 24)
        Me.dtpData.BackColor = TemaEscuro.SurfaceClaro
        Me.dtpData.CalendarMonthBackground = TemaEscuro.Surface
        Me.dtpData.CalendarForeColor = TemaEscuro.Texto
        Me.dtpData.CalendarTitleBackColor = TemaEscuro.Surface
        Me.dtpData.CalendarTitleForeColor = TemaEscuro.Accent

        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "GBS Inventory Manager"
        Me.lblTitulo.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent
        Me.lblTitulo.Location = New Point(290, 10)
        Me.lblTitulo.AutoSize = True
        Me.lblTitulo.BackColor = Color.Transparent

        Me.lblSubTitulo = New Label()
        Me.lblSubTitulo.Text = "Global Business Solution · Boston, MA · v1.0"
        Me.lblSubTitulo.Font = New Font("Segoe UI", 8)
        Me.lblSubTitulo.ForeColor = TemaEscuro.TextoMutado
        Me.lblSubTitulo.Location = New Point(292, 36)
        Me.lblSubTitulo.AutoSize = True
        Me.lblSubTitulo.BackColor = Color.Transparent

        ' Power button verde estilo moderno (simulado com botão circular)
        Me.pnlPower = New Panel()
        Me.pnlPower.Size = New Size(48, 48)
        Me.pnlPower.Location = New Point(Me.Width - 80, 6)
        Me.pnlPower.BackColor = Color.Transparent
        Me.pnlPower.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler Me.pnlPower.Paint, AddressOf desenharPowerButton

        Me.pnlTopo.Controls.AddRange({dtpData, lblTitulo, lblSubTitulo, pnlPower})

        ' ═════════════════════════════════════════════════════════════
        ' PAINEL CONTAINER — estilo moderno com borda teal
        ' ═════════════════════════════════════════════════════════════
        Me.pnlContainer = New Panel()
        Me.pnlContainer.Dock = DockStyle.Fill
        Me.pnlContainer.BackColor = TemaEscuro.Fundo
        Me.pnlContainer.Padding = New Padding(15, 5, 15, 15)

        ' ═════════════════════════════════════════════════════════════
        ' TABCONTROL
        ' ═════════════════════════════════════════════════════════════
        Me.tabPrincipal = New TabControl()
        Me.tabPrincipal.Dock = DockStyle.Fill
        Me.tabPrincipal.Font = New Font("Segoe UI", 10)

        Me.tabDashboard = New TabPage("Dashboard")
        Me.tabEstoque = New TabPage("Estoque")
        Me.tabRemessas = New TabPage("Remessas")
        Me.tabUpgrades = New TabPage("Upgrades")
        Me.tabImportacao = New TabPage("Importação")

        ' ─── ABA DASHBOARD ────────────────────────────────────────────
        Me.pnlDashboard = New Panel()
        Me.pnlDashboard.Dock = DockStyle.Fill
        Me.pnlDashboard.Padding = New Padding(40, 30, 40, 30)

        ' Adiciona o painel à aba ANTES de criar os cards (criarCardDashboard adiciona no pnlDashboard)
        Me.tabDashboard.Controls.Add(pnlDashboard)

        Me.lblTotalUnidades = criarCardDashboard("Total de Unidades", "—", 20, 20)
        Me.lblEmEstoque = criarCardDashboard("Em Estoque", "—", 260, 20)
        Me.lblCondicaoBoa = criarCardDashboard("Condição Boa", "—", 500, 20)
        Me.lblUpgrades30d = criarCardDashboard("Upgrades 30d", "—", 740, 20)
        Me.lblRemessasAtivas = criarCardDashboard("Remessas Ativas", "—", 980, 20)

        ' ─── ABA ESTOQUE ──────────────────────────────────────────────
        Me.pnlAcoes = New Panel()
        Me.pnlAcoes.Dock = DockStyle.Top
        Me.pnlAcoes.Height = 50
        Me.pnlAcoes.BackColor = TemaEscuro.Surface
        Me.pnlAcoes.Padding = New Padding(8)

        Me.lblPesquisa = New Label()
        Me.lblPesquisa.Text = "Pesquisa (UID / Serial / Modelo):"
        Me.lblPesquisa.Location = New Point(10, 15)
        Me.lblPesquisa.AutoSize = True
        Me.lblPesquisa.ForeColor = TemaEscuro.TextoMutado

        Me.txtPesquisa = New TextBox()
        Me.txtPesquisa.Location = New Point(215, 12)
        Me.txtPesquisa.Size = New Size(250, 24)

        Me.btnBuscarUID = New Button()
        Me.btnBuscarUID.Text = "Buscar"
        Me.btnBuscarUID.Location = New Point(475, 10)
        Me.btnBuscarUID.Size = New Size(85, 28)

        Me.btnScanner = New Button()
        Me.btnScanner.Text = "Scanner UID"
        Me.btnScanner.Location = New Point(570, 10)
        Me.btnScanner.Size = New Size(110, 28)

        Me.btnImportarPlanilha = New Button()
        Me.btnImportarPlanilha.Text = "Importar Planilha"
        Me.btnImportarPlanilha.Location = New Point(690, 10)
        Me.btnImportarPlanilha.Size = New Size(130, 28)

        Me.btnAtualizar = New Button()
        Me.btnAtualizar.Text = "Atualizar"
        Me.btnAtualizar.Location = New Point(830, 10)
        Me.btnAtualizar.Size = New Size(90, 28)

        Me.pnlAcoes.Controls.AddRange({lblPesquisa, txtPesquisa, btnBuscarUID,
                                        btnScanner, btnImportarPlanilha, btnAtualizar})

        Me.dgvEstoque = New DataGridView()
        Me.dgvEstoque.Dock = DockStyle.Fill

        Me.tabEstoque.Controls.Add(dgvEstoque)
        Me.tabEstoque.Controls.Add(pnlAcoes)

        ' ─── ABA REMESSAS ─────────────────────────────────────────────
        Me.dgvRemessas = New DataGridView()
        Me.dgvRemessas.Dock = DockStyle.Fill
        Me.tabRemessas.Controls.Add(dgvRemessas)

        ' ─── ABA UPGRADES ─────────────────────────────────────────────
        Me.dgvUpgrades = New DataGridView()
        Me.dgvUpgrades.Dock = DockStyle.Fill
        Me.tabUpgrades.Controls.Add(dgvUpgrades)

        ' ─── ABA IMPORTAÇÃO ───────────────────────────────────────────
        Dim lblImport As New Label() With {
            .Text = "Clique em 'Importar Planilha' na aba Estoque",
            .Location = New Point(20, 20),
            .AutoSize = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.tabImportacao.Controls.Add(lblImport)

        Me.tabPrincipal.TabPages.AddRange({tabDashboard, tabEstoque, tabRemessas, tabUpgrades, tabImportacao})
        Me.tabPrincipal.SelectedIndex = 1

        Me.pnlContainer.Controls.Add(tabPrincipal)

        ' ═════════════════════════════════════════════════════════════
        ' STATUS BAR
        ' ═════════════════════════════════════════════════════════════
        Me.lblStatus = New Label()
        Me.lblStatus.Text = "Conectado ao Oracle"
        Me.lblStatus.Dock = DockStyle.Bottom
        Me.lblStatus.Height = 22
        Me.lblStatus.TextAlign = ContentAlignment.MiddleLeft
        Me.lblStatus.BackColor = TemaEscuro.Surface
        Me.lblStatus.ForeColor = TemaEscuro.Accent
        Me.lblStatus.Padding = New Padding(12, 0, 0, 0)
        Me.lblStatus.Font = New Font("Segoe UI", 8)

        Me.lblVersao = New Label()
        Me.lblVersao.Text = "GBS_Inventory v1.0"
        Me.lblVersao.Dock = DockStyle.Bottom
        Me.lblVersao.Visible = False

        ' ═════════════════════════════════════════════════════════════
        ' MONTAGEM FINAL
        ' ═════════════════════════════════════════════════════════════
        Me.Controls.Add(pnlContainer)
        Me.Controls.Add(lblStatus)
        Me.Controls.Add(pnlTopo)

        Me.ResumeLayout(False)

    End Sub

    ''' <summary>Cria um card tipo dashboard (título + valor grande)</summary>
    Private Function criarCardDashboard(pTitulo As String, pValor As String,
                                         pX As Integer, pY As Integer) As Label

        Dim pnl As New Panel() With {
            .Location = New Point(pX, pY),
            .Size = New Size(210, 100),
            .BackColor = TemaEscuro.Surface
        }

        AddHandler pnl.Paint, Sub(sender, e)
                                  Dim g As Graphics = e.Graphics
                                  Using p As New Pen(TemaEscuro.Accent, 2)
                                      g.DrawRectangle(p, 0, 0, pnl.Width - 1, pnl.Height - 1)
                                  End Using
                              End Sub

        Dim lblTit As New Label() With {
            .Text = pTitulo,
            .Location = New Point(12, 10),
            .Size = New Size(190, 20),
            .ForeColor = TemaEscuro.TextoMutado,
            .Font = New Font("Segoe UI", 8.5F)
        }

        Dim lblVal As New Label() With {
            .Text = pValor,
            .Location = New Point(12, 35),
            .Size = New Size(190, 55),
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft
        }

        pnl.Controls.Add(lblTit)
        pnl.Controls.Add(lblVal)

        Me.pnlDashboard.Controls.Add(pnl)

        Return lblVal

    End Function

    ''' <summary>Desenha o power button verde no topo direito (estilo moderno)</summary>
    Private Sub desenharPowerButton(sender As Object, e As PaintEventArgs)

        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        ' Círculo verde
        Using brush As New SolidBrush(TemaEscuro.Verde)
            g.FillEllipse(brush, 4, 4, 40, 40)
        End Using

        ' Símbolo de power (linha vertical + arco)
        Using p As New Pen(TemaEscuro.Fundo, 3)
            g.DrawLine(p, 24, 14, 24, 24)
            g.DrawArc(p, 12, 14, 24, 22, 30, 120)
        End Using

    End Sub

End Class
