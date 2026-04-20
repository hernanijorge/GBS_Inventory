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

    ' Controles da aba Remessas
    Friend WithEvents lblRemessasTit       As Label
    Friend WithEvents btnNovaRemessa       As Button
    Friend WithEvents btnAtualizarRemessas As Button

    ' Controles da aba Upgrades
    Friend WithEvents lblUpgradesTit       As Label
    Friend WithEvents lblUpgradesTotal     As Label
    Friend WithEvents btnAtualizarUpgrades As Button

    ' Controles da aba Importação
    Friend WithEvents lblImpArquivo As Label
    Friend WithEvents btnImpSelecionar As Button
    Friend WithEvents btnImpIniciar As Button
    Friend WithEvents progImp As ProgressBar
    Friend WithEvents lblImpStatus As Label
    Friend WithEvents lblImpInseridos As Label
    Friend WithEvents lblImpAtualizados As Label
    Friend WithEvents lblImpAbas As Label
    Friend WithEvents lblImpLinhas As Label
    Friend WithEvents lblImpErros As Label
    Friend WithEvents txtImpErros As TextBox

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
        Me.tabRemessas.Padding = New Padding(0)

        Dim pnlRemessasTop As New Panel() With {
            .Dock      = DockStyle.Top,
            .Height    = 50,
            .BackColor = TemaEscuro.Surface,
            .Padding   = New Padding(8)
        }

        Me.lblRemessasTit           = New Label()
        Me.lblRemessasTit.Text      = "Remessas Ativas"
        Me.lblRemessasTit.Font      = New Font("Segoe UI", 11, FontStyle.Bold)
        Me.lblRemessasTit.ForeColor = TemaEscuro.Accent
        Me.lblRemessasTit.Location  = New Point(15, 15)
        Me.lblRemessasTit.AutoSize  = True

        Me.btnNovaRemessa            = New Button()
        Me.btnNovaRemessa.Text       = "+ Nova Remessa"
        Me.btnNovaRemessa.Location   = New Point(200, 10)
        Me.btnNovaRemessa.Size       = New Size(150, 30)

        Me.btnAtualizarRemessas            = New Button()
        Me.btnAtualizarRemessas.Text       = "Atualizar"
        Me.btnAtualizarRemessas.Location   = New Point(360, 10)
        Me.btnAtualizarRemessas.Size       = New Size(90, 30)

        pnlRemessasTop.Controls.AddRange({lblRemessasTit, btnNovaRemessa, btnAtualizarRemessas})

        Me.dgvRemessas      = New DataGridView()
        Me.dgvRemessas.Dock = DockStyle.Fill

        Me.tabRemessas.Controls.Add(dgvRemessas)
        Me.tabRemessas.Controls.Add(pnlRemessasTop)

        ' ─── ABA UPGRADES ─────────────────────────────────────────────
        Me.tabUpgrades.Padding = New Padding(0)

        Dim pnlUpgTop As New Panel() With {
            .Dock      = DockStyle.Top,
            .Height    = 50,
            .BackColor = TemaEscuro.Surface,
            .Padding   = New Padding(8)
        }

        Me.lblUpgradesTit           = New Label()
        Me.lblUpgradesTit.Text      = "Upgrades (últimos 90 dias)"
        Me.lblUpgradesTit.Font      = New Font("Segoe UI", 11, FontStyle.Bold)
        Me.lblUpgradesTit.ForeColor = TemaEscuro.Accent
        Me.lblUpgradesTit.Location  = New Point(15, 15)
        Me.lblUpgradesTit.AutoSize  = True

        Me.lblUpgradesTotal           = New Label()
        Me.lblUpgradesTotal.Text      = "Total: 0"
        Me.lblUpgradesTotal.Font      = New Font("Segoe UI", 10)
        Me.lblUpgradesTotal.ForeColor = TemaEscuro.TextoMutado
        Me.lblUpgradesTotal.Location  = New Point(290, 17)
        Me.lblUpgradesTotal.AutoSize  = True

        Me.btnAtualizarUpgrades            = New Button()
        Me.btnAtualizarUpgrades.Text       = "Atualizar"
        Me.btnAtualizarUpgrades.Location   = New Point(420, 10)
        Me.btnAtualizarUpgrades.Size       = New Size(90, 30)

        pnlUpgTop.Controls.AddRange({lblUpgradesTit, lblUpgradesTotal, btnAtualizarUpgrades})

        Me.dgvUpgrades      = New DataGridView()
        Me.dgvUpgrades.Dock = DockStyle.Fill

        Me.tabUpgrades.Controls.Add(dgvUpgrades)
        Me.tabUpgrades.Controls.Add(pnlUpgTop)

        ' ─── ABA IMPORTAÇÃO ───────────────────────────────────────────
        Me.tabImportacao.Padding = New Padding(20)

        Dim lblImpTit As New Label() With {
            .Text = "Importação de Planilha Excel",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .Location = New Point(20, 15),
            .AutoSize = True
        }

        Dim lblImpSub As New Label() With {
            .Text = "Carregue a planilha da GBS (.xlsx). Todas as abas serão processadas com UPSERT.",
            .Location = New Point(20, 45),
            .Size = New Size(900, 20),
            .ForeColor = TemaEscuro.TextoMutado
        }

        Dim lblImpArq As New Label() With {
            .Text = "Arquivo:",
            .Location = New Point(20, 90),
            .AutoSize = True,
            .ForeColor = TemaEscuro.TextoMutado
        }

        Me.lblImpArquivo = New Label()
        Me.lblImpArquivo.Text = "(nenhum selecionado)"
        Me.lblImpArquivo.Location = New Point(90, 90)
        Me.lblImpArquivo.Size = New Size(800, 20)
        Me.lblImpArquivo.ForeColor = TemaEscuro.TextoMutado
        Me.lblImpArquivo.Font = New Font("Segoe UI", 9, FontStyle.Italic)

        Me.btnImpSelecionar = New Button()
        Me.btnImpSelecionar.Text = "Selecionar Planilha..."
        Me.btnImpSelecionar.Location = New Point(20, 120)
        Me.btnImpSelecionar.Size = New Size(200, 36)

        Me.btnImpIniciar = New Button()
        Me.btnImpIniciar.Text = "Iniciar Importação"
        Me.btnImpIniciar.Location = New Point(230, 120)
        Me.btnImpIniciar.Size = New Size(200, 36)
        Me.btnImpIniciar.Enabled = False

        Me.progImp = New ProgressBar()
        Me.progImp.Location = New Point(20, 175)
        Me.progImp.Size = New Size(900, 20)
        Me.progImp.Style = ProgressBarStyle.Continuous
        Me.progImp.Visible = False

        Me.lblImpStatus = New Label()
        Me.lblImpStatus.Text = "Aguardando..."
        Me.lblImpStatus.Location = New Point(20, 200)
        Me.lblImpStatus.Size = New Size(900, 22)
        Me.lblImpStatus.ForeColor = TemaEscuro.TextoMutado

        Me.lblImpInseridos = New Label() With {.Location = New Point(20, 240), .Size = New Size(280, 22), .Text = "Inseridos: —", .Font = New Font("Segoe UI", 10, FontStyle.Bold), .ForeColor = TemaEscuro.Accent}
        Me.lblImpAtualizados = New Label() With {.Location = New Point(20, 265), .Size = New Size(280, 22), .Text = "Atualizados: —", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.Texto}
        Me.lblImpAbas = New Label() With {.Location = New Point(20, 290), .Size = New Size(280, 22), .Text = "Abas processadas: —", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.Texto}
        Me.lblImpLinhas = New Label() With {.Location = New Point(320, 240), .Size = New Size(280, 22), .Text = "Linhas lidas: —", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.Texto}
        Me.lblImpErros = New Label() With {.Location = New Point(320, 265), .Size = New Size(280, 22), .Text = "Erros: 0", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.TextoMutado}

        Me.txtImpErros = New TextBox()
        Me.txtImpErros.Location = New Point(20, 340)
        Me.txtImpErros.Size = New Size(900, 200)
        Me.txtImpErros.Multiline = True
        Me.txtImpErros.ScrollBars = ScrollBars.Vertical
        Me.txtImpErros.ReadOnly = True
        Me.txtImpErros.Font = New Font("Consolas", 8.5F)
        Me.txtImpErros.Visible = False

        Me.tabImportacao.Controls.AddRange({lblImpTit, lblImpSub, lblImpArq, lblImpArquivo,
                                             btnImpSelecionar, btnImpIniciar, progImp,
                                             lblImpStatus, lblImpInseridos, lblImpAtualizados,
                                             lblImpAbas, lblImpLinhas, lblImpErros, txtImpErros})

        ' ═════════════════════════════════════════════════════════════
        ' MONTAGEM DO TABCONTROL
        ' ═════════════════════════════════════════════════════════════
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
