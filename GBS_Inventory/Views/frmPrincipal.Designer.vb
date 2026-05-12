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
    Friend WithEvents tabCadastro As TabPage
    Friend WithEvents tabEstoque As TabPage
    Friend WithEvents tabRemessas As TabPage
    Friend WithEvents tabUpgrades   As TabPage
    Friend WithEvents tabComponents As TabPage
    Friend WithEvents tabImportacao As TabPage
    Friend WithEvents tabInvoice    As TabPage

    ' Controls — Components tab
    Friend WithEvents dgvComponents     As DataGridView
    Friend WithEvents lblCompTotal      As Label
    Friend WithEvents btnCompRefresh    As Button
    Friend WithEvents btnAddComponent   As Button
    Friend WithEvents txtCompSearch     As TextBox
    Friend WithEvents cboCompType       As ComboBox
    Friend WithEvents cboCompStatus     As ComboBox
    Friend WithEvents btnCompRelatorio  As Button
    Friend WithEvents btnCompAddToList  As Button
    Friend WithEvents btnCompClearList  As Button
    Friend WithEvents lblCompListaInfo  As Label
    Friend WithEvents pnlAcoes As Panel
    Friend WithEvents btnBuscarUID As Button
    Friend WithEvents btnScanner As Button
    Friend WithEvents btnImportarPlanilha As Button
    Friend WithEvents btnAtualizar As Button
    Friend WithEvents btnClientes As Button
    Friend WithEvents btnInvoice As Button
    Friend WithEvents txtPesquisa As TextBox
    Friend WithEvents lblPesquisa As Label
    Friend WithEvents dgvEstoque As DataGridView
    Friend WithEvents pnlDashboard As Panel
    Friend WithEvents lblTotalUnidades As Label
    Friend WithEvents lblEmEstoque As Label
    Friend WithEvents lblCondicaoBoa As Label
    Friend WithEvents lblUpgrades30d As Label
    Friend WithEvents lblRemessasAtivas As Label
    Friend WithEvents lblDashboardResumo As Label
    Friend WithEvents dgvDashboardResumo As DataGridView
    Friend WithEvents dgvDashManufacturer As DataGridView
    Friend WithEvents dgvDashModel        As DataGridView
    Friend WithEvents dgvDashCPU          As DataGridView

    Friend WithEvents lblVersao As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents dgvRemessas As DataGridView
    Friend WithEvents dgvUpgrades As DataGridView
    Friend WithEvents btnGerenciarUpgrade As Button

    ' Controles da aba Remessas
    Friend WithEvents lblRemessasTit       As Label
    Friend WithEvents btnNovaRemessa       As Button
    Friend WithEvents btnAtualizarRemessas As Button

    ' Controles da aba Upgrades
    Friend WithEvents lblUpgradesTit        As Label
    Friend WithEvents lblUpgradesTotal      As Label
    Friend WithEvents btnAtualizarUpgrades  As Button
    Friend WithEvents txtUpgradeBusca       As TextBox
    Friend WithEvents btnUpgradeBuscar      As Button
    Friend WithEvents btnUpgradeScanner     As Button
    Friend WithEvents btnUpgradeRelatorio   As Button
    Friend WithEvents btnUpgradeAddToList   As Button
    Friend WithEvents btnUpgradeClearList   As Button
    Friend WithEvents lblUpgradesListaInfo  As Label
    Friend WithEvents cboUpgBatch           As ComboBox

    ' Botão de adição manual de equipamento
    Friend WithEvents btnAddEquipamento As Button

    ' Rodapé da aba Inventory
    Friend WithEvents pnlEstoqueFooter As Panel
    Friend WithEvents lblEstoqueFooter As Label

    ' Painel de filtros da aba Inventory
    Friend WithEvents pnlFiltros         As Panel
    Friend WithEvents clbManufacturer    As CheckedListBox
    Friend WithEvents clbModel           As CheckedListBox
    Friend WithEvents clbStatus          As CheckedListBox
    Friend WithEvents clbProcessor       As CheckedListBox
    Friend WithEvents clbBatch           As CheckedListBox
    Friend WithEvents btnApplyFilter     As Button
    Friend WithEvents btnClearFilter     As Button
    Friend WithEvents btnGenerateReport  As Button
    Friend WithEvents btnExportExcel     As Button
    Friend WithEvents lblFiltrosAtivos   As Label
    Friend WithEvents btnAddToList       As Button
    Friend WithEvents btnClearList       As Button

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
    Friend WithEvents lblImpAnaliseTitle As Label
    Friend WithEvents chkImpShowIssuesOnly As CheckBox
    Friend WithEvents btnImpGerarRelatorio As Button
    Friend WithEvents btnImpEnviarRelatorio As Button
    Friend WithEvents dgvImpAnalise As DataGridView
    Friend WithEvents lblImpAnaliseFooter As Label

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        ' ═════════════════════════════════════════════════════════════
        ' FORM PRINCIPAL
        ' ═════════════════════════════════════════════════════════════
        Me.Text = "GBS Inventory Manager — Global Business Solution, Lexington"
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
        Me.lblSubTitulo.Text = "Global Business Solution · Lexington, MA · v1.0"
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

        Me.tabDashboard  = New TabPage("Dashboard")
        Me.tabCadastro   = New TabPage("Registration")
        Me.tabEstoque    = New TabPage("Inventory")
        Me.tabRemessas   = New TabPage("Shipments")
        Me.tabUpgrades   = New TabPage("Upgrades")
        Me.tabComponents = New TabPage("Components")
        Me.tabImportacao = New TabPage("Import")
        Me.tabInvoice = New TabPage("Invoice")

        ' ─── ABA DASHBOARD ────────────────────────────────────────────
        Me.pnlDashboard = New Panel()
        Me.pnlDashboard.Dock = DockStyle.Fill
        Me.pnlDashboard.Padding = New Padding(40, 30, 40, 30)

        ' Adiciona o painel à aba ANTES de criar os cards (criarCardDashboard adiciona no pnlDashboard)
        Me.tabDashboard.Controls.Add(pnlDashboard)

        Me.lblTotalUnidades = criarCardDashboard("Total Units", "—", 20, 20)
        Me.lblEmEstoque = criarCardDashboard("In Stock", "—", 260, 20)
        Me.lblCondicaoBoa = criarCardDashboard("Good Condition", "—", 500, 20)
        Me.lblUpgrades30d = criarCardDashboard("Upgrades 30d", "—", 740, 20)
        Me.lblRemessasAtivas = criarCardDashboard("Active Shipments", "—", 980, 20)

        Me.lblDashboardResumo = New Label() With {
            .Text = "Inventory by Manufacturer / Model / CPU Family",
            .Location = New Point(20, 150),
            .Size = New Size(520, 24),
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        Me.dgvDashboardResumo = New DataGridView()
        Me.dgvDashboardResumo.Location = New Point(20, 180)
        Me.dgvDashboardResumo.Size = New Size(1180, 420)
        Me.dgvDashboardResumo.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        Me.dgvDashboardResumo.ReadOnly = True
        Me.dgvDashboardResumo.MultiSelect = False
        Me.dgvDashboardResumo.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        ' ── Dashboard breakdown grids ────────────────────────────────
        Dim lblDashMfr As New Label() With {
            .Text      = "Manufacturer",
            .Location  = New Point(20, 148),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent,
            .Font      = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        Me.dgvDashManufacturer = New DataGridView()
        Me.dgvDashManufacturer.Location      = New Point(20, 172)
        Me.dgvDashManufacturer.Size          = New Size(370, 330)
        Me.dgvDashManufacturer.ReadOnly      = True
        Me.dgvDashManufacturer.MultiSelect   = False
        Me.dgvDashManufacturer.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.dgvDashManufacturer.Anchor        = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom

        Dim lblDashMdl As New Label() With {
            .Text      = "Model",
            .Location  = New Point(410, 148),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent,
            .Font      = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        Me.dgvDashModel = New DataGridView()
        Me.dgvDashModel.Location      = New Point(410, 172)
        Me.dgvDashModel.Size          = New Size(370, 330)
        Me.dgvDashModel.ReadOnly      = True
        Me.dgvDashModel.MultiSelect   = False
        Me.dgvDashModel.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.dgvDashModel.Anchor        = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom

        Dim lblDashCpu As New Label() With {
            .Text      = "CPU Family",
            .Location  = New Point(800, 148),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent,
            .Font      = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        Me.dgvDashCPU = New DataGridView()
        Me.dgvDashCPU.Location      = New Point(800, 172)
        Me.dgvDashCPU.Size          = New Size(370, 330)
        Me.dgvDashCPU.ReadOnly      = True
        Me.dgvDashCPU.MultiSelect   = False
        Me.dgvDashCPU.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.dgvDashCPU.Anchor        = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom

        Me.pnlDashboard.Controls.Add(lblDashboardResumo)
        Me.pnlDashboard.Controls.Add(dgvDashboardResumo)
        Me.pnlDashboard.Controls.AddRange({lblDashMfr, dgvDashManufacturer,
                                            lblDashMdl, dgvDashModel,
                                            lblDashCpu, dgvDashCPU})

        ' --- ABA CADASTRO ---
        Dim pnlCadastro As New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = TemaEscuro.Fundo
        }

        Dim lblCadastro As New Label() With {
            .Text = "Registrations",
            .Location = New Point(30, 30),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent
        }

        Me.btnClientes = New Button()
        Me.btnClientes.Text = "Clients"
        Me.btnClientes.Location = New Point(30, 80)
        Me.btnClientes.Size = New Size(160, 34)

        pnlCadastro.Controls.Add(lblCadastro)
        pnlCadastro.Controls.Add(btnClientes)
        Me.tabCadastro.Controls.Add(pnlCadastro)

        ' ─── ABA ESTOQUE ──────────────────────────────────────────────

        ' ── Barra de ação (busca + botões) ────────────────────────────
        Me.pnlAcoes = New Panel()
        Me.pnlAcoes.Dock = DockStyle.Top
        Me.pnlAcoes.Height = 50
        Me.pnlAcoes.BackColor = TemaEscuro.Surface
        Me.pnlAcoes.Padding = New Padding(8)

        Me.lblPesquisa = New Label()
        Me.lblPesquisa.Text = "Search (UID / Serial / Model):"
        Me.lblPesquisa.Location = New Point(10, 15)
        Me.lblPesquisa.AutoSize = True
        Me.lblPesquisa.ForeColor = TemaEscuro.TextoMutado

        Me.txtPesquisa = New TextBox()
        Me.txtPesquisa.Location = New Point(215, 12)
        Me.txtPesquisa.Size = New Size(250, 24)

        Me.btnBuscarUID = New Button()
        Me.btnBuscarUID.Text = "Search"
        Me.btnBuscarUID.Location = New Point(475, 10)
        Me.btnBuscarUID.Size = New Size(85, 28)

        Me.btnScanner = New Button()
        Me.btnScanner.Text = "Scanner UID"
        Me.btnScanner.Location = New Point(570, 10)
        Me.btnScanner.Size = New Size(110, 28)

        Me.btnImportarPlanilha = New Button()
        Me.btnImportarPlanilha.Text = "Import Spreadsheet"
        Me.btnImportarPlanilha.Location = New Point(690, 10)
        Me.btnImportarPlanilha.Size = New Size(130, 28)

        Me.btnAtualizar = New Button()
        Me.btnAtualizar.Text = "Refresh"
        Me.btnAtualizar.Location = New Point(830, 10)
        Me.btnAtualizar.Size = New Size(90, 28)

        Me.btnAddEquipamento = New Button()
        Me.btnAddEquipamento.Text      = "+ Add Equipment"
        Me.btnAddEquipamento.Location  = New Point(930, 10)
        Me.btnAddEquipamento.Size      = New Size(120, 28)
        Me.btnAddEquipamento.BackColor = TemaEscuro.Accent
        Me.btnAddEquipamento.ForeColor = TemaEscuro.Fundo
        Me.btnAddEquipamento.FlatStyle = FlatStyle.Flat

        Me.pnlAcoes.Controls.AddRange({lblPesquisa, txtPesquisa, btnBuscarUID,
                                        btnScanner, btnImportarPlanilha, btnAtualizar,
                                        btnAddEquipamento})

        ' ── Painel de filtros ─────────────────────────────────────────
        Me.pnlFiltros = New Panel()
        Me.pnlFiltros.Dock      = DockStyle.Top
        Me.pnlFiltros.Height    = 205
        Me.pnlFiltros.BackColor = TemaEscuro.Fundo
        Me.pnlFiltros.Padding   = New Padding(0)

        ' ── Card: Manufacturer ──────────────────────────────────────
        Dim pnlGrpMfr As New Panel() With {
            .Location  = New Point(6, 6),
            .Size      = New Size(230, 150),
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlGrpMfr.Paint, Sub(sender, e)
                                        Using pen As New Pen(TemaEscuro.Borda, 1)
                                            e.Graphics.DrawRectangle(pen, 0, 0, pnlGrpMfr.Width - 1, pnlGrpMfr.Height - 1)
                                        End Using
                                    End Sub
        Dim lblMfrTitle As New Label() With {
            .Text      = "Manufacturer",
            .Location  = New Point(8, 5),
            .AutoSize  = True,
            .Font      = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .BackColor = Color.Transparent
        }
        Me.clbManufacturer              = New CheckedListBox()
        Me.clbManufacturer.Location     = New Point(4, 24)
        Me.clbManufacturer.Size         = New Size(222, 120)
        Me.clbManufacturer.CheckOnClick = True
        Me.clbManufacturer.Font         = New Font("Segoe UI", 9)
        Me.clbManufacturer.BorderStyle  = BorderStyle.None
        pnlGrpMfr.Controls.AddRange({lblMfrTitle, clbManufacturer})

        ' ── Card: Model ─────────────────────────────────────────────
        Dim pnlGrpMdl As New Panel() With {
            .Location  = New Point(242, 6),
            .Size      = New Size(230, 150),
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlGrpMdl.Paint, Sub(sender, e)
                                        Using pen As New Pen(TemaEscuro.Borda, 1)
                                            e.Graphics.DrawRectangle(pen, 0, 0, pnlGrpMdl.Width - 1, pnlGrpMdl.Height - 1)
                                        End Using
                                    End Sub
        Dim lblMdlTitle As New Label() With {
            .Text      = "Model",
            .Location  = New Point(8, 5),
            .AutoSize  = True,
            .Font      = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .BackColor = Color.Transparent
        }
        Me.clbModel              = New CheckedListBox()
        Me.clbModel.Location     = New Point(4, 24)
        Me.clbModel.Size         = New Size(222, 120)
        Me.clbModel.CheckOnClick = True
        Me.clbModel.Font         = New Font("Segoe UI", 9)
        Me.clbModel.BorderStyle  = BorderStyle.None
        pnlGrpMdl.Controls.AddRange({lblMdlTitle, clbModel})

        ' ── Card: Processor ─────────────────────────────────────────
        Dim pnlGrpProc As New Panel() With {
            .Location  = New Point(478, 6),
            .Size      = New Size(230, 150),
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlGrpProc.Paint, Sub(sender, e)
                                         Using pen As New Pen(TemaEscuro.Borda, 1)
                                             e.Graphics.DrawRectangle(pen, 0, 0, pnlGrpProc.Width - 1, pnlGrpProc.Height - 1)
                                         End Using
                                     End Sub
        Dim lblProcTitle As New Label() With {
            .Text      = "Processor",
            .Location  = New Point(8, 5),
            .AutoSize  = True,
            .Font      = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .BackColor = Color.Transparent
        }
        Me.clbProcessor              = New CheckedListBox()
        Me.clbProcessor.Location     = New Point(4, 24)
        Me.clbProcessor.Size         = New Size(222, 120)
        Me.clbProcessor.CheckOnClick = True
        Me.clbProcessor.Font         = New Font("Segoe UI", 9)
        Me.clbProcessor.BorderStyle  = BorderStyle.None
        pnlGrpProc.Controls.AddRange({lblProcTitle, clbProcessor})

        ' ── Card: Status ────────────────────────────────────────────
        Dim pnlGrpSt As New Panel() With {
            .Location  = New Point(714, 6),
            .Size      = New Size(165, 150),
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlGrpSt.Paint, Sub(sender, e)
                                       Using pen As New Pen(TemaEscuro.Borda, 1)
                                           e.Graphics.DrawRectangle(pen, 0, 0, pnlGrpSt.Width - 1, pnlGrpSt.Height - 1)
                                       End Using
                                   End Sub
        Dim lblStTitle As New Label() With {
            .Text      = "Status",
            .Location  = New Point(8, 5),
            .AutoSize  = True,
            .Font      = New Font("Segoe UI", 8.5F, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .BackColor = Color.Transparent
        }
        Me.clbStatus              = New CheckedListBox()
        Me.clbStatus.Location     = New Point(4, 24)
        Me.clbStatus.Size         = New Size(157, 120)
        Me.clbStatus.CheckOnClick = True
        Me.clbStatus.Font         = New Font("Segoe UI", 9)
        Me.clbStatus.BorderStyle  = BorderStyle.None
        Me.clbStatus.Items.AddRange({"IN_STOCK", "SHIPPED", "SOLD", "SCRAPPED", "IN_REPAIR"})
        pnlGrpSt.Controls.AddRange({lblStTitle, clbStatus})

        ' ── Card: Source Batch ──────────────────────────────────────
        Dim pnlGrpBatch As New Panel() With {
            .Location  = New Point(885, 6),
            .Size      = New Size(200, 150),
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlGrpBatch.Paint, Sub(sender, e)
                                          Using pen As New Pen(TemaEscuro.Borda, 1)
                                              e.Graphics.DrawRectangle(pen, 0, 0, pnlGrpBatch.Width - 1, pnlGrpBatch.Height - 1)
                                          End Using
                                      End Sub
        Dim lblBatchTitle As New Label() With {
            .Text      = "Source Batch",
            .Location  = New Point(4, 4),
            .AutoSize  = True,
            .Font      = New Font("Segoe UI", 8, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .BackColor = Color.Transparent
        }
        Me.clbBatch              = New CheckedListBox()
        Me.clbBatch.Location     = New Point(4, 24)
        Me.clbBatch.Size         = New Size(192, 120)
        Me.clbBatch.CheckOnClick = True
        Me.clbBatch.Font         = New Font("Segoe UI", 9)
        Me.clbBatch.BorderStyle  = BorderStyle.None
        pnlGrpBatch.Controls.AddRange({lblBatchTitle, clbBatch})

        ' ── Botões e contador de filtros ativos ─────────────────────
        Me.btnApplyFilter           = New Button()
        Me.btnApplyFilter.Text      = "Apply Filter"
        Me.btnApplyFilter.Location  = New Point(6, 164)
        Me.btnApplyFilter.Size      = New Size(130, 34)
        Me.btnApplyFilter.BackColor = Color.FromArgb(0, 188, 212)
        Me.btnApplyFilter.ForeColor = Color.White
        Me.btnApplyFilter.FlatStyle = FlatStyle.Flat
        Me.btnApplyFilter.Font      = New Font("Segoe UI", 9, FontStyle.Bold)

        Me.btnClearFilter           = New Button()
        Me.btnClearFilter.Text      = "Clear"
        Me.btnClearFilter.Location  = New Point(142, 164)
        Me.btnClearFilter.Size      = New Size(80, 34)

        Me.btnGenerateReport           = New Button()
        Me.btnGenerateReport.Text      = "Generate Report"
        Me.btnGenerateReport.Location  = New Point(228, 164)
        Me.btnGenerateReport.Size      = New Size(130, 34)
        Me.btnGenerateReport.BackColor = TemaEscuro.Accent
        Me.btnGenerateReport.ForeColor = TemaEscuro.Fundo
        Me.btnGenerateReport.FlatStyle = FlatStyle.Flat

        Me.btnExportExcel           = New Button()
        Me.btnExportExcel.Text      = "Export Excel"
        Me.btnExportExcel.Location  = New Point(364, 164)
        Me.btnExportExcel.Size      = New Size(110, 34)
        Me.btnExportExcel.BackColor = Color.FromArgb(21, 128, 61)
        Me.btnExportExcel.ForeColor = Color.White
        Me.btnExportExcel.FlatStyle = FlatStyle.Flat

        Me.btnAddToList           = New Button()
        Me.btnAddToList.Text      = "+ Add to List"
        Me.btnAddToList.Location  = New Point(480, 164)
        Me.btnAddToList.Size      = New Size(130, 34)
        Me.btnAddToList.BackColor = TemaEscuro.Accent
        Me.btnAddToList.ForeColor = TemaEscuro.Fundo
        Me.btnAddToList.FlatStyle = FlatStyle.Flat
        Me.btnAddToList.Font      = New Font("Segoe UI", 8.5F, FontStyle.Bold)

        Me.btnClearList           = New Button()
        Me.btnClearList.Text      = "Clear List"
        Me.btnClearList.Location  = New Point(616, 164)
        Me.btnClearList.Size      = New Size(96, 34)
        Me.btnClearList.BackColor = Color.FromArgb(124, 45, 18)
        Me.btnClearList.ForeColor = Color.White
        Me.btnClearList.FlatStyle = FlatStyle.Flat
        Me.btnClearList.Enabled   = False

        Me.lblFiltrosAtivos           = New Label()
        Me.lblFiltrosAtivos.Text      = "No filters active"
        Me.lblFiltrosAtivos.Location  = New Point(720, 172)
        Me.lblFiltrosAtivos.AutoSize  = True
        Me.lblFiltrosAtivos.ForeColor = TemaEscuro.TextoMutado
        Me.lblFiltrosAtivos.Font      = New Font("Segoe UI", 9, FontStyle.Italic)

        Me.pnlFiltros.Controls.AddRange({pnlGrpMfr, pnlGrpMdl, pnlGrpProc, pnlGrpSt, pnlGrpBatch,
                                          btnApplyFilter, btnClearFilter,
                                          btnGenerateReport, btnExportExcel,
                                          btnAddToList, btnClearList,
                                          lblFiltrosAtivos})

        ' ── Grid de estoque ───────────────────────────────────────────
        Me.dgvEstoque = New DataGridView()
        Me.dgvEstoque.Dock = DockStyle.Fill

        ' ── Rodapé da aba Inventory ────────────────────────────────────
        Me.pnlEstoqueFooter = New Panel()
        Me.pnlEstoqueFooter.Dock      = DockStyle.Bottom
        Me.pnlEstoqueFooter.Height    = 26
        Me.pnlEstoqueFooter.BackColor = Color.FromArgb(20, 20, 28)

        Me.lblEstoqueFooter = New Label()
        Me.lblEstoqueFooter.AutoSize  = False
        Me.lblEstoqueFooter.Dock      = DockStyle.Fill
        Me.lblEstoqueFooter.Font      = New Font("Segoe UI", 9)
        Me.lblEstoqueFooter.ForeColor = Color.FromArgb(140, 140, 150)
        Me.lblEstoqueFooter.TextAlign = ContentAlignment.MiddleLeft
        Me.lblEstoqueFooter.Padding   = New Padding(10, 0, 0, 0)
        Me.lblEstoqueFooter.Text      = "Showing: 0 items"

        Me.pnlEstoqueFooter.Controls.Add(Me.lblEstoqueFooter)

        ' Separador visual entre filtros e grid
        Dim pnlSeparador As New Panel() With {
            .Dock      = DockStyle.Top,
            .Height    = 2,
            .BackColor = TemaEscuro.Borda
        }

        ' Ordem: Fill → Bottom → Top (last added = highest z-order = laid out first from edge)
        '   pnlAcoes (Top) → pnlFiltros (Top) → pnlSeparador (Top) → pnlEstoqueFooter (Bottom) → dgvEstoque (Fill)
        Me.tabEstoque.Controls.Add(dgvEstoque)
        Me.tabEstoque.Controls.Add(pnlEstoqueFooter)
        Me.tabEstoque.Controls.Add(pnlSeparador)
        Me.tabEstoque.Controls.Add(pnlFiltros)
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
        Me.lblRemessasTit.Text      = "Active Shipments"
        Me.lblRemessasTit.Font      = New Font("Segoe UI", 11, FontStyle.Bold)
        Me.lblRemessasTit.ForeColor = TemaEscuro.Accent
        Me.lblRemessasTit.Location  = New Point(15, 15)
        Me.lblRemessasTit.AutoSize  = True

        Me.btnNovaRemessa            = New Button()
        Me.btnNovaRemessa.Text       = "+ New Shipment"
        Me.btnNovaRemessa.Location   = New Point(200, 10)
        Me.btnNovaRemessa.Size       = New Size(150, 30)

        Me.btnAtualizarRemessas            = New Button()
        Me.btnAtualizarRemessas.Text       = "Refresh"
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
            .Height    = 90,
            .BackColor = TemaEscuro.Surface,
            .Padding   = New Padding(8)
        }

        Me.lblUpgradesTit           = New Label()
        Me.lblUpgradesTit.Text      = "Upgrades (last 90 days)"
        Me.lblUpgradesTit.Font      = New Font("Segoe UI", 11, FontStyle.Bold)
        Me.lblUpgradesTit.ForeColor = TemaEscuro.Accent
        Me.lblUpgradesTit.Location  = New Point(15, 12)
        Me.lblUpgradesTit.AutoSize  = True

        Me.lblUpgradesTotal           = New Label()
        Me.lblUpgradesTotal.Text      = "Total: 0"
        Me.lblUpgradesTotal.Font      = New Font("Segoe UI", 10)
        Me.lblUpgradesTotal.ForeColor = TemaEscuro.TextoMutado
        Me.lblUpgradesTotal.Location  = New Point(290, 14)
        Me.lblUpgradesTotal.AutoSize  = True

        Me.btnAtualizarUpgrades            = New Button()
        Me.btnAtualizarUpgrades.Text       = "Refresh"
        Me.btnAtualizarUpgrades.Location   = New Point(420, 8)
        Me.btnAtualizarUpgrades.Size       = New Size(90, 28)

        Me.btnGerenciarUpgrade = New Button()
        Me.btnGerenciarUpgrade.Text = "Apply Upgrade"
        Me.btnGerenciarUpgrade.Location = New Point(520, 8)
        Me.btnGerenciarUpgrade.Size = New Size(140, 28)

        Me.btnUpgradeRelatorio = New Button()
        Me.btnUpgradeRelatorio.Text = "Report"
        Me.btnUpgradeRelatorio.Location = New Point(668, 8)
        Me.btnUpgradeRelatorio.Size = New Size(90, 28)

        Me.btnUpgradeAddToList           = New Button()
        Me.btnUpgradeAddToList.Text      = "+ Add to List"
        Me.btnUpgradeAddToList.Location  = New Point(766, 8)
        Me.btnUpgradeAddToList.Size      = New Size(120, 28)
        Me.btnUpgradeAddToList.BackColor = TemaEscuro.Accent
        Me.btnUpgradeAddToList.ForeColor = TemaEscuro.Fundo
        Me.btnUpgradeAddToList.FlatStyle = FlatStyle.Flat
        Me.btnUpgradeAddToList.Font      = New Font("Segoe UI", 8.5F, FontStyle.Bold)

        Me.btnUpgradeClearList           = New Button()
        Me.btnUpgradeClearList.Text      = "Clear List"
        Me.btnUpgradeClearList.Location  = New Point(894, 8)
        Me.btnUpgradeClearList.Size      = New Size(96, 28)
        Me.btnUpgradeClearList.BackColor = Color.FromArgb(124, 45, 18)
        Me.btnUpgradeClearList.ForeColor = Color.White
        Me.btnUpgradeClearList.FlatStyle = FlatStyle.Flat
        Me.btnUpgradeClearList.Enabled   = False

        Me.lblUpgradesListaInfo          = New Label()
        Me.lblUpgradesListaInfo.Text     = ""
        Me.lblUpgradesListaInfo.Location = New Point(766, 42)
        Me.lblUpgradesListaInfo.AutoSize = True
        Me.lblUpgradesListaInfo.Font     = New Font("Segoe UI", 8.5F, FontStyle.Italic)
        Me.lblUpgradesListaInfo.ForeColor = TemaEscuro.Accent

        ' Segunda linha — busca rápida por UID/Serial
        Dim lblUpgBusca As New Label() With {
            .Text      = "UID / Serial:",
            .Location  = New Point(15, 55),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado,
            .Font      = New Font("Segoe UI", 9)
        }

        Me.txtUpgradeBusca          = New TextBox()
        Me.txtUpgradeBusca.Location = New Point(110, 52)
        Me.txtUpgradeBusca.Size     = New Size(220, 26)

        Me.btnUpgradeBuscar          = New Button()
        Me.btnUpgradeBuscar.Text     = "Search"
        Me.btnUpgradeBuscar.Location = New Point(340, 51)
        Me.btnUpgradeBuscar.Size     = New Size(85, 28)

        Me.btnUpgradeScanner          = New Button()
        Me.btnUpgradeScanner.Text     = "Scanner UID"
        Me.btnUpgradeScanner.Location = New Point(435, 51)
        Me.btnUpgradeScanner.Size     = New Size(120, 28)

        Dim lblUpgBatchFilter As New Label() With {
            .Text      = "Source Batch:",
            .Location  = New Point(575, 55),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado,
            .Font      = New Font("Segoe UI", 9)
        }

        Me.cboUpgBatch          = New ComboBox()
        Me.cboUpgBatch.Location = New Point(668, 52)
        Me.cboUpgBatch.Size     = New Size(220, 24)
        Me.cboUpgBatch.DropDownStyle = ComboBoxStyle.DropDownList

        pnlUpgTop.Controls.AddRange({lblUpgradesTit, lblUpgradesTotal, btnAtualizarUpgrades, btnGerenciarUpgrade,
                                      btnUpgradeRelatorio, btnUpgradeAddToList, btnUpgradeClearList,
                                      lblUpgradesListaInfo, lblUpgBusca, txtUpgradeBusca, btnUpgradeBuscar,
                                      btnUpgradeScanner, lblUpgBatchFilter, cboUpgBatch})

        Me.dgvUpgrades      = New DataGridView()
        Me.dgvUpgrades.Dock = DockStyle.Fill

        Me.tabUpgrades.Controls.Add(dgvUpgrades)
        Me.tabUpgrades.Controls.Add(pnlUpgTop)

        ' ─── ABA COMPONENTS ───────────────────────────────────────────
        Me.tabComponents.Padding = New Padding(0)

        Dim pnlCompTop As New Panel() With {
            .Dock      = DockStyle.Top,
            .Height    = 90,
            .BackColor = TemaEscuro.Surface,
            .Padding   = New Padding(8)
        }

        Dim lblCompTit As New Label() With {
            .Text      = "Components",
            .Font      = New Font("Segoe UI", 11, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .Location  = New Point(15, 15),
            .AutoSize  = True
        }

        Me.lblCompTotal           = New Label() With {
            .Text      = "Total: 0",
            .Font      = New Font("Segoe UI", 10),
            .ForeColor = TemaEscuro.TextoMutado,
            .Location  = New Point(180, 17),
            .AutoSize  = True
        }

        Me.btnCompRefresh          = New Button() With { .Text = "Refresh",       .Location = New Point(300, 12), .Size = New Size(90, 28) }
        Me.btnAddComponent         = New Button() With { .Text = "+ Add",         .Location = New Point(400, 12), .Size = New Size(90, 28), .BackColor = TemaEscuro.Accent, .ForeColor = TemaEscuro.Fundo, .FlatStyle = FlatStyle.Flat }

        Dim lblCSearch As New Label() With { .Text = "Search:", .Location = New Point(510, 16), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.txtCompSearch           = New TextBox() With { .Location = New Point(570, 12), .Size = New Size(160, 26) }

        Dim lblCType As New Label() With { .Text = "Type:", .Location = New Point(745, 16), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboCompType             = New ComboBox() With { .Location = New Point(785, 12), .Size = New Size(100, 26), .DropDownStyle = ComboBoxStyle.DropDownList }

        Dim lblCStat As New Label() With { .Text = "Status:", .Location = New Point(900, 16), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado }
        Me.cboCompStatus           = New ComboBox() With { .Location = New Point(948, 12), .Size = New Size(130, 26), .DropDownStyle = ComboBoxStyle.DropDownList }

        Me.btnCompRelatorio          = New Button() With { .Text = "Report",      .Location = New Point(15, 52),  .Size = New Size(90, 28) }
        Me.btnCompAddToList          = New Button() With { .Text = "+ Add to List", .Location = New Point(115, 52), .Size = New Size(110, 28),
                                                            .BackColor = TemaEscuro.Accent, .ForeColor = TemaEscuro.Fundo,
                                                            .FlatStyle = FlatStyle.Flat,
                                                            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold) }
        Me.btnCompClearList          = New Button() With { .Text = "Clear List",  .Location = New Point(235, 52), .Size = New Size(96, 28),
                                                            .BackColor = Color.FromArgb(124, 45, 18), .ForeColor = Color.White,
                                                            .FlatStyle = FlatStyle.Flat, .Enabled = False }
        Me.lblCompListaInfo          = New Label() With { .Text = "",             .Location = New Point(345, 58),
                                                           .AutoSize = True,
                                                           .Font = New Font("Segoe UI", 8.5F, FontStyle.Italic),
                                                           .ForeColor = TemaEscuro.Accent }

        pnlCompTop.Controls.AddRange({lblCompTit, lblCompTotal, btnCompRefresh, btnAddComponent,
                                       lblCSearch, txtCompSearch, lblCType, cboCompType, lblCStat, cboCompStatus,
                                       btnCompRelatorio, btnCompAddToList, btnCompClearList, lblCompListaInfo})

        Me.dgvComponents      = New DataGridView() With { .Dock = DockStyle.Fill }

        Me.tabComponents.Controls.Add(dgvComponents)
        Me.tabComponents.Controls.Add(pnlCompTop)

        ' ─── ABA IMPORTAÇÃO ───────────────────────────────────────────
        Me.tabImportacao.Padding = New Padding(20)

        Dim lblImpTit As New Label() With {
            .Text = "Excel Spreadsheet Import",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .Location = New Point(20, 15),
            .AutoSize = True
        }

        Dim lblImpSub As New Label() With {
            .Text = "Load the GBS spreadsheet (.xlsx). All sheets will be processed with UPSERT.",
            .Location = New Point(20, 45),
            .Size = New Size(900, 20),
            .ForeColor = TemaEscuro.TextoMutado
        }

        Dim lblImpArq As New Label() With {
            .Text = "File:",
            .Location = New Point(20, 90),
            .AutoSize = True,
            .ForeColor = TemaEscuro.TextoMutado
        }

        Me.lblImpArquivo = New Label()
        Me.lblImpArquivo.Text = "(none selected)"
        Me.lblImpArquivo.Location = New Point(90, 90)
        Me.lblImpArquivo.Size = New Size(800, 20)
        Me.lblImpArquivo.ForeColor = TemaEscuro.TextoMutado
        Me.lblImpArquivo.Font = New Font("Segoe UI", 9, FontStyle.Italic)

        Me.btnImpSelecionar = New Button()
        Me.btnImpSelecionar.Text = "Select Spreadsheet..."
        Me.btnImpSelecionar.Location = New Point(20, 120)
        Me.btnImpSelecionar.Size = New Size(200, 36)

        Me.btnImpIniciar = New Button()
        Me.btnImpIniciar.Text = "Start Import"
        Me.btnImpIniciar.Location = New Point(230, 120)
        Me.btnImpIniciar.Size = New Size(200, 36)
        Me.btnImpIniciar.Enabled = False

        Me.progImp = New ProgressBar()
        Me.progImp.Location = New Point(20, 175)
        Me.progImp.Size = New Size(900, 20)
        Me.progImp.Style = ProgressBarStyle.Continuous
        Me.progImp.Visible = False

        Me.lblImpStatus = New Label()
        Me.lblImpStatus.Text = "Waiting..."
        Me.lblImpStatus.Location = New Point(20, 200)
        Me.lblImpStatus.Size = New Size(900, 22)
        Me.lblImpStatus.ForeColor = TemaEscuro.TextoMutado

        Me.lblImpInseridos = New Label() With {.Location = New Point(20, 240), .Size = New Size(280, 22), .Text = "Inserted: —", .Font = New Font("Segoe UI", 10, FontStyle.Bold), .ForeColor = TemaEscuro.Accent}
        Me.lblImpAtualizados = New Label() With {.Location = New Point(20, 265), .Size = New Size(280, 22), .Text = "Updated: —", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.Texto}
        Me.lblImpAbas = New Label() With {.Location = New Point(20, 290), .Size = New Size(280, 22), .Text = "Sheets processed: —", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.Texto}
        Me.lblImpLinhas = New Label() With {.Location = New Point(320, 240), .Size = New Size(280, 22), .Text = "Lines read: —", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.Texto}
        Me.lblImpErros = New Label() With {.Location = New Point(320, 265), .Size = New Size(280, 22), .Text = "Errors: 0", .Font = New Font("Segoe UI", 10), .ForeColor = TemaEscuro.TextoMutado}

        Me.txtImpErros = New TextBox()
        Me.txtImpErros.Location = New Point(20, 340)
        Me.txtImpErros.Size = New Size(900, 200)
        Me.txtImpErros.Multiline = True
        Me.txtImpErros.ScrollBars = ScrollBars.Vertical
        Me.txtImpErros.ReadOnly = True
        Me.txtImpErros.Font = New Font("Consolas", 8.5F)
        Me.txtImpErros.Visible = False

        Me.lblImpAnaliseTitle = New Label()
        Me.lblImpAnaliseTitle.Text = "Import Quality Review"
        Me.lblImpAnaliseTitle.Location = New Point(20, 330)
        Me.lblImpAnaliseTitle.AutoSize = True
        Me.lblImpAnaliseTitle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        Me.lblImpAnaliseTitle.ForeColor = TemaEscuro.Accent
        Me.lblImpAnaliseTitle.Visible = False

        Me.chkImpShowIssuesOnly = New CheckBox()
        Me.chkImpShowIssuesOnly.Text = "Show issues only"
        Me.chkImpShowIssuesOnly.Location = New Point(210, 330)
        Me.chkImpShowIssuesOnly.AutoSize = True
        Me.chkImpShowIssuesOnly.ForeColor = TemaEscuro.Texto
        Me.chkImpShowIssuesOnly.BackColor = Color.Transparent
        Me.chkImpShowIssuesOnly.Visible = False

        Me.btnImpGerarRelatorio = New Button()
        Me.btnImpGerarRelatorio.Text = "Generate Import Report"
        Me.btnImpGerarRelatorio.Location = New Point(360, 324)
        Me.btnImpGerarRelatorio.Size = New Size(190, 30)
        Me.btnImpGerarRelatorio.Visible = False

        Me.btnImpEnviarRelatorio = New Button()
        Me.btnImpEnviarRelatorio.Text = "Send Report"
        Me.btnImpEnviarRelatorio.Location = New Point(560, 324)
        Me.btnImpEnviarRelatorio.Size = New Size(140, 30)
        Me.btnImpEnviarRelatorio.Visible = False

        Me.dgvImpAnalise = New DataGridView()
        Me.dgvImpAnalise.Location = New Point(20, 360)
        Me.dgvImpAnalise.Size = New Size(1160, 300)
        Me.dgvImpAnalise.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        Me.dgvImpAnalise.Visible = False

        Me.lblImpAnaliseFooter = New Label()
        Me.lblImpAnaliseFooter.Text = ""
        Me.lblImpAnaliseFooter.Location = New Point(20, 668)
        Me.lblImpAnaliseFooter.Size = New Size(1160, 22)
        Me.lblImpAnaliseFooter.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.lblImpAnaliseFooter.ForeColor = TemaEscuro.TextoMutado
        Me.lblImpAnaliseFooter.Visible = False

        Me.tabImportacao.Controls.AddRange({lblImpTit, lblImpSub, lblImpArq, lblImpArquivo,
                                             btnImpSelecionar, btnImpIniciar, progImp,
                                             lblImpStatus, lblImpInseridos, lblImpAtualizados,
                                             lblImpAbas, lblImpLinhas, lblImpErros, txtImpErros,
                                             lblImpAnaliseTitle, chkImpShowIssuesOnly,
                                             btnImpGerarRelatorio, btnImpEnviarRelatorio,
                                             dgvImpAnalise, lblImpAnaliseFooter})

        ' --- ABA INVOICE ---
        Dim pnlInvoice As New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = TemaEscuro.Fundo
        }

        Dim lblInvoice As New Label() With {
            .Text = "Invoice",
            .Location = New Point(30, 30),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent
        }

        Me.btnInvoice = New Button()
        Me.btnInvoice.Text = "Open Invoice"
        Me.btnInvoice.Location = New Point(30, 80)
        Me.btnInvoice.Size = New Size(160, 34)

        pnlInvoice.Controls.Add(lblInvoice)
        pnlInvoice.Controls.Add(btnInvoice)
        Me.tabInvoice.Controls.Add(pnlInvoice)

        ' ═════════════════════════════════════════════════════════════
        ' MONTAGEM DO TABCONTROL
        ' ═════════════════════════════════════════════════════════════
        Me.tabPrincipal.TabPages.AddRange({tabDashboard, tabCadastro, tabEstoque, tabRemessas, tabUpgrades, tabComponents, tabImportacao, tabInvoice})
        Me.tabPrincipal.SelectedIndex = 0
        Me.pnlContainer.Controls.Add(tabPrincipal)

        ' ═════════════════════════════════════════════════════════════
        ' STATUS BAR
        ' ═════════════════════════════════════════════════════════════
        Me.lblStatus = New Label()
        Me.lblStatus.Text = "Connected to Oracle"
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
