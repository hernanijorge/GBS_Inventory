Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.IO

''' <summary>
''' Tela principal do GBS Inventory Manager
''' Interface moderna em tema escuro com abas superiores,
''' painel com borda teal, date picker no topo e botão power verde.
''' </summary>
Public Class frmPrincipal

#Region "Atributos"

    Private oEquipController As EquipamentoController
    Private oImportController As ImportacaoController
    Private oRemessaController As RemessaController
    Private oUpgradeController As UpgradeController
    Private sArquivoSelImp As String = ""

#End Region

#Region "Inicialização"

    Public Sub New()

        InitializeComponent()

        oEquipController = New EquipamentoController()
        oImportController = New ImportacaoController()
        oRemessaController = New RemessaController()
        oUpgradeController = New UpgradeController()

        TemaEscuro.aplicar(Me)

    End Sub

    Private Sub frmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            carregarDashboard()
            carregarEstoque()
            carregarRemessas()
            carregarUpgrades()

        Catch ex As Exception

            MessageBox.Show("Erro na carga inicial: " & ex.Message, "Erro",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region

#Region "Carregamento de dados"

    Private Sub carregarDashboard()

        Try

            Dim ds As DataSet = oEquipController.obterDashboardTotais(1)

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                Dim row As DataRow = ds.Tables(0).Rows(0)
                lblTotalUnidades.Text = "Total: " & row("TOTAL_UNIDADES").ToString()
                lblEmEstoque.Text = "Em Estoque: " & row("EM_ESTOQUE").ToString()
                lblCondicaoBoa.Text = "Condição Boa: " & row("CONDICAO_BOA").ToString()
                lblUpgrades30d.Text = "Upgrades 30d: " & row("UPGRADES_30D").ToString()
                lblRemessasAtivas.Text = "Remessas: " & row("REMESSAS_ATIVAS").ToString()

            End If

        Catch ex As Exception

            ' silencioso no dashboard — apenas mantém valores em branco
            lblTotalUnidades.Text = "Sem conexão"

        End Try

    End Sub

    Private Sub carregarEstoque()

        Try

            Dim ds As DataSet = oEquipController.buscarTodos(1)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvEstoque.DataSource = ds.Tables(0)
            End If

        Catch ex As Exception

            MessageBox.Show("Erro ao carregar estoque: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region

#Region "Ações dos botões"

    Private Sub btnBuscarUID_Click(sender As Object, e As EventArgs) Handles btnBuscarUID.Click

        Dim sUID As String = txtPesquisa.Text.Trim()

        If String.IsNullOrEmpty(sUID) Then
            carregarEstoque()
            Return
        End If

        Try

            Dim ds As DataSet = oEquipController.buscarPorFiltro(sUID, "", "", "")

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvEstoque.DataSource = ds.Tables(0)
            End If

        Catch ex As Exception

            MessageBox.Show("Erro na pesquisa: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnImportarPlanilha_Click(sender As Object, e As EventArgs) Handles btnImportarPlanilha.Click

        Dim frm As New frmImportacao()
        frm.ShowDialog(Me)
        carregarEstoque()
        carregarDashboard()

    End Sub

    Private Sub btnScanner_Click(sender As Object, e As EventArgs) Handles btnScanner.Click

        Dim frm As New frmScanner()
        frm.ShowDialog(Me)

    End Sub

    Private Sub btnAtualizar_Click(sender As Object, e As EventArgs) Handles btnAtualizar.Click

        carregarDashboard()
        carregarEstoque()
        carregarRemessas()
        carregarUpgrades()

    End Sub

    Private Sub txtPesquisa_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPesquisa.KeyDown

        If e.KeyCode = Keys.Enter Then
            btnBuscarUID_Click(sender, e)
        End If

    End Sub

#End Region

#Region "Aba Remessas"

    Private Sub carregarRemessas()

        Try

            Dim ds As DataSet = oRemessaController.buscarRemessasAtivas()

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvRemessas.DataSource = ds.Tables(0)
            End If

        Catch ex As Exception

            lblRemessasTit.Text = "Remessas Ativas (erro ao carregar)"

        End Try

    End Sub

    Private Sub btnNovaRemessa_Click(sender As Object, e As EventArgs) Handles btnNovaRemessa.Click

        Try
            Dim frm As New frmRemessa()
            frm.ShowDialog(Me)
            carregarRemessas()
            carregarDashboard()
            carregarEstoque()
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir tela de remessa: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnAtualizarRemessas_Click(sender As Object, e As EventArgs) Handles btnAtualizarRemessas.Click
        carregarRemessas()
    End Sub

#End Region

#Region "Aba Upgrades"

    Private Sub carregarUpgrades()

        Try

            Dim ds As DataSet = oUpgradeController.buscarRecentes(90)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvUpgrades.DataSource = ds.Tables(0)
                lblUpgradesTotal.Text = "Total: " & ds.Tables(0).Rows.Count.ToString()
            End If

        Catch ex As Exception

            lblUpgradesTit.Text = "Upgrades (erro ao carregar)"

        End Try

    End Sub

    Private Sub btnAtualizarUpgrades_Click(sender As Object, e As EventArgs) Handles btnAtualizarUpgrades.Click
        carregarUpgrades()
    End Sub

#End Region

#Region "Aba Importação"

    Private Sub btnImpSelecionar_Click(sender As Object, e As EventArgs) Handles btnImpSelecionar.Click

        Using ofd As New OpenFileDialog()
            ofd.Filter = "Planilhas Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*"
            ofd.Title = "Selecione a planilha da GBS"

            If ofd.ShowDialog() = DialogResult.OK Then
                sArquivoSelImp = ofd.FileName
                lblImpArquivo.Text = Path.GetFileName(ofd.FileName)
                lblImpArquivo.ForeColor = TemaEscuro.Accent
                btnImpIniciar.Enabled = True
            End If
        End Using

    End Sub

    Private Sub btnImpIniciar_Click(sender As Object, e As EventArgs) Handles btnImpIniciar.Click

        If String.IsNullOrEmpty(sArquivoSelImp) Then Return

        btnImpIniciar.Enabled = False
        btnImpSelecionar.Enabled = False

        lblImpStatus.Text = "Importando..."
        lblImpStatus.ForeColor = TemaEscuro.Accent
        progImp.Value = 0
        progImp.Visible = True

        Application.DoEvents()

        Try

            Dim callback As Action(Of Integer, Integer, String) =
                Sub(atual, total, uid)
                    Me.Invoke(Sub()
                                  progImp.Maximum = Math.Max(total, 1)
                                  progImp.Value = Math.Min(atual, total)
                                  lblImpStatus.Text = $"Processando {atual}/{total} — UID {uid}"
                                  Application.DoEvents()
                              End Sub)
                End Sub

            oImportController.importarPlanilha(sArquivoSelImp, callback)

            progImp.Value = progImp.Maximum
            lblImpStatus.Text = "Concluído!"
            lblImpStatus.ForeColor = TemaEscuro.Accent

            lblImpInseridos.Text = "Inseridos: " & oImportController.TotalInseridos
            lblImpAtualizados.Text = "Atualizados: " & oImportController.TotalAtualizados
            lblImpAbas.Text = "Abas processadas: " & oImportController.TotalAbas
            lblImpLinhas.Text = "Linhas lidas: " & oImportController.TotalLinhasPlanilha
            lblImpErros.Text = "Erros: " & oImportController.TotalErros

            If oImportController.TotalErros > 0 Then
                txtImpErros.Lines = oImportController.Erros.ToArray()
                txtImpErros.Visible = True
                lblImpErros.ForeColor = TemaEscuro.Vermelho
            End If

            carregarDashboard()
            carregarEstoque()

            MessageBox.Show($"Importação concluída!" & vbCrLf & vbCrLf &
                            $"Inseridos: {oImportController.TotalInseridos}" & vbCrLf &
                            $"Atualizados: {oImportController.TotalAtualizados}" & vbCrLf &
                            $"Erros: {oImportController.TotalErros}",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show("Erro na importação: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblImpStatus.Text = "Falha"
            lblImpStatus.ForeColor = TemaEscuro.Vermelho

        Finally

            btnImpIniciar.Enabled = True
            btnImpSelecionar.Enabled = True

        End Try

    End Sub

#End Region

End Class
