Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data

''' <summary>
''' Tela principal do GBS Inventory Manager
''' Interface moderna em tema escuro com abas superiores,
''' painel com borda teal, date picker no topo e botão power verde.
''' </summary>
Public Class frmPrincipal

#Region "Atributos"

    Private oEquipController As EquipamentoController

#End Region

#Region "Inicialização"

    Public Sub New()

        InitializeComponent()

        oEquipController = New EquipamentoController()

        TemaEscuro.aplicar(Me)

    End Sub

    Private Sub frmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            carregarDashboard()
            carregarEstoque()

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
                lblTotalUnidades.Text = "Total: "      & row("TOTAL_UNIDADES").ToString()
                lblEmEstoque.Text     = "Em Estoque: " & row("EM_ESTOQUE").ToString()
                lblCondicaoBoa.Text   = "Condição Boa: " & row("CONDICAO_BOA").ToString()
                lblUpgrades30d.Text   = "Upgrades 30d: " & row("UPGRADES_30D").ToString()
                lblRemessasAtivas.Text= "Remessas: "   & row("REMESSAS_ATIVAS").ToString()

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

    End Sub

    Private Sub txtPesquisa_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPesquisa.KeyDown

        If e.KeyCode = Keys.Enter Then
            btnBuscarUID_Click(sender, e)
        End If

    End Sub

#End Region

End Class
