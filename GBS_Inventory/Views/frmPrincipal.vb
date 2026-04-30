Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.IO
Imports Oracle.ManagedDataAccess.Client

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
    Private dtResumoManufacturer As DataTable
    Private dtResumoModel As DataTable
    Private dtResumoCpuFamily As DataTable
    Private dtEstoqueCompleto As DataTable

#End Region

#Region "Inicialização"

    Public Sub New()

        InitializeComponent()

        oEquipController = New EquipamentoController()
        oImportController = New ImportacaoController()
        oRemessaController = New RemessaController()
        oUpgradeController = New UpgradeController()

        TemaEscuro.aplicarHelius(Me)
        ConfigurarCardsDashboard()
        ConfigurarMenuContextoEstoque()

    End Sub

    Private Sub frmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            carregarDashboard()
            carregarEstoque()
            carregarRemessas()
            carregarUpgrades()

        Catch ex As Exception

            MessageBox.Show("Initial load error: " & ex.Message, "Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region

#Region "Carregamento de dados"

    Private Sub carregarDashboard()

        Try

            Dim ds As DataSet = oEquipController.obterDashboardTotais(1)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                Dim row As DataRow = ds.Tables(0).Rows(0)
                lblTotalUnidades.Text = "Total: " & ObterNumeroColuna(row, {"TOTAL_UNIDADES"}).ToString()
                lblEmEstoque.Text = "Em Estoque: " & ObterNumeroColuna(row, {"EM_ESTOQUE"}).ToString()
                lblCondicaoBoa.Text = "Condição Boa: " & ObterNumeroColuna(row, {"CONDICAO_BOA"}).ToString()
                lblUpgrades30d.Text = "Upgrades 30d: " & ObterNumeroColuna(row, {"UPGRADES_30D"}).ToString()
                lblRemessasAtivas.Text = "Remessas: " & ObterNumeroColuna(row, {"REMESSAS_ATIVAS"}).ToString()

            End If

            Dim dsEstoque As DataSet = oEquipController.buscarTodos(1)
            If dsEstoque IsNot Nothing AndAlso dsEstoque.Tables.Count > 0 Then
                carregarDashboardResumoCards(dsEstoque.Tables(0))
            Else
                lblTotalUnidades.Text = "0"
                lblEmEstoque.Text = "0"
                lblRemessasAtivas.Text = "0"
            End If

        Catch ex As Exception

            ' silencioso no dashboard — apenas mantém valores em branco
            lblTotalUnidades.Text = "0"
            lblEmEstoque.Text = "0"
            lblRemessasAtivas.Text = "0"

        End Try


    End Sub

    Private Sub carregarEstoque()

        Try

            Dim ds As DataSet = oEquipController.buscarTodos(1)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dtEstoqueCompleto = ds.Tables(0)
                popularCombosFilter()
                aplicarFiltrosEstoque()
            End If

        Catch ex As Exception

            MessageBox.Show("Error loading inventory: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub popularCombosFilter()

        If dtEstoqueCompleto Is Nothing Then Return

        Dim colMarca As String = ObterNomeColuna(dtEstoqueCompleto, {"MANUFACTURER", "MARCA"})
        PopularComboDistinto(cboFilterManufacturer, colMarca)
        ' cboFilterModel é populado pelo cboFilterManufacturer_SelectedIndexChanged

    End Sub

    Private Sub recarregarModelosPorManufacturer()

        If dtEstoqueCompleto Is Nothing Then Return

        Dim colMarca  As String = ObterNomeColuna(dtEstoqueCompleto, {"MANUFACTURER", "MARCA"})
        Dim colModelo As String = ObterNomeColuna(dtEstoqueCompleto, {"MODEL", "MODELO"})
        If String.IsNullOrEmpty(colModelo) Then Return

        Dim sMarca As String = If(cboFilterManufacturer.SelectedIndex > 0,
                                  cboFilterManufacturer.SelectedItem.ToString(), "")
        Dim selAtual As String = If(cboFilterModel.SelectedIndex > 0,
                                    cboFilterModel.SelectedItem.ToString(), "")

        Dim modelos As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each row As DataRow In dtEstoqueCompleto.Rows
            If Not String.IsNullOrEmpty(sMarca) AndAlso Not String.IsNullOrEmpty(colMarca) Then
                If Not row(colMarca).ToString().Equals(sMarca, StringComparison.OrdinalIgnoreCase) Then Continue For
            End If
            If Not IsDBNull(row(colModelo)) Then
                Dim v As String = row(colModelo).ToString().Trim()
                If Not String.IsNullOrEmpty(v) Then modelos.Add(v)
            End If
        Next

        cboFilterModel.Items.Clear()
        cboFilterModel.Items.Add("(All)")
        For Each m As String In modelos
            cboFilterModel.Items.Add(m)
        Next

        Dim idx As Integer = cboFilterModel.Items.IndexOf(selAtual)
        cboFilterModel.SelectedIndex = If(idx > 0, idx, 0)

    End Sub

    Private Sub PopularComboDistinto(pCombo As ComboBox, pColuna As String)

        Dim selAtual As String = If(pCombo.SelectedIndex > 0, pCombo.SelectedItem.ToString(), "")
        pCombo.Items.Clear()
        pCombo.Items.Add("(All)")

        If String.IsNullOrEmpty(pColuna) OrElse dtEstoqueCompleto Is Nothing Then
            pCombo.SelectedIndex = 0
            Return
        End If

        Dim valores As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each row As DataRow In dtEstoqueCompleto.Rows
            If Not IsDBNull(row(pColuna)) Then
                Dim v As String = row(pColuna).ToString().Trim()
                If Not String.IsNullOrEmpty(v) Then valores.Add(v)
            End If
        Next

        For Each v As String In valores
            pCombo.Items.Add(v)
        Next

        Dim idx As Integer = pCombo.Items.IndexOf(selAtual)
        pCombo.SelectedIndex = If(idx > 0, idx, 0)

    End Sub

    Private Sub aplicarFiltrosEstoque()

        If dtEstoqueCompleto Is Nothing Then Return

        Dim texto  As String = txtPesquisa.Text.Trim().ToUpperInvariant()
        Dim sMarca As String = If(cboFilterManufacturer.SelectedIndex > 0, cboFilterManufacturer.SelectedItem.ToString(), "")
        Dim sModel As String = If(cboFilterModel.SelectedIndex > 0,        cboFilterModel.SelectedItem.ToString(), "")
        Dim sStatus As String = If(cboFilterStatus.SelectedIndex > 0,      cboFilterStatus.SelectedItem.ToString(), "")
        Dim colMarca  As String = ObterNomeColuna(dtEstoqueCompleto, {"MANUFACTURER", "MARCA"})
        Dim colModelo As String = ObterNomeColuna(dtEstoqueCompleto, {"MODEL", "MODELO"})

        Dim dtFiltrada As DataTable = dtEstoqueCompleto.Clone()

        For Each row As DataRow In dtEstoqueCompleto.Rows

            ' texto livre
            If Not String.IsNullOrEmpty(texto) Then
                Dim encontrou As Boolean = False
                For Each col As String In {"INTERNAL_UID", "SERIAL_NUMBER", "MODEL", "MODELO",
                                           "MARCA", "MANUFACTURER", "PROCESSADOR", "CPU_MODEL", "CPU_FAMILY"}
                    If dtEstoqueCompleto.Columns.Contains(col) AndAlso Not IsDBNull(row(col)) Then
                        If row(col).ToString().ToUpperInvariant().Contains(texto) Then
                            encontrou = True
                            Exit For
                        End If
                    End If
                Next
                If Not encontrou Then Continue For
            End If

            ' manufacturer
            If Not String.IsNullOrEmpty(sMarca) AndAlso Not String.IsNullOrEmpty(colMarca) Then
                If Not row(colMarca).ToString().Equals(sMarca, StringComparison.OrdinalIgnoreCase) Then Continue For
            End If

            ' model
            If Not String.IsNullOrEmpty(sModel) AndAlso Not String.IsNullOrEmpty(colModelo) Then
                If Not row(colModelo).ToString().Equals(sModel, StringComparison.OrdinalIgnoreCase) Then Continue For
            End If

            ' status
            If Not String.IsNullOrEmpty(sStatus) AndAlso dtEstoqueCompleto.Columns.Contains("STATUS") Then
                If Not row("STATUS").ToString().Equals(sStatus, StringComparison.OrdinalIgnoreCase) Then Continue For
            End If

            dtFiltrada.ImportRow(row)
        Next

        dgvEstoque.DataSource = dtFiltrada
        adicionarColunaCheckBox()

    End Sub

    Private Sub ExportarExcel()

        Dim itens As List(Of DataRow) = ColetarItensRelatorio()
        If itens.Count = 0 Then Return

        Try
            Dim outputPath As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"

            If Not System.IO.Directory.Exists(outputPath) Then
                System.IO.Directory.CreateDirectory(outputPath)
            End If

            Dim caminho As String = ReportService.GerarExcel(itens, outputPath)

            System.Diagnostics.Process.Start(caminho)

            MessageBox.Show("Excel exported successfully!" & vbCrLf & vbCrLf &
                            "File: " & caminho & vbCrLf &
                            "Items: " & itens.Count.ToString(),
                            "Export Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error exporting Excel: " & ex.Message, "Export Excel",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub adicionarColunaCheckBox()

        If dgvEstoque.Columns.Contains("_SEL") Then Return

        Dim chk As New DataGridViewCheckBoxColumn()
        chk.Name        = "_SEL"
        chk.HeaderText  = ""
        chk.Width       = 30
        chk.ReadOnly    = False
        chk.FillWeight  = 1
        dgvEstoque.Columns.Insert(0, chk)

    End Sub

    Private Function ColetarItensRelatorio() As List(Of DataRow)

        Dim marcadas As New List(Of DataRow)
        If dgvEstoque.Columns.Contains("_SEL") Then
            For Each gridRow As DataGridViewRow In dgvEstoque.Rows
                Dim cell As DataGridViewCheckBoxCell = TryCast(gridRow.Cells("_SEL"), DataGridViewCheckBoxCell)
                If cell IsNot Nothing AndAlso cell.Value IsNot Nothing AndAlso CBool(cell.Value) Then
                    Dim drv As DataRowView = TryCast(gridRow.DataBoundItem, DataRowView)
                    If drv IsNot Nothing Then marcadas.Add(drv.Row)
                End If
            Next
        End If

        If marcadas.Count > 0 Then Return marcadas

        Dim todos As New List(Of DataRow)
        For Each gridRow As DataGridViewRow In dgvEstoque.Rows
            Dim drv As DataRowView = TryCast(gridRow.DataBoundItem, DataRowView)
            If drv IsNot Nothing Then todos.Add(drv.Row)
        Next
        Return todos

    End Function

    Private Sub GerarRelatorio()

        Dim itens As List(Of DataRow) = ColetarItensRelatorio()
        If itens.Count = 0 Then Return

        Try
            Dim outputPath As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"

            Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
            If String.IsNullOrWhiteSpace(logoPath) Then logoPath = ""

            If Not System.IO.Directory.Exists(outputPath) Then
                System.IO.Directory.CreateDirectory(outputPath)
            End If

            Dim caminho As String = ReportService.GerarRelatorio(itens, outputPath, logoPath)

            System.Diagnostics.Process.Start(caminho)

            MessageBox.Show("Report generated successfully!" & vbCrLf & vbCrLf &
                            "File: " & caminho & vbCrLf &
                            "Items: " & itens.Count.ToString(),
                            "Generate Report", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error generating report: " & ex.Message, "Generate Report",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Function ObterNomeColuna(pTabela As DataTable, pCandidatos As String()) As String

        For Each nome As String In pCandidatos
            If pTabela.Columns.Contains(nome) Then Return nome
        Next
        Return ""

    End Function

#End Region

#Region "Ações dos botões"

    Private Sub btnBuscarUID_Click(sender As Object, e As EventArgs) Handles btnBuscarUID.Click
        aplicarFiltrosEstoque()
    End Sub

    Private Sub cboFilterManufacturer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFilterManufacturer.SelectedIndexChanged
        recarregarModelosPorManufacturer()
    End Sub

    Private Sub btnApplyFilter_Click(sender As Object, e As EventArgs) Handles btnApplyFilter.Click
        aplicarFiltrosEstoque()
    End Sub

    Private Sub btnClearFilter_Click(sender As Object, e As EventArgs) Handles btnClearFilter.Click
        txtPesquisa.Text = ""
        cboFilterManufacturer.SelectedIndex = 0
        cboFilterModel.SelectedIndex = 0
        cboFilterStatus.SelectedIndex = 0
        aplicarFiltrosEstoque()
    End Sub

    Private Sub btnGenerateReport_Click(sender As Object, e As EventArgs) Handles btnGenerateReport.Click
        GerarRelatorio()
    End Sub

    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click
        ExportarExcel()
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

    Private Sub btnClientes_Click(sender As Object, e As EventArgs) Handles btnClientes.Click

        Try
            Dim frm As New frmClientes()
            frm.ShowDialog(Me)
        Catch ex As Exception
            MessageBox.Show("Error opening clients: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnInvoice_Click(sender As Object, e As EventArgs) Handles btnInvoice.Click

        Try
            Dim frm As New frmInvoice()
            frm.ShowDialog(Me)
        Catch ex As Exception
            MessageBox.Show("Error opening invoice: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

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

            lblRemessasTit.Text = "Active Shipments (load error)"

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
            MessageBox.Show("Error opening shipment screen: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnAtualizarRemessas_Click(sender As Object, e As EventArgs) Handles btnAtualizarRemessas.Click
        carregarRemessas()
    End Sub

    Private Sub dgvRemessas_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRemessas.CellClick

        If e.RowIndex < 0 Then Return

        Try
            Dim frm As New frmRemessa()
            frm.ShowDialog(Me)
            carregarRemessas()
            carregarEstoque()
            carregarDashboard()
        Catch ex As Exception
            MessageBox.Show("Error opening shipment: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

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

            lblUpgradesTit.Text = "Upgrades (load error)"

        End Try

    End Sub

    Private Sub btnAtualizarUpgrades_Click(sender As Object, e As EventArgs) Handles btnAtualizarUpgrades.Click
        carregarUpgrades()
    End Sub

    Private Sub btnGerenciarUpgrade_Click(sender As Object, e As EventArgs) Handles btnGerenciarUpgrade.Click
        AbrirUpgradeDoEstoqueSelecionado()
    End Sub

    Private Sub txtUpgradeBusca_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUpgradeBusca.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            BuscarEquipamentoParaUpgrade()
        End If
    End Sub

    Private Sub btnUpgradeBuscar_Click(sender As Object, e As EventArgs) Handles btnUpgradeBuscar.Click
        BuscarEquipamentoParaUpgrade()
    End Sub

    Private Sub btnUpgradeScanner_Click(sender As Object, e As EventArgs) Handles btnUpgradeScanner.Click
        Dim frm As New frmScanner()
        frm.ShowDialog(Me)
        carregarUpgrades()
        carregarDashboard()
        carregarEstoque()
    End Sub

    Private Sub BuscarEquipamentoParaUpgrade()

        Dim sUID As String = txtUpgradeBusca.Text.Trim()
        If String.IsNullOrEmpty(sUID) Then Return

        Try

            Dim ds As DataSet = oEquipController.buscarPorUID(sUID)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                Dim row As DataRow = ds.Tables(0).Rows(0)
                Dim id As Integer = CInt(row("ID_EQUIPAMENTO"))
                Dim uid As String = row("INTERNAL_UID").ToString()

                Dim frm As New frmUpgrade(id, uid)
                If frm.ShowDialog(Me) = DialogResult.OK Then
                    carregarUpgrades()
                    carregarDashboard()
                    carregarEstoque()
                End If

            Else

                MessageBox.Show("Equipment not found: " & sUID, "Upgrade",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)

            End If

        Catch ex As Exception

            MessageBox.Show("Search error: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dgvEstoque_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvEstoque.CellDoubleClick
        If e.RowIndex < 0 Then Return
        AbrirHistoricoEquipamento()
    End Sub

    Private Sub AbrirHistoricoEquipamento()

        If dgvEstoque Is Nothing OrElse dgvEstoque.CurrentRow Is Nothing Then Return

        Dim row As DataGridViewRow = dgvEstoque.CurrentRow
        Dim id  As Integer = 0
        Dim uid As String  = ""

        If row.DataGridView.Columns.Contains("ID_EQUIPAMENTO") AndAlso row.Cells("ID_EQUIPAMENTO").Value IsNot Nothing Then
            Integer.TryParse(row.Cells("ID_EQUIPAMENTO").Value.ToString(), id)
        End If

        If row.DataGridView.Columns.Contains("INTERNAL_UID") AndAlso row.Cells("INTERNAL_UID").Value IsNot Nothing Then
            uid = row.Cells("INTERNAL_UID").Value.ToString()
        End If

        If id <= 0 Then
            MessageBox.Show("Equipment ID not found in selected row.", "History",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim frm As New frmHistoricoEquipamento(id, uid)
            frm.ShowDialog(Me)
        Catch ex As Exception
            MessageBox.Show("Error opening history: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub AbrirUpgradeDoEstoqueSelecionado()

        If dgvEstoque Is Nothing OrElse dgvEstoque.CurrentRow Is Nothing Then
            MessageBox.Show("Select equipment in the Inventory tab to register an upgrade.", "Upgrade",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim row As DataGridViewRow = dgvEstoque.CurrentRow
        Dim id As Integer = 0
        Dim uid As String = ""

        If row.DataGridView.Columns.Contains("ID_EQUIPAMENTO") AndAlso row.Cells("ID_EQUIPAMENTO").Value IsNot Nothing Then
            Integer.TryParse(row.Cells("ID_EQUIPAMENTO").Value.ToString(), id)
        End If

        If row.DataGridView.Columns.Contains("INTERNAL_UID") AndAlso row.Cells("INTERNAL_UID").Value IsNot Nothing Then
            uid = row.Cells("INTERNAL_UID").Value.ToString()
        End If

        If id <= 0 Then
            MessageBox.Show("Equipment ID not found in current selection.", "Upgrade",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim frm As New frmUpgrade(id, uid)
        If frm.ShowDialog(Me) = DialogResult.OK Then
            carregarEstoque()
            carregarUpgrades()
            carregarDashboard()
        End If

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

        lblImpStatus.Text = "Importing..."
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
                                  lblImpStatus.Text = $"Processing {atual}/{total} — UID {uid}"
                                  Application.DoEvents()
                              End Sub)
                End Sub

            oImportController.importarPlanilha(sArquivoSelImp, callback)

            progImp.Value = progImp.Maximum
            lblImpStatus.Text = "Done!"
            lblImpStatus.ForeColor = TemaEscuro.Accent

            lblImpInseridos.Text = "Inserted: " & oImportController.TotalInseridos
            lblImpAtualizados.Text = "Updated: " & oImportController.TotalAtualizados
            lblImpAbas.Text = "Sheets processed: " & oImportController.TotalAbas
            lblImpLinhas.Text = "Lines read: " & oImportController.TotalLinhasPlanilha
            lblImpErros.Text = "Errors: " & oImportController.TotalErros

            If oImportController.TotalErros > 0 Then
                txtImpErros.Lines = oImportController.Erros.ToArray()
                txtImpErros.Visible = True
                lblImpErros.ForeColor = TemaEscuro.Vermelho
            End If

            carregarDashboard()
            carregarEstoque()

            MessageBox.Show($"Import completed!" & vbCrLf & vbCrLf &
                            $"Inserted: {oImportController.TotalInseridos}" & vbCrLf &
                            $"Updated: {oImportController.TotalAtualizados}" & vbCrLf &
                            $"Errors: {oImportController.TotalErros}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show("Import error: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblImpStatus.Text = "Failed"
            lblImpStatus.ForeColor = TemaEscuro.Vermelho

        Finally

            btnImpIniciar.Enabled = True
            btnImpSelecionar.Enabled = True

        End Try

    End Sub

#End Region


    Private Sub carregarDashboardResumoCards(pTabela As DataTable)

        dtResumoManufacturer = CriarResumoAgrupado(pTabela, "MANUFACTURER")
        dtResumoModel = CriarResumoAgrupado(pTabela, "MODEL")
        dtResumoCpuFamily = CriarResumoAgrupado(pTabela, "CPU_FAMILY")

        lblTotalUnidades.Text = dtResumoManufacturer.Rows.Count.ToString()
        lblEmEstoque.Text = dtResumoModel.Rows.Count.ToString()
        lblRemessasAtivas.Text = dtResumoCpuFamily.Rows.Count.ToString()

        dgvDashManufacturer.DataSource = dtResumoManufacturer
        dgvDashModel.DataSource        = dtResumoModel
        dgvDashCPU.DataSource          = dtResumoCpuFamily

    End Sub

    Private Function CriarResumoAgrupado(pTabela As DataTable, pTipo As String) As DataTable

        Dim dt As New DataTable()
        dt.Columns.Add("ITEM", GetType(String))
        dt.Columns.Add("QTD_IN_STOCK", GetType(Integer))

        Dim mapaEstoque As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim mapaTodos As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim encontrouStatusReconhecido As Boolean = False

        For Each row As DataRow In pTabela.Rows
            Dim status As String = ObterTextoColuna(row, {"STATUS", "STATUS_EQUIPAMENTO", "STATUS_DESCRICAO"})
            Dim statusNormalizado As String = NormalizarStatus(status)
            If EhStatusConhecido(statusNormalizado) Then
                encontrouStatusReconhecido = True
            End If

            Dim chave As String = ""

            Select Case pTipo
                Case "MANUFACTURER"
                    chave = ObterTextoColuna(row, {"MANUFACTURER", "MARCA"})
                Case "MODEL"
                    chave = ObterTextoColuna(row, {"MODEL", "MODELO"})
                Case Else
                    chave = ObterTextoColuna(row, {"CPU_FAMILY", "CPU_MODEL", "PROCESSADOR"})
            End Select

            If String.IsNullOrWhiteSpace(chave) Then
                chave = "-"
            End If

            If Not mapaTodos.ContainsKey(chave) Then
                mapaTodos(chave) = 0
            End If
            mapaTodos(chave) += 1

            If EhStatusEstoque(statusNormalizado) Then
                If Not mapaEstoque.ContainsKey(chave) Then
                    mapaEstoque(chave) = 0
                End If
                mapaEstoque(chave) += 1
            End If
        Next

        Dim mapaFinal As Dictionary(Of String, Integer) = mapaEstoque
        If mapaFinal.Count = 0 AndAlso mapaTodos.Count > 0 AndAlso Not encontrouStatusReconhecido Then
            mapaFinal = mapaTodos
        End If

        For Each kv As KeyValuePair(Of String, Integer) In mapaFinal
            dt.Rows.Add(kv.Key, kv.Value)
        Next

        Dim dv As New DataView(dt)
        dv.Sort = "QTD_IN_STOCK DESC, ITEM ASC"

        Return dv.ToTable()

    End Function

    Private Sub ConfigurarCardsDashboard()

        ConfigurarCardInterativo(lblTotalUnidades, "MANUFACTURER", "Manufacturer (IN_STOCK)")
        ConfigurarCardInterativo(lblEmEstoque, "MODEL", "Model (IN_STOCK)")
        ConfigurarCardInterativo(lblRemessasAtivas, "CPU_FAMILY", "CPU Family (IN_STOCK)")

        If lblCondicaoBoa IsNot Nothing AndAlso lblCondicaoBoa.Parent IsNot Nothing Then
            lblCondicaoBoa.Parent.Visible = False
        End If

        If lblUpgrades30d IsNot Nothing AndAlso lblUpgrades30d.Parent IsNot Nothing Then
            lblUpgrades30d.Parent.Visible = False
        End If

        If lblRemessasAtivas IsNot Nothing AndAlso lblRemessasAtivas.Parent IsNot Nothing Then
            lblRemessasAtivas.Parent.Location = New Point(500, 20)
        End If

        If lblDashboardResumo IsNot Nothing Then
            lblDashboardResumo.Visible = False
        End If

        If dgvDashboardResumo IsNot Nothing Then
            dgvDashboardResumo.Visible = False
        End If

    End Sub

    Private Sub ConfigurarCardInterativo(pValor As Label, pChave As String, pTitulo As String)

        If pValor Is Nothing OrElse pValor.Parent Is Nothing Then
            Return
        End If

        Dim pnl As Control = pValor.Parent
        pnl.Tag = pChave
        pValor.Tag = pChave
        pValor.Cursor = Cursors.Hand
        pnl.Cursor = Cursors.Hand

        For Each ctrl As Control In pnl.Controls
            If TypeOf ctrl Is Label Then
                ctrl.Tag = pChave
                ctrl.Cursor = Cursors.Hand
            End If
        Next

        For Each ctrl As Control In pnl.Controls
            Dim lbl As Label = TryCast(ctrl, Label)
            If lbl IsNot Nothing AndAlso Not Object.ReferenceEquals(lbl, pValor) Then
                lbl.Text = pTitulo
                Exit For
            End If
        Next

        AddHandler pnl.Click, AddressOf dashboardCard_Click
        AddHandler pValor.Click, AddressOf dashboardCard_Click
        For Each ctrl As Control In pnl.Controls
            If TypeOf ctrl Is Label Then
                AddHandler ctrl.Click, AddressOf dashboardCard_Click
            End If
        Next

    End Sub

    Private Sub dashboardCard_Click(sender As Object, e As EventArgs)

        Dim ctrl As Control = TryCast(sender, Control)
        If ctrl Is Nothing Then Return

        Dim chave As String = ""
        If ctrl.Tag IsNot Nothing Then
            chave = ctrl.Tag.ToString()
        End If

        Select Case chave
            Case "MANUFACTURER"
                MostrarPopupResumo("Manufacturer in stock", dtResumoManufacturer)
            Case "MODEL"
                MostrarPopupResumo("Model in stock", dtResumoModel)
            Case "CPU_FAMILY"
                MostrarPopupResumo("CPU Family in stock", dtResumoCpuFamily)
        End Select

    End Sub

    Private Sub MostrarPopupResumo(pTitulo As String, pDados As DataTable)

        If pDados Is Nothing OrElse pDados.Rows.Count = 0 Then
            MessageBox.Show("No data to display.", "Dashboard",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim frm As New Form()
        frm.Text = pTitulo
        frm.StartPosition = FormStartPosition.CenterParent
        frm.Size = New Size(760, 520)
        frm.BackColor = TemaEscuro.HeliusFundo
        frm.ForeColor = TemaEscuro.HeliusTexto

        Dim dgv As New DataGridView()
        dgv.Dock = DockStyle.Fill
        dgv.ReadOnly = True
        dgv.MultiSelect = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.DataSource = pDados

        frm.Controls.Add(dgv)
        TemaEscuro.aplicarHelius(frm)
        frm.ShowDialog(Me)

    End Sub

    Private Function ObterTextoColuna(pRow As DataRow, pColunas As String()) As String

        If pRow Is Nothing OrElse pColunas Is Nothing Then
            Return ""
        End If

        For Each nome As String In pColunas
            If pRow.Table.Columns.Contains(nome) AndAlso Not IsDBNull(pRow(nome)) Then
                Return pRow(nome).ToString().Trim()
            End If
        Next

        Return ""

    End Function

    Private Function ObterNumeroColuna(pRow As DataRow, pColunas As String()) As Integer

        Dim s As String = ObterTextoColuna(pRow, pColunas)
        Dim v As Integer
        If Integer.TryParse(s, v) Then
            Return v
        End If
        Return 0

    End Function

    Private Function EhStatusEstoque(pStatusNormalizado As String) As Boolean

        Return pStatusNormalizado = "INSTOCK" OrElse
               pStatusNormalizado = "EMESTOQUE" OrElse
               pStatusNormalizado = "ESTOQUE"

    End Function

    Private Function EhStatusConhecido(pStatusNormalizado As String) As Boolean

        If String.IsNullOrWhiteSpace(pStatusNormalizado) Then
            Return False
        End If

        Select Case pStatusNormalizado
            Case "INSTOCK", "EMESTOQUE", "ESTOQUE",
                 "SOLD", "VENDIDO",
                 "SHIPPED", "ENVIADO",
                 "INTRANSIT", "EMTRANSITO",
                 "DELIVERED", "ENTREGUE",
                 "RETURNED", "DEVOLVIDO", "DEVOLVIDA",
                 "LABELCREATED", "PICKEDUP"
                Return True
        End Select

        Return False

    End Function

    Private Function NormalizarStatus(pStatus As String) As String

        If String.IsNullOrWhiteSpace(pStatus) Then
            Return ""
        End If

        Dim s As String = pStatus.Trim().ToUpperInvariant()
        s = s.Replace("_", "")
        s = s.Replace("-", "")
        s = s.Replace(" ", "")
        s = s.Replace("Á", "A").Replace("À", "A").Replace("Â", "A").Replace("Ã", "A")
        s = s.Replace("É", "E").Replace("Ê", "E")
        s = s.Replace("Í", "I")
        s = s.Replace("Ó", "O").Replace("Ô", "O").Replace("Õ", "O")
        s = s.Replace("Ú", "U")
        s = s.Replace("Ç", "C")
        Return s

    End Function

#Region "Menu de Contexto — Grid Inventory"

    Private Sub ConfigurarMenuContextoEstoque()

        Dim ctx As New ContextMenuStrip()

        ' ── Change Status → submenu ────────────────────────────────────
        Dim mnuStatus As New ToolStripMenuItem("Change Status")
        For Each sta As String In {"IN_STOCK", "SHIPPED", "SOLD", "SCRAPPED", "IN_REPAIR"}
            Dim item As New ToolStripMenuItem(sta)
            item.Tag = sta
            AddHandler item.Click, AddressOf mnuStatus_Click
            mnuStatus.DropDownItems.Add(item)
        Next
        ctx.Items.Add(mnuStatus)

        ' ── Edit Notes ─────────────────────────────────────────────────
        Dim mnuNotes As New ToolStripMenuItem("Edit Notes")
        AddHandler mnuNotes.Click, AddressOf mnuEditNotes_Click
        ctx.Items.Add(mnuNotes)

        ctx.Items.Add(New ToolStripSeparator())

        ' ── View History ───────────────────────────────────────────────
        Dim mnuHistory As New ToolStripMenuItem("View History")
        AddHandler mnuHistory.Click, AddressOf mnuViewHistory_Click
        ctx.Items.Add(mnuHistory)

        AddHandler ctx.Opening, AddressOf ctxEstoque_Opening
        dgvEstoque.ContextMenuStrip = ctx

    End Sub

    ' Seleciona a linha ao clicar com botão direito antes do menu abrir
    Private Sub dgvEstoque_MouseDown(sender As Object, e As MouseEventArgs) Handles dgvEstoque.MouseDown
        If e.Button <> MouseButtons.Right Then Return
        Dim hit As DataGridView.HitTestInfo = dgvEstoque.HitTest(e.X, e.Y)
        If hit.RowIndex >= 0 AndAlso hit.ColumnIndex >= 0 Then
            dgvEstoque.ClearSelection()
            dgvEstoque.Rows(hit.RowIndex).Selected = True
            dgvEstoque.CurrentCell = dgvEstoque.Rows(hit.RowIndex).Cells(hit.ColumnIndex)
        End If
    End Sub

    Private Sub ctxEstoque_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs)
        If dgvEstoque.CurrentRow Is Nothing OrElse dgvEstoque.CurrentRow.Index < 0 Then
            e.Cancel = True
        End If
    End Sub

    ' Retorna (ID_EQUIPAMENTO, INTERNAL_UID, válido) da linha selecionada
    Private Function ObterDadosLinhaSelecionada() As (Id As Integer, UID As String, Valid As Boolean)
        If dgvEstoque Is Nothing OrElse dgvEstoque.CurrentRow Is Nothing Then
            Return (0, "", False)
        End If
        Dim row As DataGridViewRow = dgvEstoque.CurrentRow
        Dim id  As Integer = 0
        Dim uid As String  = ""
        If row.DataGridView.Columns.Contains("ID_EQUIPAMENTO") AndAlso
           row.Cells("ID_EQUIPAMENTO").Value IsNot Nothing Then
            Integer.TryParse(row.Cells("ID_EQUIPAMENTO").Value.ToString(), id)
        End If
        If row.DataGridView.Columns.Contains("INTERNAL_UID") AndAlso
           row.Cells("INTERNAL_UID").Value IsNot Nothing Then
            uid = row.Cells("INTERNAL_UID").Value.ToString()
        End If
        Return (id, uid, id > 0)
    End Function

    Private Sub mnuStatus_Click(sender As Object, e As EventArgs)
        Dim item As ToolStripMenuItem = TryCast(sender, ToolStripMenuItem)
        If item Is Nothing Then Return
        Dim novoStatus As String = item.Tag.ToString()
        Dim dados = ObterDadosLinhaSelecionada()
        If Not dados.Valid Then Return

        Dim res As DialogResult = MessageBox.Show(
            "Change status of UID [" & dados.UID & "] to [" & novoStatus & "]?" & vbCrLf &
            "This action will be logged.",
            "Change Status", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res <> DialogResult.Yes Then Return

        ' Executa o UPDATE imediatamente (antes do menu fechar)
        Dim erroMsg As String = Nothing
        Try
            Dim cs As String = System.Configuration.ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim p0 As New OracleParameter("P_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
            Dim p1 As New OracleParameter("P_ID",     OracleDbType.Int32,    ParameterDirection.Input)
            p0.Value = novoStatus
            p1.Value = dados.Id
            OracleHelper.ExecuteNonQuery(cs, CommandType.Text,
                "UPDATE TBL_EQUIPAMENTO SET STATUS = :P_STATUS, DATA_ATUALIZACAO = SYSDATE" &
                " WHERE ID_EQUIPAMENTO = :P_ID", p0, p1)
        Catch ex As Exception
            erroMsg = ex.Message
        End Try

        ' Adia o reload e o MessageBox para depois do ContextMenuStrip fechar completamente.
        ' Alterar DataSource dentro do handler do menu causa crash pois o DGV ainda está
        ' processando eventos do ContextMenuStrip no mesmo ciclo de mensagens.
        Dim uid    As String = dados.UID
        Dim status As String = novoStatus
        Me.BeginInvoke(Sub()
            If erroMsg IsNot Nothing Then
                MessageBox.Show("Error updating status: " & erroMsg, "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            Try
                carregarEstoque()
                carregarDashboard()
                MessageBox.Show("Status updated successfully!" & vbCrLf &
                                "UID: " & uid & "  →  " & status,
                                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error reloading grid: " & ex.Message, "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub)
    End Sub

    Private Sub mnuEditNotes_Click(sender As Object, e As EventArgs)
        Dim dados = ObterDadosLinhaSelecionada()
        If Not dados.Valid Then Return

        Dim obsAtual As String = ""
        Dim row As DataGridViewRow = dgvEstoque.CurrentRow
        For Each colName As String In {"OBSERVACAO", "NOTES"}
            If row.DataGridView.Columns.Contains(colName) AndAlso
               row.Cells(colName).Value IsNot Nothing Then
                obsAtual = row.Cells(colName).Value.ToString()
                Exit For
            End If
        Next

        Dim novaObs As String = InputBox("Notes for UID: " & dados.UID, "Edit Notes", obsAtual)
        If novaObs = obsAtual Then Return

        ' Executa o UPDATE antes do menu fechar
        Dim erroMsg As String = Nothing
        Try
            Dim cs As String = System.Configuration.ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim p0 As New OracleParameter("P_OBS", OracleDbType.Varchar2, ParameterDirection.Input)
            Dim p1 As New OracleParameter("P_ID",  OracleDbType.Int32,    ParameterDirection.Input)
            p0.Value = If(String.IsNullOrWhiteSpace(novaObs), DBNull.Value, CObj(novaObs.Trim()))
            p1.Value = dados.Id

            Try
                OracleHelper.ExecuteNonQuery(cs, CommandType.Text,
                    "UPDATE TBL_EQUIPAMENTO SET OBSERVACAO = :P_OBS, DATA_ATUALIZACAO = SYSDATE" &
                    " WHERE ID_EQUIPAMENTO = :P_ID", p0, p1)
            Catch exObs As Exception
                If Not exObs.ToString().ToUpperInvariant().Contains("ORA-00904") Then Throw
                Dim p0b As New OracleParameter("P_OBS", OracleDbType.Varchar2, ParameterDirection.Input)
                Dim p1b As New OracleParameter("P_ID",  OracleDbType.Int32,    ParameterDirection.Input)
                p0b.Value = If(String.IsNullOrWhiteSpace(novaObs), DBNull.Value, CObj(novaObs.Trim()))
                p1b.Value = dados.Id
                OracleHelper.ExecuteNonQuery(cs, CommandType.Text,
                    "UPDATE TBL_EQUIPAMENTO SET NOTES = :P_OBS, DATA_ATUALIZACAO = SYSDATE" &
                    " WHERE ID_EQUIPAMENTO = :P_ID", p0b, p1b)
            End Try
        Catch ex As Exception
            erroMsg = ex.Message
        End Try

        ' Reload adiado — mesma razão do mnuStatus_Click
        Me.BeginInvoke(Sub()
            If erroMsg IsNot Nothing Then
                MessageBox.Show("Error updating notes: " & erroMsg, "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            Try
                carregarEstoque()
                MessageBox.Show("Notes updated successfully!", "Success",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Error reloading grid: " & ex.Message, "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub)
    End Sub

    Private Sub mnuViewHistory_Click(sender As Object, e As EventArgs)
        AbrirHistoricoEquipamento()
    End Sub

#End Region

End Class
