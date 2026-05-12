Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Net
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
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
    Private dtUpgradesCompleto As DataTable
    Private dtImportQuality As DataTable
    Private _todasSelecionadas As Boolean = False
    Private _listaRelatorio        As New ListaRelatorio()
    Private _listaUpgradeRelatorio As New ListaRelatorio("ID_UPGRADE")

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
                lblEmEstoque.Text = "In Stock: " & ObterNumeroColuna(row, {"EM_ESTOQUE"}).ToString()
                lblCondicaoBoa.Text = "Good Condition: " & ObterNumeroColuna(row, {"CONDICAO_BOA"}).ToString()
                lblUpgrades30d.Text = "Upgrades 30d: " & ObterNumeroColuna(row, {"UPGRADES_30D"}).ToString()
                lblRemessasAtivas.Text = "Shipments: " & ObterNumeroColuna(row, {"REMESSAS_ATIVAS"}).ToString()

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
        PopularCheckedListBox(clbManufacturer, colMarca)
        PopularCheckedListBox(clbBatch, "SOURCE_BATCH")
        recarregarModelosPorManufacturer()

    End Sub

    Private Sub recarregarModelosPorManufacturer()

        If dtEstoqueCompleto Is Nothing Then Return

        Dim colModelo As String = ObterNomeColuna(dtEstoqueCompleto, {"MODEL", "MODELO"})
        If String.IsNullOrEmpty(colModelo) Then Return

        Dim marcas As New List(Of String)()
        For Each item As Object In clbManufacturer.CheckedItems
            marcas.Add(item.ToString())
        Next

        PopularCheckedListBox(clbModel, colModelo, pFiltroMarca:=If(marcas.Count > 0, marcas, Nothing))
        recarregarProcessoresPorModelo()

    End Sub

    Private Sub recarregarProcessoresPorModelo()

        If dtEstoqueCompleto Is Nothing Then Return

        Dim colProc As String = ObterNomeColuna(dtEstoqueCompleto, {"PROCESSADOR", "CPU_MODEL"})
        If String.IsNullOrEmpty(colProc) Then Return

        Dim marcas As New List(Of String)()
        For Each item As Object In clbManufacturer.CheckedItems
            marcas.Add(item.ToString())
        Next

        Dim modelos As New List(Of String)()
        For Each item As Object In clbModel.CheckedItems
            modelos.Add(item.ToString())
        Next

        PopularCheckedListBox(clbProcessor, colProc,
                              pFiltroMarca:=If(marcas.Count > 0, marcas, Nothing),
                              pFiltroModelo:=If(modelos.Count > 0, modelos, Nothing))

    End Sub

    Private Sub PopularCheckedListBox(pClb As CheckedListBox, pColuna As String,
                                      Optional pFiltroMarca As List(Of String) = Nothing,
                                      Optional pFiltroModelo As List(Of String) = Nothing)

        Dim checked As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each item As Object In pClb.CheckedItems
            checked.Add(item.ToString())
        Next

        pClb.Items.Clear()

        If String.IsNullOrEmpty(pColuna) OrElse dtEstoqueCompleto Is Nothing Then Return
        If Not dtEstoqueCompleto.Columns.Contains(pColuna) Then Return

        Dim colMarca As String = ObterNomeColuna(dtEstoqueCompleto, {"MANUFACTURER", "MARCA"})
        Dim colModeloF As String = ObterNomeColuna(dtEstoqueCompleto, {"MODEL", "MODELO"})
        Dim valores As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In dtEstoqueCompleto.Rows

            If pFiltroMarca IsNot Nothing AndAlso pFiltroMarca.Count > 0 AndAlso Not String.IsNullOrEmpty(colMarca) Then
                Dim marca As String = If(IsDBNull(row(colMarca)), "", row(colMarca).ToString())
                Dim marcaOk As Boolean = False
                For Each m As String In pFiltroMarca
                    If String.Equals(m, marca, StringComparison.OrdinalIgnoreCase) Then marcaOk = True : Exit For
                Next
                If Not marcaOk Then Continue For
            End If

            If pFiltroModelo IsNot Nothing AndAlso pFiltroModelo.Count > 0 AndAlso Not String.IsNullOrEmpty(colModeloF) Then
                Dim modelo As String = If(IsDBNull(row(colModeloF)), "", row(colModeloF).ToString())
                Dim modeloOk As Boolean = False
                For Each m As String In pFiltroModelo
                    If String.Equals(m, modelo, StringComparison.OrdinalIgnoreCase) Then modeloOk = True : Exit For
                Next
                If Not modeloOk Then Continue For
            End If

            If Not IsDBNull(row(pColuna)) Then
                Dim v As String = row(pColuna).ToString().Trim()
                If Not String.IsNullOrEmpty(v) Then valores.Add(v)
            End If

        Next

        For Each v As String In valores
            pClb.Items.Add(v, checked.Contains(v))
        Next

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

        Dim texto As String = txtPesquisa.Text.Trim().ToUpperInvariant()

        Dim marcas As New List(Of String)()
        For Each item As Object In clbManufacturer.CheckedItems
            marcas.Add(item.ToString())
        Next

        Dim modelos As New List(Of String)()
        For Each item As Object In clbModel.CheckedItems
            modelos.Add(item.ToString())
        Next

        Dim statuses As New List(Of String)()
        For Each item As Object In clbStatus.CheckedItems
            statuses.Add(item.ToString())
        Next

        Dim processors As New List(Of String)()
        For Each item As Object In clbProcessor.CheckedItems
            processors.Add(item.ToString())
        Next

        Dim batches As New List(Of String)()
        For Each item As Object In clbBatch.CheckedItems
            batches.Add(item.ToString())
        Next

        Dim colMarca As String = ObterNomeColuna(dtEstoqueCompleto, {"MANUFACTURER", "MARCA"})
        Dim colModelo As String = ObterNomeColuna(dtEstoqueCompleto, {"MODEL", "MODELO"})
        Dim colProc As String = ObterNomeColuna(dtEstoqueCompleto, {"PROCESSADOR", "CPU_MODEL"})

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

            ' manufacturer — OR among checked, empty = all
            If marcas.Count > 0 AndAlso Not String.IsNullOrEmpty(colMarca) Then
                Dim v As String = If(IsDBNull(row(colMarca)), "", row(colMarca).ToString())
                Dim ok As Boolean = False
                For Each m As String In marcas
                    If String.Equals(m, v, StringComparison.OrdinalIgnoreCase) Then ok = True : Exit For
                Next
                If Not ok Then Continue For
            End If

            ' model — OR among checked, empty = all
            If modelos.Count > 0 AndAlso Not String.IsNullOrEmpty(colModelo) Then
                Dim v As String = If(IsDBNull(row(colModelo)), "", row(colModelo).ToString())
                Dim ok As Boolean = False
                For Each m As String In modelos
                    If String.Equals(m, v, StringComparison.OrdinalIgnoreCase) Then ok = True : Exit For
                Next
                If Not ok Then Continue For
            End If

            ' status — OR among checked, empty = all
            If statuses.Count > 0 AndAlso dtEstoqueCompleto.Columns.Contains("STATUS") Then
                Dim v As String = If(IsDBNull(row("STATUS")), "", row("STATUS").ToString())
                Dim ok As Boolean = False
                For Each s As String In statuses
                    If String.Equals(s, v, StringComparison.OrdinalIgnoreCase) Then ok = True : Exit For
                Next
                If Not ok Then Continue For
            End If

            ' processor — OR among checked, empty = all
            If processors.Count > 0 AndAlso Not String.IsNullOrEmpty(colProc) AndAlso
               dtEstoqueCompleto.Columns.Contains(colProc) Then
                Dim v As String = If(IsDBNull(row(colProc)), "", row(colProc).ToString())
                Dim ok As Boolean = False
                For Each p As String In processors
                    If String.Equals(p, v, StringComparison.OrdinalIgnoreCase) Then ok = True : Exit For
                Next
                If Not ok Then Continue For
            End If

            ' source batch — OR among checked, empty = all
            If batches.Count > 0 AndAlso dtEstoqueCompleto.Columns.Contains("SOURCE_BATCH") Then
                Dim v As String = If(IsDBNull(row("SOURCE_BATCH")), "", row("SOURCE_BATCH").ToString())
                Dim ok As Boolean = False
                For Each b As String In batches
                    If String.Equals(b, v, StringComparison.OrdinalIgnoreCase) Then ok = True : Exit For
                Next
                If Not ok Then Continue For
            End If

            dtFiltrada.ImportRow(row)
        Next

        _todasSelecionadas = False
        dgvEstoque.DataSource = dtFiltrada
        adicionarColunaCheckBox()
        ConfigurarColunasGrid()
        AtualizarRodapeEstoque()
        AtualizarContadorFiltros()

    End Sub

    Private Sub AtualizarRodapeEstoque()

        If lblEstoqueFooter Is Nothing Then Return

        Dim dtAtual As DataTable = TryCast(dgvEstoque.DataSource, DataTable)
        Dim showing As Integer = If(dtAtual IsNot Nothing, dtAtual.Rows.Count, 0)

        Dim inStock As Integer = 0
        Dim sold As Integer = 0
        Dim inRepair As Integer = 0
        Dim selected As Integer = 0

        If dtAtual IsNot Nothing AndAlso dtAtual.Columns.Contains("STATUS") Then
            For Each row As DataRow In dtAtual.Rows
                Dim st As String = If(IsDBNull(row("STATUS")), "",
                                      row("STATUS").ToString().ToUpperInvariant().Trim())
                Select Case st
                    Case "IN_STOCK" : inStock += 1
                    Case "SOLD" : sold += 1
                    Case "IN_REPAIR" : inRepair += 1
                End Select
            Next
        End If

        If dgvEstoque.Columns.Contains("_SEL") Then
            For Each gridRow As DataGridViewRow In dgvEstoque.Rows
                Dim cell As DataGridViewCheckBoxCell = TryCast(gridRow.Cells("_SEL"), DataGridViewCheckBoxCell)
                If cell IsNot Nothing AndAlso cell.Value IsNot Nothing AndAlso CBool(cell.Value) Then
                    selected += 1
                End If
            Next
        End If

        Dim totalDB As Integer = If(dtEstoqueCompleto IsNot Nothing, dtEstoqueCompleto.Rows.Count, 0)

        Dim modoTexto As String
        If _listaRelatorio IsNot Nothing AndAlso _listaRelatorio.EstaAtiva Then
            modoTexto = "  |  Mode: Custom list (" & _listaRelatorio.Count.ToString() & " items)"
        Else
            modoTexto = "  |  Mode: Current filter (" & showing.ToString() & " items)"
        End If

        lblEstoqueFooter.Text =
            "Showing: " & showing.ToString() & "  |  " &
            "IN_STOCK: " & inStock.ToString() & "  |  " &
            "SOLD: " & sold.ToString() & "  |  " &
            "IN_REPAIR: " & inRepair.ToString() & "  |  " &
            "Selected: " & selected.ToString() & "  |  " &
            "Total in DB: " & totalDB.ToString() &
            modoTexto

    End Sub

    Private Sub ExportarExcel()

        Dim itens As List(Of DataRow) = ColetarItensRelatorio()
        If itens.Count = 0 Then Return

        Try
            Dim outputPath As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
            Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"

            If Not System.IO.Directory.Exists(outputPath) Then
                System.IO.Directory.CreateDirectory(outputPath)
            End If

            Dim caminho As String = ReportService.GerarExcel(itens, outputPath, logoPath)

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
        chk.Name = "_SEL"
        chk.HeaderText = ""
        chk.Width = 30
        chk.ReadOnly = False
        chk.FillWeight = 1
        dgvEstoque.Columns.Insert(0, chk)

    End Sub

    Private Sub ConfigurarColunasGrid()
        ' Hide redundant STATUS alias
        If dgvEstoque.Columns.Contains("STATUS_DESCRICAO") Then
            dgvEstoque.Columns("STATUS_DESCRICAO").Visible = False
        End If

        Dim headers As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"ID_EQUIPAMENTO", "ID"},
            {"INTERNAL_UID", "Internal UID"},
            {"SERIAL_NUMBER", "Serial Number"},
            {"MARCA", "Manufacturer"},
            {"MANUFACTURER", "Manufacturer"},
            {"MODEL", "Model"},
            {"MODELO", "Model"},
            {"PROCESSADOR", "Processor"},
            {"CPU_MODEL", "Processor"},
            {"CPU_FAMILY", "CPU Family"},
            {"RAM_GB", "RAM (GB)"},
            {"STORAGE_GB", "Storage (GB)"},
            {"CONDITION_STATUS", "Battery Condition"},
            {"STATUS", "Status"},
            {"OBSERVACAO", "Notes"},
            {"NOTES", "Notes"},
            {"DATA_CADASTRO", "Created"},
            {"DATA_ATUALIZACAO", "Updated"},
            {"SOURCE_BATCH", "Source Batch"}
        }

        For Each col As DataGridViewColumn In dgvEstoque.Columns
            If headers.ContainsKey(col.Name) Then
                col.HeaderText = headers(col.Name)
            End If
        Next

        ' Configure CONDITION_STATUS column
        If dgvEstoque.Columns.Contains("CONDITION_STATUS") Then
            With dgvEstoque.Columns("CONDITION_STATUS")
                .HeaderText = "Battery Condition"
                .DataPropertyName = "CONDITION_STATUS"
                .Visible = True
                .MinimumWidth = 120
                .FillWeight = 70
            End With
            If dgvEstoque.Columns.Contains("STORAGE_GB") Then
                Dim idx As Integer = dgvEstoque.Columns("STORAGE_GB").DisplayIndex
                dgvEstoque.Columns("CONDITION_STATUS").DisplayIndex = idx + 1
            End If
        End If
    End Sub

    Private Function ColetarItensRelatorio() As List(Of DataRow)

        ' Prioridade 1: lista personalizada em memória
        If _listaRelatorio.EstaAtiva Then
            Return _listaRelatorio.Itens
        End If

        ' Prioridade 2: linhas marcadas no grid
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

        ' Prioridade 3: todos os itens visíveis no grid
        Dim todos As New List(Of DataRow)
        For Each gridRow As DataGridViewRow In dgvEstoque.Rows
            Dim drv As DataRowView = TryCast(gridRow.DataBoundItem, DataRowView)
            If drv IsNot Nothing Then todos.Add(drv.Row)
        Next
        Return todos

    End Function

#Region "Lista de Relatório"

    Private Sub AdicionarSelecionadosNaLista()

        Dim adicionados As Integer = 0
        Dim duplicados As Integer = 0

        If dgvEstoque.Columns.Contains("_SEL") Then
            For Each gridRow As DataGridViewRow In dgvEstoque.Rows
                Dim cell As DataGridViewCheckBoxCell = TryCast(gridRow.Cells("_SEL"), DataGridViewCheckBoxCell)
                If cell IsNot Nothing AndAlso cell.Value IsNot Nothing AndAlso CBool(cell.Value) Then
                    Dim drv As DataRowView = TryCast(gridRow.DataBoundItem, DataRowView)
                    If drv IsNot Nothing Then
                        If _listaRelatorio.Adicionar(drv.Row) Then
                            adicionados += 1
                        Else
                            duplicados += 1
                        End If
                    End If
                End If
            Next
        End If

        ' Se nenhuma marcada, adiciona a linha focada
        If adicionados = 0 AndAlso duplicados = 0 AndAlso
           dgvEstoque.CurrentRow IsNot Nothing Then
            Dim drv As DataRowView = TryCast(dgvEstoque.CurrentRow.DataBoundItem, DataRowView)
            If drv IsNot Nothing Then
                If _listaRelatorio.Adicionar(drv.Row) Then
                    adicionados = 1
                Else
                    duplicados = 1
                End If
            End If
        End If

        AtualizarModoRelatorio()

        Dim msg As String = adicionados.ToString() & " item(s) added to the list"
        If duplicados > 0 Then msg &= " (" & duplicados.ToString() & " duplicate(s) ignored)"
        lblStatus.Text = msg

    End Sub

    Private Sub LimparListaRelatorio()

        If Not _listaRelatorio.EstaAtiva Then
            lblStatus.Text = "Report list is already empty."
            Return
        End If

        Dim res As DialogResult = MessageBox.Show(
            "Clear the report list with " & _listaRelatorio.Count.ToString() & " item(s)?",
            "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If res = DialogResult.Yes Then
            _listaRelatorio.Limpar()
            AtualizarModoRelatorio()
            lblStatus.Text = "Report list cleared."
        End If

    End Sub

    Private Sub AtualizarModoRelatorio()
        If btnClearList IsNot Nothing Then
            btnClearList.Enabled = _listaRelatorio.EstaAtiva
        End If
        AtualizarRodapeEstoque()
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If tabPrincipal.SelectedTab Is tabEstoque Then
            Select Case keyData
                Case Keys.Control Or Keys.L
                    AdicionarSelecionadosNaLista()
                    Return True
                Case Keys.Control Or Keys.Shift Or Keys.L
                    LimparListaRelatorio()
                    Return True
                Case Keys.F9
                    GerarRelatorio()
                    Return True
            End Select
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

#End Region

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

            Dim baseNome As String = "Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")

            Dim caminhoDoc As String = ReportService.GerarRelatorio(itens, outputPath, logoPath, baseNome)
            Dim caminhoPdf As String = ReportService.GerarPdf(itens, outputPath, logoPath, baseNome)
            Dim caminhoXlsx As String = ReportService.GerarExcel(itens, outputPath, logoPath, baseNome)

            Dim frmEmail As New frmEnviarRelatorio(caminhoDoc, caminhoPdf, caminhoXlsx, itens)
            frmEmail.ShowDialog(Me)

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

    Private Sub clbManufacturer_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbManufacturer.ItemCheck
        Me.BeginInvoke(Sub()
                           recarregarModelosPorManufacturer()
                           AtualizarContadorFiltros()
                       End Sub)
    End Sub

    Private Sub clbModel_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbModel.ItemCheck
        Me.BeginInvoke(Sub()
                           recarregarProcessoresPorModelo()
                           AtualizarContadorFiltros()
                       End Sub)
    End Sub

    Private Sub clbStatus_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbStatus.ItemCheck
        Me.BeginInvoke(New Action(AddressOf AtualizarContadorFiltros))
    End Sub

    Private Sub clbBatch_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbBatch.ItemCheck
        Me.BeginInvoke(New Action(AddressOf AtualizarContadorFiltros))
    End Sub

    Private Sub clbProcessor_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbProcessor.ItemCheck
        Me.BeginInvoke(New Action(AddressOf AtualizarContadorFiltros))
    End Sub

    Private Sub AtualizarContadorFiltros()
        If lblFiltrosAtivos Is Nothing Then Return
        Dim count As Integer = clbManufacturer.CheckedItems.Count +
                               clbModel.CheckedItems.Count +
                               clbStatus.CheckedItems.Count +
                               clbProcessor.CheckedItems.Count +
                               clbBatch.CheckedItems.Count
        lblFiltrosAtivos.Text = If(count = 0, "No filters active",
                                   If(count = 1, "1 filter active",
                                      count.ToString() & " filters active"))
        lblFiltrosAtivos.ForeColor = If(count = 0, TemaEscuro.TextoMutado, TemaEscuro.Accent)
    End Sub

    Private Sub btnApplyFilter_Click(sender As Object, e As EventArgs) Handles btnApplyFilter.Click
        aplicarFiltrosEstoque()
    End Sub

    Private Sub btnClearFilter_Click(sender As Object, e As EventArgs) Handles btnClearFilter.Click
        txtPesquisa.Text = ""
        For i As Integer = 0 To clbManufacturer.Items.Count - 1
            clbManufacturer.SetItemChecked(i, False)
        Next
        For i As Integer = 0 To clbModel.Items.Count - 1
            clbModel.SetItemChecked(i, False)
        Next
        For i As Integer = 0 To clbStatus.Items.Count - 1
            clbStatus.SetItemChecked(i, False)
        Next
        For i As Integer = 0 To clbProcessor.Items.Count - 1
            clbProcessor.SetItemChecked(i, False)
        Next
        For i As Integer = 0 To clbBatch.Items.Count - 1
            clbBatch.SetItemChecked(i, False)
        Next
        aplicarFiltrosEstoque()
    End Sub

    Private Sub btnGenerateReport_Click(sender As Object, e As EventArgs) Handles btnGenerateReport.Click
        GerarRelatorio()
    End Sub

    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click
        ExportarExcel()
    End Sub

    Private Sub btnAddToList_Click(sender As Object, e As EventArgs) Handles btnAddToList.Click
        AdicionarSelecionadosNaLista()
    End Sub

    Private Sub btnClearList_Click(sender As Object, e As EventArgs) Handles btnClearList.Click
        LimparListaRelatorio()
    End Sub

    Private Sub btnAddEquipamento_Click(sender As Object, e As EventArgs) Handles btnAddEquipamento.Click

        Try
            Dim frm As New frmAddEquipamento()
            If frm.ShowDialog(Me) = DialogResult.OK Then
                carregarEstoque()
                carregarDashboard()
                MessageBox.Show("Equipment added successfully!" & vbCrLf & "UID: " & frm.SavedUID,
                                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error opening Add Equipment form: " & ex.Message, "Error",
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

    Private Sub dgvEstoque_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvEstoque.ColumnHeaderMouseClick
        If e.ColumnIndex < 0 Then Return
        If dgvEstoque.Columns(e.ColumnIndex).Name <> "_SEL" Then Return

        _todasSelecionadas = Not _todasSelecionadas
        For Each row As DataGridViewRow In dgvEstoque.Rows
            row.Cells("_SEL").Value = _todasSelecionadas
        Next
        dgvEstoque.EndEdit()
        AtualizarRodapeEstoque()
    End Sub

    Private Sub dgvEstoque_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvEstoque.CurrentCellDirtyStateChanged
        If dgvEstoque.IsCurrentCellDirty AndAlso dgvEstoque.CurrentCell IsNot Nothing AndAlso
           dgvEstoque.Columns(dgvEstoque.CurrentCell.ColumnIndex).Name = "_SEL" Then
            dgvEstoque.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub dgvEstoque_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvEstoque.CellValueChanged
        If e.ColumnIndex >= 0 AndAlso e.ColumnIndex < dgvEstoque.Columns.Count AndAlso
           dgvEstoque.Columns(e.ColumnIndex).Name = "_SEL" Then
            AtualizarRodapeEstoque()
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
                dtUpgradesCompleto = ds.Tables(0)
                PopularCboUpgBatch()
                aplicarFiltrosUpgrade()
            End If

        Catch ex As Exception

            lblUpgradesTit.Text = "Upgrades (load error)"

        End Try

    End Sub

    Private Sub PopularCboUpgBatch()
        Dim selAtual As String = If(cboUpgBatch.SelectedIndex > 0, cboUpgBatch.SelectedItem.ToString(), "")
        cboUpgBatch.Items.Clear()
        cboUpgBatch.Items.Add("(All)")
        If dtUpgradesCompleto IsNot Nothing AndAlso dtUpgradesCompleto.Columns.Contains("SOURCE_BATCH") Then
            Dim valores As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each row As DataRow In dtUpgradesCompleto.Rows
                If Not IsDBNull(row("SOURCE_BATCH")) Then
                    Dim v As String = row("SOURCE_BATCH").ToString().Trim()
                    If Not String.IsNullOrEmpty(v) Then valores.Add(v)
                End If
            Next
            For Each v As String In valores
                cboUpgBatch.Items.Add(v)
            Next
        End If
        Dim idx As Integer = cboUpgBatch.Items.IndexOf(selAtual)
        cboUpgBatch.SelectedIndex = If(idx > 0, idx, 0)
    End Sub

    Private Sub aplicarFiltrosUpgrade()
        If dtUpgradesCompleto Is Nothing Then Return
        Dim dv As DataView = dtUpgradesCompleto.DefaultView
        Dim batch As String = If(cboUpgBatch.SelectedIndex > 0, cboUpgBatch.SelectedItem.ToString(), "")
        If Not String.IsNullOrEmpty(batch) AndAlso dtUpgradesCompleto.Columns.Contains("SOURCE_BATCH") Then
            dv.RowFilter = "SOURCE_BATCH = '" & batch.Replace("'", "''") & "'"
        Else
            dv.RowFilter = ""
        End If
        dgvUpgrades.DataSource = dv
        lblUpgradesTotal.Text = "Total: " & dv.Count.ToString()
    End Sub

    Private Sub cboUpgBatch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboUpgBatch.SelectedIndexChanged
        aplicarFiltrosUpgrade()
    End Sub

    Private Sub btnAtualizarUpgrades_Click(sender As Object, e As EventArgs) Handles btnAtualizarUpgrades.Click
        carregarUpgrades()
    End Sub

    Private Sub btnGerenciarUpgrade_Click(sender As Object, e As EventArgs) Handles btnGerenciarUpgrade.Click
        AbrirUpgradeDoEstoqueSelecionado()
    End Sub

    Private Sub btnUpgradeAddToList_Click(sender As Object, e As EventArgs) Handles btnUpgradeAddToList.Click
        AdicionarUpgradeNaLista()
    End Sub

    Private Sub btnUpgradeClearList_Click(sender As Object, e As EventArgs) Handles btnUpgradeClearList.Click
        LimparListaUpgrade()
    End Sub

    Private Sub AdicionarUpgradeNaLista()

        If dgvUpgrades Is Nothing OrElse dgvUpgrades.CurrentRow Is Nothing Then Return

        Dim drv As DataRowView = TryCast(dgvUpgrades.CurrentRow.DataBoundItem, DataRowView)
        If drv Is Nothing Then Return

        If _listaUpgradeRelatorio.Adicionar(drv.Row) Then
            AtualizarModoRelatorioUpgrade()
            lblStatus.Text = "1 upgrade added to list (" & _listaUpgradeRelatorio.Count.ToString() & " total)"
        Else
            lblStatus.Text = "This upgrade is already in the list."
        End If

    End Sub

    Private Sub LimparListaUpgrade()

        If Not _listaUpgradeRelatorio.EstaAtiva Then
            lblStatus.Text = "Upgrade list is already empty."
            Return
        End If

        Dim res As DialogResult = MessageBox.Show(
            "Clear the upgrade list with " & _listaUpgradeRelatorio.Count.ToString() & " item(s)?",
            "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If res = DialogResult.Yes Then
            _listaUpgradeRelatorio.Limpar()
            AtualizarModoRelatorioUpgrade()
            lblStatus.Text = "Upgrade list cleared."
        End If

    End Sub

    Private Sub AtualizarModoRelatorioUpgrade()
        If btnUpgradeClearList IsNot Nothing Then
            btnUpgradeClearList.Enabled = _listaUpgradeRelatorio.EstaAtiva
        End If
        If lblUpgradesListaInfo IsNot Nothing Then
            lblUpgradesListaInfo.Text = If(_listaUpgradeRelatorio.EstaAtiva,
                "List mode: " & _listaUpgradeRelatorio.Count.ToString() & " upgrade(s) selected",
                "")
        End If
    End Sub

    Private Sub btnUpgradeRelatorio_Click(sender As Object, e As EventArgs) Handles btnUpgradeRelatorio.Click
        GerarRelatorioUpgradesCliente()
    End Sub

    Private Sub GerarRelatorioUpgradesCliente()

        Try
            Dim ds As DataSet = oUpgradeController.buscarComCliente(90)

            If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                MessageBox.Show("No upgrades found.", "Upgrade Report",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim rows As New List(Of DataRow)(ds.Tables(0).Rows.Cast(Of DataRow)())

            If _listaUpgradeRelatorio.EstaAtiva Then
                Dim idsNaLista As New HashSet(Of String)(
                    _listaUpgradeRelatorio.Itens _
                    .Where(Function(r) r.Table.Columns.Contains("ID_UPGRADE") AndAlso Not IsDBNull(r("ID_UPGRADE"))) _
                    .Select(Function(r) r("ID_UPGRADE").ToString()))
                rows = rows.Where(Function(r)
                    Return r.Table.Columns.Contains("ID_UPGRADE") AndAlso
                           Not IsDBNull(r("ID_UPGRADE")) AndAlso
                           idsNaLista.Contains(r("ID_UPGRADE").ToString())
                End Function).ToList()
                If rows.Count = 0 Then
                    MessageBox.Show("None of the selected upgrades were found in the last 90 days.", "Upgrade Report",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If
            End If

            Dim outputPath As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"
            If Not IO.Directory.Exists(outputPath) Then IO.Directory.CreateDirectory(outputPath)

            Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
            Dim baseNome As String = "UpgradeReport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim caminhoXlsx As String = ReportService.GerarExcelUpgradesPorCliente(rows, outputPath, logoPath, baseNome)

            ' Abre o Excel para revisão
            System.Diagnostics.Process.Start(caminhoXlsx)

            ' Pergunta se quer enviar por email
            Dim resp As DialogResult = MessageBox.Show(
                "Excel generated successfully." & vbCrLf & vbCrLf &
                "Do you want to send it by email?",
                "Upgrade Report", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If resp = DialogResult.Yes Then
                Dim frmEmail As New frmEnviarRelatorio("", "", caminhoXlsx, rows)
                frmEmail.txtSubject.Text = "Hardware Upgrade Report — " & DateTime.Now.ToString("yyyy-MM-dd")
                frmEmail.txtBody.Text =
                    "Please find attached the hardware upgrade report." & vbCrLf & vbCrLf &
                    "Best regards," & vbCrLf & "GBS — Global Business Solution"
                frmEmail.chkAttachWord.Checked = False
                frmEmail.chkAttachPdf.Checked  = False
                frmEmail.ShowDialog(Me)
            End If

        Catch ex As Exception
            MessageBox.Show("Error generating upgrade report: " & ex.Message, "Upgrade Report",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

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
        Dim id As Integer = 0
        Dim uid As String = ""

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
        OcultarAnaliseImportacao()

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

            AtualizarSourceBatchImportado()
            carregarDashboard()
            carregarEstoque()
            CarregarAnaliseImportacao()

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

    Private Sub OcultarAnaliseImportacao()
        dtImportQuality = Nothing
        If lblImpAnaliseTitle IsNot Nothing Then lblImpAnaliseTitle.Visible = False
        If chkImpShowIssuesOnly IsNot Nothing Then
            chkImpShowIssuesOnly.Checked = False
            chkImpShowIssuesOnly.Visible = False
        End If
        If btnImpGerarRelatorio IsNot Nothing Then btnImpGerarRelatorio.Visible = False
        If btnImpEnviarRelatorio IsNot Nothing Then btnImpEnviarRelatorio.Visible = False
        If dgvImpAnalise IsNot Nothing Then
            dgvImpAnalise.DataSource = Nothing
            dgvImpAnalise.Visible = False
        End If
        If lblImpAnaliseFooter IsNot Nothing Then
            lblImpAnaliseFooter.Text = ""
            lblImpAnaliseFooter.Visible = False
        End If
    End Sub

    Private Sub AtualizarSourceBatchImportado()
        If oImportController Is Nothing OrElse oImportController.UIDsImportados.Count = 0 Then Return

        Dim batchName As String = ObterImportBatchName()
        If String.IsNullOrWhiteSpace(batchName) Then Return

        Try
            Dim cs As String = System.Configuration.ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim pBatch As New OracleParameter("P_BATCH", OracleDbType.Varchar2, ParameterDirection.Input)
            pBatch.Value = batchName

            OracleHelper.ExecuteNonQuery(cs, CommandType.Text,
                "UPDATE TBL_EQUIPAMENTO" &
                "   SET SOURCE_BATCH = :P_BATCH, DATA_ATUALIZACAO = SYSDATE" &
                " WHERE " & BuildImportInClause(oImportController.UIDsImportados),
                pBatch)
        Catch
            ' Source batch is helpful for reports, but import should not fail if this update is blocked.
        End Try
    End Sub

    Private Sub CarregarAnaliseImportacao()
        If oImportController Is Nothing OrElse oImportController.UIDsImportados.Count = 0 Then Return

        Try
            Dim cs As String = System.Configuration.ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim filtroUids As String = BuildImportInClause(oImportController.UIDsImportados)

            Dim sqlFull As String =
                "SELECT ID_EQUIPAMENTO, INTERNAL_UID, SERIAL_NUMBER, MARCA, MODEL, PROCESSADOR," &
                "       RAM_GB, STORAGE_GB, CONDITION_STATUS, STATUS, SOURCE_BATCH," &
                "       OBSERVACAO, DATA_CADASTRO, DATA_ATUALIZACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & filtroUids &
                " ORDER BY INTERNAL_UID"

            Dim sqlFallback As String =
                "SELECT ID_EQUIPAMENTO, INTERNAL_UID, SERIAL_NUMBER, MARCA, MODEL, PROCESSADOR," &
                "       RAM_GB, STORAGE_GB, CAST(NULL AS VARCHAR2(20)) AS CONDITION_STATUS," &
                "       STATUS, CAST(NULL AS VARCHAR2(100)) AS SOURCE_BATCH," &
                "       OBSERVACAO, DATA_CADASTRO, DATA_ATUALIZACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & filtroUids &
                " ORDER BY INTERNAL_UID"

            Dim dtSrc As DataTable = Nothing
            Dim lastEx As Exception = Nothing
            For Each sql As String In New String() {sqlFull, sqlFallback}
                Try
                    Dim ds As DataSet = OracleHelper.ExecuteDataset(cs, CommandType.Text, sql)
                    If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                        dtSrc = ds.Tables(0)
                        Exit For
                    End If
                Catch ex As Exception
                    lastEx = ex
                    If Not ex.Message.Contains("ORA-00904") Then Throw
                End Try
            Next
            If dtSrc Is Nothing Then
                If lastEx IsNot Nothing Then Throw lastEx
                Return
            End If

            dtImportQuality = New DataTable()
            dtImportQuality.Columns.Add("ID_EQUIPAMENTO", GetType(Integer))
            dtImportQuality.Columns.Add("INTERNAL_UID", GetType(String))
            dtImportQuality.Columns.Add("FINAL_OK", GetType(Boolean))
            dtImportQuality.Columns.Add("SERIAL_NUMBER", GetType(String))
            dtImportQuality.Columns.Add("MARCA", GetType(String))
            dtImportQuality.Columns.Add("MODEL", GetType(String))
            dtImportQuality.Columns.Add("PROCESSADOR", GetType(String))
            dtImportQuality.Columns.Add("RAM_GB", GetType(String))
            dtImportQuality.Columns.Add("RAM_UPGRADE_TO", GetType(String))
            dtImportQuality.Columns.Add("STORAGE_GB", GetType(String))
            dtImportQuality.Columns.Add("STORAGE_UPGRADE_TO", GetType(String))
            dtImportQuality.Columns.Add("CONDITION_STATUS", GetType(String))
            dtImportQuality.Columns.Add("BATTERY_SOURCE", GetType(String))
            dtImportQuality.Columns.Add("BATTERY_CHECK", GetType(String))
            dtImportQuality.Columns.Add("BATTERY_REPLACE", GetType(Boolean))
            dtImportQuality.Columns.Add("STATUS", GetType(String))
            dtImportQuality.Columns.Add("SOURCE_BATCH", GetType(String))
            dtImportQuality.Columns.Add("OBSERVATION", GetType(String))
            dtImportQuality.Columns.Add("DATA_CADASTRO", GetType(String))
            dtImportQuality.Columns.Add("DATA_ATUALIZACAO", GetType(String))
            dtImportQuality.Columns.Add("PROBLEM", GetType(Boolean))
            dtImportQuality.Columns.Add("ISSUE", GetType(String))
            dtImportQuality.Columns.Add("ISSUE_REASON", GetType(String))

            Dim batchName As String = ObterImportBatchName()

            For Each row As DataRow In dtSrc.Rows
                Dim uid As String = ImportColStr(row, "INTERNAL_UID")
                Dim cond As String = ImportColStr(row, "CONDITION_STATUS").ToUpperInvariant().Trim()
                Dim status As String = ImportColStr(row, "STATUS").ToUpperInvariant().Trim()
                Dim obs As String = ImportColStr(row, "OBSERVACAO").Trim()
                Dim sourceBatch As String = ImportColStr(row, "SOURCE_BATCH").Trim()
                If String.IsNullOrWhiteSpace(sourceBatch) Then sourceBatch = batchName
                Dim batteryFromSheet As String = cond
                If oImportController IsNot Nothing AndAlso
                   oImportController.BatteryCheckByUid IsNot Nothing AndAlso
                   oImportController.BatteryCheckByUid.ContainsKey(uid) Then
                    batteryFromSheet = oImportController.BatteryCheckByUid(uid)
                End If
                Dim motivo As String = MontarMotivoImportacao(cond, obs, status)
                Dim problem As Boolean = Not String.IsNullOrWhiteSpace(motivo)

                dtImportQuality.Rows.Add(
                    ImportInt(row, "ID_EQUIPAMENTO"),
                    uid,
                    Not problem,
                    ImportColStr(row, "SERIAL_NUMBER"),
                    ImportColStr(row, "MARCA"),
                    ImportColStr(row, "MODEL"),
                    ImportColStr(row, "PROCESSADOR"),
                    ImportColStr(row, "RAM_GB"),
                    "",
                    ImportColStr(row, "STORAGE_GB"),
                    "",
                    cond,
                    batteryFromSheet,
                    "",
                    False,
                    status,
                    sourceBatch,
                    obs,
                    ImportDateStr(row, "DATA_CADASTRO"),
                    ImportDateStr(row, "DATA_ATUALIZACAO"),
                    problem,
                    If(problem, "Yes", "No"),
                    motivo
                )
            Next

            AplicarFiltroAnaliseImportacao()
            lblImpAnaliseTitle.Visible = True
            chkImpShowIssuesOnly.Visible = True
            btnImpGerarRelatorio.Visible = True
            btnImpEnviarRelatorio.Visible = True
            dgvImpAnalise.Visible = True
            lblImpAnaliseFooter.Visible = True

        Catch ex As Exception
            MessageBox.Show("Error loading import quality review: " & ex.Message,
                            "Import Review", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub AplicarFiltroAnaliseImportacao()
        If dtImportQuality Is Nothing Then Return

        Dim view As New DataView(dtImportQuality)
        If chkImpShowIssuesOnly IsNot Nothing AndAlso chkImpShowIssuesOnly.Checked Then
            view.RowFilter = "FINAL_OK = False"
        End If

        dgvImpAnalise.DataSource = view
        ConfigurarGridAnaliseImportacao()
        AtualizarFooterAnaliseImportacao()
    End Sub

    Private Sub ConfigurarGridAnaliseImportacao()
        If dgvImpAnalise Is Nothing OrElse dgvImpAnalise.Columns.Count = 0 Then Return

        dgvImpAnalise.ReadOnly = False
        dgvImpAnalise.EditMode = DataGridViewEditMode.EditOnEnter
        dgvImpAnalise.AllowUserToAddRows = False
        dgvImpAnalise.AllowUserToDeleteRows = False
        dgvImpAnalise.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvImpAnalise.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvImpAnalise.ScrollBars = ScrollBars.Both

        For Each col As DataGridViewColumn In dgvImpAnalise.Columns
            col.ReadOnly = True
        Next

        If dgvImpAnalise.Columns.Contains("FINAL_OK") Then
            With dgvImpAnalise.Columns("FINAL_OK")
                .HeaderText = "Final OK?"
                .ReadOnly = False
                .DisplayIndex = 0
                .Width = 80
            End With
        End If
        If dgvImpAnalise.Columns.Contains("PROBLEM") Then
            dgvImpAnalise.Columns("PROBLEM").Visible = False
        End If
        If dgvImpAnalise.Columns.Contains("ISSUE") Then
            dgvImpAnalise.Columns("ISSUE").Visible = False
        End If

        DefinirColunaImport("ID_EQUIPAMENTO", "ID", 95, True, 1)
        DefinirColunaImport("INTERNAL_UID", "Internal UID", 105, True, 2)
        DefinirColunaImport("SERIAL_NUMBER", "Serial Number", 120, True, 3)
        DefinirColunaImport("MARCA", "Manufacturer", 95, True, 4)
        DefinirColunaImport("MODEL", "MODEL", 155, True, 5)
        DefinirColunaImport("PROCESSADOR", "Processor", 120, True, 6)
        DefinirColunaImport("RAM_GB", "RAM (GB)", 70, True, 7)
        DefinirColunaImport("RAM_UPGRADE_TO", "RAM Check", 110, False, 8)
        DefinirColunaImport("STORAGE_GB", "Storage (GB)", 90, True, 9)
        DefinirColunaImport("STORAGE_UPGRADE_TO", "Storage Check", 130, False, 10)

        If dgvImpAnalise.Columns.Contains("CONDITION_STATUS") Then
            dgvImpAnalise.Columns("CONDITION_STATUS").Visible = False
        End If
        DefinirColunaImport("BATTERY_SOURCE", "Battery From Sheet", 145, True, 11)
        DefinirColunaImport("BATTERY_CHECK", "Battery Check", 130, False, 12)
        DefinirColunaImport("BATTERY_REPLACE", "Battery Needed?", 120, False, 13)
        DefinirColunaImport("STATUS", "Status", 95, True, 14)
        DefinirColunaImport("SOURCE_BATCH", "Source Batch", 170, True, 15)

        If dgvImpAnalise.Columns.Contains("ISSUE_REASON") Then
            With dgvImpAnalise.Columns("ISSUE_REASON")
                .HeaderText = "ISSUE REASON"
                .ReadOnly = False
                .Width = 220
                .DisplayIndex = 16
            End With
        End If
        If dgvImpAnalise.Columns.Contains("OBSERVATION") Then
            With dgvImpAnalise.Columns("OBSERVATION")
                .ReadOnly = False
                .Width = 320
                .DisplayIndex = 17
            End With
        End If
        DefinirColunaImport("DATA_CADASTRO", "Created", 125, True, 18)
        DefinirColunaImport("DATA_ATUALIZACAO", "Updated", 130, True, 19)
    End Sub

    Private Sub DefinirColunaImport(nome As String,
                                    titulo As String,
                                    largura As Integer,
                                    pReadOnly As Boolean,
                                    displayIndex As Integer,
                                    Optional formato As String = "")
        If Not dgvImpAnalise.Columns.Contains(nome) Then Return

        With dgvImpAnalise.Columns(nome)
            .HeaderText = titulo
            .ReadOnly = pReadOnly
            .Width = largura
            If displayIndex >= 0 AndAlso displayIndex < dgvImpAnalise.Columns.Count Then
                .DisplayIndex = displayIndex
            End If
            If Not String.IsNullOrWhiteSpace(formato) Then
                .DefaultCellStyle.Format = formato
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
        End With
    End Sub

    Private Sub AtualizarFooterAnaliseImportacao()
        If dtImportQuality Is Nothing Then Return

        Dim total As Integer = dtImportQuality.Rows.Count
        Dim issues As Integer = ContarProblemasImportacao()
        Dim ok As Integer = total - issues
        lblImpAnaliseFooter.Text = $"Total: {total}  |  Final OK: {ok}  |  Final BAD / Issues: {issues}"
        lblImpAnaliseFooter.ForeColor = If(issues > 0, TemaEscuro.Vermelho, TemaEscuro.Verde)
    End Sub

    Private Sub chkImpShowIssuesOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkImpShowIssuesOnly.CheckedChanged
        AplicarFiltroAnaliseImportacao()
    End Sub

    Private Sub dgvImpAnalise_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvImpAnalise.CurrentCellDirtyStateChanged
        If dgvImpAnalise Is Nothing OrElse Not dgvImpAnalise.IsCurrentCellDirty Then Return
        If TypeOf dgvImpAnalise.CurrentCell Is DataGridViewCheckBoxCell Then
            dgvImpAnalise.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub dgvImpAnalise_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvImpAnalise.CellValueChanged
        If e.RowIndex < 0 OrElse dgvImpAnalise Is Nothing Then Return
        If e.ColumnIndex < 0 OrElse e.ColumnIndex >= dgvImpAnalise.Columns.Count Then Return

        Dim rowView As DataRowView = TryCast(dgvImpAnalise.Rows(e.RowIndex).DataBoundItem, DataRowView)
        If rowView Is Nothing Then Return

        Dim colName As String = dgvImpAnalise.Columns(e.ColumnIndex).Name
        If colName = "FINAL_OK" Then
            Dim finalOk As Boolean = False
            If Not IsDBNull(rowView.Row("FINAL_OK")) Then finalOk = CBool(rowView.Row("FINAL_OK"))
            Dim problem As Boolean = Not finalOk
            rowView.Row("PROBLEM") = problem
            rowView.Row("ISSUE") = If(problem, "Yes", "No")
            If problem AndAlso String.IsNullOrWhiteSpace(ImportRowStr(rowView.Row, "ISSUE_REASON")) Then
                rowView.Row("ISSUE_REASON") = "Manual review"
            End If
        End If

        AtualizarFooterAnaliseImportacao()
        dgvImpAnalise.InvalidateRow(e.RowIndex)
    End Sub

    Private Sub dgvImpAnalise_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvImpAnalise.DataError
        e.ThrowException = False
        MessageBox.Show("Invalid value for this column. Please check numbers and try again.",
                        "Import Review", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub dgvImpAnalise_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvImpAnalise.CellFormatting
        If e.RowIndex < 0 OrElse dgvImpAnalise Is Nothing Then Return
        Dim rowView As DataRowView = TryCast(dgvImpAnalise.Rows(e.RowIndex).DataBoundItem, DataRowView)
        If rowView Is Nothing OrElse Not rowView.Row.Table.Columns.Contains("FINAL_OK") Then Return

        Try
            If Not CBool(rowView.Row("FINAL_OK")) Then
                e.CellStyle.BackColor = Color.FromArgb(255, 204, 204)
                e.CellStyle.ForeColor = Color.FromArgb(140, 20, 20)
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 160, 160)
                e.CellStyle.SelectionForeColor = Color.FromArgb(80, 0, 0)
            End If
        Catch
        End Try
    End Sub

    Private Sub btnImpGerarRelatorio_Click(sender As Object, e As EventArgs) Handles btnImpGerarRelatorio.Click
        GerarRelatorioQualidadeImportacao(False)
    End Sub

    Private Sub btnImpEnviarRelatorio_Click(sender As Object, e As EventArgs) Handles btnImpEnviarRelatorio.Click
        GerarRelatorioQualidadeImportacao(True)
    End Sub

    Private Sub GerarRelatorioQualidadeImportacao(enviarEmail As Boolean)
        If dtImportQuality Is Nothing OrElse dtImportQuality.Rows.Count = 0 Then
            MessageBox.Show("Run an import first.", "Import Report",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            Dim outputPath As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"
            If Not Directory.Exists(outputPath) Then Directory.CreateDirectory(outputPath)

            Dim baseNome As String = "ImportQuality_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim caminhoXlsx As String = GerarExcelRelatorioQualidadeImportacao(outputPath, baseNome)

            If enviarEmail Then
                Dim itens As New List(Of DataRow)()
                For Each row As DataRow In dtImportQuality.Rows
                    itens.Add(row)
                Next

                Dim frmEmail As New frmEnviarRelatorio("", "", caminhoXlsx, itens)
                frmEmail.txtSubject.Text = "Import Quality Report - " & DateTime.Now.ToString("yyyy-MM-dd")
                frmEmail.txtBody.Text = MontarResumoEmailQualidadeImportacao()
                frmEmail.chkAttachWord.Checked = False
                frmEmail.chkAttachPdf.Checked = False
                frmEmail.chkAttachExcel.Checked = True
                frmEmail.ShowDialog(Me)
            Else
                MessageBox.Show("Import Excel report generated successfully!" & vbCrLf & vbCrLf & caminhoXlsx,
                                "Import Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error generating import report: " & ex.Message,
                            "Import Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GerarExcelRelatorioQualidadeImportacao(outputPath As String,
                                                            baseNome As String) As String
        Dim caminhoFinal As String = Path.Combine(outputPath, baseNome & ".xlsx")
        Dim total As Integer = dtImportQuality.Rows.Count
        Dim issues As Integer = ContarProblemasImportacao()
        Dim ok As Integer = total - issues
        Dim issueRate As Decimal = If(total = 0, 0D, Math.Round((issues * 100D) / total, 1))

        Dim colunas As (Key As String, Header As String, Width As Double)() = {
            ("FINAL_RESULT", "Final Result", 13),
            ("ID_EQUIPAMENTO", "ID", 10),
            ("INTERNAL_UID", "Internal UID", 14),
            ("SERIAL_NUMBER", "Serial Number", 18),
            ("MARCA", "Manufacturer", 16),
            ("MODEL", "Model", 22),
            ("PROCESSADOR", "Processor", 16),
            ("RAM_GB", "RAM GB", 10),
            ("RAM_UPGRADE_TO", "RAM Check", 14),
            ("STORAGE_GB", "Storage GB", 12),
            ("STORAGE_UPGRADE_TO", "Storage Check", 16),
            ("BATTERY_SOURCE", "Battery From Sheet", 18),
            ("BATTERY_CHECK", "Battery Check", 16),
            ("BATTERY_REPLACE", "Battery Needed", 15),
            ("STATUS", "Status", 13),
            ("SOURCE_BATCH", "Batch", 22),
            ("ISSUE_REASON", "Issue Reason", 26),
            ("OBSERVATION", "Observation", 38),
            ("DATA_CADASTRO", "Created", 18),
            ("DATA_ATUALIZACAO", "Updated", 18)
        }

        Using pkg As New ExcelPackage()
            Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets.Add("Import Quality")
            Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
            Dim headerRow As Integer = 6

            ws.Row(1).Height = 40
            ws.Row(2).Height = 20
            ws.Row(3).Height = 20
            ws.Row(4).Height = 18
            ws.Row(5).Height = 8

            If Not String.IsNullOrWhiteSpace(logoPath) AndAlso File.Exists(logoPath) Then
                Try
                    Dim logo = ws.Drawings.AddPicture("GBS_Logo", New FileInfo(logoPath))
                    logo.SetPosition(0, 4, 0, 4)
                    logo.SetSize(180, 60)
                Catch
                    ' Keep the report usable even if the logo cannot be loaded.
                End Try
            End If

            ws.Cells(1, 4, 1, 10).Merge = True
            ws.Cells(1, 4).Value = "GBS Import Quality Report"
            ws.Cells(1, 4).Style.Font.Bold = True
            ws.Cells(1, 4).Style.Font.Size = 16
            ws.Cells(1, 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center

            ws.Cells(2, 4, 2, 12).Merge = True
            ws.Cells(2, 4).Value = "Source file: " & Path.GetFileName(sArquivoSelImp) &
                                   " | Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            ws.Cells(2, 4).Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(107, 114, 128))

            ws.Cells(3, 4, 3, 12).Merge = True
            ws.Cells(3, 4).Value = $"Total: {total} | Final OK: {ok} | Final BAD / Issues: {issues} | Issue rate: {issueRate:0.0}%"
            ws.Cells(3, 4).Style.Font.Bold = True

            For i As Integer = 0 To colunas.Length - 1
                Dim cell = ws.Cells(headerRow, i + 1)
                cell.Value = colunas(i).Header
                cell.Style.Font.Bold = True
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(55, 65, 81))
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                ws.Column(i + 1).Width = colunas(i).Width
            Next

            For rowIdx As Integer = 0 To dtImportQuality.Rows.Count - 1
                Dim row As DataRow = dtImportQuality.Rows(rowIdx)
                Dim excelRow As Integer = headerRow + rowIdx + 1
                Dim finalOk As Boolean = CBool(row("FINAL_OK"))

                For colIdx As Integer = 0 To colunas.Length - 1
                    Dim key As String = colunas(colIdx).Key
                    Dim value As Object = ""

                    If key = "FINAL_RESULT" Then
                        value = If(finalOk, "GOOD", "BAD")
                    ElseIf key = "BATTERY_REPLACE" Then
                        value = If(CBool(row("BATTERY_REPLACE")), "Yes", "No")
                    ElseIf row.Table.Columns.Contains(key) AndAlso Not IsDBNull(row(key)) Then
                        value = row(key).ToString()
                    End If

                    ws.Cells(excelRow, colIdx + 1).Value = value
                Next

                If Not finalOk Then
                    Dim rng = ws.Cells(excelRow, 1, excelRow, colunas.Length)
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(254, 226, 226))
                    rng.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(127, 29, 29))
                ElseIf rowIdx Mod 2 = 1 Then
                    Dim rng = ws.Cells(excelRow, 1, excelRow, colunas.Length)
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(249, 250, 251))
                End If
            Next

            Dim lastRow As Integer = headerRow + Math.Max(dtImportQuality.Rows.Count, 1)
            ws.Cells(headerRow, 1, lastRow, colunas.Length).AutoFilter = True
            ws.Cells(headerRow, 1, lastRow, colunas.Length).Style.VerticalAlignment = ExcelVerticalAlignment.Top
            ws.Cells(headerRow, 1, lastRow, colunas.Length).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
            ws.Cells(headerRow, 1, lastRow, colunas.Length).Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(229, 231, 235))
            ws.Column(18).Style.WrapText = True
            ws.View.FreezePanes(headerRow + 1, 1)

            pkg.SaveAs(New FileInfo(caminhoFinal))
        End Using

        Return caminhoFinal
    End Function

    Private Function MontarHtmlRelatorioQualidadeImportacao() As String
        Dim total As Integer = dtImportQuality.Rows.Count
        Dim issues As Integer = ContarProblemasImportacao()
        Dim ok As Integer = total - issues
        Dim issueRate As Decimal = If(total = 0, 0D, Math.Round((issues * 100D) / total, 1))

        Dim sb As New StringBuilder()
        sb.AppendLine("<html><head><meta charset='utf-8' />")
        sb.AppendLine("<style>")
        sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;font-size:10pt;color:#1f2937;margin:36px;}")
        sb.AppendLine("h2{margin:0 0 4px 0;color:#111827;}")
        sb.AppendLine(".sub{color:#6b7280;margin-bottom:18px;}")
        sb.AppendLine(".cards{display:flex;gap:10px;margin:16px 0 18px 0;}")
        sb.AppendLine(".card{border:1px solid #d1d5db;padding:10px 14px;min-width:130px;}")
        sb.AppendLine(".label{color:#6b7280;font-size:9pt;}.value{font-size:18pt;font-weight:700;}")
        sb.AppendLine("table{width:100%;border-collapse:collapse;table-layout:fixed;}")
        sb.AppendLine("th{background:#374151;color:#fff;text-align:left;padding:7px 6px;font-size:8pt;}")
        sb.AppendLine("td{border:1px solid #e5e7eb;padding:5px 6px;font-size:8pt;word-wrap:break-word;vertical-align:top;}")
        sb.AppendLine(".bad td{background:#fee2e2;color:#7f1d1d;}")
        sb.AppendLine("</style></head><body>")
        AppendImportLogoHtml(sb)
        sb.AppendLine("<h2>GBS Import Quality Report</h2>")
        sb.AppendLine("<div class='sub'>Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                      " | Source file: " & WebUtility.HtmlEncode(Path.GetFileName(sArquivoSelImp)) & "</div>")
        sb.AppendLine("<div class='cards'>")
        sb.AppendLine("<div class='card'><div class='label'>Total imported</div><div class='value'>" & total & "</div></div>")
        sb.AppendLine("<div class='card'><div class='label'>Final OK</div><div class='value'>" & ok & "</div></div>")
        sb.AppendLine("<div class='card'><div class='label'>Final BAD / Issues</div><div class='value'>" & issues & "</div></div>")
        sb.AppendLine("<div class='card'><div class='label'>Issue rate</div><div class='value'>" & issueRate.ToString("0.0") & "%</div></div>")
        sb.AppendLine("</div>")
        sb.AppendLine("<table>")
        sb.AppendLine("<tr><th>Final Result</th><th>Internal UID</th><th>Serial Number</th><th>Manufacturer</th><th>Model</th><th>Processor</th><th>RAM</th><th>RAM Check</th><th>Storage</th><th>Storage Check</th><th>Battery Check</th><th>Battery Needed</th><th>Status</th><th>Batch</th><th>Issue Reason</th><th>Observation</th></tr>")

        For Each row As DataRow In dtImportQuality.Rows
            Dim finalOk As Boolean = CBool(row("FINAL_OK"))
            Dim problem As Boolean = Not finalOk
            sb.Append(If(problem, "<tr class='bad'>", "<tr>"))
            AppendImportTd(sb, If(finalOk, "GOOD", "BAD"))
            AppendImportTd(sb, ImportRowStr(row, "INTERNAL_UID"))
            AppendImportTd(sb, ImportRowStr(row, "SERIAL_NUMBER"))
            AppendImportTd(sb, ImportRowStr(row, "MARCA"))
            AppendImportTd(sb, ImportRowStr(row, "MODEL"))
            AppendImportTd(sb, ImportRowStr(row, "PROCESSADOR"))
            AppendImportTd(sb, ImportRowStr(row, "RAM_GB"))
            AppendImportTd(sb, ImportRowStr(row, "RAM_UPGRADE_TO"))
            AppendImportTd(sb, ImportRowStr(row, "STORAGE_GB"))
            AppendImportTd(sb, ImportRowStr(row, "STORAGE_UPGRADE_TO"))
            AppendImportTd(sb, ImportRowStr(row, "BATTERY_CHECK"))
            AppendImportTd(sb, If(CBool(row("BATTERY_REPLACE")), "Yes", "No"))
            AppendImportTd(sb, ImportRowStr(row, "STATUS"))
            AppendImportTd(sb, ImportRowStr(row, "SOURCE_BATCH"))
            AppendImportTd(sb, ImportRowStr(row, "ISSUE_REASON"))
            AppendImportTd(sb, ImportRowStr(row, "OBSERVATION"))
            sb.AppendLine("</tr>")
        Next

        sb.AppendLine("</table>")
        sb.AppendLine("</body></html>")
        Return sb.ToString()
    End Function

    Private Function MontarResumoEmailQualidadeImportacao() As String
        Dim total As Integer = dtImportQuality.Rows.Count
        Dim issues As Integer = ContarProblemasImportacao()
        Dim ok As Integer = total - issues
        Dim issueRate As Decimal = If(total = 0, 0D, Math.Round((issues * 100D) / total, 1))

        Dim sb As New StringBuilder()
        sb.AppendLine("Please find attached the import quality report for the latest customer spreadsheet.")
        sb.AppendLine()
        sb.AppendLine("Summary:")
        sb.AppendLine("  - Source file: " & Path.GetFileName(sArquivoSelImp))
        sb.AppendLine("  - Total imported: " & total.ToString())
        sb.AppendLine("  - Final OK: " & ok.ToString())
        sb.AppendLine("  - Final BAD / Issues: " & issues.ToString())
        sb.AppendLine("  - Issue rate: " & issueRate.ToString("0.0") & "%")
        Return sb.ToString()
    End Function

    Private Function ContarProblemasImportacao() As Integer
        Dim total As Integer = 0
        If dtImportQuality Is Nothing Then Return 0
        For Each row As DataRow In dtImportQuality.Rows
            If Not CBool(row("FINAL_OK")) Then total += 1
        Next
        Return total
    End Function

    Private Function MontarMotivoImportacao(conditionStatus As String, observacao As String, status As String) As String
        Dim motivos As New List(Of String)()
        If conditionStatus = "FAIR" OrElse conditionStatus = "POOR" Then
            motivos.Add("Battery condition: " & conditionStatus)
        End If
        If Not String.IsNullOrWhiteSpace(observacao) Then
            motivos.Add("Observation")
        End If
        If status = "IN_REPAIR" Then
            motivos.Add("Status: IN_REPAIR")
        End If
        Return String.Join("; ", motivos)
    End Function

    Private Sub AppendImportLogoHtml(sb As StringBuilder)
        Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
        If String.IsNullOrWhiteSpace(logoPath) OrElse Not File.Exists(logoPath) Then Return

        Try
            Dim b64 As String = Convert.ToBase64String(File.ReadAllBytes(logoPath))
            Dim mime As String = If(Path.GetExtension(logoPath).ToLowerInvariant() = ".png", "image/png", "image/jpeg")
            sb.AppendLine("<div style='margin-bottom:14px;'><img src='data:" & mime & ";base64," &
                          b64 & "' style='max-height:60px;max-width:180px;' /></div>")
        Catch
        End Try
    End Sub

    Private Sub AppendImportTd(sb As StringBuilder, valor As String)
        sb.Append("<td>").Append(WebUtility.HtmlEncode(valor)).Append("</td>")
    End Sub

    Private Function ImportColStr(row As DataRow, colName As String) As String
        If Not row.Table.Columns.Contains(colName) Then Return ""
        If IsDBNull(row(colName)) Then Return ""
        Return row(colName).ToString()
    End Function

    Private Function ImportInt(row As DataRow, colName As String) As Integer
        If Not row.Table.Columns.Contains(colName) Then Return 0
        If IsDBNull(row(colName)) Then Return 0

        Dim id As Integer
        If Integer.TryParse(row(colName).ToString(), id) Then Return id
        Return 0
    End Function

    Private Function ImportDateStr(row As DataRow, colName As String) As String
        If Not row.Table.Columns.Contains(colName) Then Return ""
        If IsDBNull(row(colName)) Then Return ""

        Dim dt As DateTime
        If DateTime.TryParse(row(colName).ToString(), dt) Then
            Return dt.ToString("MM/dd/yyyy HH:mm")
        End If

        Return row(colName).ToString()
    End Function

    Private Function ImportRowStr(row As DataRow, colName As String) As String
        If Not row.Table.Columns.Contains(colName) Then Return ""
        If IsDBNull(row(colName)) Then Return ""
        Return row(colName).ToString()
    End Function

    Private Function ObterImportBatchName() As String
        Dim nome As String = ""

        If Not String.IsNullOrWhiteSpace(sArquivoSelImp) Then
            nome = Path.GetFileNameWithoutExtension(sArquivoSelImp)
        End If

        If String.IsNullOrWhiteSpace(nome) Then
            nome = "IMPORT_" & DateTime.Now.ToString("yyyyMMdd_HHmm")
        End If

        If nome.Length > 50 Then nome = nome.Substring(0, 50)
        Return nome
    End Function

    Private Function BuildImportInClause(uids As List(Of String)) As String
        If uids Is Nothing OrElse uids.Count = 0 Then Return "1=0"

        Dim sb As New StringBuilder()
        Dim i As Integer = 0
        While i < uids.Count
            If sb.Length > 0 Then sb.Append(" OR ")
            sb.Append("INTERNAL_UID IN (")

            Dim first As Boolean = True
            Dim limit As Integer = Math.Min(i + 999, uids.Count)
            While i < limit
                If Not first Then sb.Append(",")
                sb.Append("'").Append(uids(i).Replace("'", "''")).Append("'")
                first = False
                i += 1
            End While

            sb.Append(")")
        End While

        Return sb.ToString()
    End Function

#End Region


    Private Sub carregarDashboardResumoCards(pTabela As DataTable)

        dtResumoManufacturer = CriarResumoAgrupado(pTabela, "MANUFACTURER")
        dtResumoModel = CriarResumoAgrupado(pTabela, "MODEL")
        dtResumoCpuFamily = CriarResumoAgrupado(pTabela, "CPU_FAMILY")

        lblTotalUnidades.Text = dtResumoManufacturer.Rows.Count.ToString()
        lblEmEstoque.Text = dtResumoModel.Rows.Count.ToString()
        lblRemessasAtivas.Text = dtResumoCpuFamily.Rows.Count.ToString()

        dgvDashManufacturer.DataSource = dtResumoManufacturer
        dgvDashModel.DataSource = dtResumoModel
        dgvDashCPU.DataSource = dtResumoCpuFamily

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
