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
    Private dtResumoManufacturer As DataTable
    Private dtResumoModel As DataTable
    Private dtResumoCpuFamily As DataTable

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
                dgvEstoque.DataSource = ds.Tables(0)
            End If

        Catch ex As Exception

            MessageBox.Show("Error loading inventory: " & ex.Message, "Error",
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

            MessageBox.Show("Search error: " & ex.Message, "Error",
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
        AbrirUpgradeDoEstoqueSelecionado()
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

End Class
