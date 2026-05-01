Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.Diagnostics
Imports System.Globalization
Imports System.Media
Imports GBS_Inventory.Models

''' <summary>
''' Formulário de gerenciamento de Remessas (FedEx, UPS, USPS).
''' Lista remessas ativas e permite criar novas, atualizar status e abrir URL de rastreamento.
''' </summary>
Partial Public Class frmRemessa

    Private oController As RemessaController
    Private vIdRemessaSelecionada As Integer = 0
    Private vItensParaIncluir As New List(Of ItemRemessa)()
    Private _timerScanFeedback As Timer

    Public Sub New()

        InitializeComponent()

        oController = New RemessaController()

        TemaEscuro.aplicarHelius(Me)

        _timerScanFeedback = New Timer() With {.Interval = 2000}
        AddHandler _timerScanFeedback.Tick, AddressOf TimerScanFeedback_Tick

    End Sub

    Private Sub frmRemessa_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cboCarrier.Items.AddRange({"FEDEX", "UPS", "USPS", "DHL", "OTHER"})
        cboCarrier.SelectedIndex = 0

        cboDirection.Items.AddRange({"OUTBOUND", "INBOUND"})
        cboDirection.SelectedIndex = 0

        cboNovoStatus.Items.AddRange({"LABEL_CREATED", "PICKED_UP", "IN_TRANSIT",
                                       "OUT_FOR_DELIVERY", "DELIVERED", "DELAYED",
                                       "EXCEPTION", "RETURNED"})
        cboNovoStatus.SelectedIndex = 0

        configurarGridItens()
        carregarRemessas()
        atualizarGridItens()

        txtQuickScan.Focus()

    End Sub

    Private Sub carregarRemessas()

        Try

            Dim ds As DataSet = oController.buscarRemessasAtivas()

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvRemessas.DataSource = ds.Tables(0)
                lblTotal.Text = "Active total: " & ds.Tables(0).Rows.Count.ToString()
            End If

        Catch ex As Exception

            MessageBox.Show("Erro ao carregar remessas: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub dgvRemessas_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRemessas.CellClick

        If e.RowIndex < 0 Then Return

        Dim row As DataGridViewRow = dgvRemessas.Rows(e.RowIndex)
        vIdRemessaSelecionada = CInt(row.Cells("ID_REMESSA").Value)

        lblSelecionada.Text = $"Selected: {row.Cells("REMESSA_REF").Value} - Tracking: {row.Cells("TRACKING_NUMBER").Value}"

        btnAtualizarStatus.Enabled = True
        btnAbrirRastreio.Enabled   = True

    End Sub

    Private Sub btnNovaRemessa_Click(sender As Object, e As EventArgs) Handles btnNovaRemessa.Click

        If String.IsNullOrWhiteSpace(txtTracking.Text) OrElse String.IsNullOrWhiteSpace(txtRecipientName.Text) Then
            MessageBox.Show("Fill in Tracking Number and Recipient Name.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            Dim r As New Remessa() With {
                .Direction        = cboDirection.SelectedItem.ToString(),
                .Carrier          = cboCarrier.SelectedItem.ToString(),
                .TrackingNumber   = txtTracking.Text.Trim(),
                .RecipientName    = txtRecipientName.Text.Trim(),
                .RecipientAddress = txtRecipientAddress.Text.Trim(),
                .ServiceLevel     = txtServiceLevel.Text.Trim(),
                .Notes            = txtNotes.Text.Trim()
            }

            Dim vPeso As Decimal
            If Decimal.TryParse(txtPeso.Text.Replace(",", "."), Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, vPeso) Then
                r.WeightLbs = vPeso
            End If

            Dim vCusto As Decimal
            If Decimal.TryParse(txtCustoEnvio.Text.Replace(",", "."), Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, vCusto) Then
                r.ShippingCostUsd = vCusto
            End If

            Dim sRefGerada As String = ""

            Dim idGerado As Integer = oController.incluirRemessa(r, vItensParaIncluir, sRefGerada)

            MessageBox.Show($"Shipment created!{vbCrLf}Ref: {sRefGerada}{vbCrLf}ID: {idGerado}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            limparCampos()
            carregarRemessas()

        Catch ex As Exception

            MessageBox.Show("Error creating shipment: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnBuscarEquip_Click(sender As Object, e As EventArgs) Handles btnBuscarEquip.Click

        Try

            Dim ds As DataSet = oController.buscarEquipamentosDisponiveis(txtBuscarEquip.Text.Trim())

            If ds Is Nothing OrElse ds.Tables.Count = 0 Then
                dgvResultadoBusca.DataSource = Nothing
                Return
            End If

            dgvResultadoBusca.DataSource = ds.Tables(0)
            dgvResultadoBusca.ReadOnly = True
            dgvResultadoBusca.MultiSelect = False
            dgvResultadoBusca.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgvResultadoBusca.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            If dgvResultadoBusca.Columns.Contains("ID_EQUIPAMENTO") Then
                dgvResultadoBusca.Columns("ID_EQUIPAMENTO").Visible = False
            End If

        Catch ex As Exception

            MessageBox.Show("Error searching equipment: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnAdicionarItem_Click(sender As Object, e As EventArgs) Handles btnAdicionarItem.Click

        If dgvResultadoBusca.SelectedRows.Count = 0 Then
            MessageBox.Show("Select an equipment in the search results.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim row As DataGridViewRow = dgvResultadoBusca.SelectedRows(0)
        Dim idEquipamento As Integer = ObterInt(row.Cells("ID_EQUIPAMENTO").Value)

        If idEquipamento <= 0 Then
            MessageBox.Show("Invalid equipment.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If ItemJaSelecionado(idEquipamento) Then
            MessageBox.Show("This equipment has already been added.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim item As New ItemRemessa() With {
            .IdEquipamento = idEquipamento,
            .InternalUID = ObterTextoCelula(row, "INTERNAL_UID"),
            .Manufacturer = ObterTextoCelula(row, "MARCA"),
            .Model = ObterTextoCelula(row, "MODEL"),
            .ConditionAtShip = "GOOD",
            .SalePriceUsd = ParseDecimalNullable(txtPrecoItem.Text),
            .Notes = txtNotaItem.Text.Trim()
        }

        vItensParaIncluir.Add(item)
        atualizarGridItens()

        txtPrecoItem.Clear()
        txtNotaItem.Clear()

    End Sub

    Private Sub btnRemoverItem_Click(sender As Object, e As EventArgs) Handles btnRemoverItem.Click

        If dgvItensSelecionados.SelectedRows.Count = 0 Then
            MessageBox.Show("Select an item to remove.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim idx As Integer = dgvItensSelecionados.SelectedRows(0).Index

        If idx >= 0 AndAlso idx < vItensParaIncluir.Count Then
            vItensParaIncluir.RemoveAt(idx)
            atualizarGridItens()
        End If

    End Sub

    Private Sub btnAtualizarStatus_Click(sender As Object, e As EventArgs) Handles btnAtualizarStatus.Click

        If vIdRemessaSelecionada = 0 Then Return

        Try

            Dim sNovoStatus As String = cboNovoStatus.SelectedItem.ToString()
            Dim dataEntrega As Date? = Nothing

            If sNovoStatus = "DELIVERED" Then
                dataEntrega = Date.Today
            End If

            oController.atualizarStatus(vIdRemessaSelecionada, sNovoStatus, dataEntrega)

            MessageBox.Show("Status updated successfully!", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

            carregarRemessas()

        Catch ex As Exception

            MessageBox.Show("Error updating status: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnAbrirRastreio_Click(sender As Object, e As EventArgs) Handles btnAbrirRastreio.Click

        If vIdRemessaSelecionada = 0 Then Return

        Try

            Dim ds As DataSet = oController.buscarRemessa(vIdRemessaSelecionada)
            If ds.Tables(0).Rows.Count = 0 Then Return

            Dim row As DataRow = ds.Tables(0).Rows(0)

            Dim r As New Remessa() With {
                .Carrier        = row("CARRIER").ToString(),
                .TrackingNumber = row("TRACKING_NUMBER").ToString()
            }

            If String.IsNullOrEmpty(r.UrlRastreio) Then
                MessageBox.Show("Carrier not supported for automatic tracking.", "Warning",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Process.Start(r.UrlRastreio)

        Catch ex As Exception

            MessageBox.Show("Error opening tracking: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub limparCampos()
        txtTracking.Clear()
        txtRecipientName.Clear()
        txtRecipientAddress.Clear()
        txtServiceLevel.Clear()
        txtPeso.Clear()
        txtCustoEnvio.Clear()
        txtNotes.Clear()
        txtBuscarEquip.Clear()
        txtPrecoItem.Clear()
        txtNotaItem.Clear()
        dgvResultadoBusca.DataSource = Nothing
        vItensParaIncluir.Clear()
        atualizarGridItens()
    End Sub

    Private Sub txtQuickScan_KeyDown(sender As Object, e As KeyEventArgs) Handles txtQuickScan.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Dim scanValue As String = txtQuickScan.Text.Trim()
            If Not String.IsNullOrEmpty(scanValue) Then
                ExecutarQuickScan(scanValue)
            End If
        End If
    End Sub

    Private Sub ExecutarQuickScan(scanValue As String)
        Try
            Dim ds As DataSet = oController.buscarEquipamentosDisponiveis(scanValue)

            If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                SystemSounds.Exclamation.Play()
                MostrarFeedbackScan("Not found: " & scanValue, False)
                txtQuickScan.Clear()
                Return
            End If

            If ds.Tables(0).Rows.Count > 1 Then
                SystemSounds.Beep.Play()
                MostrarFeedbackScan("Multiple results — use manual search", Nothing)
                txtQuickScan.Clear()
                Return
            End If

            Dim row As DataRow = ds.Tables(0).Rows(0)
            Dim idEquipamento As Integer = 0
            If Not IsDBNull(row("ID_EQUIPAMENTO")) Then
                Integer.TryParse(row("ID_EQUIPAMENTO").ToString(), idEquipamento)
            End If

            If idEquipamento <= 0 Then
                SystemSounds.Exclamation.Play()
                MostrarFeedbackScan("Invalid equipment", False)
                txtQuickScan.Clear()
                Return
            End If

            Dim uid As String = ObterStrRow(row, "INTERNAL_UID")

            If ItemJaSelecionado(idEquipamento) Then
                SystemSounds.Beep.Play()
                MostrarFeedbackScan("Already added: " & uid, Nothing)
                txtQuickScan.Clear()
                Return
            End If

            Dim item As New ItemRemessa() With {
                .IdEquipamento = idEquipamento,
                .InternalUID   = uid,
                .Manufacturer  = ObterStrRow(row, "MARCA"),
                .Model         = ObterStrRow(row, "MODEL"),
                .ConditionAtShip = "GOOD",
                .SalePriceUsd  = Nothing,
                .Notes         = ""
            }

            vItensParaIncluir.Add(item)
            atualizarGridItens()

            SystemSounds.Beep.Play()
            MostrarFeedbackScan("Added: " & uid, True)

        Catch ex As Exception
            SystemSounds.Exclamation.Play()
            MostrarFeedbackScan("Error: " & ex.Message, False)
        End Try

        txtQuickScan.Clear()
    End Sub

    Private Sub MostrarFeedbackScan(msg As String, isSuccess As Boolean?)
        If isSuccess.HasValue Then
            lblQuickScanStatus.ForeColor = If(isSuccess.Value, TemaEscuro.Verde, TemaEscuro.Vermelho)
        Else
            lblQuickScanStatus.ForeColor = TemaEscuro.TextoMutado
        End If
        lblQuickScanStatus.Text = msg
        _timerScanFeedback.Stop()
        _timerScanFeedback.Start()
    End Sub

    Private Sub TimerScanFeedback_Tick(sender As Object, e As EventArgs)
        _timerScanFeedback.Stop()
        lblQuickScanStatus.Text = ""
        txtQuickScan.Focus()
    End Sub

    Private Function ObterStrRow(pRow As DataRow, colName As String) As String
        If Not pRow.Table.Columns.Contains(colName) Then Return ""
        If IsDBNull(pRow(colName)) Then Return ""
        Return pRow(colName).ToString()
    End Function

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub

    Private Sub configurarGridItens()

        dgvItensSelecionados.AutoGenerateColumns = False
        dgvItensSelecionados.Columns.Clear()
        dgvItensSelecionados.ReadOnly = True
        dgvItensSelecionados.MultiSelect = False
        dgvItensSelecionados.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvItensSelecionados.AllowUserToAddRows = False
        dgvItensSelecionados.AllowUserToDeleteRows = False
        dgvItensSelecionados.AllowUserToResizeRows = False

        Dim colUid As New DataGridViewTextBoxColumn() With {
            .Name = "COL_UID",
            .HeaderText = "UID",
            .Width = 140
        }

        Dim colMarca As New DataGridViewTextBoxColumn() With {
            .Name = "COL_MARCA",
            .HeaderText = "Brand",
            .Width = 120
        }

        Dim colModelo As New DataGridViewTextBoxColumn() With {
            .Name = "COL_MODELO",
            .HeaderText = "Model",
            .Width = 160
        }

        Dim colCondicao As New DataGridViewTextBoxColumn() With {
            .Name = "COL_CONDICAO",
            .HeaderText = "Condition",
            .Width = 100
        }

        Dim colUsd As New DataGridViewTextBoxColumn() With {
            .Name = "COL_USD",
            .HeaderText = "USD",
            .Width = 100,
            .DefaultCellStyle = New DataGridViewCellStyle() With {
                .Alignment = DataGridViewContentAlignment.MiddleRight,
                .Format = "0.00"
            }
        }

        dgvItensSelecionados.Columns.AddRange({colUid, colMarca, colModelo, colCondicao, colUsd})
        dgvItensSelecionados.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgvResultadoBusca.ReadOnly = True
        dgvResultadoBusca.MultiSelect = False
        dgvResultadoBusca.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvResultadoBusca.AllowUserToAddRows = False
        dgvResultadoBusca.AllowUserToDeleteRows = False
        dgvResultadoBusca.AllowUserToResizeRows = False

    End Sub

    Private Sub atualizarGridItens()

        dgvItensSelecionados.Rows.Clear()

        For Each item As ItemRemessa In vItensParaIncluir
            dgvItensSelecionados.Rows.Add(item.InternalUID,
                                          item.Manufacturer,
                                          item.Model,
                                          item.ConditionAtShip,
                                          If(item.SalePriceUsd.HasValue, item.SalePriceUsd.Value, 0D))
        Next

        lblContadorItens.Text = "Items: " & vItensParaIncluir.Count.ToString()

    End Sub

    Private Function ItemJaSelecionado(pIdEquipamento As Integer) As Boolean

        For Each item As ItemRemessa In vItensParaIncluir
            If item.IdEquipamento = pIdEquipamento Then
                Return True
            End If
        Next

        Return False

    End Function

    Private Function ParseDecimalNullable(pValor As String) As Decimal?

        If String.IsNullOrWhiteSpace(pValor) Then
            Return Nothing
        End If

        Dim s As String = pValor.Trim().Replace(",", ".")
        Dim v As Decimal

        If Decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
            Return v
        End If

        Return Nothing

    End Function

    Private Function ObterTextoCelula(pRow As DataGridViewRow, pColuna As String) As String

        If pRow Is Nothing OrElse pRow.Cells Is Nothing Then Return ""
        If Not pRow.DataGridView.Columns.Contains(pColuna) Then Return ""

        Dim valor As Object = pRow.Cells(pColuna).Value
        If valor Is Nothing OrElse IsDBNull(valor) Then Return ""

        Return valor.ToString()

    End Function

    Private Function ObterInt(pValor As Object) As Integer

        If pValor Is Nothing OrElse IsDBNull(pValor) Then Return 0

        Dim i As Integer
        If Integer.TryParse(pValor.ToString(), i) Then
            Return i
        End If

        Return 0

    End Function

End Class
