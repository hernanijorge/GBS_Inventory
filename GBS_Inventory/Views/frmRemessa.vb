Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.Diagnostics
Imports GBS_Inventory.Models

''' <summary>
''' Formulário de gerenciamento de Remessas (FedEx, UPS, USPS).
''' Lista remessas ativas e permite criar novas, atualizar status e abrir URL de rastreamento.
''' </summary>
Public Class frmRemessa

    Private oController As RemessaController
    Private vIdRemessaSelecionada As Integer = 0

    Public Sub New()

        InitializeComponent()

        oController = New RemessaController()

        TemaEscuro.aplicar(Me)

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

        carregarRemessas()

    End Sub

    Private Sub carregarRemessas()

        Try

            Dim ds As DataSet = oController.buscarRemessasAtivas()

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvRemessas.DataSource = ds.Tables(0)
                lblTotal.Text = "Total ativas: " & ds.Tables(0).Rows.Count.ToString()
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

        lblSelecionada.Text = $"Selecionada: {row.Cells("REMESSA_REF").Value} · Tracking: {row.Cells("TRACKING_NUMBER").Value}"

        btnAtualizarStatus.Enabled = True
        btnAbrirRastreio.Enabled   = True

    End Sub

    Private Sub btnNovaRemessa_Click(sender As Object, e As EventArgs) Handles btnNovaRemessa.Click

        If String.IsNullOrWhiteSpace(txtTracking.Text) OrElse String.IsNullOrWhiteSpace(txtRecipientName.Text) Then
            MessageBox.Show("Preencha Tracking Number e Nome do Destinatário.", "Atenção",
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

            Dim idGerado As Integer = oController.incluirRemessa(r, New List(Of ItemRemessa)(), sRefGerada)

            MessageBox.Show($"Remessa criada!{vbCrLf}Ref: {sRefGerada}{vbCrLf}ID: {idGerado}",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

            limparCampos()
            carregarRemessas()

        Catch ex As Exception

            MessageBox.Show("Erro ao criar remessa: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

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

            MessageBox.Show("Status atualizado com sucesso!", "Sucesso",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

            carregarRemessas()

        Catch ex As Exception

            MessageBox.Show("Erro ao atualizar status: " & ex.Message, "Erro",
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
                MessageBox.Show("Carrier não suportado para rastreamento automático.", "Atenção",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Process.Start(r.UrlRastreio)

        Catch ex As Exception

            MessageBox.Show("Erro ao abrir rastreio: " & ex.Message, "Erro",
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
    End Sub

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub

End Class
