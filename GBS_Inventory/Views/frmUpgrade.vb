Imports System.Windows.Forms
Imports System.Drawing
Imports GBS_Inventory.Models

''' <summary>
''' Diálogo para registrar um upgrade (RAM, SSD, etc) em um equipamento.
''' </summary>
Public Class frmUpgrade

    Private oController As UpgradeController
    Private vIdEquipamento As Integer
    Private sInternalUID As String

    Public Sub New(pIdEquipamento As Integer, pInternalUID As String)

        InitializeComponent()

        oController    = New UpgradeController()
        vIdEquipamento = pIdEquipamento
        sInternalUID   = pInternalUID

        TemaEscuro.aplicarHelius(Me)

        lblUID.Text = "UID: " & pInternalUID

    End Sub

    Private Sub frmUpgrade_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cboComponente.Items.Clear()
        cboComponente.Items.AddRange({"RAM", "SSD", "HDD", "BATTERY", "SCREEN", "KEYBOARD", "COVER", "GPU", "OTHER"})
        cboComponente.SelectedIndex = 0

        cboOrigem.Items.Clear()
        cboOrigem.Items.AddRange({"SHIPMENT_SURPLUS", "NEW_PURCHASE", "TRANSFERRED", "WARRANTY", "OTHER"})
        cboOrigem.SelectedIndex = 0

    End Sub

    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click

        If String.IsNullOrWhiteSpace(txtValorDepois.Text) Then
            MessageBox.Show("Fill in the VALUE AFTER (e.g.: 16 GB, 512 GB).", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtValorDepois.Focus()
            Return
        End If

        Try

            Dim up As New Upgrade() With {
                .IdEquipamento = vIdEquipamento,
                .InternalUID   = sInternalUID,
                .ComponentType = cboComponente.SelectedItem.ToString(),
                .ValueBefore   = txtValorAntes.Text.Trim(),
                .ValueAfter    = txtValorDepois.Text.Trim(),
                .PartSerial    = txtPartSerial.Text.Trim(),
                .SourceOrigem  = cboOrigem.SelectedItem.ToString(),
                .Technician    = txtTecnico.Text.Trim(),
                .Notes         = txtNotas.Text.Trim()
            }

            Dim vCusto As Decimal
            If Decimal.TryParse(txtCusto.Text.Replace(",", "."), Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, vCusto) Then
                up.CostUsd = vCusto
            End If

            oController.incluir(up)

            MessageBox.Show("Upgrade registered successfully!", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception

            MessageBox.Show("Error saving upgrade: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class
