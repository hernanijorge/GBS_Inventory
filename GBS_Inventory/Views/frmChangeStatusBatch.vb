Imports System.Windows.Forms
Imports System.Drawing

' Batch status change for the Components grid. The caller reads NovoStatus / Nota
' after DialogResult.OK and performs the update.
Public Class frmChangeStatusBatch

    Private ReadOnly StatusOpcoes As String() = {"IN_STOCK", "RESERVED", "SHIPPED", "SOLD", "SCRAPPED", "IN_REPAIR"}
    Private ReadOnly StatusSaidaEstoque As String() = {"SOLD", "SHIPPED"}

    Public ReadOnly Property NovoStatus As String
        Get
            Return If(cboNovoStatus.SelectedItem IsNot Nothing, cboNovoStatus.SelectedItem.ToString(), "")
        End Get
    End Property

    Public ReadOnly Property Nota As String
        Get
            Return txtNota.Text.Trim()
        End Get
    End Property

    Public Sub New(pQuantidade As Integer)
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        btnOK.Font = New Font("Segoe UI", 10, FontStyle.Bold)

        lblInfo.Text = "Change status of " & pQuantidade.ToString() & " selected item" &
                       If(pQuantidade = 1, "", "s") & " to:"
        cboNovoStatus.Items.AddRange(StatusOpcoes)
        cboNovoStatus.SelectedIndex = 0
    End Sub

    Private Sub cboNovoStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNovoStatus.SelectedIndexChanged
        lblAviso.Visible = StatusSaidaEstoque.Contains(NovoStatus)
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        If String.IsNullOrEmpty(NovoStatus) Then Return
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

End Class
