Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data

''' <summary>
''' Tela de Scanner — campo grande para ler Internal UID do código de barras
''' ou digitar manualmente. Ao ler, exibe card com todas as informações do item.
''' </summary>
Public Class frmScanner

    Private oController As EquipamentoController
    Private oUpgradeController As UpgradeController
    Private vIdEquipamentoAtual As Integer = 0

    Public Sub New()
        InitializeComponent()
        oController        = New EquipamentoController()
        oUpgradeController = New UpgradeController()
        TemaEscuro.aplicarHelius(Me)
    End Sub

    Private Sub frmScanner_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtUID.Focus()
    End Sub

    Private Sub txtUID_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUID.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            buscarUID()
        End If
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        buscarUID()
    End Sub

    Private Sub buscarUID()

        Dim sUID As String = txtUID.Text.Trim()
        If String.IsNullOrEmpty(sUID) Then
            lblResultado.Text = "⚠ Enter or scan an Internal UID"
            lblResultado.ForeColor = TemaEscuro.Vermelho
            limparDetalhes()
            Return
        End If

        Try

            Dim ds As DataSet = oController.buscarPorUID(sUID)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                Dim row As DataRow = ds.Tables(0).Rows(0)
                vIdEquipamentoAtual = CInt(row("ID_EQUIPAMENTO"))

                lblResultado.Text     = "✓ Equipment found"
                lblResultado.ForeColor = TemaEscuro.Accent

                lblManufacturer.Text = row("MARCA").ToString() & " " & row("MODEL").ToString()
                lblSerial.Text       = "Serial: "    & row("SERIAL_NUMBER").ToString()
                lblCPU.Text          = "CPU: "       & row("PROCESSADOR").ToString()
                lblRAM.Text          = "Memory: "    & row("RAM_GB").ToString() & " GB"
                lblStorage.Text      = "Storage: "   & row("STORAGE_GB").ToString() & " GB"
                lblCondition.Text    = ""
                lblStatus.Text       = "Status: "    & row("STATUS_DESCRICAO").ToString()

                Dim upgrades As Integer = 0
                Try
                    If Not IsDBNull(row("TOTAL_UPGRADES")) Then
                        upgrades = CInt(row("TOTAL_UPGRADES"))
                    End If
                Catch
                End Try
                lblUpgrades.Text     = "Upgrades: " & upgrades.ToString()

                ' Carrega grid de upgrades
                carregarUpgrades(vIdEquipamentoAtual)

                btnAddUpgrade.Enabled = True

            Else

                lblResultado.Text = "✗ UID not found: " & sUID
                lblResultado.ForeColor = TemaEscuro.Vermelho
                limparDetalhes()

            End If

            txtUID.SelectAll()

        Catch ex As Exception

            MessageBox.Show("Search error: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub carregarUpgrades(pIdEquipamento As Integer)
        Try
            Dim ds As DataSet = oUpgradeController.buscarPorEquipamento(pIdEquipamento)
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvUpgrades.DataSource = ds.Tables(0)
            End If
        Catch ex As Exception
            ' silencioso
        End Try
    End Sub

    Private Sub limparDetalhes()
        lblManufacturer.Text = ""
        lblSerial.Text       = ""
        lblCPU.Text          = ""
        lblRAM.Text          = ""
        lblStorage.Text      = ""
        lblCondition.Text    = ""
        lblStatus.Text       = ""
        lblUpgrades.Text     = ""
        dgvUpgrades.DataSource = Nothing
        btnAddUpgrade.Enabled  = False
        vIdEquipamentoAtual    = 0
    End Sub

    Private Sub btnAddUpgrade_Click(sender As Object, e As EventArgs) Handles btnAddUpgrade.Click
        If vIdEquipamentoAtual = 0 Then Return

        Dim frm As New frmUpgrade(vIdEquipamentoAtual, txtUID.Text.Trim())
        If frm.ShowDialog(Me) = DialogResult.OK Then
            carregarUpgrades(vIdEquipamentoAtual)
            buscarUID()
        End If
    End Sub

End Class
