Imports System.Windows.Forms
Imports System.Drawing
Imports GBS_Inventory.Models

''' <summary>
''' Dialog to register a component upgrade on a piece of equipment.
''' Modes: Install from Inventory, Remove Component, Extract to Stock, Manual Entry.
''' </summary>
Public Class frmUpgrade

    Private oController     As UpgradeController
    Private oCompController As ComponentController
    Private vIdEquipamento  As Integer
    Private sInternalUID    As String

    Public Sub New(pIdEquipamento As Integer, pInternalUID As String)

        InitializeComponent()

        oController     = New UpgradeController()
        oCompController = New ComponentController()
        vIdEquipamento  = pIdEquipamento
        sInternalUID    = pInternalUID

        TemaEscuro.aplicarHelius(Me)

        lblUID.Text = "UID: " & pInternalUID

    End Sub

    Private Sub frmUpgrade_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cboComponente.Items.Clear()
        cboComponente.Items.AddRange({"RAM", "SSD", "HDD", "BATTERY", "SCREEN", "KEYBOARD", "COVER", "GPU", "OTHER"})
        cboComponente.SelectedIndex = 0

        cboOrigem.Items.Clear()
        Try
            Dim ds As DataSet = oController.buscarOrigensDistintas()
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim v As String = row(0).ToString().Trim()
                    If Not String.IsNullOrEmpty(v) Then cboOrigem.Items.Add(v)
                Next
            End If
        Catch
        End Try
        If cboOrigem.Items.Count > 0 Then cboOrigem.SelectedIndex = 0

        AplicarModoAcao()

    End Sub

    ' ── Mode helpers ────────────────────────────────────────────────────────

    Private Function GetCurrentMode() As String
        If rbInstall.Checked Then Return "INSTALL"
        If rbRemove.Checked  Then Return "REMOVE"
        If rbExtract.Checked Then Return "EXTRACT"
        Return "MANUAL"
    End Function

    Private Sub AplicarModoAcao()
        Dim mode     As String  = GetCurrentMode()
        Dim isManual  As Boolean = (mode = "MANUAL")
        Dim isExtract As Boolean = (mode = "EXTRACT")
        Dim isRemove  As Boolean = (mode = "REMOVE")
        Dim isStock   As Boolean = (mode = "INSTALL" OrElse mode = "REMOVE")

        ' Stock selector vs Part Source vs Extract panel
        lblCompStock.Visible      = isStock
        cboComponentStock.Visible = isStock
        lblOrigem.Visible         = isManual
        cboOrigem.Visible         = isManual
        pnlExtract.Visible        = isExtract

        ' Disposition (REMOVE only)
        pnlDisposition.Visible = isRemove

        ' Labels
        If mode = "INSTALL" Then
            lblCompStock.Text    = "Select IN-STOCK Component *"
            lblValorDepois.Text  = "Value AFTER (e.g.: 16 GB) *"
            lblValorAntes.Text   = "Value BEFORE (e.g.: 8 GB)"
        ElseIf mode = "REMOVE" Then
            lblCompStock.Text    = "Select INSTALLED Component *"
            lblValorDepois.Text  = "Value AFTER (e.g.: 0 GB)"
            lblValorAntes.Text   = "Value BEFORE (e.g.: 16 GB)"
        ElseIf mode = "EXTRACT" Then
            lblValorAntes.Text   = "Machine RAM/Storage BEFORE (e.g.: 32 GB)"
            lblValorDepois.Text  = "Machine RAM/Storage AFTER  (e.g.: 8 GB)"
        Else
            lblValorDepois.Text  = "Value AFTER (e.g.: 16 GB) *"
            lblValorAntes.Text   = "Value BEFORE (e.g.: 8 GB)"
        End If

        If isStock Then
            CarregarComponentesEstoque(mode)
        End If
    End Sub

    Private Sub CarregarComponentesEstoque(mode As String)
        cboComponentStock.DataSource    = Nothing
        cboComponentStock.DisplayMember = ""
        cboComponentStock.ValueMember   = ""
        cboComponentStock.Items.Clear()

        Dim tipo As String = If(cboComponente.SelectedItem IsNot Nothing, cboComponente.SelectedItem.ToString(), "")
        If String.IsNullOrEmpty(tipo) Then Return

        Try
            Dim ds As DataSet = If(mode = "INSTALL",
                oController.buscarComponentesParaInstalar(tipo),
                oController.buscarComponentesInstalados(tipo))

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                cboComponentStock.DataSource    = ds.Tables(0)
                cboComponentStock.DisplayMember = "DISPLAY_TEXT"
                cboComponentStock.ValueMember   = "ID_COMPONENT"
            Else
                Dim needed As String = If(mode = "INSTALL", "IN_STOCK", "INSTALLED")
                MessageBox.Show("No " & needed & " " & tipo & " components found in inventory.",
                                "No Components", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading components: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ── Event handlers ──────────────────────────────────────────────────────

    Private Sub rbInstall_CheckedChanged(sender As Object, e As EventArgs) Handles rbInstall.CheckedChanged
        If rbInstall.Checked Then AplicarModoAcao()
    End Sub

    Private Sub rbRemove_CheckedChanged(sender As Object, e As EventArgs) Handles rbRemove.CheckedChanged
        If rbRemove.Checked Then AplicarModoAcao()
    End Sub

    Private Sub rbExtract_CheckedChanged(sender As Object, e As EventArgs) Handles rbExtract.CheckedChanged
        If rbExtract.Checked Then AplicarModoAcao()
    End Sub

    Private Sub rbManual_CheckedChanged(sender As Object, e As EventArgs) Handles rbManual.CheckedChanged
        If rbManual.Checked Then AplicarModoAcao()
    End Sub

    Private Sub cboComponente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboComponente.SelectedIndexChanged
        Dim mode As String = GetCurrentMode()
        If mode = "INSTALL" OrElse mode = "REMOVE" Then
            CarregarComponentesEstoque(mode)
        End If
    End Sub

    Private Sub cboComponentStock_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboComponentStock.SelectedIndexChanged
        If cboComponentStock.SelectedItem Is Nothing Then Return
        Dim row As DataRowView = TryCast(cboComponentStock.SelectedItem, DataRowView)
        If row Is Nothing Then Return

        Dim cap As String = ""
        If Not IsDBNull(row("CAPACITY_GB")) Then
            cap = row("CAPACITY_GB").ToString() & " GB"
        End If

        Dim mode As String = GetCurrentMode()
        If mode = "INSTALL" Then
            txtValorDepois.Text = cap
        ElseIf mode = "REMOVE" Then
            txtValorAntes.Text = cap
        End If
    End Sub

    ' ── Save ────────────────────────────────────────────────────────────────

    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click

        Dim mode      As String  = GetCurrentMode()
        Dim isManual  As Boolean = (mode = "MANUAL")
        Dim isExtract As Boolean = (mode = "EXTRACT")
        Dim isStock   As Boolean = (mode = "INSTALL" OrElse mode = "REMOVE")

        ' ── Validation ──
        If isStock Then
            If cboComponentStock.DataSource Is Nothing OrElse cboComponentStock.SelectedValue Is Nothing Then
                MessageBox.Show("Please select a component from the inventory list.",
                                "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cboComponentStock.Focus()
                Return
            End If
        ElseIf isExtract Then
            If nudExtractCapacity.Value < 1 Then
                MessageBox.Show("Enter the capacity (GB) of the components to extract.",
                                "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                nudExtractCapacity.Focus()
                Return
            End If
        Else
            If String.IsNullOrWhiteSpace(txtValorDepois.Text) Then
                MessageBox.Show("Fill in the VALUE AFTER (e.g.: 16 GB, 512 GB).",
                                "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtValorDepois.Focus()
                Return
            End If
        End If

        Try
            ' ── Extract to Stock: create N component records ──
            If isExtract Then
                Dim qty    As Integer = CInt(nudExtractQty.Value)
                Dim capGB  As Integer = CInt(nudExtractCapacity.Value)
                Dim speed  As Integer = CInt(nudExtractSpeed.Value)
                Dim gen    As String  = txtExtractGeneration.Text.Trim()
                Dim brand  As String  = txtExtractBrand.Text.Trim()
                Dim tipo   As String  = cboComponente.SelectedItem.ToString()

                For i As Integer = 1 To qty
                    Dim comp As New Models.Component() With {
                        .ComponentType   = tipo,
                        .CapacityGB      = capGB,
                        .SpeedMhz        = If(speed > 0, CType(speed, Integer?), Nothing),
                        .Generation      = If(String.IsNullOrEmpty(gen), Nothing, gen),
                        .Brand           = If(String.IsNullOrEmpty(brand), Nothing, brand),
                        .ConditionStatus = "USED",
                        .Status          = "IN_STOCK",
                        .SourceBatch     = sInternalUID,
                        .Notes           = "Extracted from machine " & sInternalUID
                    }
                    oCompController.add(comp)
                Next

                ' Log the upgrade (value before/after updates the machine spec)
                Dim descricao As String = qty & "x " & capGB & " GB" &
                                          If(String.IsNullOrEmpty(gen), "", " " & gen) &
                                          " extracted to stock."
                If Not String.IsNullOrWhiteSpace(txtNotas.Text) Then
                    descricao &= " " & txtNotas.Text.Trim()
                End If

                Dim upEx As New Upgrade() With {
                    .IdEquipamento = vIdEquipamento,
                    .InternalUID   = sInternalUID,
                    .ComponentType = cboComponente.SelectedItem.ToString(),
                    .ValueBefore   = txtValorAntes.Text.Trim(),
                    .ValueAfter    = txtValorDepois.Text.Trim(),
                    .PartSerial    = "",
                    .SourceOrigem  = "",
                    .Technician    = txtTecnico.Text.Trim(),
                    .Notes         = descricao,
                    .IdComponent   = Nothing,
                    .ActionType    = "EXTRACT",
                    .CompNewStatus = Nothing
                }
                oController.incluir(upEx)

                MessageBox.Show(qty & " component(s) added to stock successfully!",
                                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
                Return
            End If

            ' ── Install / Remove / Manual ──
            Dim idComp As Integer? = Nothing
            If isStock Then
                Dim sv As Object = cboComponentStock.SelectedValue
                If sv IsNot Nothing AndAlso Not IsDBNull(sv) Then idComp = CInt(sv)
            End If

            Dim compNewStatus As String = Nothing
            If mode = "REMOVE" Then
                compNewStatus = If(rbScrap.Checked, "SCRAPPED", "IN_STOCK")
            End If

            Dim up As New Upgrade() With {
                .IdEquipamento = vIdEquipamento,
                .InternalUID   = sInternalUID,
                .ComponentType = cboComponente.SelectedItem.ToString(),
                .ValueBefore   = txtValorAntes.Text.Trim(),
                .ValueAfter    = txtValorDepois.Text.Trim(),
                .PartSerial    = txtPartSerial.Text.Trim(),
                .SourceOrigem  = If(isManual, cboOrigem.Text.Trim(), ""),
                .Technician    = txtTecnico.Text.Trim(),
                .Notes         = txtNotas.Text.Trim(),
                .IdComponent   = idComp,
                .ActionType    = mode,
                .CompNewStatus = compNewStatus
            }

            Dim vCusto As Decimal
            If Decimal.TryParse(txtCusto.Text.Replace(",", "."),
                                Globalization.NumberStyles.Any,
                                Globalization.CultureInfo.InvariantCulture, vCusto) AndAlso vCusto > 0 Then
                up.CostUsd = vCusto
            End If

            oController.incluir(up)

            Dim modeLabel As String = If(mode = "INSTALL", "installed", If(mode = "REMOVE", "removed", "registered"))
            MessageBox.Show("Component " & modeLabel & " successfully!",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
