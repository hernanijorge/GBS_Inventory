Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.Configuration
Imports GBS_Inventory.Models
Imports GBS_Inventory.OracleHelper

Public Class frmAddComponent

    Private Class TypeItem
        Public Property Value As String
        Public Property Label As String
        Public Sub New(pValue As String, pLabel As String)
            Value = pValue
            Label = pLabel
        End Sub
        Public Overrides Function ToString() As String
            Return Label
        End Function
    End Class

    Public Property SavedUID   As String  = ""
    Public Property SavedCount As Integer = 0

    Private oController As ComponentController

    Private ReadOnly CapacityRAM  As String() = {"4", "8", "16", "32", "64", "128"}
    Private ReadOnly CapacitySSD  As String() = {"128", "256", "512", "1000", "2000", "4000"}
    Private ReadOnly CapacityHDD  As String() = {"320", "500", "1000", "2000", "4000", "8000"}
    Private ReadOnly GenerationRAM As String() = {"DDR4", "DDR5", "DDR3"}
    Private ReadOnly GenerationSSD As String() = {"NVMe", "SATA"}
    Private ReadOnly GenerationHDD As String() = {"SATA"}
    Private ReadOnly SpeedRAM     As String() = {"2133", "2400", "2666", "3200", "3600", "4800", "5200", "5600"}

    Public Sub New()
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        ConfigurarEstilos()
        oController = New ComponentController()
    End Sub

    Private Sub frmAddComponent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cboType.Items.AddRange({
            New TypeItem("RAM", "RAM"),
            New TypeItem("SSD", "SSD"),
            New TypeItem("HDD", "HDD"),
            New TypeItem("DESKTOP", "Desktop"),
            New TypeItem("MINI_DESKTOP", "Mini Desktop")
        })
        cboCondition.Items.AddRange({"GOOD", "FAIR", "POOR", "UNTESTED"})
        cboStatus.Items.AddRange({"IN_STOCK", "INSTALLED", "SOLD", "SCRAPPED"})
        cboType.SelectedIndex      = 0
        cboCondition.SelectedIndex = 0
        cboStatus.SelectedIndex    = 0
        CarregarBrands()
        CarregarSourceBatches()
        txtModel.AutoCompleteMode         = AutoCompleteMode.SuggestAppend
        txtModel.AutoCompleteSource       = AutoCompleteSource.CustomSource
        txtModel.AutoCompleteCustomSource = carregarSugestoesModelComponente()
    End Sub

    Private Sub ConfigurarEstilos()
        btnSave.Font = New Font("Segoe UI", 10, FontStyle.Bold)
    End Sub

    Private Sub CarregarBrands()
        Try
            Dim ds As DataSet = oController.fetchBrands()
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim v As String = row(0).ToString().Trim()
                    If Not String.IsNullOrEmpty(v) Then cboBrand.Items.Add(v)
                Next
            End If
        Catch
        End Try
    End Sub

    ' Suggestions are loaded once per dialog open (not per keystroke).
    Private Function carregarSugestoesModelComponente() As AutoCompleteStringCollection
        Dim col As New AutoCompleteStringCollection()
        Try
            Dim ds As DataSet = oController.fetchModels()
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim v As String = row(0).ToString().Trim()
                    If Not String.IsNullOrEmpty(v) Then col.Add(v)
                Next
            End If
        Catch
            ' silent — no suggestions; the user can still type freely
        End Try
        Return col
    End Function

    Private Sub CarregarSourceBatches()
        Try
            Dim cs As String = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim ds As DataSet = OracleHelper.ExecuteDataset(cs, System.Data.CommandType.Text,
                "SELECT DISTINCT SOURCE_BATCH FROM TBL_EQUIPAMENTO WHERE SOURCE_BATCH IS NOT NULL ORDER BY SOURCE_BATCH")
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim v As String = row(0).ToString().Trim()
                    If Not String.IsNullOrEmpty(v) Then cboSourceBatch.Items.Add(v)
                Next
            End If
        Catch
        End Try
    End Sub

    Private Sub cboType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboType.SelectedIndexChanged
        Dim selType As TypeItem = TryCast(cboType.SelectedItem, TypeItem)
        Dim t As String = If(selType IsNot Nothing, selType.Value, "")

        Dim isDesktop As Boolean = (t = "MINI_DESKTOP" OrElse t = "DESKTOP")

        lblCap.Text            = If(isDesktop, "RAM (GB)", "Capacity (GB) *")
        lblGen.Visible         = Not isDesktop
        cboGeneration.Visible  = Not isDesktop
        lblSpeedLabel.Visible  = Not isDesktop
        cboSpeed.Visible       = Not isDesktop
        lblCpu.Visible         = isDesktop
        txtCpu.Visible         = isDesktop
        lblStorage.Visible     = isDesktop
        txtStorage.Visible     = isDesktop

        txtCpu.Clear()
        txtStorage.Clear()

        cboCapacity.Items.Clear()
        cboGeneration.Items.Clear()
        cboSpeed.Items.Clear()
        cboCapacity.Text = ""
        Select Case t
            Case "RAM"
                cboCapacity.Items.AddRange(CapacityRAM)
                cboGeneration.Items.AddRange(GenerationRAM)
                cboSpeed.Items.AddRange(SpeedRAM)
                cboSpeed.Enabled     = True
                lblSpeedLabel.ForeColor = TemaEscuro.TextoMutado
            Case "SSD"
                cboCapacity.Items.AddRange(CapacitySSD)
                cboGeneration.Items.AddRange(GenerationSSD)
                cboSpeed.Enabled     = False
                lblSpeedLabel.ForeColor = Color.Gray
            Case "HDD"
                cboCapacity.Items.AddRange(CapacityHDD)
                cboGeneration.Items.AddRange(GenerationHDD)
                cboSpeed.Enabled     = False
                lblSpeedLabel.ForeColor = Color.Gray
        End Select
        If cboCapacity.Items.Count > 0  Then cboCapacity.SelectedIndex  = 0
        If cboGeneration.Items.Count > 0 Then cboGeneration.SelectedIndex = 0
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim selType As TypeItem = TryCast(cboType.SelectedItem, TypeItem)
        Dim tipo As String = If(selType IsNot Nothing, selType.Value, "")
        Dim isDesktop As Boolean = (tipo = "MINI_DESKTOP" OrElse tipo = "DESKTOP")
        Dim capTxt As String = cboCapacity.Text.Trim()
        Dim cpuTxt As String = txtCpu.Text.Trim()
        Dim storageTxt As String = txtStorage.Text.Trim()

        Dim erros As New List(Of String)()
        If String.IsNullOrEmpty(tipo) Then erros.Add("  · Type")

        Dim capGb As Integer = 0
        Dim storageGb As Integer = 0
        If isDesktop Then
            ' RAM / Storage are optional and never block the save: blank, 0 or
            ' unparseable text is saved as NULL (TryParse leaves 0 on failure)
            Integer.TryParse(capTxt, capGb)
            Integer.TryParse(storageTxt, storageGb)
            If String.IsNullOrEmpty(cpuTxt) Then erros.Add("  · CPU")
        Else
            If String.IsNullOrEmpty(capTxt) Then erros.Add("  · Capacity")
            If Not Integer.TryParse(capTxt, capGb) OrElse capGb <= 0 Then erros.Add("  · Capacity (invalid number)")
        End If

        If erros.Count > 0 Then
            MessageBox.Show("Required fields missing:" & vbCrLf & String.Join(vbCrLf, erros),
                            "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim comp As New Component() With {
            .ComponentType   = tipo,
            .CapacityGB      = If(capGb > 0, CType(capGb, Integer?), Nothing),
            .Generation      = If(Not isDesktop AndAlso cboGeneration.SelectedIndex >= 0, cboGeneration.SelectedItem.ToString(), ""),
            .Brand           = cboBrand.Text.Trim(),
            .Model           = txtModel.Text.Trim(),
            .PartNumber      = txtPartNumber.Text.Trim(),
            .Cpu             = If(isDesktop, cpuTxt, ""),
            .StorageGb       = If(isDesktop AndAlso storageGb > 0, CType(storageGb, Integer?), Nothing),
            .ConditionStatus = If(cboCondition.SelectedIndex >= 0, cboCondition.SelectedItem.ToString(), "GOOD"),
            .Status          = If(cboStatus.SelectedIndex >= 0, cboStatus.SelectedItem.ToString(), "IN_STOCK"),
            .SourceBatch     = cboSourceBatch.Text.Trim(),
            .Notes           = txtNotes.Text.Trim()
        }

        If Not isDesktop AndAlso cboSpeed.Enabled AndAlso cboSpeed.SelectedIndex >= 0 Then
            Dim spd As Integer
            If Integer.TryParse(cboSpeed.SelectedItem.ToString(), spd) Then comp.SpeedMhz = spd
        End If

        Dim qty As Integer = CInt(nudQuantity.Value)

        btnSave.Enabled = False
        Cursor = Cursors.WaitCursor

        Try
            Dim lastUID As String = ""
            For i As Integer = 1 To qty
                lastUID = oController.add(comp)
            Next
            SavedUID     = lastUID
            SavedCount   = qty
            DialogResult = DialogResult.OK
            Close()
        Catch ex As Exception
            MessageBox.Show("Error saving component:" & vbCrLf & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnSave.Enabled = True
            Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

End Class
