Imports System.Windows.Forms
Imports System.Drawing
Imports GBS_Inventory.Models

' Options shown before the Components consolidated report. The caller reads Opcoes
' (or the individual properties) after DialogResult.OK.
Public Class frmReportOptions

    Public ReadOnly Property AgruparPorModel As Boolean
        Get
            Return rbGroupWithModel.Checked
        End Get
    End Property

    ' Known types only (DESKTOP, MINI_DESKTOP, RAM, SSD, HDD); "Other types" is IncluirOutrosTipos
    Public ReadOnly Property TiposIncluidos As List(Of String)
        Get
            Return TiposPorCheck().Where(Function(p) p.Value.Checked).Select(Function(p) p.Key).ToList()
        End Get
    End Property

    Public ReadOnly Property IncluirOutrosTipos As Boolean
        Get
            Return chkOthers.Checked
        End Get
    End Property

    Public ReadOnly Property ColunasStatus As List(Of String)
        Get
            Return StatusPorCheck().Where(Function(p) p.Value.Checked).Select(Function(p) p.Key).ToList()
        End Get
    End Property

    Public ReadOnly Property Opcoes As ComponentSummaryOptions
        Get
            Return New ComponentSummaryOptions() With {
                .AgruparPorModel    = AgruparPorModel,
                .TiposIncluidos     = TiposIncluidos,
                .IncluirOutrosTipos = IncluirOutrosTipos,
                .ColunasStatus      = ColunasStatus
            }
        End Get
    End Property

    Public Sub New()
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        btnGenerate.Font = New Font("Segoe UI", 10, FontStyle.Bold)
    End Sub

    Private Function TiposPorCheck() As List(Of KeyValuePair(Of String, CheckBox))
        Return New List(Of KeyValuePair(Of String, CheckBox)) From {
            New KeyValuePair(Of String, CheckBox)("DESKTOP",      chkDesktop),
            New KeyValuePair(Of String, CheckBox)("MINI_DESKTOP", chkMiniDesktop),
            New KeyValuePair(Of String, CheckBox)("RAM",          chkRam),
            New KeyValuePair(Of String, CheckBox)("SSD",          chkSsd),
            New KeyValuePair(Of String, CheckBox)("HDD",          chkHdd)
        }
    End Function

    Private Function StatusPorCheck() As List(Of KeyValuePair(Of String, CheckBox))
        Return New List(Of KeyValuePair(Of String, CheckBox)) From {
            New KeyValuePair(Of String, CheckBox)("IN_STOCK",  chkColInStock),
            New KeyValuePair(Of String, CheckBox)("INSTALLED", chkColInstalled),
            New KeyValuePair(Of String, CheckBox)("SOLD",      chkColSold),
            New KeyValuePair(Of String, CheckBox)("SCRAPPED",  chkColScrapped)
        }
    End Function

    ' Checks exactly pTipos (and Other types when pOutros); everything else is cleared
    Private Sub MarcarTipos(pTipos As String(), pOutros As Boolean)
        For Each p In TiposPorCheck()
            p.Value.Checked = pTipos.Contains(p.Key)
        Next
        chkOthers.Checked = pOutros
    End Sub

    Private Sub btnSelectAll_Click(sender As Object, e As EventArgs) Handles btnSelectAll.Click
        MarcarTipos(ComponentSummaryOptions.TiposConhecidos, True)
    End Sub

    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        MarcarTipos({}, False)
    End Sub

    Private Sub btnOnlyMachines_Click(sender As Object, e As EventArgs) Handles btnOnlyMachines.Click
        MarcarTipos(ComponentSummaryOptions.TiposMaquinas, False)
    End Sub

    Private Sub btnOnlyParts_Click(sender As Object, e As EventArgs) Handles btnOnlyParts.Click
        MarcarTipos(ComponentSummaryOptions.TiposPecas, False)
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If TiposIncluidos.Count = 0 AndAlso Not IncluirOutrosTipos Then
            MessageBox.Show("Select at least one component type.", "Report Options",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If ColunasStatus.Count = 0 Then
            MessageBox.Show("Select at least one status column.", "Report Options",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

End Class
