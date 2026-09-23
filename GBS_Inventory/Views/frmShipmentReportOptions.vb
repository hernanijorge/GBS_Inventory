Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports GBS_Inventory.Models

' Options for the Shipments "Report by Recipient". The caller reads Opcoes (or the individual
' properties) after DialogResult.OK.
Public Class frmShipmentReportOptions

    Private Const TODOS As String = "(All)"

    Private ReadOnly _controller As RemessaController
    Private _carregando As Boolean = True

    ' Shipments the user unchecked. Kept across list refreshes (filter changes) so a shipment
    ' excluded once stays excluded; shipments that newly enter the list come in checked.
    Private ReadOnly _desmarcados As New HashSet(Of Integer)()

    Private Class RemessaItem
        Public Property Id    As Integer
        Public Property Texto As String
        Public Overrides Function ToString() As String
            Return Texto
        End Function
    End Class

#Region "Options read by the caller"

    Public ReadOnly Property DataInicio As Date
        Get
            Return dtpFrom.Value.Date
        End Get
    End Property

    Public ReadOnly Property DataFim As Date
        Get
            Return dtpTo.Value.Date
        End Get
    End Property

    ' Status group keys (DELIVERED, IN_TRANSIT, PENDING, ISSUES) — see ShipmentReportOptions.GruposStatus
    Public ReadOnly Property StatusIncluidos As List(Of String)
        Get
            Dim grupos As New List(Of String)()
            If chkDelivered.Checked Then grupos.Add("DELIVERED")
            If chkInTransit.Checked Then grupos.Add("IN_TRANSIT")
            If chkPending.Checked   Then grupos.Add("PENDING")
            If chkIssues.Checked    Then grupos.Add("ISSUES")
            Return grupos
        End Get
    End Property

    ' Nothing = all recipients
    Public ReadOnly Property RecipientFiltro As String
        Get
            Dim sel As String = If(cboRecipient.SelectedItem IsNot Nothing, cboRecipient.SelectedItem.ToString(), TODOS)
            Return If(sel = TODOS, Nothing, sel)
        End Get
    End Property

    Public ReadOnly Property Detalhado As Boolean
        Get
            Return rbDetailed.Checked
        End Get
    End Property

    Public ReadOnly Property Agrupamento As String
        Get
            If rbGroupByPrefix.Checked    Then Return ShipmentReportOptions.AGRUP_PREFIXO
            If rbGroupManualAlias.Checked Then Return ShipmentReportOptions.AGRUP_ALIAS
            Return ShipmentReportOptions.AGRUP_NENHUM
        End Get
    End Property

    Public ReadOnly Property PrefixLen As Integer
        Get
            Return CInt(nudPrefixLen.Value)
        End Get
    End Property

    ' Nothing when every listed shipment is checked (report takes all matching shipments)
    Public ReadOnly Property IdsSelecionados As List(Of Integer)
        Get
            If clbShipments.CheckedItems.Count = clbShipments.Items.Count Then Return Nothing
            Return clbShipments.CheckedItems.Cast(Of RemessaItem)().Select(Function(i) i.Id).ToList()
        End Get
    End Property

    Public ReadOnly Property Opcoes As ShipmentReportOptions
        Get
            Dim o As ShipmentReportOptions = OpcoesDeFiltro()
            o.Detalhado       = Detalhado
            o.IdsSelecionados = IdsSelecionados
            Return o
        End Get
    End Property

    ' Filters only (no manual selection) — used for the checklist and the preview
    Private Function OpcoesDeFiltro() As ShipmentReportOptions
        Return New ShipmentReportOptions() With {
            .DataInicio      = DataInicio,
            .DataFim         = DataFim,
            .GruposIncluidos = StatusIncluidos,
            .RecipientFiltro = RecipientFiltro,
            .Agrupamento     = Agrupamento,
            .PrefixLen       = PrefixLen
        }
    End Function

#End Region

#Region "Init"

    Public Sub New(pController As RemessaController)
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        btnGenerate.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        _controller = pController

        dtpFrom.Value = Date.Today.AddDays(-90)
        dtpTo.Value   = Date.Today

        cboRecipient.Items.Add(TODOS)
        Try
            Dim ds As DataSet = _controller.buscarDestinatarios()
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim v As String = row(0).ToString().Trim()
                    If v <> "" Then cboRecipient.Items.Add(v)
                Next
            End If
        Catch
            ' silent — only (All) available; report still works
        End Try
        cboRecipient.SelectedIndex = 0

        AtualizarStatusMapeamento()
        AtualizarModoAgrupamento()
        _carregando = False
        AtualizarRemessasEPreview()
    End Sub

#End Region

#Region "Checklist + preview"

    Private Sub AtualizarRemessasEPreview()
        If _carregando Then Return
        AtualizarPreview()
        AtualizarListaRemessas()
    End Sub

    Private Sub AtualizarPreview()
        If StatusIncluidos.Count = 0 OrElse DataInicio > DataFim Then
            lblPreview.Text = ""
            Return
        End If
        Try
            Dim p = _controller.buscarPreviewAgrupamento(OpcoesDeFiltro())
            If Agrupamento = ShipmentReportOptions.AGRUP_NENHUM Then
                lblPreview.Text = "Preview: " & p.Nomes & " distinct clients (no merge)"
            Else
                lblPreview.Text = "Preview: " & p.Nomes & " distinct client names will become " & p.Grupos & " groups after merge"
            End If
        Catch
            lblPreview.Text = ""
        End Try
    End Sub

    Private Sub AtualizarListaRemessas()
        _carregando = True
        clbShipments.BeginUpdate()
        Try
            clbShipments.Items.Clear()
            If StatusIncluidos.Count > 0 AndAlso DataInicio <= DataFim Then
                Dim ds As DataSet = _controller.buscarRemessasParaSelecao(OpcoesDeFiltro())
                For Each r As DataRow In ds.Tables(0).Rows
                    Dim id As Integer = Convert.ToInt32(r("ID_REMESSA"))
                    Dim texto As String = CDate(r("DATA_CADASTRO")).ToString("yyyy-MM-dd") & "  " &
                                          r("REMESSA_REF").ToString().PadRight(22) & " " &
                                          r("RECIPIENT").ToString() & "  (" & r("TOTAL_ITEMS").ToString() & " items)"
                    clbShipments.Items.Add(New RemessaItem With {.Id = id, .Texto = texto}, Not _desmarcados.Contains(id))
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading shipments: " & ex.Message, "Report by Recipient",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            clbShipments.EndUpdate()
            _carregando = False
        End Try
        AtualizarContagemSelecao()
    End Sub

    Private Sub AtualizarContagemSelecao()
        lblSelecao.Text = clbShipments.CheckedItems.Count & " of " & clbShipments.Items.Count & " selected"
    End Sub

    Private Sub clbShipments_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles clbShipments.ItemCheck
        If _carregando Then Return
        Dim item As RemessaItem = DirectCast(clbShipments.Items(e.Index), RemessaItem)
        If e.NewValue = CheckState.Checked Then _desmarcados.Remove(item.Id) Else _desmarcados.Add(item.Id)
        ' ItemCheck fires before the state changes; count after it is applied
        BeginInvoke(Sub() AtualizarContagemSelecao())
    End Sub

    Private Sub MarcarTodos(pMarcar As Boolean)
        _carregando = True
        For i As Integer = 0 To clbShipments.Items.Count - 1
            clbShipments.SetItemChecked(i, pMarcar)
            Dim id As Integer = DirectCast(clbShipments.Items(i), RemessaItem).Id
            If pMarcar Then _desmarcados.Remove(id) Else _desmarcados.Add(id)
        Next
        _carregando = False
        AtualizarContagemSelecao()
    End Sub

    Private Sub btnCheckAll_Click(sender As Object, e As EventArgs) Handles btnCheckAll.Click
        MarcarTodos(True)
    End Sub

    Private Sub btnUncheckAll_Click(sender As Object, e As EventArgs) Handles btnUncheckAll.Click
        MarcarTodos(False)
    End Sub

    ' Any filter change refreshes the list and the preview
    Private Sub Filtro_Changed(sender As Object, e As EventArgs) Handles dtpFrom.ValueChanged, dtpTo.ValueChanged,
            chkDelivered.CheckedChanged, chkInTransit.CheckedChanged, chkPending.CheckedChanged, chkIssues.CheckedChanged,
            cboRecipient.SelectedIndexChanged
        AtualizarRemessasEPreview()
    End Sub

#End Region

#Region "Client grouping"

    Private Sub Agrupamento_Changed(sender As Object, e As EventArgs) Handles rbGroupNone.CheckedChanged,
            rbGroupByPrefix.CheckedChanged, rbGroupManualAlias.CheckedChanged
        ' CheckedChanged fires for the radio turning off too; react once, on the one turning on
        Dim rb As RadioButton = TryCast(sender, RadioButton)
        If rb IsNot Nothing AndAlso Not rb.Checked Then Return
        AtualizarModoAgrupamento()
        AtualizarRemessasEPreview()
    End Sub

    Private Sub nudPrefixLen_ValueChanged(sender As Object, e As EventArgs) Handles nudPrefixLen.ValueChanged
        AtualizarModoAgrupamento()
        AtualizarRemessasEPreview()
    End Sub

    Private Sub AtualizarModoAgrupamento()
        pnlPrefix.Visible = rbGroupByPrefix.Checked
        pnlAlias.Visible  = rbGroupManualAlias.Checked

        If PrefixLen <= 4 Then
            lblPrefixHint.Text      = "Short prefix: unrelated clients may be merged together"
            lblPrefixHint.ForeColor = TemaEscuro.Destaque
        Else
            lblPrefixHint.Text      = "Names sharing the first " & PrefixLen & " characters will be grouped"
            lblPrefixHint.ForeColor = TemaEscuro.TextoMutado
        End If
    End Sub

    Private Sub AtualizarStatusMapeamento()
        Try
            Dim mapa As Dictionary(Of String, List(Of String)) = _controller.buscarAliases()
            Dim qtd As Integer = mapa.Values.Sum(Function(l) l.Count)
            lblMappingStatus.Text = qtd & " mapping" & If(qtd = 1, "", "s") & " defined" &
                                    If(mapa.Count > 0, " (" & mapa.Count & " client" & If(mapa.Count = 1, "", "s") & ")", "")
        Catch
            lblMappingStatus.Text = "Mapping unavailable"
        End Try
    End Sub

    Private Sub btnEditMapping_Click(sender As Object, e As EventArgs) Handles btnEditMapping.Click
        Using frm As New frmClientAliasMapping(_controller, PrefixLen)
            If frm.ShowDialog(Me) = DialogResult.OK Then
                AtualizarStatusMapeamento()
                AtualizarRemessasEPreview()
            End If
        End Using
    End Sub

#End Region

#Region "Footer"

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If DataInicio > DataFim Then
            MessageBox.Show("'From' date cannot be after 'To' date.", "Report by Recipient",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If StatusIncluidos.Count = 0 Then
            MessageBox.Show("Select at least one status.", "Report by Recipient",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If clbShipments.Items.Count > 0 AndAlso clbShipments.CheckedItems.Count = 0 Then
            MessageBox.Show("Select at least one shipment.", "Report by Recipient",
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

#End Region

End Class
