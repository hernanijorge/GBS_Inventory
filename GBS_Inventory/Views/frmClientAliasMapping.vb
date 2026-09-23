Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data

' Edits TBL_CLIENT_ALIAS: canonical client name -> spellings merged into it.
' Everything is edited in memory; the table is only replaced on "Save Mapping".
Public Class frmClientAliasMapping

    Private ReadOnly _controller As RemessaController
    Private ReadOnly _prefixLen  As Integer
    Private ReadOnly _mapa As New SortedDictionary(Of String, List(Of String))(StringComparer.Ordinal)
    Private _alterado As Boolean = False

    Public Sub New(pController As RemessaController, pPrefixLen As Integer)
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        btnSave.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        _controller = pController
        _prefixLen  = pPrefixLen
        btnAutoDetect.Text = "Auto-detect from Prefix (" & _prefixLen & ")"

        For Each g As DataGridView In {dgvClienteReal, dgvAliases}
            g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            g.ColumnHeadersVisible = True
        Next

        Try
            For Each par In _controller.buscarAliases()
                _mapa(par.Key) = New List(Of String)(par.Value)
            Next
            For Each r As DataRow In _controller.buscarDestinatarios().Tables(0).Rows
                cboNovoAlias.Items.Add(r(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show("Error loading client mapping: " & ex.Message, "Client Alias Mapping",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        AtualizarClientes(Nothing)
    End Sub

#Region "Grid refresh"

    Private ReadOnly Property ClienteSelecionado As String
        Get
            If dgvClienteReal.CurrentRow Is Nothing Then Return Nothing
            Return TryCast(dgvClienteReal.CurrentRow.Cells(0).Value, String)
        End Get
    End Property

    Private Sub AtualizarClientes(pSelecionar As String)
        dgvClienteReal.Rows.Clear()
        For Each k As String In _mapa.Keys
            Dim i As Integer = dgvClienteReal.Rows.Add(k)
            If k = pSelecionar Then dgvClienteReal.CurrentCell = dgvClienteReal.Rows(i).Cells(0)
        Next
        AtualizarAliases()
        Dim qtd As Integer = _mapa.Values.Sum(Function(l) l.Count)
        lblInfo.Text = _mapa.Count & " client(s), " & qtd & " alias(es)" & If(_alterado, " — unsaved changes", "")
    End Sub

    Private Sub AtualizarAliases()
        dgvAliases.Rows.Clear()
        Dim c As String = ClienteSelecionado
        If c Is Nothing OrElse Not _mapa.ContainsKey(c) Then Return
        For Each a As String In _mapa(c).OrderBy(Function(s) s, StringComparer.Ordinal)
            dgvAliases.Rows.Add(a)
        Next
    End Sub

    Private Sub dgvClienteReal_SelectionChanged(sender As Object, e As EventArgs) Handles dgvClienteReal.SelectionChanged
        AtualizarAliases()
    End Sub

#End Region

#Region "Validation helpers"

    Private Shared Function Normalizar(pNome As String) As String
        Return If(pNome, "").Trim().ToUpperInvariant()
    End Function

    ' Canonical name that already owns pNome as an alias, or Nothing
    Private Function DonoDoAlias(pNome As String) As String
        For Each par In _mapa
            If par.Value.Contains(pNome) Then Return par.Key
        Next
        Return Nothing
    End Function

    Private Sub MarcarAlterado(pSelecionar As String)
        _alterado = True
        AtualizarClientes(pSelecionar)
    End Sub

#End Region

#Region "Clients"

    Private Sub btnAddClient_Click(sender As Object, e As EventArgs) Handles btnAddClient.Click
        Dim nome As String = Normalizar(InputBox("Canonical client name (as it should appear in the report):",
                                                 "Add Client", If(cboNovoAlias.Text, "")))
        If nome = "" Then Return
        If nome.Length > 200 Then
            MessageBox.Show("Name is too long (max 200 characters).", "Add Client", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim dono As String = DonoDoAlias(nome)
        If dono IsNot Nothing Then
            MessageBox.Show("'" & nome & "' is already an alias of '" & dono & "'." & vbCrLf &
                            "Remove it from there first.", "Add Client", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If Not _mapa.ContainsKey(nome) Then _mapa(nome) = New List(Of String)()
        MarcarAlterado(nome)
    End Sub

    Private Sub btnRemoveClient_Click(sender As Object, e As EventArgs) Handles btnRemoveClient.Click
        Dim c As String = ClienteSelecionado
        If c Is Nothing Then Return
        If MessageBox.Show("Remove '" & c & "' and its " & _mapa(c).Count & " alias(es) from the mapping?",
                           "Remove Client", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        _mapa.Remove(c)
        MarcarAlterado(Nothing)
    End Sub

#End Region

#Region "Aliases"

    Private Sub btnAddAlias_Click(sender As Object, e As EventArgs) Handles btnAddAlias.Click
        Dim c As String = ClienteSelecionado
        If c Is Nothing Then
            MessageBox.Show("Select (or add) a canonical client first.", "Add Alias", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim nome As String = Normalizar(cboNovoAlias.Text)
        If nome = "" Then Return
        If nome = c Then
            MessageBox.Show("'" & nome & "' is the canonical name itself — it is always included.", "Add Alias",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        If _mapa.ContainsKey(nome) Then
            MessageBox.Show("'" & nome & "' is a canonical client. Remove it first to make it an alias.", "Add Alias",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim dono As String = DonoDoAlias(nome)
        If dono IsNot Nothing Then
            MessageBox.Show("'" & nome & "' is already an alias of '" & dono & "'.", "Add Alias",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        _mapa(c).Add(nome)
        cboNovoAlias.Text = ""
        MarcarAlterado(c)
    End Sub

    Private Sub btnRemoveAlias_Click(sender As Object, e As EventArgs) Handles btnRemoveAlias.Click
        Dim c As String = ClienteSelecionado
        If c Is Nothing OrElse dgvAliases.CurrentRow Is Nothing Then Return
        _mapa(c).Remove(dgvAliases.CurrentRow.Cells(0).Value.ToString())
        MarcarAlterado(c)
    End Sub

#End Region

#Region "Auto-detect"

    ' Suggests one client per group of recipient spellings sharing the first N characters:
    ' shortest spelling becomes canonical, the rest become its aliases. Existing entries win;
    ' nothing is saved until "Save Mapping".
    Private Sub btnAutoDetect_Click(sender As Object, e As EventArgs) Handles btnAutoDetect.Click
        Dim nomes As List(Of String) = cboNovoAlias.Items.Cast(Of Object)().Select(Function(o) o.ToString()).ToList()
        Dim grupos = nomes.GroupBy(Function(n) If(n.Length > _prefixLen, n.Substring(0, _prefixLen), n)) _
                          .Where(Function(g) g.Count() > 1)

        Dim novosGrupos As Integer = 0, novosAliases As Integer = 0
        For Each g In grupos
            Dim ordenados = g.OrderBy(Function(n) n.Length).ThenBy(Function(n) n, StringComparer.Ordinal).ToList()
            Dim canonico As String = ordenados(0)
            If DonoDoAlias(canonico) IsNot Nothing Then Continue For   ' already merged elsewhere
            If Not _mapa.ContainsKey(canonico) Then
                _mapa(canonico) = New List(Of String)()
                novosGrupos += 1
            End If
            For Each n As String In ordenados.Skip(1)
                If _mapa.ContainsKey(n) OrElse DonoDoAlias(n) IsNot Nothing Then Continue For
                _mapa(canonico).Add(n)
                novosAliases += 1
            Next
        Next

        If novosAliases = 0 Then
            MessageBox.Show("No new groups found with a " & _prefixLen & "-character prefix.", "Auto-detect",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        MarcarAlterado(Nothing)
        MessageBox.Show(novosGrupos & " new client(s) and " & novosAliases & " alias(es) suggested." & vbCrLf &
                        "Review them (remove wrong ones) and click Save Mapping.", "Auto-detect",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

#End Region

#Region "Footer"

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            ' Clients with no alias add nothing to the report; drop them instead of saving
            Dim paraSalvar As Dictionary(Of String, List(Of String)) =
                _mapa.Where(Function(p) p.Value.Count > 0).ToDictionary(Function(p) p.Key, Function(p) p.Value)
            _controller.salvarAliases(paraSalvar)
            DialogResult = DialogResult.OK
            Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Client Alias Mapping", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If _alterado AndAlso MessageBox.Show("Discard unsaved mapping changes?", "Client Alias Mapping",
                                             MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

#End Region

End Class
