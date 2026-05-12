Imports System.Configuration
Imports System.Data
Imports System.Windows.Forms
Imports GBS_Inventory.Models

''' <summary>
''' Tela de cadastro de clientes.
''' </summary>
Public Class frmClientes

    Private oController As ClienteController
    Private vIdSelecionado As Integer = 0

    Public Sub New()
        InitializeComponent()
        oController = New ClienteController()

        Dim variante As String = ConfigurationManager.AppSettings("UiVariantClientesInvoice")
        If String.Equals(variante, "A", StringComparison.OrdinalIgnoreCase) Then
            TemaEscuro.aplicar(Me)
        Else
            TemaEscuro.aplicarHelius(Me)
        End If

    End Sub

    Private Sub frmClientes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        carregarClientes()
    End Sub

    Private Sub carregarClientes(Optional pFiltro As String = "")

        Try

            Dim ds As DataSet

            If String.IsNullOrWhiteSpace(pFiltro) Then
                ds = oController.buscarTodos(Nothing)
            Else
                ds = oController.buscarPorFiltro(pFiltro.Trim())
            End If

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                dgvClientes.DataSource = ds.Tables(0)
                ConfigurarColunasGrid()
            Else
                dgvClientes.DataSource = Nothing
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading customers: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnPesquisar_Click(sender As Object, e As EventArgs) Handles btnPesquisar.Click
        carregarClientes(txtPesquisa.Text)
    End Sub

    Private Sub txtPesquisa_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPesquisa.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnPesquisar_Click(sender, e)
        End If
    End Sub

    Private Sub dgvClientes_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvClientes.CellClick

        If e.RowIndex < 0 Then Return

        Dim row As DataGridViewRow = dgvClientes.Rows(e.RowIndex)

        vIdSelecionado = CInt(row.Cells("ID_CLIENTE").Value)
        lblSelecionado.Text = "Customer selected: " & row.Cells("NOME_RAZAO").Value.ToString()

        txtNomeRazao.Text = ObterTexto(row, "NOME_RAZAO")
        txtNomeFantasia.Text = ObterTexto(row, "NOME_FANTASIA")
        txtDocumento.Text = ObterTexto(row, "DOCUMENTO")
        txtEmail.Text = ObterTexto(row, "EMAIL")
        txtTelefone.Text = ObterTexto(row, "TELEFONE")
        txtEndereco1.Text = ObterTexto(row, "ENDERECO1")
        txtEndereco2.Text = ObterTexto(row, "ENDERECO2")
        txtCidade.Text = ObterTexto(row, "CIDADE")
        txtEstado.Text = ObterTexto(row, "ESTADO")
        txtZipCode.Text = ObterTexto(row, "ZIP_CODE")
        txtPais.Text = ObterTexto(row, "PAIS")
        txtObservacoes.Text = ObterTexto(row, "OBSERVACOES")
        chkAtivo.Checked = (ObterTexto(row, "ATIVO") = "Y")

    End Sub

    Private Sub btnNovo_Click(sender As Object, e As EventArgs) Handles btnNovo.Click
        limparTela()
    End Sub

    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click

        If String.IsNullOrWhiteSpace(txtNomeRazao.Text) Then
            MessageBox.Show("Enter Name / Company.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim cliente As Cliente = ObterClienteTela()

        Try

            If vIdSelecionado = 0 Then
                Dim idNovo As Integer = oController.incluir(cliente)
                MessageBox.Show("Customer created successfully. ID: " & idNovo.ToString(), "Success",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                vIdSelecionado = idNovo
            Else
                cliente.IdCliente = vIdSelecionado
                oController.alterar(cliente)
                MessageBox.Show("Customer updated successfully.", "Success",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            carregarClientes(txtPesquisa.Text)

        Catch ex As Exception
            MessageBox.Show("Error saving customer: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnExcluir_Click(sender As Object, e As EventArgs) Handles btnExcluir.Click

        If vIdSelecionado = 0 Then
            MessageBox.Show("Select a customer to delete.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If MessageBox.Show("Delete/deactivate this customer?", "Confirm",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
            Return
        End If

        Try
            oController.excluir(vIdSelecionado)
            MessageBox.Show("Customer removed successfully.", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            limparTela()
            carregarClientes(txtPesquisa.Text)
        Catch ex As Exception
            MessageBox.Show("Error deleting customer: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnAtualizar_Click(sender As Object, e As EventArgs) Handles btnAtualizar.Click
        carregarClientes(txtPesquisa.Text)
    End Sub

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub

    Private Sub limparTela()

        vIdSelecionado = 0
        lblSelecionado.Text = "New customer"

        txtNomeRazao.Clear()
        txtNomeFantasia.Clear()
        txtDocumento.Clear()
        txtEmail.Clear()
        txtTelefone.Clear()
        txtEndereco1.Clear()
        txtEndereco2.Clear()
        txtCidade.Clear()
        txtEstado.Clear()
        txtZipCode.Clear()
        txtPais.Text = "USA"
        txtObservacoes.Clear()
        chkAtivo.Checked = True

    End Sub

    Private Function ObterClienteTela() As Cliente

        Return New Cliente() With {
            .NomeRazao = txtNomeRazao.Text.Trim(),
            .NomeFantasia = txtNomeFantasia.Text.Trim(),
            .Documento = txtDocumento.Text.Trim(),
            .Email = txtEmail.Text.Trim(),
            .Telefone = txtTelefone.Text.Trim(),
            .Endereco1 = txtEndereco1.Text.Trim(),
            .Endereco2 = txtEndereco2.Text.Trim(),
            .Cidade = txtCidade.Text.Trim(),
            .Estado = txtEstado.Text.Trim(),
            .ZipCode = txtZipCode.Text.Trim(),
            .Pais = txtPais.Text.Trim(),
            .Observacoes = txtObservacoes.Text.Trim(),
            .Ativo = If(chkAtivo.Checked, "Y", "N")
        }

    End Function

    Private Sub ConfigurarColunasGrid()
        Dim mapa As New Dictionary(Of String, String) From {
            {"ID_CLIENTE",    "ID"},
            {"NOME_RAZAO",    "Name / Company"},
            {"NOME_FANTASIA", "Trade Name"},
            {"DOCUMENTO",     "Document"},
            {"EMAIL",         "Email"},
            {"TELEFONE",      "Phone"},
            {"ENDERECO1",     "Address 1"},
            {"ENDERECO2",     "Address 2"},
            {"CIDADE",        "City"},
            {"ESTADO",        "State"},
            {"ZIP_CODE",      "ZIP Code"},
            {"PAIS",          "Country"},
            {"ATIVO",         "Active"},
            {"OBSERVACOES",   "Notes"},
            {"DATA_CADASTRO", "Created"},
            {"DATA_ALTERACAO","Updated"}
        }
        For Each par In mapa
            If dgvClientes.Columns.Contains(par.Key) Then
                dgvClientes.Columns(par.Key).HeaderText = par.Value
            End If
        Next
    End Sub

    Private Function ObterTexto(pRow As DataGridViewRow, pColuna As String) As String

        If Not pRow.DataGridView.Columns.Contains(pColuna) Then Return ""
        Dim valor As Object = pRow.Cells(pColuna).Value
        If valor Is Nothing OrElse IsDBNull(valor) Then Return ""
        Return valor.ToString()

    End Function

End Class
