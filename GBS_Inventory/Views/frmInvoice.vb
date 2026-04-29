Imports System.Configuration
Imports System.Data
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Imports GBS_Inventory.Models

''' <summary>
''' Tela de geracao e envio de invoice.
''' </summary>
Public Class frmInvoice

    Private oClienteController As ClienteController
    Private oRemessaController As RemessaController
    Private oInvoiceController As InvoiceController

    Private dtClientes As DataTable
    Private dtItens As DataTable

    Private vIdInvoiceAtual As Integer = 0
    Private vInvoiceNumeroAtual As String = ""
    Private vArquivoWord As String = ""
    Private vArquivoPdf As String = ""

    Public Sub New()
        InitializeComponent()

        oClienteController = New ClienteController()
        oRemessaController = New RemessaController()
        oInvoiceController = New InvoiceController()

        Dim variante As String = ConfigurationManager.AppSettings("UiVariantClientesInvoice")
        If String.Equals(variante, "A", StringComparison.OrdinalIgnoreCase) Then
            TemaEscuro.aplicar(Me)
        Else
            TemaEscuro.aplicarHelius(Me)
        End If
    End Sub

    Private Sub frmInvoice_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dtIssue.Value = Date.Today
        dtDue.Value = Date.Today.AddDays(15)

        txtLogo.Text = ConfigurationManager.AppSettings("InvoiceLogoPath")
        txtLogo.Text = ObterLogoPathEfetivo()

        Dim pastaSaida As String = ConfigurationManager.AppSettings("InvoiceOutputPath")
        If String.IsNullOrWhiteSpace(pastaSaida) Then
            pastaSaida = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GBS_Invoices")
        End If
        txtOutput.Text = pastaSaida

        carregarClientes()
        carregarRemessas()
        prepararGridItens()
        atualizarTotais()

    End Sub

    Private Sub carregarClientes()

        Try

            Dim ds As DataSet = oClienteController.buscarTodos("Y")

            If ds Is Nothing OrElse ds.Tables.Count = 0 Then
                cboCliente.DataSource = Nothing
                dtClientes = Nothing
                Return
            End If

            dtClientes = ds.Tables(0)

            cboCliente.DataSource = dtClientes
            cboCliente.DisplayMember = "NOME_RAZAO"
            cboCliente.ValueMember = "ID_CLIENTE"

        Catch ex As Exception
            MessageBox.Show("Error loading customers: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub carregarRemessas()

        Try

            Dim ds As DataSet = oRemessaController.buscarRemessasAtivas()

            If ds Is Nothing OrElse ds.Tables.Count = 0 Then
                cboRemessa.DataSource = Nothing
                Return
            End If

            cboRemessa.DataSource = ds.Tables(0)
            cboRemessa.DisplayMember = "REMESSA_REF"
            cboRemessa.ValueMember = "ID_REMESSA"

        Catch ex As Exception
            MessageBox.Show("Error loading shipments: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub prepararGridItens()

        dtItens = New DataTable()
        dtItens.Columns.Add("ID_EQUIPAMENTO", GetType(Integer))
        dtItens.Columns.Add("INTERNAL_UID", GetType(String))
        dtItens.Columns.Add("DESCRICAO", GetType(String))
        dtItens.Columns.Add("QTY", GetType(Decimal))
        dtItens.Columns.Add("UNIT_PRICE_USD", GetType(Decimal))
        dtItens.Columns.Add("LINE_TOTAL_USD", GetType(Decimal))
        dtItens.Columns.Add("NOTES", GetType(String))

        dgvItens.DataSource = dtItens
        dgvItens.ReadOnly = False

        dgvItens.Columns("ID_EQUIPAMENTO").Visible = False
        dgvItens.Columns("INTERNAL_UID").ReadOnly = True
        dgvItens.Columns("DESCRICAO").ReadOnly = False
        dgvItens.Columns("QTY").ReadOnly = False
        dgvItens.Columns("UNIT_PRICE_USD").ReadOnly = False
        dgvItens.Columns("LINE_TOTAL_USD").ReadOnly = True
        dgvItens.Columns("NOTES").ReadOnly = False

        dgvItens.Columns("QTY").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvItens.Columns("UNIT_PRICE_USD").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvItens.Columns("LINE_TOTAL_USD").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

    End Sub

    Private Sub btnCarregarItens_Click(sender As Object, e As EventArgs) Handles btnCarregarItens.Click

        Dim idRemessa As Integer
        If Not tentarObterIdCombo(cboRemessa, idRemessa) Then
            MessageBox.Show("Select a shipment.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            Dim ds As DataSet = oRemessaController.buscarItensRemessa(idRemessa)

            dtItens.Rows.Clear()

            If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                atualizarTotais()
                MessageBox.Show("Shipment has no items.", "Warning",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            For Each row As DataRow In ds.Tables(0).Rows

                Dim mfr As String = If(IsDBNull(row("MARCA")), "", row("MARCA").ToString())
                Dim modelo As String = If(IsDBNull(row("MODEL")), "", row("MODEL").ToString())
                Dim uid As String = If(IsDBNull(row("INTERNAL_UID")), "", row("INTERNAL_UID").ToString())
                Dim descricao As String = (mfr & " " & modelo & " (" & uid & ")").Trim()
                Dim unitPrice As Decimal = 0D

                If Not IsDBNull(row("SALE_PRICE_USD")) Then
                    unitPrice = Convert.ToDecimal(row("SALE_PRICE_USD"))
                End If

                Dim dr As DataRow = dtItens.NewRow()
                dr("ID_EQUIPAMENTO") = DBNull.Value
                dr("INTERNAL_UID") = uid
                dr("DESCRICAO") = descricao
                dr("QTY") = 1D
                dr("UNIT_PRICE_USD") = unitPrice
                dr("LINE_TOTAL_USD") = unitPrice
                dr("NOTES") = If(IsDBNull(row("NOTES")), "", row("NOTES").ToString())

                dtItens.Rows.Add(dr)

            Next

            atualizarTotais()

        Catch ex As Exception
            MessageBox.Show("Erro ao carregar itens da remessa: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub dgvItens_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvItens.CellEndEdit

        If e.RowIndex < 0 Then Return

        Dim row As DataGridViewRow = dgvItens.Rows(e.RowIndex)
        Dim qtd As Decimal = ParseDecimal(row.Cells("QTY").Value)
        Dim unit As Decimal = ParseDecimal(row.Cells("UNIT_PRICE_USD").Value)

        row.Cells("QTY").Value = qtd
        row.Cells("UNIT_PRICE_USD").Value = unit
        row.Cells("LINE_TOTAL_USD").Value = qtd * unit

        atualizarTotais()

    End Sub

    Private Sub txtValor_TextChanged(sender As Object, e As EventArgs) _
        Handles txtDiscount.TextChanged, txtShipping.TextChanged, txtTax.TextChanged
        atualizarTotais()
    End Sub

    Private Sub atualizarTotais()

        Dim subtotal As Decimal = 0D

        If dtItens IsNot Nothing Then
            For Each row As DataRow In dtItens.Rows
                subtotal += ParseDecimal(row("LINE_TOTAL_USD"))
            Next
        End If

        Dim discount As Decimal = ParseDecimal(txtDiscount.Text)
        Dim shipping As Decimal = ParseDecimal(txtShipping.Text)
        Dim tax As Decimal = ParseDecimal(txtTax.Text)
        Dim total As Decimal = subtotal - discount + shipping + tax

        lblSubtotal.Text = "Subtotal: " & subtotal.ToString("0.00", CultureInfo.InvariantCulture)
        lblTotal.Text = "Total: " & total.ToString("0.00", CultureInfo.InvariantCulture)

    End Sub

    Private Sub cboCliente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboCliente.SelectedIndexChanged

        If cboCliente.SelectedItem Is Nothing Then Return

        Dim drv As DataRowView = TryCast(cboCliente.SelectedItem, DataRowView)
        If drv Is Nothing Then Return

        txtEmail.Text = If(IsDBNull(drv("EMAIL")), "", drv("EMAIL").ToString())

    End Sub

    Private Sub btnSelecionarLogo_Click(sender As Object, e As EventArgs) Handles btnSelecionarLogo.Click

        Using ofd As New OpenFileDialog()
            ofd.Filter = "Imagens (*.png;*.jpg;*.jpeg;*.gif)|*.png;*.jpg;*.jpeg;*.gif|Todos os arquivos (*.*)|*.*"
            If ofd.ShowDialog() = DialogResult.OK Then
                txtLogo.Text = ofd.FileName
            End If
        End Using

    End Sub

    Private Sub btnSelecionarPasta_Click(sender As Object, e As EventArgs) Handles btnSelecionarPasta.Click

        Using fbd As New FolderBrowserDialog()
            If Directory.Exists(txtOutput.Text) Then
                fbd.SelectedPath = txtOutput.Text
            End If

            If fbd.ShowDialog() = DialogResult.OK Then
                txtOutput.Text = fbd.SelectedPath
            End If
        End Using

    End Sub

    Private Sub btnGerar_Click(sender As Object, e As EventArgs) Handles btnGerar.Click

        Dim clienteId As Integer
        If Not tentarObterIdCombo(cboCliente, clienteId) Then
            MessageBox.Show("Select a customer.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim remessaId As Integer
        If Not tentarObterIdCombo(cboRemessa, remessaId) Then
            MessageBox.Show("Select a shipment.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If dtItens Is Nothing OrElse dtItens.Rows.Count = 0 Then
            MessageBox.Show("Load shipment items before generating the invoice.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            Dim cliente As Cliente = carregarCliente(clienteId)
            Dim itens As List(Of InvoiceItem) = montarItens()

            Dim invoice As New Invoice() With {
                .IdCliente = clienteId,
                .IdRemessa = remessaId,
                .IssueDate = dtIssue.Value.Date,
                .DueDate = dtDue.Value.Date,
                .DiscountUsd = ParseDecimal(txtDiscount.Text),
                .ShippingUsd = ParseDecimal(txtShipping.Text),
                .TaxUsd = ParseDecimal(txtTax.Text),
                .ClientEmail = txtEmail.Text.Trim(),
                .Notes = txtNotes.Text.Trim(),
                .Currency = "USD"
            }

            Dim subtotal As Decimal = 0D
            For Each it As InvoiceItem In itens
                subtotal += it.LineTotalUsd
            Next
            invoice.SubtotalUsd = subtotal
            invoice.TotalUsd = subtotal - invoice.DiscountUsd + invoice.ShippingUsd + invoice.TaxUsd

            Dim invoiceNumero As String = ""
            Dim idGerado As Integer = oInvoiceController.incluirComItens(invoice, itens, invoiceNumero)

            invoice.IdInvoice = idGerado
            invoice.InvoiceNumber = invoiceNumero

            Dim pastaSaida As String = txtOutput.Text.Trim()
            If String.IsNullOrWhiteSpace(pastaSaida) Then
                pastaSaida = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "GBS_Invoices")
            End If

            If Not Directory.Exists(pastaSaida) Then
                Directory.CreateDirectory(pastaSaida)
            End If

            Dim nomeBase As String = LimparNomeArquivo(invoiceNumero)
            Dim arqWord As String = Path.Combine(pastaSaida, nomeBase & ".doc")
            Dim arqPdf As String = Path.Combine(pastaSaida, nomeBase & ".pdf")
            Dim logo As String = ObterLogoPathEfetivo()
            txtLogo.Text = logo

            InvoiceDocumentoService.GerarArquivos(invoice, cliente, itens, logo, arqWord, arqPdf)
            oInvoiceController.atualizarPaths(idGerado, logo, arqWord, arqPdf)

            vIdInvoiceAtual = idGerado
            vInvoiceNumeroAtual = invoiceNumero
            vArquivoWord = arqWord
            vArquivoPdf = arqPdf

            txtInvoiceNumber.Text = invoiceNumero

            MessageBox.Show("Invoice generated successfully!" & vbCrLf &
                            "Word: " & arqWord & vbCrLf &
                            "PDF: " & arqPdf,
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error generating invoice: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnEnviar_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click

        Dim somentePdf As Boolean = chkEnviarSomentePdf.Checked

        If vIdInvoiceAtual = 0 OrElse String.IsNullOrWhiteSpace(vArquivoPdf) OrElse
           (Not somentePdf AndAlso String.IsNullOrWhiteSpace(vArquivoWord)) Then
            MessageBox.Show("Generate the invoice before sending the email.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try

            Dim destino As String = txtEmail.Text.Trim()
            Dim assunto As String = "Invoice " & vInvoiceNumeroAtual & " - Global Business Solution"
            Dim corpo As String
            If somentePdf Then
                corpo = "Hello," & vbCrLf & vbCrLf &
                        "Please find attached your invoice in PDF format." & vbCrLf & vbCrLf &
                        "Best regards," & vbCrLf &
                        "Global Business Solution"
            Else
                corpo = "Hello," & vbCrLf & vbCrLf &
                        "Please find attached your invoice in PDF and Word formats." & vbCrLf & vbCrLf &
                        "Best regards," & vbCrLf &
                        "Global Business Solution"
            End If

            Dim anexos As New List(Of String) From {vArquivoPdf}

            If Not somentePdf Then
                anexos.Add(vArquivoWord)
            End If

            If Not somentePdf AndAlso Not String.IsNullOrWhiteSpace(txtLogo.Text) AndAlso File.Exists(txtLogo.Text) Then
                anexos.Add(txtLogo.Text)
            End If

            InvoiceEmailService.EnviarInvoice(destino, assunto, corpo, anexos)
            oInvoiceController.atualizarStatus(vIdInvoiceAtual, "SENT")

            MessageBox.Show("Email sent successfully.", "Success",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error sending email: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub

    Private Function montarItens() As List(Of InvoiceItem)

        Dim lista As New List(Of InvoiceItem)()

        For Each row As DataRow In dtItens.Rows
            Dim item As New InvoiceItem() With {
                .IdEquipamento = Nothing,
                .Description = row("DESCRICAO").ToString(),
                .Qty = ParseDecimal(row("QTY")),
                .UnitPriceUsd = ParseDecimal(row("UNIT_PRICE_USD")),
                .Notes = If(IsDBNull(row("NOTES")), "", row("NOTES").ToString())
            }
            item.RecalcularTotal()
            row("LINE_TOTAL_USD") = item.LineTotalUsd
            lista.Add(item)
        Next

        Return lista

    End Function

    Private Function carregarCliente(pId As Integer) As Cliente

        Dim ds As DataSet = oClienteController.buscarPorId(pId)

        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Customer not found.")
        End If

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Return New Cliente() With {
            .IdCliente = pId,
            .NomeRazao = ToTexto(row, "NOME_RAZAO"),
            .NomeFantasia = ToTexto(row, "NOME_FANTASIA"),
            .Documento = ToTexto(row, "DOCUMENTO"),
            .Email = ToTexto(row, "EMAIL"),
            .Telefone = ToTexto(row, "TELEFONE"),
            .Endereco1 = ToTexto(row, "ENDERECO1"),
            .Endereco2 = ToTexto(row, "ENDERECO2"),
            .Cidade = ToTexto(row, "CIDADE"),
            .Estado = ToTexto(row, "ESTADO"),
            .ZipCode = ToTexto(row, "ZIP_CODE"),
            .Pais = ToTexto(row, "PAIS")
        }

    End Function

    Private Function ParseDecimal(pValor As Object) As Decimal

        If pValor Is Nothing OrElse IsDBNull(pValor) Then Return 0D

        Dim s As String = pValor.ToString().Trim()
        If String.IsNullOrWhiteSpace(s) Then Return 0D

        Dim v As Decimal

        If Decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
            Return v
        End If

        s = s.Replace(",", ".")

        If Decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
            Return v
        End If

        Return 0D

    End Function

    Private Function ToTexto(pRow As DataRow, pColuna As String) As String
        If Not pRow.Table.Columns.Contains(pColuna) Then Return ""
        If IsDBNull(pRow(pColuna)) Then Return ""
        Return pRow(pColuna).ToString()
    End Function

    Private Function LimparNomeArquivo(pNome As String) As String

        Dim invalido As Char() = Path.GetInvalidFileNameChars()
        Dim s As String = pNome

        For Each c As Char In invalido
            s = s.Replace(c, "_"c)
        Next

        Return s

    End Function

    Private Function tentarObterIdCombo(pCombo As ComboBox, ByRef pId As Integer) As Boolean

        pId = 0

        If pCombo Is Nothing OrElse pCombo.SelectedValue Is Nothing Then
            Return False
        End If

        If TypeOf pCombo.SelectedValue Is Integer Then
            pId = CInt(pCombo.SelectedValue)
            Return True
        End If

        Return Integer.TryParse(pCombo.SelectedValue.ToString(), pId)

    End Function

    Private Function ObterLogoPathEfetivo() As String

        Dim atual As String = txtLogo.Text.Trim()
        If Not String.IsNullOrWhiteSpace(atual) AndAlso File.Exists(atual) Then
            Return atual
        End If

        Dim candidatos As String() = {
            "C:\GBS\logo.png",
            "C:\GBS\logo.jpg",
            "C:\GBS\logo.jpeg"
        }

        For Each c As String In candidatos
            If File.Exists(c) Then
                Return c
            End If
        Next

        Return ""

    End Function

End Class
