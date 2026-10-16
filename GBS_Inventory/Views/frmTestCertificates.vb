Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.IO
Imports System.IO.Compression

''' <summary>
''' Certificados de teste de uma remessa: pre-visualiza cada maquina com o veredito do
''' diagnostico mais recente e gera UM PDF (capa + PDFs individuais) para enviar ao cliente.
''' Nao altera nada na remessa.
''' </summary>
Public Class frmTestCertificates

    Private Const LimiteAnexoMB As Integer = 20

    Private ReadOnly _idRemessa As Long
    Private ReadOnly _controller As New DiagnosticoController()
    Private _cabecalho As DataRow
    Private _linhas As DataTable
    Private _caminhoGerado As String = ""
    Private ReadOnly _fonteDestaque As New Font("Segoe UI", 9, FontStyle.Bold)

    Public Sub New(pIdRemessa As Long)
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        _idRemessa = pIdRemessa
    End Sub

    Private Sub frmTestCertificates_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            _cabecalho = _controller.fetchShipmentHeader(_idRemessa)
            If _cabecalho Is Nothing Then
                MessageBox.Show("Shipment not found.", "Test Certificates", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Close()
                Return
            End If
            _linhas = _controller.fetchByShipment(_idRemessa)
        Catch ex As Exception
            MessageBox.Show("Error loading shipment: " & ex.Message, "Test Certificates", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Close()
            Return
        End Try

        Dim codigo As String = Texto(_cabecalho, "CODIGO_REMESSA")
        Me.Text = "Test Certificates - " & codigo
        lblTitulo.Text = "Test Certificates — Shipment " & codigo
        Dim email As String = Texto(_cabecalho, "CLIENTE_EMAIL")
        If Texto(_cabecalho, "ID_CLIENTE") = "" Then
            lblCliente.Text = "Customer: " & Texto(_cabecalho, "CLIENTE_NOME") &
                              "   –  No customer linked to this shipment (Shipments tab → Set Customer)"
            lblCliente.ForeColor = TemaEscuro.Destaque
        ElseIf email = "" Then
            lblCliente.Text = "Customer: " & Texto(_cabecalho, "CLIENTE_NOME") &
                              "   –  Customer email not found – register the customer's email"
            lblCliente.ForeColor = TemaEscuro.Destaque
        Else
            lblCliente.Text = "Customer: " & Texto(_cabecalho, "CLIENTE_NOME") & "   <" & email & ">"
        End If
        lblContador.Text = CertificadoService.Resumo(_linhas.Rows.Cast(Of DataRow)())

        dgvMaquinas.DataSource = _linhas
        ConfigurarGrid()
        AtualizarIncluidos()
    End Sub

#Region "Grid"

    Private Sub ConfigurarGrid()
        ' O tema forca Fill; aqui larguras fixas com Model em Fill
        dgvMaquinas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgvMaquinas.RowHeadersVisible = False

        Dim spec As (Key As String, Header As String, Largura As Integer, Centro As Boolean)() = {
            ("INTERNAL_UID",   "Internal UID", 100, False),
            ("SERVICE_TAG",    "Service Tag",  115, False),
            ("MODEL",          "Model",          0, False),
            ("VERDICT",        "Verdict",      150, True),
            ("BATTERY_HEALTH", "Battery %",     85, True),
            ("SSD_LIFE",       "SSD %",         75, True),
            ("DATA_TESTE",     "Test Date",    140, False),
            ("PDF_BYTES",      "PDF",           60, True)
        }

        For Each col As DataGridViewColumn In dgvMaquinas.Columns
            col.Visible = False
        Next
        For i As Integer = 0 To spec.Length - 1
            If Not dgvMaquinas.Columns.Contains(spec(i).Key) Then Continue For
            Dim col As DataGridViewColumn = dgvMaquinas.Columns(spec(i).Key)
            col.Visible = True
            col.HeaderText = spec(i).Header
            col.DisplayIndex = i
            If spec(i).Largura = 0 Then
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                col.MinimumWidth = 150
            Else
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                col.Width = spec(i).Largura
            End If
            Dim alin As DataGridViewContentAlignment = If(spec(i).Centro, DataGridViewContentAlignment.MiddleCenter, DataGridViewContentAlignment.MiddleLeft)
            col.DefaultCellStyle.Alignment = alin
            col.HeaderCell.Style.Alignment = alin
            col.DefaultCellStyle.NullValue = DiagnosticoReportService.TextoNulo
        Next

        For Each key As String In {"BATTERY_HEALTH", "SSD_LIFE"}
            dgvMaquinas.Columns(key).DefaultCellStyle.Format = DiagnosticoReportService.FormatoPercentual
            dgvMaquinas.Columns(key).DefaultCellStyle.FormatProvider = Globalization.CultureInfo.InvariantCulture
        Next
        dgvMaquinas.Columns("DATA_TESTE").DefaultCellStyle.Format = DiagnosticoReportService.FormatoData
        dgvMaquinas.Columns("DATA_TESTE").DefaultCellStyle.FormatProvider = Globalization.CultureInfo.InvariantCulture
        dgvMaquinas.Columns("VERDICT").DefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
    End Sub

    ' Veredito com as cores da aba Diagnostics; sem teste = "NOT TESTED" cinza.
    ' Linhas FAIL / INCOMPLETE / NOT TESTED ficam em negrito para chamar atencao.
    Private Sub dgvMaquinas_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvMaquinas.CellFormatting
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        Dim nome As String = dgvMaquinas.Columns(e.ColumnIndex).Name
        Dim drv As DataRowView = TryCast(dgvMaquinas.Rows(e.RowIndex).DataBoundItem, DataRowView)
        If drv Is Nothing Then Return

        Dim codigo As String = Texto(drv.Row, "VERDICT")
        Dim problema As Boolean = Not CertificadoService.Aprovado(drv.Row)

        Select Case nome
            Case "VERDICT"
                e.Value = If(codigo = "", "NOT TESTED", DiagnosticoReportService.VerdictTexto(codigo))
                e.CellStyle.BackColor = CertificadoService.CorVerdictCapa(codigo)
                e.CellStyle.ForeColor = Color.FromArgb(17, 24, 39)
                e.FormattingApplied = True
            Case "PDF_BYTES"
                Dim temPdf As Boolean = Not IsDBNull(drv.Row("PDF_BYTES")) AndAlso CDec(drv.Row("PDF_BYTES")) > 0
                e.Value = If(temPdf, "Yes", "No")
                e.FormattingApplied = True
        End Select

        If problema AndAlso nome <> "VERDICT" Then
            e.CellStyle.Font = _fonteDestaque
            e.CellStyle.ForeColor = Color.FromArgb(253, 186, 116)
        End If
    End Sub

#End Region

#Region "Gerar"

    Private Sub rbOpcao_CheckedChanged(sender As Object, e As EventArgs) Handles rbSoAprovados.CheckedChanged, rbTodos.CheckedChanged
        If _linhas IsNot Nothing Then AtualizarIncluidos()
    End Sub

    Private Function LinhasIncluidas() As List(Of DataRow)
        Dim todas As IEnumerable(Of DataRow) = _linhas.Rows.Cast(Of DataRow)()
        If rbSoAprovados.Checked Then todas = todas.Where(Function(r) CertificadoService.Aprovado(r))
        Return todas.ToList()
    End Function

    Private Sub AtualizarIncluidos()
        Dim inc As List(Of DataRow) = LinhasIncluidas()
        Dim comPdf As Integer = inc.Where(Function(r) TemPdf(r)).Count()
        lblIncluidos.Text = $"Will include {inc.Count} machine(s) on the cover and {comPdf} certificate(s)."
        btnGenerate.Enabled = inc.Count > 0
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Dim inc As List(Of DataRow) = LinhasIncluidas()
        If inc.Count = 0 Then Return

        Dim pastaTmp As String = Path.Combine(Path.GetTempPath(), "GBS_Certificates", Guid.NewGuid().ToString("N"))
        Try
            Cursor = Cursors.WaitCursor
            btnGenerate.Enabled = False

            ' PDFs individuais: BLOB -> arquivo (um subdiretorio por diagnostico, sem colisao de nomes)
            Dim pdfs As New Dictionary(Of Long, String)
            For Each r As DataRow In inc
                If Not TemPdf(r) Then Continue For
                Dim id As Long = CLng(r("ID_DIAGNOSTICO"))
                Dim caminho As String = _controller.savePdf(id, Path.Combine(pastaTmp, id.ToString()))
                If caminho <> "" Then pdfs(id) = caminho
            Next

            Dim codigo As String = Texto(_cabecalho, "CODIGO_REMESSA")
            Dim nome As String = $"GBS_Test_Certificates_{codigo}_{DateTime.Now:yyyyMMdd}.pdf"
            For Each c As Char In Path.GetInvalidFileNameChars()
                nome = nome.Replace(c, "_"c)
            Next
            Dim saida As String = Path.Combine(PastaRelatorios(), "Certificates", nome)
            If ArquivoEmUso(saida) Then
                saida = Path.Combine(Path.GetDirectoryName(saida), Path.GetFileNameWithoutExtension(nome) & "_" & DateTime.Now.ToString("HHmmss") & ".pdf")
            End If

            Dim res As CertificadoService.Resultado =
                CertificadoService.GerarPdfCombinado(_cabecalho, inc, pdfs, saida, LogoRelatorios())
            _caminhoGerado = res.Caminho

            Cursor = Cursors.Default
            Dim tamanhoMB As Double = New FileInfo(_caminhoGerado).Length / 1024.0 / 1024.0
            lblArquivo.Text = $"{Path.GetFileName(_caminhoGerado)} ({tamanhoMB:0.0} MB)"
            btnOpenPdf.Enabled = True
            btnSaveAs.Enabled = True
            btnEmail.Enabled = True

            If res.CertificadosComFalha.Count > 0 Then
                MessageBox.Show("Some certificates could not be embedded (a notice page was added instead):" & vbCrLf &
                                String.Join(", ", res.CertificadosComFalha), "Test Certificates",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show("Error generating certificates: " & ex.Message, "Test Certificates", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnGenerate.Enabled = True
            Try : Directory.Delete(pastaTmp, True) : Catch : End Try
        End Try
    End Sub

#End Region

#Region "Acoes"

    Private Sub btnOpenPdf_Click(sender As Object, e As EventArgs) Handles btnOpenPdf.Click
        If File.Exists(_caminhoGerado) Then Process.Start(_caminhoGerado)
    End Sub

    Private Sub btnSaveAs_Click(sender As Object, e As EventArgs) Handles btnSaveAs.Click
        If Not File.Exists(_caminhoGerado) Then Return
        Using dlg As New SaveFileDialog() With {
            .FileName = Path.GetFileName(_caminhoGerado),
            .Filter = "PDF files (*.pdf)|*.pdf",
            .DefaultExt = "pdf"
        }
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Try
                File.Copy(_caminhoGerado, dlg.FileName, True)
            Catch ex As Exception
                MessageBox.Show("Error saving file: " & ex.Message, "Save As", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub btnEmail_Click(sender As Object, e As EventArgs) Handles btnEmail.Click
        If Not File.Exists(_caminhoGerado) Then Return

        Dim anexo As String = _caminhoGerado
        Dim tamanhoMB As Double = New FileInfo(anexo).Length / 1024.0 / 1024.0
        If tamanhoMB > LimiteAnexoMB Then
            Dim resp As DialogResult = MessageBox.Show(
                $"The PDF is {tamanhoMB:0.0} MB, above the {LimiteAnexoMB} MB that most email servers accept." & vbCrLf & vbCrLf &
                "Create a ZIP and attach it instead?" & vbCrLf &
                "(Yes = attach ZIP, No = attach the PDF anyway, Cancel = do not send)",
                "Large attachment", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
            If resp = DialogResult.Cancel Then Return
            If resp = DialogResult.Yes Then
                Try
                    anexo = CriarZip(_caminhoGerado)
                    Dim zipMB As Double = New FileInfo(anexo).Length / 1024.0 / 1024.0
                    If zipMB > LimiteAnexoMB Then
                        MessageBox.Show($"The ZIP is still {zipMB:0.0} MB. Consider generating with 'only passed machines' " &
                                        "or sending the file by another channel.", "Large attachment",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error creating ZIP: " & ex.Message, "Large attachment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try
            End If
        End If

        Dim codigo As String = Texto(_cabecalho, "CODIGO_REMESSA")
        Dim inc As List(Of DataRow) = LinhasIncluidas()
        Dim assunto As String = "Test Certificates - Shipment " & codigo
        Dim corpo As String =
            "Hello," & vbCrLf & vbCrLf &
            $"Please find attached the hardware test certificates for shipment {codigo}." & vbCrLf &
            $"Summary: {CertificadoService.Resumo(inc)}." & vbCrLf & vbCrLf &
            "The first page lists every device; the individual test reports follow in the same order." & vbCrLf & vbCrLf &
            "Please let us know if you have any questions."

        Using frm As New frmEnviarRelatorio("", anexo, "", New List(Of DataRow), assunto, corpo,
                                            Texto(_cabecalho, "CLIENTE_EMAIL"))
            frm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Shared Function CriarZip(pArquivo As String) As String
        Dim zip As String = Path.ChangeExtension(pArquivo, ".zip")
        If File.Exists(zip) Then File.Delete(zip)
        Using arq As ZipArchive = ZipFile.Open(zip, ZipArchiveMode.Create)
            arq.CreateEntryFromFile(pArquivo, Path.GetFileName(pArquivo), CompressionLevel.Optimal)
        End Using
        Return zip
    End Function

#End Region

#Region "Utilitarios"

    Private Shared Function TemPdf(r As DataRow) As Boolean
        Return Not IsDBNull(r("ID_DIAGNOSTICO")) AndAlso Not IsDBNull(r("PDF_BYTES")) AndAlso CDec(r("PDF_BYTES")) > 0
    End Function

    Private Shared Function ArquivoEmUso(pCaminho As String) As Boolean
        If Not File.Exists(pCaminho) Then Return False
        Try
            Using fs As New FileStream(pCaminho, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            End Using
            Return False
        Catch ex As IOException
            Return True
        End Try
    End Function

    Private Shared Function PastaRelatorios() As String
        Dim p As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
        Return If(String.IsNullOrWhiteSpace(p), "C:\GBS\Reports", p)
    End Function

    Private Shared Function LogoRelatorios() As String
        Return If(System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath"), "")
    End Function

    Private Shared Function Texto(r As DataRow, pCol As String) As String
        If r Is Nothing OrElse Not r.Table.Columns.Contains(pCol) OrElse IsDBNull(r(pCol)) Then Return ""
        Return r(pCol).ToString().Trim()
    End Function

#End Region

End Class
