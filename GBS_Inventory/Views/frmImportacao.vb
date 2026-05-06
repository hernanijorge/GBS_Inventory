Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Data
Imports System.Text
Imports System.Net
Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Tela de Importação de Planilha Excel da GBS.
''' </summary>
Public Class frmImportacao

    Private oController          As ImportacaoController
    Private sArquivoSelecionado  As String    = ""
    Private _dtResultado         As DataTable = Nothing
    Private _chkShowIssuesOnly   As CheckBox
    Private _btnGerarRelatorio   As Button
    Private _btnEnviarRelatorio  As Button
    Private _dgvResultado        As DataGridView
    Private _lblResultadoFooter  As Label
    Private _lblAnaliseTitle     As Label

    Public Sub New()
        InitializeComponent()
        oController = New ImportacaoController()
        TemaEscuro.aplicarHelius(Me)
        InicializarGridAnalise()
    End Sub

    Private Sub btnSelecionar_Click(sender As Object, e As EventArgs) Handles btnSelecionar.Click

        Using ofd As New OpenFileDialog()
            ofd.Filter = "Planilhas Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*"
            ofd.Title  = "Selecione a planilha da GBS"

            If ofd.ShowDialog() = DialogResult.OK Then
                sArquivoSelecionado = ofd.FileName
                lblArquivo.Text     = Path.GetFileName(ofd.FileName)
                lblArquivo.ForeColor = TemaEscuro.Accent
                btnImportar.Enabled = True
            End If
        End Using

    End Sub

    Private Sub btnImportar_Click(sender As Object, e As EventArgs) Handles btnImportar.Click

        If String.IsNullOrEmpty(sArquivoSelecionado) Then
            MessageBox.Show("Select a spreadsheet first.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        btnImportar.Enabled   = False
        btnSelecionar.Enabled = False
        btnFechar.Enabled     = False

        lblStatus.Text        = "Importing..."
        lblStatus.ForeColor   = TemaEscuro.Accent
        progress.Value        = 0
        progress.Visible      = True

        Application.DoEvents()

        Try

            Dim callback As Action(Of Integer, Integer, String) = _
                Sub(atual, total, uid)
                    Me.Invoke(Sub()
                        progress.Maximum = total
                        progress.Value   = Math.Min(atual, total)
                        lblStatus.Text   = $"Processing {atual}/{total} — UID {uid}"
                        Application.DoEvents()
                    End Sub)
                End Sub

            oController.importarPlanilha(sArquivoSelecionado, callback)

            progress.Value     = progress.Maximum
            lblStatus.Text     = "Done!"
            lblStatus.ForeColor= TemaEscuro.Accent

            lblInseridos.Text   = "Inserted: "     & oController.TotalInseridos.ToString()
            lblAtualizados.Text = "Updated: "      & oController.TotalAtualizados.ToString()
            lblAbas.Text        = "Sheets processed: " & oController.TotalAbas.ToString()
            lblLinhas.Text      = "Lines read: "   & oController.TotalLinhasPlanilha.ToString()
            lblErros.Text       = "Errors: "       & oController.TotalErros.ToString()

            lblSkipped.Text = "Skipped: " & oController.TotalSkipped.ToString()

            If oController.TotalErros > 0 Then
                txtErros.Lines     = oController.Erros.ToArray()
                txtErros.Visible   = True
                lblErros.ForeColor = TemaEscuro.Vermelho
            End If

            CarregarGridAnalise()

            MessageBox.Show($"Import completed!{vbCrLf & vbCrLf}" &
                            $"Inserted: {oController.TotalInseridos}{vbCrLf}" &
                            $"Updated: {oController.TotalAtualizados}{vbCrLf}" &
                            $"Errors: {oController.TotalErros}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show("Import error: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblStatus.Text     = "Import failed"
            lblStatus.ForeColor= TemaEscuro.Vermelho

        Finally

            btnImportar.Enabled   = True
            btnSelecionar.Enabled = True
            btnFechar.Enabled     = True

        End Try

    End Sub

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub

#Region "Grid de Análise"

    Private Sub InicializarGridAnalise()

        _lblAnaliseTitle           = New Label()
        _lblAnaliseTitle.Text      = "Import Analysis"
        _lblAnaliseTitle.Font      = New Font("Segoe UI", 10, FontStyle.Bold)
        _lblAnaliseTitle.ForeColor = TemaEscuro.Accent
        _lblAnaliseTitle.Location  = New Point(20, 574)
        _lblAnaliseTitle.AutoSize  = True
        _lblAnaliseTitle.Visible   = False

        _chkShowIssuesOnly           = New CheckBox()
        _chkShowIssuesOnly.Text      = "Show issues only"
        _chkShowIssuesOnly.Location  = New Point(185, 576)
        _chkShowIssuesOnly.AutoSize  = True
        _chkShowIssuesOnly.ForeColor = TemaEscuro.Texto
        _chkShowIssuesOnly.BackColor = Color.Transparent
        _chkShowIssuesOnly.Visible   = False
        AddHandler _chkShowIssuesOnly.CheckedChanged, AddressOf OnFiltroChanged

        _btnGerarRelatorio          = New Button()
        _btnGerarRelatorio.Text     = "Generate Report"
        _btnGerarRelatorio.Location = New Point(330, 570)
        _btnGerarRelatorio.Size     = New Size(165, 28)
        _btnGerarRelatorio.Visible  = False
        TemaEscuro.estilizarBotaoHelius(_btnGerarRelatorio)
        AddHandler _btnGerarRelatorio.Click, AddressOf OnGerarRelatorioImportacao

        _btnEnviarRelatorio          = New Button()
        _btnEnviarRelatorio.Text     = "Send Report"
        _btnEnviarRelatorio.Location = New Point(505, 570)
        _btnEnviarRelatorio.Size     = New Size(175, 28)
        _btnEnviarRelatorio.Visible  = False
        TemaEscuro.estilizarBotaoHelius(_btnEnviarRelatorio)
        AddHandler _btnEnviarRelatorio.Click, AddressOf OnEnviarRelatorioImportacao

        _dgvResultado          = New DataGridView()
        _dgvResultado.Location = New Point(20, 600)
        _dgvResultado.Size     = New Size(660, 220)
        _dgvResultado.Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or
                                  AnchorStyles.Right Or AnchorStyles.Bottom
        _dgvResultado.Visible  = False
        TemaEscuro.estilizarGridHelius(_dgvResultado)
        _dgvResultado.ReadOnly = False
        _dgvResultado.EditMode = DataGridViewEditMode.EditOnEnter
        AddHandler _dgvResultado.CellFormatting, AddressOf OnGridCellFormatting
        AddHandler _dgvResultado.CellValueChanged, AddressOf OnGridCellValueChanged
        AddHandler _dgvResultado.CurrentCellDirtyStateChanged, AddressOf OnGridCurrentCellDirtyStateChanged

        _lblResultadoFooter           = New Label()
        _lblResultadoFooter.Text      = ""
        _lblResultadoFooter.Location  = New Point(20, 828)
        _lblResultadoFooter.Size      = New Size(660, 20)
        _lblResultadoFooter.Anchor    = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        _lblResultadoFooter.ForeColor = TemaEscuro.TextoMutado
        _lblResultadoFooter.Visible   = False

        Me.Controls.AddRange(New Control() {
            _lblAnaliseTitle, _chkShowIssuesOnly, _btnGerarRelatorio,
            _btnEnviarRelatorio, _dgvResultado, _lblResultadoFooter
        })

    End Sub

    Private Sub CarregarGridAnalise()
        If oController.UIDsImportados.Count = 0 Then Return

        Try
            Dim cs As String = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim w  As String = BuildInClause(oController.UIDsImportados)

            Dim sqlFull As String =
                "SELECT INTERNAL_UID, MARCA AS MANUFACTURER, MODEL, SERIAL_NUMBER," &
                "       CONDITION_STATUS, STATUS, SOURCE_BATCH, OBSERVACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & w & " ORDER BY SOURCE_BATCH, INTERNAL_UID"

            Dim sqlNoExt As String =
                "SELECT INTERNAL_UID, MARCA AS MANUFACTURER, MODEL, SERIAL_NUMBER," &
                "       CAST(NULL AS VARCHAR2(20)) AS CONDITION_STATUS," &
                "       STATUS, CAST(NULL AS VARCHAR2(100)) AS SOURCE_BATCH, OBSERVACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & w & " ORDER BY INTERNAL_UID"

            Dim sqlEn As String =
                "SELECT INTERNAL_UID, MANUFACTURER, MODEL, SERIAL_NUMBER," &
                "       CAST(NULL AS VARCHAR2(20)) AS CONDITION_STATUS," &
                "       STATUS, CAST(NULL AS VARCHAR2(100)) AS SOURCE_BATCH," &
                "       CAST(NULL AS VARCHAR2(4000)) AS OBSERVACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & w & " ORDER BY INTERNAL_UID"

            Dim dtSrc As DataTable = Nothing
            Dim lastEx As Exception = Nothing
            For Each sql As String In New String() {sqlFull, sqlNoExt, sqlEn}
                Try
                    Dim ds As DataSet = OracleHelper.ExecuteDataset(cs, CommandType.Text, sql)
                    If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                        dtSrc = ds.Tables(0)
                        Exit For
                    End If
                Catch ex As Exception
                    lastEx = ex
                    If Not ex.Message.Contains("ORA-00904") Then Throw
                End Try
            Next
            If dtSrc Is Nothing Then
                If lastEx IsNot Nothing Then Throw lastEx
                Return
            End If

            ' Build result DataTable with ISSUE column
            _dtResultado = New DataTable()
            _dtResultado.Columns.Add("INTERNAL_UID",    GetType(String))
            _dtResultado.Columns.Add("MANUFACTURER",    GetType(String))
            _dtResultado.Columns.Add("MODEL",           GetType(String))
            _dtResultado.Columns.Add("SERIAL_NUMBER",   GetType(String))
            _dtResultado.Columns.Add("CONDITION_STATUS",GetType(String))
            _dtResultado.Columns.Add("STATUS",          GetType(String))
            _dtResultado.Columns.Add("SOURCE_BATCH",    GetType(String))
            _dtResultado.Columns.Add("OBSERVATION",     GetType(String))
            _dtResultado.Columns.Add("PROBLEM",         GetType(Boolean))
            _dtResultado.Columns.Add("ISSUE",           GetType(String))
            _dtResultado.Columns.Add("ISSUE_REASON",    GetType(String))
            _dtResultado.Columns.Add("_FLAG",           GetType(Boolean))

            For Each row As DataRow In dtSrc.Rows
                Dim cond As String = ColStr(row, "CONDITION_STATUS").ToUpperInvariant().Trim()
                Dim obs  As String = ColStr(row, "OBSERVACAO")
                Dim sts  As String = ColStr(row, "STATUS").ToUpperInvariant().Trim()
                Dim issueReason As String = MontarMotivoProblema(cond, obs, sts)
                Dim flag As Boolean = Not String.IsNullOrWhiteSpace(issueReason)
                _dtResultado.Rows.Add(
                    ColStr(row, "INTERNAL_UID"),
                    ColStr(row, "MANUFACTURER"),
                    ColStr(row, "MODEL"),
                    ColStr(row, "SERIAL_NUMBER"),
                    cond,
                    sts,
                    ColStr(row, "SOURCE_BATCH"),
                    obs,
                    flag,
                    If(flag, "Yes", "No"),
                    issueReason,
                    flag
                )
            Next

            AplicarFiltro()

            _lblAnaliseTitle.Visible  = True
            _chkShowIssuesOnly.Visible = True
            _btnGerarRelatorio.Visible = True
            _btnEnviarRelatorio.Visible = True
            _dgvResultado.Visible     = True
            _lblResultadoFooter.Visible = True

        Catch ex As Exception
            MessageBox.Show("Error loading analysis grid: " & ex.Message, "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub AplicarFiltro()
        If _dtResultado Is Nothing Then Return

        Dim view As New DataView(_dtResultado)
        If _chkShowIssuesOnly IsNot Nothing AndAlso _chkShowIssuesOnly.Checked Then
            view.RowFilter = "PROBLEM = True"
        End If

        _dgvResultado.DataSource = view
        ConfigurarGridAnalise()
        AtualizarFooterAnalise()
    End Sub

    Private Sub ConfigurarGridAnalise()
        If _dgvResultado Is Nothing OrElse _dgvResultado.Columns.Count = 0 Then Return

        For Each col As DataGridViewColumn In _dgvResultado.Columns
            col.ReadOnly = True
        Next

        If _dgvResultado.Columns.Contains("_FLAG") Then
            _dgvResultado.Columns("_FLAG").Visible = False
        End If

        If _dgvResultado.Columns.Contains("PROBLEM") Then
            With _dgvResultado.Columns("PROBLEM")
                .HeaderText = "Problem?"
                .ReadOnly = False
                .Width = 70
            End With
        End If
        If _dgvResultado.Columns.Contains("ISSUE") Then _dgvResultado.Columns("ISSUE").Width = 55
        If _dgvResultado.Columns.Contains("INTERNAL_UID") Then _dgvResultado.Columns("INTERNAL_UID").Width = 90
        If _dgvResultado.Columns.Contains("SERIAL_NUMBER") Then _dgvResultado.Columns("SERIAL_NUMBER").Width = 105
        If _dgvResultado.Columns.Contains("CONDITION_STATUS") Then _dgvResultado.Columns("CONDITION_STATUS").HeaderText = "BATTERY CONDITION"
        If _dgvResultado.Columns.Contains("ISSUE_REASON") Then _dgvResultado.Columns("ISSUE_REASON").HeaderText = "ISSUE REASON"
    End Sub

    Private Sub AtualizarFooterAnalise()
        If _dtResultado Is Nothing Then Return

        Dim total  As Integer = _dtResultado.Rows.Count
        Dim issues As Integer = 0
        For Each r As DataRow In _dtResultado.Rows
            If CBool(r("PROBLEM")) Then issues += 1
        Next
        Dim ok As Integer = total - issues
        _lblResultadoFooter.Text = $"Total: {total}  |  OK / According to spreadsheet: {ok}  |  Issues: {issues}"
        _lblResultadoFooter.ForeColor = If(issues > 0, TemaEscuro.Vermelho, TemaEscuro.Verde)
    End Sub

    Private Sub OnFiltroChanged(sender As Object, e As EventArgs)
        AplicarFiltro()
    End Sub

    Private Sub OnGridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 Then Return
        If _dgvResultado Is Nothing OrElse _dgvResultado.DataSource Is Nothing Then Return
        Dim rowView As DataRowView = TryCast(_dgvResultado.Rows(e.RowIndex).DataBoundItem, DataRowView)
        If rowView Is Nothing OrElse Not rowView.Row.Table.Columns.Contains("PROBLEM") Then Return
        Try
            If CBool(rowView.Row("PROBLEM")) Then
                e.CellStyle.BackColor          = Color.FromArgb(255, 204, 204)
                e.CellStyle.ForeColor          = Color.FromArgb(140, 20, 20)
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 160, 160)
                e.CellStyle.SelectionForeColor = Color.FromArgb(80, 0, 0)
            End If
        Catch
        End Try
    End Sub

    Private Sub OnGridCurrentCellDirtyStateChanged(sender As Object, e As EventArgs)
        If _dgvResultado Is Nothing OrElse Not _dgvResultado.IsCurrentCellDirty Then Return
        If TypeOf _dgvResultado.CurrentCell Is DataGridViewCheckBoxCell Then
            _dgvResultado.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub OnGridCellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 OrElse _dgvResultado Is Nothing Then Return
        If Not _dgvResultado.Columns.Contains("PROBLEM") Then Return
        If _dgvResultado.Columns(e.ColumnIndex).Name <> "PROBLEM" Then Return

        Dim rowView As DataRowView = TryCast(_dgvResultado.Rows(e.RowIndex).DataBoundItem, DataRowView)
        If rowView Is Nothing Then Return

        Dim problem As Boolean = False
        If Not IsDBNull(rowView.Row("PROBLEM")) Then problem = CBool(rowView.Row("PROBLEM"))
        rowView.Row("ISSUE") = If(problem, "Yes", "No")
        rowView.Row("_FLAG") = problem
        AtualizarFooterAnalise()
        _dgvResultado.InvalidateRow(e.RowIndex)
    End Sub

    Private Sub OnGerarRelatorioImportacao(sender As Object, e As EventArgs)
        GerarRelatorioQualidade(False)
    End Sub

    Private Sub OnEnviarRelatorioImportacao(sender As Object, e As EventArgs)
        GerarRelatorioQualidade(True)
    End Sub

    Private Sub GerarRelatorioQualidade(enviarEmail As Boolean)
        If _dtResultado Is Nothing OrElse _dtResultado.Rows.Count = 0 Then
            MessageBox.Show("Run an import first.", "Import Report",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            Dim outputPath As String = ConfigurationManager.AppSettings("ReportsOutputPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"
            If Not Directory.Exists(outputPath) Then Directory.CreateDirectory(outputPath)

            Dim baseNome As String = "ImportQuality_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim caminhoDoc As String = Path.Combine(outputPath, baseNome & ".doc")
            File.WriteAllText(caminhoDoc, MontarHtmlRelatorioQualidade(), Encoding.UTF8)

            If enviarEmail Then
                Dim itens As New List(Of DataRow)()
                For Each row As DataRow In _dtResultado.Rows
                    itens.Add(row)
                Next

                Dim frmEmail As New frmEnviarRelatorio(caminhoDoc, "", "", itens)
                frmEmail.txtSubject.Text = "Import Quality Report - " & DateTime.Now.ToString("yyyy-MM-dd")
                frmEmail.txtBody.Text = MontarResumoEmailQualidade()
                frmEmail.ShowDialog(Me)
            Else
                MessageBox.Show("Import report generated successfully!" & vbCrLf & vbCrLf &
                                caminhoDoc,
                                "Import Report", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Error generating import report: " & ex.Message, "Import Report",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function MontarHtmlRelatorioQualidade() As String
        Dim total As Integer = _dtResultado.Rows.Count
        Dim issues As Integer = ContarProblemas()
        Dim ok As Integer = total - issues
        Dim issueRate As Decimal = If(total = 0, 0D, Math.Round((issues * 100D) / total, 1))

        Dim sb As New StringBuilder()
        sb.AppendLine("<html><head><meta charset='utf-8' />")
        sb.AppendLine("<style>")
        sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;font-size:10pt;color:#1f2937;margin:36px;}")
        sb.AppendLine("h2{margin:0 0 4px 0;color:#111827;}")
        sb.AppendLine(".sub{color:#6b7280;margin-bottom:18px;}")
        sb.AppendLine(".cards{display:flex;gap:10px;margin:16px 0 18px 0;}")
        sb.AppendLine(".card{border:1px solid #d1d5db;padding:10px 14px;min-width:130px;}")
        sb.AppendLine(".label{color:#6b7280;font-size:9pt;}.value{font-size:18pt;font-weight:700;}")
        sb.AppendLine("table{width:100%;border-collapse:collapse;table-layout:fixed;}")
        sb.AppendLine("th{background:#374151;color:#fff;text-align:left;padding:7px 6px;font-size:8pt;}")
        sb.AppendLine("td{border:1px solid #e5e7eb;padding:5px 6px;font-size:8pt;word-wrap:break-word;vertical-align:top;}")
        sb.AppendLine(".bad td{background:#fee2e2;color:#7f1d1d;}")
        sb.AppendLine("</style></head><body>")

        AppendLogoHtml(sb)
        sb.AppendLine("<h2>GBS Import Quality Report</h2>")
        sb.AppendLine("<div class='sub'>Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                      " | Source file: " & WebUtility.HtmlEncode(Path.GetFileName(sArquivoSelecionado)) & "</div>")
        sb.AppendLine("<div class='cards'>")
        sb.AppendLine("<div class='card'><div class='label'>Total imported</div><div class='value'>" & total & "</div></div>")
        sb.AppendLine("<div class='card'><div class='label'>OK / According to spreadsheet</div><div class='value'>" & ok & "</div></div>")
        sb.AppendLine("<div class='card'><div class='label'>With issue</div><div class='value'>" & issues & "</div></div>")
        sb.AppendLine("<div class='card'><div class='label'>Issue rate</div><div class='value'>" & issueRate.ToString("0.0") & "%</div></div>")
        sb.AppendLine("</div>")

        sb.AppendLine("<table>")
        sb.AppendLine("<tr><th>Internal UID</th><th>Manufacturer</th><th>Model</th><th>Serial Number</th><th>Battery Condition</th><th>Status</th><th>Batch</th><th>Problem?</th><th>Issue Reason</th><th>Observation</th></tr>")
        For Each row As DataRow In _dtResultado.Rows
            Dim problem As Boolean = CBool(row("PROBLEM"))
            sb.Append(If(problem, "<tr class='bad'>", "<tr>"))
            AppendTd(sb, RowStr(row, "INTERNAL_UID"))
            AppendTd(sb, RowStr(row, "MANUFACTURER"))
            AppendTd(sb, RowStr(row, "MODEL"))
            AppendTd(sb, RowStr(row, "SERIAL_NUMBER"))
            AppendTd(sb, RowStr(row, "CONDITION_STATUS"))
            AppendTd(sb, RowStr(row, "STATUS"))
            AppendTd(sb, RowStr(row, "SOURCE_BATCH"))
            AppendTd(sb, If(problem, "Yes", "No"))
            AppendTd(sb, RowStr(row, "ISSUE_REASON"))
            AppendTd(sb, RowStr(row, "OBSERVATION"))
            sb.AppendLine("</tr>")
        Next
        sb.AppendLine("</table>")
        sb.AppendLine("</body></html>")
        Return sb.ToString()
    End Function

    Private Function MontarResumoEmailQualidade() As String
        Dim total As Integer = _dtResultado.Rows.Count
        Dim issues As Integer = ContarProblemas()
        Dim ok As Integer = total - issues
        Dim issueRate As Decimal = If(total = 0, 0D, Math.Round((issues * 100D) / total, 1))

        Dim sb As New StringBuilder()
        sb.AppendLine("Please find attached the import quality report for the latest spreadsheet.")
        sb.AppendLine()
        sb.AppendLine("Summary:")
        sb.AppendLine("  - Source file: " & Path.GetFileName(sArquivoSelecionado))
        sb.AppendLine("  - Total imported: " & total.ToString())
        sb.AppendLine("  - OK / According to spreadsheet: " & ok.ToString())
        sb.AppendLine("  - With issue: " & issues.ToString())
        sb.AppendLine("  - Issue rate: " & issueRate.ToString("0.0") & "%")
        Return sb.ToString()
    End Function

    Private Function ContarProblemas() As Integer
        Dim total As Integer = 0
        If _dtResultado Is Nothing Then Return 0
        For Each row As DataRow In _dtResultado.Rows
            If CBool(row("PROBLEM")) Then total += 1
        Next
        Return total
    End Function

    Private Function MontarMotivoProblema(conditionStatus As String, observacao As String, status As String) As String
        Dim motivos As New List(Of String)()
        If conditionStatus = "FAIR" OrElse conditionStatus = "POOR" Then
            motivos.Add("Battery condition: " & conditionStatus)
        End If
        If Not String.IsNullOrWhiteSpace(observacao) Then
            motivos.Add("Observation")
        End If
        If status = "IN_REPAIR" Then
            motivos.Add("Status: IN_REPAIR")
        End If
        Return String.Join("; ", motivos)
    End Function

    Private Sub AppendLogoHtml(sb As StringBuilder)
        Dim logoPath As String = ConfigurationManager.AppSettings("InvoiceLogoPath")
        If String.IsNullOrWhiteSpace(logoPath) Then Return
        If Not File.Exists(logoPath) Then Return

        Try
            Dim b64 As String = Convert.ToBase64String(File.ReadAllBytes(logoPath))
            Dim mime As String = If(Path.GetExtension(logoPath).ToLowerInvariant() = ".png", "image/png", "image/jpeg")
            sb.AppendLine("<div style='margin-bottom:14px;'><img src='data:" & mime & ";base64," &
                          b64 & "' style='max-height:60px;max-width:180px;' /></div>")
        Catch
        End Try
    End Sub

    Private Sub AppendTd(sb As StringBuilder, valor As String)
        sb.Append("<td>").Append(WebUtility.HtmlEncode(valor)).Append("</td>")
    End Sub

    Private Function RowStr(row As DataRow, colName As String) As String
        If Not row.Table.Columns.Contains(colName) Then Return ""
        If IsDBNull(row(colName)) Then Return ""
        Return row(colName).ToString()
    End Function

    Private Function ColStr(row As DataRow, colName As String) As String
        If Not row.Table.Columns.Contains(colName) Then Return ""
        If IsDBNull(row(colName)) Then Return ""
        Return row(colName).ToString()
    End Function

    Private Function BuildInClause(uids As List(Of String)) As String
        If uids.Count = 0 Then Return "1=0"
        Dim sb As New System.Text.StringBuilder()
        Dim i As Integer = 0
        While i < uids.Count
            If i > 0 Then sb.Append(" OR ")
            sb.Append("INTERNAL_UID IN (")
            Dim first As Boolean = True
            Dim limit As Integer = Math.Min(i + 999, uids.Count)
            While i < limit
                If Not first Then sb.Append(",")
                sb.Append("'").Append(uids(i).Replace("'", "''")).Append("'")
                first = False
                i += 1
            End While
            sb.Append(")")
        End While
        Return sb.ToString()
    End Function

#End Region

End Class
