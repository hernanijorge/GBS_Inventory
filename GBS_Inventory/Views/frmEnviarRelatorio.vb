Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.IO
Imports System.Net

Public Class frmEnviarRelatorio

    Private _caminhoDoc  As String
    Private _caminhoPdf  As String
    Private _caminhoXlsx As String
    Private _itens       As List(Of DataRow)

    Public Sub New(caminhoDoc  As String,
                   caminhoPdf  As String,
                   caminhoXlsx As String,
                   itens       As List(Of DataRow))

        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)

        _caminhoDoc  = caminhoDoc
        _caminhoPdf  = caminhoPdf
        _caminhoXlsx = caminhoXlsx
        _itens       = itens

        PreencherCaminhos()
        PreencherCampos()

    End Sub

    ' ── Fill file path labels and enable/disable Open buttons ─────────────────

    Private Sub PreencherCaminhos()

        lblWordPath.Text  = If(Not String.IsNullOrEmpty(_caminhoDoc),
                                Path.GetFileName(_caminhoDoc), "(not generated)")
        lblPdfPath.Text   = If(Not String.IsNullOrEmpty(_caminhoPdf),
                                Path.GetFileName(_caminhoPdf), "(not generated)")
        lblExcelPath.Text = If(Not String.IsNullOrEmpty(_caminhoXlsx),
                                Path.GetFileName(_caminhoXlsx), "(not generated)")

        btnOpenWord.Enabled  = Not String.IsNullOrEmpty(_caminhoDoc)  AndAlso File.Exists(_caminhoDoc)
        btnOpenPdf.Enabled   = Not String.IsNullOrEmpty(_caminhoPdf)  AndAlso File.Exists(_caminhoPdf)
        btnOpenExcel.Enabled = Not String.IsNullOrEmpty(_caminhoXlsx) AndAlso File.Exists(_caminhoXlsx)

        chkAttachWord.Enabled  = btnOpenWord.Enabled
        chkAttachPdf.Enabled   = btnOpenPdf.Enabled
        chkAttachExcel.Enabled = btnOpenExcel.Enabled

        If Not chkAttachWord.Enabled  Then chkAttachWord.Checked  = False
        If Not chkAttachPdf.Enabled   Then chkAttachPdf.Checked   = False
        If Not chkAttachExcel.Enabled Then chkAttachExcel.Checked = False

    End Sub

    ' ── Pre-fill subject and body from report data ────────────────────────────

    Private Sub PreencherCampos()

        Dim data  As String  = DateTime.Now.ToString("yyyy-MM-dd")
        Dim total As Integer = _itens.Count

        txtSubject.Text = $"GBS Inventory Report - {data} - {total} items"

        Dim mfrs As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim sts  As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each row As DataRow In _itens
            For Each col As String In {"MARCA", "MANUFACTURER"}
                If row.Table.Columns.Contains(col) AndAlso Not IsDBNull(row(col)) Then
                    Dim v As String = row(col).ToString().Trim()
                    If v.Length > 0 Then mfrs.Add(v) : Exit For
                End If
            Next
            If row.Table.Columns.Contains("STATUS") AndAlso Not IsDBNull(row("STATUS")) Then
                Dim v As String = row("STATUS").ToString().Trim()
                If v.Length > 0 Then sts.Add(v)
            End If
        Next

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine($"Please find attached the inventory report generated on {data}.")
        sb.AppendLine()
        sb.AppendLine("Summary:")
        sb.AppendLine($"  - Total items: {total}")
        sb.AppendLine($"  - Manufacturers: {String.Join(", ", mfrs)}")
        sb.AppendLine($"  - Status: {String.Join(", ", sts)}")

        txtBody.Text = sb.ToString()

    End Sub

    ' ── Open file buttons ─────────────────────────────────────────────────────

    Private Sub btnOpenWord_Click(sender As Object, e As EventArgs) Handles btnOpenWord.Click
        If File.Exists(_caminhoDoc)  Then System.Diagnostics.Process.Start(_caminhoDoc)
    End Sub

    Private Sub btnOpenPdf_Click(sender As Object, e As EventArgs) Handles btnOpenPdf.Click
        If File.Exists(_caminhoPdf)  Then System.Diagnostics.Process.Start(_caminhoPdf)
    End Sub

    Private Sub btnOpenExcel_Click(sender As Object, e As EventArgs) Handles btnOpenExcel.Click
        If File.Exists(_caminhoXlsx) Then System.Diagnostics.Process.Start(_caminhoXlsx)
    End Sub

    ' ── Send Email ────────────────────────────────────────────────────────────

    Private Sub btnSendEmail_Click(sender As Object, e As EventArgs) Handles btnSendEmail.Click

        ' Validate TO — must contain @ and .
        Dim destino As String = txtTo.Text.Trim()
        If String.IsNullOrWhiteSpace(destino) OrElse
           Not destino.Contains("@") OrElse
           Not destino.Contains(".") Then
            MessageBox.Show("Please enter a valid email address in the To field.", "Validation",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTo.Focus()
            Return
        End If

        ' Validate CC — optional; warn if filled with invalid text but never block
        Dim ccEfetivo As String = txtCC.Text.Trim()
        If Not String.IsNullOrWhiteSpace(ccEfetivo) AndAlso
           Not (ccEfetivo.Contains("@") AndAlso ccEfetivo.Contains(".")) Then
            Dim res As DialogResult = MessageBox.Show(
                "CC field has an invalid email — it will be ignored. Continue?",
                "CC Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If res = DialogResult.No Then Return
            ccEfetivo = ""
        End If

        Dim anexos As New List(Of String)()
        If chkAttachWord.Checked  AndAlso File.Exists(_caminhoDoc)  Then anexos.Add(_caminhoDoc)
        If chkAttachPdf.Checked   AndAlso File.Exists(_caminhoPdf)  Then anexos.Add(_caminhoPdf)
        If chkAttachExcel.Checked AndAlso File.Exists(_caminhoXlsx) Then anexos.Add(_caminhoXlsx)

        If anexos.Count = 0 Then
            MessageBox.Show("Please select at least one attachment.", "Validation",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Build HTML body. The email service adds the logo as CID inline for Gmail/Outlook compatibility.
        Dim corpoFinal As String =
            "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:11pt;color:#1f2937;'>" &
            "<p>" & WebUtility.HtmlEncode(txtBody.Text).Replace(vbCrLf, "<br/>").Replace(vbLf, "<br/>") & "</p>" &
            "</body></html>"
        Dim isHtml As Boolean = True

        Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
        Try
            Me.Cursor            = Cursors.WaitCursor
            btnSendEmail.Enabled = False
            btnClose.Enabled     = False

            InvoiceEmailService.EnviarInvoice(
                destino,
                txtSubject.Text.Trim(),
                corpoFinal,
                anexos,
                ccEfetivo,
                isHtml,
                logoPath,
                True)

            Me.Cursor = Cursors.Default
            MessageBox.Show($"Email sent to {destino} with {anexos.Count} attachment(s).",
                            "Email Sent", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()

        Catch ex As Exception
            Me.Cursor            = Cursors.Default
            btnSendEmail.Enabled = True
            btnClose.Enabled     = True
            MessageBox.Show("Error sending email: " & ex.Message,
                            "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    ' ── Close ─────────────────────────────────────────────────────────────────

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

End Class
