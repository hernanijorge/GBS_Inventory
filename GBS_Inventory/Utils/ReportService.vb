Imports System.Data
Imports System.IO
Imports System.Text
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

''' <summary>
''' Gera relatório de inventário em Word (.doc HTML) com logo, cabeçalho e tabela formatada.
''' </summary>
Public Class ReportService

    Private Const PastaDefault As String = "C:\GBS\Reports\"

    Public Shared Function GerarRelatorio(pItens As List(Of DataRow),
                                           pOutputPath As String,
                                           pLogoPath As String) As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nomeArq As String = "Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".doc"
        Dim caminhoFinal As String = Path.Combine(pasta, nomeArq)

        Dim logoEfetivo As String = ObterLogoPath(pLogoPath)
        Dim html As String = MontarHtml(pItens, logoEfetivo)
        File.WriteAllText(caminhoFinal, html, Encoding.UTF8)

        Return caminhoFinal

    End Function

    ' ── Excel export ─────────────────────────────────────────────────────────

    Public Shared Function GerarExcel(pItens As List(Of DataRow),
                                       pOutputPath As String) As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nomeArq As String = "Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
        Dim caminhoFinal As String = Path.Combine(pasta, nomeArq)

        ' Colunas: chave interna, header, largura em caracteres
        Dim colunas As (Key As String, Header As String, Largura As Double)() = {
            ("INTERNAL_UID",  "Internal UID",  16),
            ("MARCA",         "Manufacturer",  18),
            ("MODEL",         "Model",         16),
            ("SERIAL_NUMBER", "Serial Number", 20),
            ("PROCESSADOR",   "Processor",     20),
            ("RAM_GB",        "RAM GB",        10),
            ("STORAGE_GB",    "Storage GB",    12),
            ("STATUS",        "Status",        14),
            ("OBSERVACAO",    "Notes",         40)
        }

        ' Resolve fallbacks de nome de coluna
        Dim colsReais As New Dictionary(Of String, String)
        For Each c In colunas
            colsReais(c.Key) = c.Key
        Next

        If pItens.Count > 0 Then
            Dim tbl As DataTable = pItens(0).Table
            If Not tbl.Columns.Contains("MARCA")       AndAlso tbl.Columns.Contains("MANUFACTURER") Then colsReais("MARCA")       = "MANUFACTURER"
            If Not tbl.Columns.Contains("MODEL")       AndAlso tbl.Columns.Contains("MODELO")       Then colsReais("MODEL")       = "MODELO"
            If Not tbl.Columns.Contains("PROCESSADOR") AndAlso tbl.Columns.Contains("CPU_MODEL")    Then colsReais("PROCESSADOR") = "CPU_MODEL"
            If Not tbl.Columns.Contains("OBSERVACAO")  AndAlso tbl.Columns.Contains("NOTES")        Then colsReais("OBSERVACAO")  = "NOTES"
        End If

        Using pkg As New ExcelPackage()
            Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets.Add("Inventory Report")

            ' ── Cabeçalho ─────────────────────────────────────────────
            For i As Integer = 0 To colunas.Length - 1
                Dim cell As ExcelRange = ws.Cells(1, i + 1)
                cell.Value = colunas(i).Header
                cell.Style.Font.Bold = True
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(55, 65, 81))
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            Next

            ' ── Dados ─────────────────────────────────────────────────
            Dim corAlternado As System.Drawing.Color = System.Drawing.Color.FromArgb(249, 250, 251)

            For rowIdx As Integer = 0 To pItens.Count - 1
                Dim dr As DataRow = pItens(rowIdx)
                Dim tbl As DataTable = dr.Table
                For colIdx As Integer = 0 To colunas.Length - 1
                    Dim colReal As String = colsReais(colunas(colIdx).Key)
                    Dim val As String = ""
                    If tbl.Columns.Contains(colReal) AndAlso Not IsDBNull(dr(colReal)) Then
                        val = dr(colReal).ToString()
                    End If
                    ws.Cells(rowIdx + 2, colIdx + 1).Value = val
                Next

                If rowIdx Mod 2 = 1 Then
                    ws.Cells(rowIdx + 2, 1, rowIdx + 2, colunas.Length).Style.Fill.PatternType = ExcelFillStyle.Solid
                    ws.Cells(rowIdx + 2, 1, rowIdx + 2, colunas.Length).Style.Fill.BackgroundColor.SetColor(corAlternado)
                End If
            Next

            ' ── Larguras e wrap ───────────────────────────────────────
            For i As Integer = 0 To colunas.Length - 1
                ws.Column(i + 1).Width = colunas(i).Largura
            Next
            ws.Column(colunas.Length).Style.WrapText = True

            ' ── Freeze linha de cabeçalho ─────────────────────────────
            ws.View.FreezePanes(2, 1)

            pkg.SaveAs(New FileInfo(caminhoFinal))
        End Using

        Return caminhoFinal

    End Function

    ' ── HTML builder ─────────────────────────────────────────────────────────

    Private Shared Function MontarHtml(pItens As List(Of DataRow), pLogoPath As String) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("<html>")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset='utf-8' />")
        sb.AppendLine("<style>")
        sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;font-size:10pt;color:#1f2937;margin:36px;}")
        sb.AppendLine("h2{color:#1f2937;margin-bottom:4px;}")
        sb.AppendLine(".sub{font-size:9pt;color:#6b7280;margin-bottom:22px;}")
        sb.AppendLine("table{table-layout:fixed;width:100%;border-collapse:collapse;margin-top:10px;}")
        sb.AppendLine("th{background:#374151;color:#ffffff;padding:7px 6px;text-align:left;font-size:8.5pt;overflow:hidden;}")
        sb.AppendLine("td{border:1px solid #e5e7eb;padding:5px 6px;font-size:8.5pt;vertical-align:top;word-wrap:break-word;white-space:normal;overflow:visible;}")
        sb.AppendLine("tr:nth-child(even) td{background:#f9fafb;}")
        sb.AppendLine("</style>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")

        ' Logo em Base64
        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Dim bytes As Byte() = File.ReadAllBytes(pLogoPath)
            Dim b64 As String = Convert.ToBase64String(bytes)
            Dim mime As String = ObterMime(pLogoPath)
            sb.AppendLine("<div style='margin-bottom:14px;'>")
            sb.AppendLine("<img src='data:" & mime & ";base64," & b64 &
                          "' style='max-height:70px;max-width:220px;' />")
            sb.AppendLine("</div>")
        End If

        ' Título e meta
        sb.AppendLine("<h2>GBS Inventory Report</h2>")
        sb.AppendLine("<div class='sub'>Generated: " &
                      DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                      " &nbsp;|&nbsp; Total items: " & pItens.Count.ToString() & "</div>")

        ' Definição das colunas: chave interna, label, largura fixa
        Dim colunas As (Key As String, Label As String, Width As String)() = {
            ("INTERNAL_UID",  "Internal UID",  "80px"),
            ("MARCA",         "Manufacturer",  "90px"),
            ("MODEL",         "Model",         "70px"),
            ("SERIAL_NUMBER", "Serial Number", "90px"),
            ("PROCESSADOR",   "Processor",     "90px"),
            ("RAM_GB",        "RAM GB",        "50px"),
            ("STORAGE_GB",    "Storage GB",    "60px"),
            ("STATUS",        "Status",        "70px"),
            ("OBSERVACAO",    "Notes",        "200px")
        }

        ' Resolve nomes reais com fallback para esquemas alternativos
        Dim colsReais As New Dictionary(Of String, String)
        For Each c In colunas
            colsReais(c.Key) = c.Key
        Next

        If pItens.Count > 0 Then
            Dim tbl As DataTable = pItens(0).Table
            If Not tbl.Columns.Contains("MARCA")       AndAlso tbl.Columns.Contains("MANUFACTURER") Then colsReais("MARCA")       = "MANUFACTURER"
            If Not tbl.Columns.Contains("MODEL")       AndAlso tbl.Columns.Contains("MODELO")       Then colsReais("MODEL")       = "MODELO"
            If Not tbl.Columns.Contains("PROCESSADOR") AndAlso tbl.Columns.Contains("CPU_MODEL")    Then colsReais("PROCESSADOR") = "CPU_MODEL"
            If Not tbl.Columns.Contains("OBSERVACAO")  AndAlso tbl.Columns.Contains("NOTES")        Then colsReais("OBSERVACAO")  = "NOTES"
        End If

        ' Tabela com larguras fixas por coluna
        sb.AppendLine("<table>")

        sb.Append("<tr>")
        For Each c In colunas
            sb.Append("<th style='width:" & c.Width & "'>" & c.Label & "</th>")
        Next
        sb.AppendLine("</tr>")

        For Each row As DataRow In pItens
            Dim tbl As DataTable = row.Table
            sb.Append("<tr>")
            For Each c In colunas
                Dim colReal As String = colsReais(c.Key)
                Dim val As String = ""
                If tbl.Columns.Contains(colReal) AndAlso Not IsDBNull(row(colReal)) Then
                    val = row(colReal).ToString()
                End If
                Dim estiloNotes As String = If(c.Key = "OBSERVACAO", " style='max-width:200px;word-wrap:break-word;white-space:normal;overflow:visible;'", "")
                sb.Append("<td" & estiloNotes & ">" & HtmlEncode(val) & "</td>")
            Next
            sb.AppendLine("</tr>")
        Next

        sb.AppendLine("</table>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()

    End Function

    ' ── Helpers ──────────────────────────────────────────────────────────────

    Private Shared Function HtmlEncode(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("&", "&amp;").
                 Replace("<", "&lt;").
                 Replace(">", "&gt;").
                 Replace("""", "&quot;")
    End Function

    Private Shared Function ObterMime(pPath As String) As String
        Select Case Path.GetExtension(pPath).ToLowerInvariant()
            Case ".png"        : Return "image/png"
            Case ".jpg", ".jpeg" : Return "image/jpeg"
            Case ".gif"        : Return "image/gif"
            Case ".bmp"        : Return "image/bmp"
            Case Else          : Return "image/jpeg"
        End Select
    End Function

    Private Shared Function ObterLogoPath(pLogoPath As String) As String
        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Return pLogoPath
        End If

        For Each candidato As String In {"C:\GBS\logo.png", "C:\GBS\logo.jpg", "C:\GBS\logo.jpeg"}
            If File.Exists(candidato) Then Return candidato
        Next

        Return ""
    End Function

End Class
