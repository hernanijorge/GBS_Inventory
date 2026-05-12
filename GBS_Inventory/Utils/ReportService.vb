Imports System.Data
Imports System.IO
Imports System.Text
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports iTextSharp.text
Imports iTextSharp.text.pdf

''' <summary>
''' Gera relatório de inventário em Word (.doc HTML), PDF e Excel com logo, cabeçalho e tabela formatada.
''' </summary>
Public Class ReportService

    Private Const PastaDefault As String = "C:\GBS\Reports\"

    Public Shared Function GerarRelatorio(pItens As List(Of DataRow),
                                           pOutputPath As String,
                                           pLogoPath As String,
                                           Optional baseNome As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(baseNome),
                                 "Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), baseNome)
        Dim nomeArq As String = nome & ".doc"
        Dim caminhoFinal As String = Path.Combine(pasta, nomeArq)

        Dim logoEfetivo As String = ObterLogoPath(pLogoPath)
        Dim html As String = MontarHtml(pItens, logoEfetivo)
        File.WriteAllText(caminhoFinal, html, Encoding.UTF8)

        Return caminhoFinal

    End Function

    ' ── Excel export ─────────────────────────────────────────────────────────

    Public Shared Function GerarExcel(pItens As List(Of DataRow),
                                       pOutputPath As String,
                                       Optional pLogoPath As String = "",
                                       Optional baseNome As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(baseNome),
                                 "Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), baseNome)
        Dim nomeArq As String = nome & ".xlsx"
        Dim caminhoFinal As String = Path.Combine(pasta, nomeArq)

        ' Colunas: chave interna, header, largura em caracteres
        Dim colunas As (Key As String, Header As String, Largura As Double)() = {
            ("INTERNAL_UID",     "Internal UID",      16),
            ("MARCA",            "Manufacturer",       18),
            ("MODEL",            "Model",              16),
            ("SERIAL_NUMBER",    "Serial Number",      20),
            ("PROCESSADOR",      "Processor",          20),
            ("RAM_GB",           "RAM GB",             10),
            ("STORAGE_GB",       "Storage GB",         12),
            ("CONDITION_STATUS", "Battery Condition",  18),
            ("OBSERVACAO",       "Notes",              40)
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
            If Not tbl.Columns.Contains("CONDITION_STATUS") AndAlso tbl.Columns.Contains("BATTERY_CONDITION") Then colsReais("CONDITION_STATUS") = "BATTERY_CONDITION"
        End If

        Using pkg As New ExcelPackage()
            Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets.Add("Inventory Report")
            Dim linhaCabecalho As Integer = 5
            Dim logoEfetivo As String = ObterLogoPath(pLogoPath)

            ws.Row(1).Height = 40
            ws.Row(2).Height = 20
            ws.Row(3).Height = 20
            ws.Row(4).Height = 20

            ws.Cells(1, 4, 1, 8).Merge = True
            ws.Cells(1, 4).Value = "GBS Inventory Report"
            ws.Cells(1, 4).Style.Font.Bold = True
            ws.Cells(1, 4).Style.Font.Size = 16
            ws.Cells(1, 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center

            ws.Cells(2, 4, 2, 8).Merge = True
            ws.Cells(2, 4).Value = "Generated: " &
                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                                    " | Total items: " & pItens.Count.ToString()
            ws.Cells(2, 4).Style.Font.Size = 10
            ws.Cells(2, 4).Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(107, 114, 128))

            If Not String.IsNullOrWhiteSpace(logoEfetivo) AndAlso File.Exists(logoEfetivo) Then
                Try
                    ' Logo anchored to A1 — 180×60 px, does not overlap column D onward
                    Dim logo = ws.Drawings.AddPicture("GBS_Logo", New FileInfo(logoEfetivo))
                    logo.SetPosition(0, 4, 0, 4)
                    logo.SetSize(180, 60)
                Catch
                    ' Mantem o relatorio funcionando mesmo se a imagem nao carregar.
                End Try
            End If

            ' ── Cabeçalho ─────────────────────────────────────────────
            For i As Integer = 0 To colunas.Length - 1
                Dim cell As ExcelRange = ws.Cells(linhaCabecalho, i + 1)
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
                    ws.Cells(linhaCabecalho + rowIdx + 1, colIdx + 1).Value = val
                Next

                If rowIdx Mod 2 = 1 Then
                    ws.Cells(linhaCabecalho + rowIdx + 1, 1, linhaCabecalho + rowIdx + 1, colunas.Length).Style.Fill.PatternType = ExcelFillStyle.Solid
                    ws.Cells(linhaCabecalho + rowIdx + 1, 1, linhaCabecalho + rowIdx + 1, colunas.Length).Style.Fill.BackgroundColor.SetColor(corAlternado)
                End If
            Next

            ' ── Larguras e wrap ───────────────────────────────────────
            For i As Integer = 0 To colunas.Length - 1
                ws.Column(i + 1).Width = colunas(i).Largura
            Next
            ws.Column(colunas.Length).Style.WrapText = True

            If pItens.Count > 0 Then
                ws.Cells(linhaCabecalho, 1, linhaCabecalho + pItens.Count, colunas.Length).AutoFilter = True
            End If

            ' ── Freeze linha de cabeçalho ─────────────────────────────
            ws.View.FreezePanes(linhaCabecalho + 1, 1)

            pkg.SaveAs(New FileInfo(caminhoFinal))
        End Using

        Return caminhoFinal

    End Function

    ' ── PDF export ───────────────────────────────────────────────────────────

    Public Shared Function GerarPdf(pItens As List(Of DataRow),
                                     pOutputPath As String,
                                     pLogoPath As String,
                                     Optional baseNome As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(baseNome),
                                 "Report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), baseNome)
        Dim nomeArq As String = nome & ".pdf"
        Dim caminhoFinal As String = Path.Combine(pasta, nomeArq)

        Dim colunas As (Key As String, Header As String, Largura As Single)() = {
            ("INTERNAL_UID",     "Internal UID",      56),
            ("MARCA",            "Manufacturer",       68),
            ("MODEL",            "Model",              58),
            ("SERIAL_NUMBER",    "Serial Number",      72),
            ("PROCESSADOR",      "Processor",          80),
            ("RAM_GB",           "RAM GB",             32),
            ("STORAGE_GB",       "Storage GB",         42),
            ("CONDITION_STATUS", "Battery Condition",  68),
            ("OBSERVACAO",       "Notes",             120)
        }

        Dim colsReais As New Dictionary(Of String, String)
        For Each c In colunas
            colsReais(c.Key) = c.Key
        Next
        If pItens.Count > 0 Then
            Dim tbl As DataTable = pItens(0).Table
            If Not tbl.Columns.Contains("MARCA")            AndAlso tbl.Columns.Contains("MANUFACTURER")    Then colsReais("MARCA")            = "MANUFACTURER"
            If Not tbl.Columns.Contains("MODEL")            AndAlso tbl.Columns.Contains("MODELO")          Then colsReais("MODEL")            = "MODELO"
            If Not tbl.Columns.Contains("PROCESSADOR")      AndAlso tbl.Columns.Contains("CPU_MODEL")       Then colsReais("PROCESSADOR")      = "CPU_MODEL"
            If Not tbl.Columns.Contains("OBSERVACAO")       AndAlso tbl.Columns.Contains("NOTES")           Then colsReais("OBSERVACAO")       = "NOTES"
            If Not tbl.Columns.Contains("CONDITION_STATUS") AndAlso tbl.Columns.Contains("BATTERY_CONDITION") Then colsReais("CONDITION_STATUS") = "BATTERY_CONDITION"
        End If

        Dim doc As New Document(PageSize.A4.Rotate(), 20, 20, 30, 20)
        Using fs As New FileStream(caminhoFinal, FileMode.Create, FileAccess.Write)
            PdfWriter.GetInstance(doc, fs)
            doc.Open()

            ' Logo
            Dim logoEfetivo As String = ObterLogoPath(pLogoPath)
            If Not String.IsNullOrWhiteSpace(logoEfetivo) AndAlso File.Exists(logoEfetivo) Then
                Try
                    Dim img As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(logoEfetivo)
                    img.ScaleToFit(140, 50)
                    img.Alignment = Element.ALIGN_LEFT
                    doc.Add(img)
                Catch
                End Try
            End If

            ' Title + meta
            Dim fontTitulo As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 13, New BaseColor(31, 41, 55))
            Dim fontMeta   As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 8,  New BaseColor(107, 114, 128))
            Dim pTitulo As New Paragraph("GBS Inventory Report", fontTitulo)
            pTitulo.SpacingBefore = 4
            doc.Add(pTitulo)
            Dim pMeta As New Paragraph("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                                        "  |  Total items: " & pItens.Count.ToString(), fontMeta)
            pMeta.SpacingAfter = 8
            doc.Add(pMeta)

            ' Table
            Dim totalCols As Integer = colunas.Length
            Dim tabela As New PdfPTable(totalCols)
            tabela.WidthPercentage = 100
            Dim widths(totalCols - 1) As Single
            For i As Integer = 0 To totalCols - 1
                widths(i) = colunas(i).Largura
            Next
            tabela.SetWidths(widths)

            Dim fontHdr  As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7.5F, BaseColor.WHITE)
            Dim fontData As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 7, New BaseColor(31, 41, 55))
            Dim hdrBg As New BaseColor(55, 65, 81)
            Dim altBg As New BaseColor(249, 250, 251)

            For Each c In colunas
                Dim cell As New PdfPCell(New Phrase(c.Header, fontHdr))
                cell.BackgroundColor     = hdrBg
                cell.HorizontalAlignment = Element.ALIGN_CENTER
                cell.Padding             = 5
                tabela.AddCell(cell)
            Next

            Dim rowIdx As Integer = 0
            For Each row As DataRow In pItens
                Dim dtRow As DataTable = row.Table
                Dim bg As BaseColor = If(rowIdx Mod 2 = 1, altBg, BaseColor.WHITE)
                For Each c In colunas
                    Dim colReal As String = colsReais(c.Key)
                    Dim val As String = ""
                    If dtRow.Columns.Contains(colReal) AndAlso Not IsDBNull(row(colReal)) Then
                        val = row(colReal).ToString()
                    End If
                    Dim cell As New PdfPCell(New Phrase(val, fontData))
                    cell.BackgroundColor = bg
                    cell.Padding         = 4
                    tabela.AddCell(cell)
                Next
                rowIdx += 1
            Next

            doc.Add(tabela)
            doc.Close()
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
            ("INTERNAL_UID",     "Internal UID",      "80px"),
            ("MARCA",            "Manufacturer",       "90px"),
            ("MODEL",            "Model",              "70px"),
            ("SERIAL_NUMBER",    "Serial Number",      "90px"),
            ("PROCESSADOR",      "Processor",          "90px"),
            ("RAM_GB",           "RAM GB",             "50px"),
            ("STORAGE_GB",       "Storage GB",         "60px"),
            ("CONDITION_STATUS", "Battery Condition",  "80px"),
            ("OBSERVACAO",       "Notes",             "200px")
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
            If Not tbl.Columns.Contains("CONDITION_STATUS") AndAlso tbl.Columns.Contains("BATTERY_CONDITION") Then colsReais("CONDITION_STATUS") = "BATTERY_CONDITION"
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

    ' ── Upgrades por cliente ─────────────────────────────────────────────────
    ' Uma aba por customer. Colunas: Device | S/N | Date | Component | Before | After | Technician | Notes

    Public Shared Function GerarExcelUpgradesPorCliente(
            pRows      As List(Of DataRow),
            pOutputPath As String,
            Optional pLogoPath As String = "",
            Optional baseNome  As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(baseNome),
                                 "UpgradeReport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), baseNome)
        Dim caminhoFinal As String = Path.Combine(pasta, nome & ".xlsx")

        Dim corHeader  As System.Drawing.Color = System.Drawing.Color.FromArgb(55, 65, 81)
        Dim corZebra   As System.Drawing.Color = System.Drawing.Color.FromArgb(249, 250, 251)
        Dim logoEfetivo As String = ObterLogoPath(pLogoPath)

        Dim colunas As (Key As String, Header As String, Largura As Double)() = {
            ("INTERNAL_UID",   "Internal UID", 16),
            ("MODEL",          "Model",        16),
            ("SERIAL_NUMBER",  "Serial Number", 20),
            ("DATA_UPGRADE",   "Date",         14),
            ("COMPONENT_TYPE", "Component",    14),
            ("VALUE_BEFORE",   "Before",       12),
            ("VALUE_AFTER",    "After",        12),
            ("TECHNICIAN",     "Technician",   16),
            ("NOTES",          "Notes",        40)
        }

        ' Agrupar por CUSTOMER
        Dim grupos As New Dictionary(Of String, List(Of DataRow))(StringComparer.OrdinalIgnoreCase)
        For Each dr As DataRow In pRows
            Dim customer As String = "Unassigned"
            If dr.Table.Columns.Contains("CUSTOMER") AndAlso Not IsDBNull(dr("CUSTOMER")) Then
                customer = dr("CUSTOMER").ToString().Trim()
                If String.IsNullOrWhiteSpace(customer) Then customer = "Unassigned"
            End If
            If Not grupos.ContainsKey(customer) Then grupos(customer) = New List(Of DataRow)
            grupos(customer).Add(dr)
        Next

        Using pkg As New ExcelPackage()

            For Each kvp As KeyValuePair(Of String, List(Of DataRow)) In grupos

                ' Nome da aba: max 31 chars, sem caracteres inválidos
                Dim sheetName As String = kvp.Key
                For Each c As Char In {"/"c, "\"c, "?"c, "*"c, "["c, "]"c, ":"c}
                    sheetName = sheetName.Replace(c, "-"c)
                Next
                If sheetName.Length > 31 Then sheetName = sheetName.Substring(0, 31)

                Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets.Add(sheetName)
                Dim linhaInicio As Integer = 4

                ' Logo
                If Not String.IsNullOrWhiteSpace(logoEfetivo) AndAlso File.Exists(logoEfetivo) Then
                    Try
                        Dim logo = ws.Drawings.AddPicture("logo_" & sheetName, New FileInfo(logoEfetivo))
                        logo.SetPosition(0, 4, 0, 4)
                        logo.SetSize(140, 46)
                    Catch
                    End Try
                End If

                ' Título e sub-título
                ws.Row(1).Height = 40
                ws.Cells(1, 4, 1, 7).Merge = True
                ws.Cells(1, 4).Value = "Hardware Upgrade Report"
                ws.Cells(1, 4).Style.Font.Bold = True
                ws.Cells(1, 4).Style.Font.Size = 14
                ws.Cells(1, 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center

                ws.Cells(2, 4, 2, 7).Merge = True
                ws.Cells(2, 4).Value = "Customer: " & kvp.Key & "   |   " &
                                       kvp.Value.Count.ToString() & " upgrade(s)   |   Generated: " &
                                       DateTime.Now.ToString("yyyy-MM-dd HH:mm")
                ws.Cells(2, 4).Style.Font.Size = 9
                ws.Cells(2, 4).Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(107, 114, 128))

                ' Cabeçalho de colunas
                For i As Integer = 0 To colunas.Length - 1
                    Dim cell As ExcelRange = ws.Cells(linhaInicio, i + 1)
                    cell.Value = colunas(i).Header
                    cell.Style.Font.Bold = True
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                    cell.Style.Fill.BackgroundColor.SetColor(corHeader)
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                Next

                ' Dados
                Dim rowIdx As Integer = 0
                For Each dr As DataRow In kvp.Value
                    Dim linha As Integer = linhaInicio + rowIdx + 1
                    For colIdx As Integer = 0 To colunas.Length - 1
                        Dim colKey As String = colunas(colIdx).Key
                        Dim val As String = ""
                        If dr.Table.Columns.Contains(colKey) AndAlso Not IsDBNull(dr(colKey)) Then
                            val = dr(colKey).ToString()
                        End If
                        ws.Cells(linha, colIdx + 1).Value = val
                    Next
                    If rowIdx Mod 2 = 1 Then
                        ws.Cells(linha, 1, linha, colunas.Length).Style.Fill.PatternType = ExcelFillStyle.Solid
                        ws.Cells(linha, 1, linha, colunas.Length).Style.Fill.BackgroundColor.SetColor(corZebra)
                    End If
                    rowIdx += 1
                Next

                ' Larguras, filtro e freeze
                For i As Integer = 0 To colunas.Length - 1
                    ws.Column(i + 1).Width = colunas(i).Largura
                Next
                ws.Column(colunas.Length).Style.WrapText = True
                ws.Cells(linhaInicio, 1, linhaInicio + rowIdx, colunas.Length).AutoFilter = True
                ws.View.FreezePanes(linhaInicio + 1, 1)

            Next

            pkg.SaveAs(New FileInfo(caminhoFinal))
        End Using

        Return caminhoFinal

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

    ' ── Shipment Item Reports ────────────────────────────────────────────────

    Public Shared Function GerarExcelRemessa(pItens As List(Of DataRow),
                                              pOutputPath As String,
                                              Optional pLogoPath As String = "",
                                              Optional baseNome As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(baseNome),
                                 "ShipmentReport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), baseNome)
        Dim caminhoFinal As String = Path.Combine(pasta, nome & ".xlsx")

        Dim colunas As (Key As String, Header As String, Largura As Double)() = {
            ("INTERNAL_UID",     "Internal UID",      16),
            ("MARCA",            "Manufacturer",       18),
            ("MODEL",            "Model",              16),
            ("SERIAL_NUMBER",    "Serial Number",      20),
            ("PROCESSADOR",      "Processor",          20),
            ("RAM_GB",           "RAM GB",             10),
            ("STORAGE_GB",       "Storage GB",         12),
            ("CONDITION_STATUS", "Battery Condition",  18)
        }

        Using pkg As New ExcelPackage()
            Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets.Add("Shipment Report")
            Dim linhaCabecalho As Integer = 5
            Dim logoEfetivo As String = ObterLogoPath(pLogoPath)

            ws.Row(1).Height = 40
            ws.Row(2).Height = 20
            ws.Row(3).Height = 20
            ws.Row(4).Height = 20

            ws.Cells(1, 4, 1, 8).Merge = True
            ws.Cells(1, 4).Value = "GBS Shipment Report"
            ws.Cells(1, 4).Style.Font.Bold = True
            ws.Cells(1, 4).Style.Font.Size = 16
            ws.Cells(1, 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center

            ws.Cells(2, 4, 2, 8).Merge = True
            ws.Cells(2, 4).Value = "Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                                    " | Total items: " & pItens.Count.ToString()
            ws.Cells(2, 4).Style.Font.Size = 10
            ws.Cells(2, 4).Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(107, 114, 128))

            If Not String.IsNullOrWhiteSpace(logoEfetivo) AndAlso File.Exists(logoEfetivo) Then
                Try
                    Dim logo = ws.Drawings.AddPicture("GBS_Logo", New FileInfo(logoEfetivo))
                    logo.SetPosition(0, 4, 0, 4)
                    logo.SetSize(180, 60)
                Catch
                End Try
            End If

            For i As Integer = 0 To colunas.Length - 1
                Dim cell As ExcelRange = ws.Cells(linhaCabecalho, i + 1)
                cell.Value = colunas(i).Header
                cell.Style.Font.Bold = True
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(55, 65, 81))
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            Next

            Dim corAlternado As System.Drawing.Color = System.Drawing.Color.FromArgb(249, 250, 251)

            For rowIdx As Integer = 0 To pItens.Count - 1
                Dim dr As DataRow = pItens(rowIdx)
                Dim tbl As DataTable = dr.Table
                For colIdx As Integer = 0 To colunas.Length - 1
                    Dim colKey As String = colunas(colIdx).Key
                    Dim val As String = ""
                    If tbl.Columns.Contains(colKey) AndAlso Not IsDBNull(dr(colKey)) Then
                        val = dr(colKey).ToString()
                    End If
                    ws.Cells(linhaCabecalho + rowIdx + 1, colIdx + 1).Value = val
                Next
                If rowIdx Mod 2 = 1 Then
                    ws.Cells(linhaCabecalho + rowIdx + 1, 1, linhaCabecalho + rowIdx + 1, colunas.Length).Style.Fill.PatternType = ExcelFillStyle.Solid
                    ws.Cells(linhaCabecalho + rowIdx + 1, 1, linhaCabecalho + rowIdx + 1, colunas.Length).Style.Fill.BackgroundColor.SetColor(corAlternado)
                End If
            Next

            For i As Integer = 0 To colunas.Length - 1
                ws.Column(i + 1).Width = colunas(i).Largura
            Next
            ws.Column(colunas.Length).Style.WrapText = True

            If pItens.Count > 0 Then
                ws.Cells(linhaCabecalho, 1, linhaCabecalho + pItens.Count, colunas.Length).AutoFilter = True
            End If

            ws.View.FreezePanes(linhaCabecalho + 1, 1)
            pkg.SaveAs(New FileInfo(caminhoFinal))
        End Using

        Return caminhoFinal

    End Function

    Public Shared Function GerarPdfRemessa(pItens As List(Of DataRow),
                                            pOutputPath As String,
                                            Optional pLogoPath As String = "",
                                            Optional baseNome As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(baseNome),
                                 "ShipmentReport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), baseNome)
        Dim caminhoFinal As String = Path.Combine(pasta, nome & ".pdf")

        Dim colunas As (Key As String, Header As String, Largura As Single)() = {
            ("INTERNAL_UID",     "Internal UID",      56),
            ("MARCA",            "Manufacturer",       68),
            ("MODEL",            "Model",              58),
            ("SERIAL_NUMBER",    "Serial Number",      72),
            ("PROCESSADOR",      "Processor",          80),
            ("RAM_GB",           "RAM GB",             32),
            ("STORAGE_GB",       "Storage GB",         42),
            ("CONDITION_STATUS", "Battery Condition",  68)
        }

        Dim doc As New Document(PageSize.A4.Rotate(), 20, 20, 30, 20)
        Using fs As New FileStream(caminhoFinal, FileMode.Create, FileAccess.Write)
            PdfWriter.GetInstance(doc, fs)
            doc.Open()

            Dim logoEfetivo As String = ObterLogoPath(pLogoPath)
            If Not String.IsNullOrWhiteSpace(logoEfetivo) AndAlso File.Exists(logoEfetivo) Then
                Try
                    Dim img As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(logoEfetivo)
                    img.ScaleToFit(140, 50)
                    img.Alignment = Element.ALIGN_LEFT
                    doc.Add(img)
                Catch
                End Try
            End If

            Dim fontTitulo As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 13, New BaseColor(31, 41, 55))
            Dim fontMeta   As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 8,  New BaseColor(107, 114, 128))
            Dim pTitulo As New Paragraph("GBS Shipment Report", fontTitulo)
            pTitulo.SpacingBefore = 4
            doc.Add(pTitulo)
            Dim pMeta As New Paragraph("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                                        "  |  Total items: " & pItens.Count.ToString(), fontMeta)
            pMeta.SpacingAfter = 8
            doc.Add(pMeta)

            Dim totalCols As Integer = colunas.Length
            Dim tabela As New PdfPTable(totalCols)
            tabela.WidthPercentage = 100
            Dim widths(totalCols - 1) As Single
            For i As Integer = 0 To totalCols - 1
                widths(i) = colunas(i).Largura
            Next
            tabela.SetWidths(widths)

            Dim fontHdr  As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7.5F, BaseColor.WHITE)
            Dim fontData As iTextSharp.text.Font = FontFactory.GetFont(FontFactory.HELVETICA, 7, New BaseColor(31, 41, 55))
            Dim hdrBg As New BaseColor(55, 65, 81)
            Dim altBg As New BaseColor(249, 250, 251)

            For Each c In colunas
                Dim cell As New PdfPCell(New Phrase(c.Header, fontHdr))
                cell.BackgroundColor     = hdrBg
                cell.HorizontalAlignment = Element.ALIGN_CENTER
                cell.Padding             = 5
                tabela.AddCell(cell)
            Next

            Dim rowIdx As Integer = 0
            For Each row As DataRow In pItens
                Dim dtRow As DataTable = row.Table
                Dim bg As BaseColor = If(rowIdx Mod 2 = 1, altBg, BaseColor.WHITE)
                For Each c In colunas
                    Dim val As String = ""
                    If dtRow.Columns.Contains(c.Key) AndAlso Not IsDBNull(row(c.Key)) Then
                        val = row(c.Key).ToString()
                    End If
                    Dim cell As New PdfPCell(New Phrase(val, fontData))
                    cell.BackgroundColor = bg
                    cell.Padding         = 4
                    tabela.AddCell(cell)
                Next
                rowIdx += 1
            Next

            doc.Add(tabela)
            doc.Close()
        End Using

        Return caminhoFinal

    End Function

    Public Shared Function GerarRelatorioRemessa(pItens As List(Of DataRow),
                                                  pOutputPath As String,
                                                  pLogoPath As String,
                                                  Optional baseNome As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(baseNome),
                                 "ShipmentReport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), baseNome)
        Dim caminhoFinal As String = Path.Combine(pasta, nome & ".doc")

        Dim logoEfetivo As String = ObterLogoPath(pLogoPath)
        Dim html As String = MontarHtmlRemessa(pItens, logoEfetivo)
        File.WriteAllText(caminhoFinal, html, Encoding.UTF8)

        Return caminhoFinal

    End Function

    Private Shared Function MontarHtmlRemessa(pItens As List(Of DataRow), pLogoPath As String) As String

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

        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Dim bytes As Byte() = File.ReadAllBytes(pLogoPath)
            Dim b64 As String = Convert.ToBase64String(bytes)
            Dim mime As String = ObterMime(pLogoPath)
            sb.AppendLine("<div style='margin-bottom:14px;'>")
            sb.AppendLine("<img src='data:" & mime & ";base64," & b64 &
                          "' style='max-height:70px;max-width:220px;' />")
            sb.AppendLine("</div>")
        End If

        sb.AppendLine("<h2>GBS Shipment Report</h2>")
        sb.AppendLine("<div class='sub'>Generated: " &
                      DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                      " &nbsp;|&nbsp; Total items: " & pItens.Count.ToString() & "</div>")

        Dim colunas As (Key As String, Label As String, Width As String)() = {
            ("INTERNAL_UID",     "Internal UID",      "80px"),
            ("MARCA",            "Manufacturer",       "90px"),
            ("MODEL",            "Model",              "80px"),
            ("SERIAL_NUMBER",    "Serial Number",      "90px"),
            ("PROCESSADOR",      "Processor",          "90px"),
            ("RAM_GB",           "RAM GB",             "50px"),
            ("STORAGE_GB",       "Storage GB",         "60px"),
            ("CONDITION_STATUS", "Battery Condition",  "80px")
        }

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
                Dim val As String = ""
                If tbl.Columns.Contains(c.Key) AndAlso Not IsDBNull(row(c.Key)) Then
                    val = row(c.Key).ToString()
                End If
                Dim estiloNotes As String = If(c.Key = "NOTES", " style='max-width:120px;word-wrap:break-word;'", "")
                sb.Append("<td" & estiloNotes & ">" & HtmlEncode(val) & "</td>")
            Next
            sb.AppendLine("</tr>")
        Next

        sb.AppendLine("</table>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()

    End Function

End Class
