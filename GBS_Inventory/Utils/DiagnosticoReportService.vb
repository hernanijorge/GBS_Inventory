Imports System.Data
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

''' <summary>
''' Excel da aba Diagnostics, gerado a partir das linhas do grid (PROC_LIST).
''' Mesmo layout dos relatorios do ReportService (logo, titulo, cabecalho escuro, zebra).
''' </summary>
Public Class DiagnosticoReportService

    Private Const PastaDefault As String = "C:\GBS\Reports\"

    Public Shared Function GerarExcel(pItens As List(Of DataRow),
                                      pOutputPath As String,
                                      Optional pLogoPath As String = "",
                                      Optional pBaseNome As String = "") As String

        Dim pasta As String = If(Not String.IsNullOrWhiteSpace(pOutputPath), pOutputPath, PastaDefault)
        If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)

        Dim nome As String = If(String.IsNullOrWhiteSpace(pBaseNome),
                                "Diagnostics_" & DateTime.Now.ToString("yyyyMMdd_HHmmss"), pBaseNome)
        For Each c As Char In Path.GetInvalidFileNameChars()
            nome = nome.Replace(c, "_"c)
        Next
        Dim caminhoFinal As String = Path.Combine(pasta, nome & ".xlsx")

        ' Colunas: chave (PROC_LIST), header, largura em caracteres
        Dim colunas As (Key As String, Header As String, Largura As Double)() = {
            ("DATA_TESTE",       "Test Date",        17),
            ("INTERNAL_UID",     "Internal UID",     14),
            ("SERVICE_TAG",      "Service Tag",      16),
            ("MANUFACTURER",     "Manufacturer",     14),
            ("MODEL",            "Model",            26),
            ("CPU",              "Processor",        30),
            ("RAM_GIB",          "RAM GiB",           9),
            ("STORAGE",          "Storage",          28),
            ("VERDICT",          "Verdict",          20),
            ("BATTERY_HEALTH",   "Battery %",        10),
            ("BATTERY_CYCLES",   "Battery Cycles",   13),
            ("SSD_LIFE",         "SSD Life %",       10),
            ("FAIL_ITEMS",       "Failed",           30),
            ("WARN_ITEMS",       "Warnings",         30),
            ("NOT_TESTED_ITEMS", "Not Tested",       30),
            ("TECHNICIAN",       "Technician",       14),
            ("ORIGIN",           "Origin",           13),
            ("TOOL_VERSION",     "Tool Version",     12)
        }
        Dim colVerdict As Integer = Array.FindIndex(colunas, Function(c) c.Key = "VERDICT") + 1

        Using pkg As New ExcelPackage()
            Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets.Add("Diagnostics")
            Dim linhaCabecalho As Integer = 5

            ws.Row(1).Height = 40
            ws.Cells(1, 4, 1, 10).Merge = True
            ws.Cells(1, 4).Value = "GBS Hardware Diagnostics Report"
            ws.Cells(1, 4).Style.Font.Bold = True
            ws.Cells(1, 4).Style.Font.Size = 16
            ws.Cells(1, 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center

            ws.Cells(2, 4, 2, 10).Merge = True
            ws.Cells(2, 4).Value = "Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") &
                                   " | Total devices: " & pItens.Count.ToString()
            ws.Cells(2, 4).Style.Font.Size = 10
            ws.Cells(2, 4).Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(107, 114, 128))

            Dim logo As String = ObterLogoPath(pLogoPath)
            If logo.Length > 0 Then
                Try
                    Dim pic = ws.Drawings.AddPicture("GBS_Logo", New FileInfo(logo))
                    pic.SetPosition(0, 4, 0, 4)
                    pic.SetSize(180, 60)
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
                Dim linha As Integer = linhaCabecalho + rowIdx + 1

                If rowIdx Mod 2 = 1 Then
                    ws.Cells(linha, 1, linha, colunas.Length).Style.Fill.PatternType = ExcelFillStyle.Solid
                    ws.Cells(linha, 1, linha, colunas.Length).Style.Fill.BackgroundColor.SetColor(corAlternado)
                End If

                For colIdx As Integer = 0 To colunas.Length - 1
                    Dim key As String = colunas(colIdx).Key
                    If Not dr.Table.Columns.Contains(key) OrElse IsDBNull(dr(key)) Then Continue For
                    Dim cell As ExcelRange = ws.Cells(linha, colIdx + 1)

                    Select Case key
                        Case "DATA_TESTE"
                            cell.Value = CDate(dr(key))
                            cell.Style.Numberformat.Format = "yyyy-mm-dd hh:mm"
                        Case "RAM_GIB", "BATTERY_HEALTH", "BATTERY_CYCLES", "SSD_LIFE"
                            cell.Value = Convert.ToDecimal(dr(key))
                            cell.Style.Numberformat.Format = "0.#"
                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        Case "VERDICT"
                            Dim v As String = dr(key).ToString()
                            cell.Value = v.Replace("_", " ")
                            cell.Style.Font.Bold = True
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                            cell.Style.Fill.BackgroundColor.SetColor(CorVerdict(v))
                        Case Else
                            cell.Value = dr(key).ToString()
                    End Select
                Next
            Next

            For i As Integer = 0 To colunas.Length - 1
                ws.Column(i + 1).Width = colunas(i).Largura
            Next
            For Each key As String In {"FAIL_ITEMS", "WARN_ITEMS", "NOT_TESTED_ITEMS", "STORAGE", "CPU"}
                Dim idx As Integer = Array.FindIndex(colunas, Function(c) c.Key = key) + 1
                ws.Column(idx).Style.WrapText = True
            Next

            If pItens.Count > 0 Then
                ws.Cells(linhaCabecalho, 1, linhaCabecalho + pItens.Count, colunas.Length).AutoFilter = True
            End If
            ws.View.FreezePanes(linhaCabecalho + 1, 1)

            pkg.SaveAs(New FileInfo(caminhoFinal))
        End Using

        Return caminhoFinal

    End Function

    ''' <summary>Cor de fundo do veredito (tons claros para impressao).</summary>
    Public Shared Function CorVerdict(pVerdict As String) As System.Drawing.Color
        Select Case If(pVerdict, "").ToUpperInvariant()
            Case "PASS" : Return System.Drawing.Color.FromArgb(187, 247, 208)
            Case "PASS_WITH_WARNINGS" : Return System.Drawing.Color.FromArgb(254, 240, 138)
            Case "FAIL" : Return System.Drawing.Color.FromArgb(254, 202, 202)
            Case Else : Return System.Drawing.Color.FromArgb(229, 231, 235)
        End Select
    End Function

    Private Shared Function ObterLogoPath(pLogoPath As String) As String
        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then Return pLogoPath
        For Each candidato As String In {"C:\GBS\logo.png", "C:\GBS\logo.jpg", "C:\GBS\logo.jpeg"}
            If File.Exists(candidato) Then Return candidato
        Next
        Return ""
    End Function

End Class
