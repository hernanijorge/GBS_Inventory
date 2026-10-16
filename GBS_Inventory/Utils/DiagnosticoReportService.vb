Imports System.Data
Imports System.Globalization
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

''' <summary>
''' Colunas e Excel da aba Diagnostics. A mesma especificacao (Colunas) e usada pelo grid
''' e pelo Excel, para os dois ficarem com a mesma ordem, cabecalhos e formato.
''' </summary>
Public Class DiagnosticoReportService

    Private Const PastaDefault As String = "C:\GBS\Reports\"

    Public Enum TipoColuna
        Texto
        Data
        Percentual
        Verdict
    End Enum

    ''' <summary>
    ''' Coluna do grid / Excel. Key = campo do PACK_DIAGNOSTICO.PROC_LIST.
    ''' LarguraGrid = pixels (0 = Fill); LarguraExcel = caracteres.
    ''' </summary>
    Public Class Coluna
        Public ReadOnly Key As String
        Public ReadOnly Header As String
        Public ReadOnly LarguraGrid As Integer
        Public ReadOnly LarguraExcel As Double
        Public ReadOnly Centralizada As Boolean
        Public ReadOnly Tipo As TipoColuna

        Public Sub New(pKey As String, pHeader As String, pLarguraGrid As Integer, pLarguraExcel As Double,
                       pCentralizada As Boolean, pTipo As TipoColuna)
            Key = pKey : Header = pHeader : LarguraGrid = pLarguraGrid : LarguraExcel = pLarguraExcel
            Centralizada = pCentralizada : Tipo = pTipo
        End Sub
    End Class

    ' Colunas do grid, na ordem de exibicao. O Excel usa as mesmas, na mesma ordem.
    Public Shared ReadOnly Colunas As Coluna() = {
        New Coluna("DATA_TESTE",     "Date",         140, 19, False, TipoColuna.Data),
        New Coluna("INTERNAL_UID",   "Internal UID", 100, 14, False, TipoColuna.Texto),
        New Coluna("SERVICE_TAG",    "Service Tag",  110, 16, False, TipoColuna.Texto),
        New Coluna("MODEL",          "Model",          0, 30, False, TipoColuna.Texto),
        New Coluna("VERDICT",        "Verdict",      150, 20, True,  TipoColuna.Verdict),
        New Coluna("BATTERY_HEALTH", "Battery %",     85, 11, True,  TipoColuna.Percentual),
        New Coluna("SSD_LIFE",       "SSD %",         75, 10, True,  TipoColuna.Percentual),
        New Coluna("TECHNICIAN",     "Technician",   110, 15, False, TipoColuna.Texto),
        New Coluna("ORIGIN",         "Origin",       105, 13, True,  TipoColuna.Texto)
    }

    ' Detalhes que so vao para o Excel, depois das colunas do grid
    Private Shared ReadOnly ColunasExtrasExcel As Coluna() = {
        New Coluna("MANUFACTURER",     "Manufacturer",   0, 14, False, TipoColuna.Texto),
        New Coluna("CPU",              "Processor",      0, 30, False, TipoColuna.Texto),
        New Coluna("RAM_GIB",          "RAM GiB",        0,  9, True,  TipoColuna.Texto),
        New Coluna("STORAGE",          "Storage",        0, 28, False, TipoColuna.Texto),
        New Coluna("BATTERY_CYCLES",   "Battery Cycles", 0, 13, True,  TipoColuna.Texto),
        New Coluna("FAIL_ITEMS",       "Failed",         0, 30, False, TipoColuna.Texto),
        New Coluna("WARN_ITEMS",       "Notes",          0, 30, False, TipoColuna.Texto),
        New Coluna("NOT_TESTED_ITEMS", "Not Tested",     0, 30, False, TipoColuna.Texto),
        New Coluna("TOOL_VERSION",     "Tool Version",   0, 12, False, TipoColuna.Texto)
    }

    Public Const FormatoData As String = "MM/dd/yyyy h:mm tt"
    Public Const FormatoPercentual As String = "0.0'%'"
    Public Const TextoNulo As String = "-"

    ' Codigo gravado no banco -> texto exibido
    Public Shared ReadOnly VerdictTextos As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"PASS", "PASSED"},
        {"PASS_WITH_WARNINGS", "PASSED WITH NOTES"},
        {"FAIL", "FAILED"},
        {"INCOMPLETE", "INCOMPLETE"}
    }

    Public Shared Function VerdictTexto(pCodigo As String) As String
        Dim t As String = Nothing
        If pCodigo IsNot Nothing AndAlso VerdictTextos.TryGetValue(pCodigo, t) Then Return t
        Return If(pCodigo, "")
    End Function

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

        Dim colunasExcel As Coluna() = Colunas.Concat(ColunasExtrasExcel).ToArray()

        Using pkg As New ExcelPackage()
            Dim ws As ExcelWorksheet = pkg.Workbook.Worksheets.Add("Diagnostics")
            Dim linhaCabecalho As Integer = 5

            ws.Row(1).Height = 40
            ws.Cells(1, 4, 1, 9).Merge = True
            ws.Cells(1, 4).Value = "GBS Hardware Diagnostics Report"
            ws.Cells(1, 4).Style.Font.Bold = True
            ws.Cells(1, 4).Style.Font.Size = 16
            ws.Cells(1, 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center

            ws.Cells(2, 4, 2, 9).Merge = True
            ws.Cells(2, 4).Value = "Generated: " & DateTime.Now.ToString(FormatoData, CultureInfo.InvariantCulture) &
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

            ' ── Cabeçalho (mesmo alinhamento das células) ──────────────
            For i As Integer = 0 To colunasExcel.Length - 1
                Dim cell As ExcelRange = ws.Cells(linhaCabecalho, i + 1)
                cell.Value = colunasExcel(i).Header
                cell.Style.Font.Bold = True
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(55, 65, 81))
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                cell.Style.HorizontalAlignment = Alinhamento(colunasExcel(i))
            Next

            ' ── Dados ─────────────────────────────────────────────────
            Dim corAlternado As System.Drawing.Color = System.Drawing.Color.FromArgb(249, 250, 251)

            For rowIdx As Integer = 0 To pItens.Count - 1
                Dim dr As DataRow = pItens(rowIdx)
                Dim linha As Integer = linhaCabecalho + rowIdx + 1

                If rowIdx Mod 2 = 1 Then
                    ws.Cells(linha, 1, linha, colunasExcel.Length).Style.Fill.PatternType = ExcelFillStyle.Solid
                    ws.Cells(linha, 1, linha, colunasExcel.Length).Style.Fill.BackgroundColor.SetColor(corAlternado)
                End If

                For colIdx As Integer = 0 To colunasExcel.Length - 1
                    Dim col As Coluna = colunasExcel(colIdx)
                    Dim cell As ExcelRange = ws.Cells(linha, colIdx + 1)
                    cell.Style.HorizontalAlignment = Alinhamento(col)

                    If Not dr.Table.Columns.Contains(col.Key) OrElse IsDBNull(dr(col.Key)) Then
                        cell.Value = TextoNulo
                        Continue For
                    End If

                    Select Case col.Tipo
                        Case TipoColuna.Data
                            cell.Value = CDate(dr(col.Key))
                            cell.Style.Numberformat.Format = "mm/dd/yyyy h:mm AM/PM"
                        Case TipoColuna.Percentual
                            cell.Value = Convert.ToDecimal(dr(col.Key))
                            cell.Style.Numberformat.Format = "0.0\%"   ' % literal (valor ja e percentual)
                        Case TipoColuna.Verdict
                            Dim v As String = dr(col.Key).ToString()
                            cell.Value = VerdictTexto(v)
                            cell.Style.Font.Bold = True
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid
                            cell.Style.Fill.BackgroundColor.SetColor(CorVerdict(v))
                        Case Else
                            cell.Value = dr(col.Key).ToString()
                    End Select
                Next
            Next

            For i As Integer = 0 To colunasExcel.Length - 1
                ws.Column(i + 1).Width = colunasExcel(i).LarguraExcel
            Next
            For Each key As String In {"FAIL_ITEMS", "WARN_ITEMS", "NOT_TESTED_ITEMS", "STORAGE", "CPU"}
                Dim idx As Integer = Array.FindIndex(colunasExcel, Function(c) c.Key = key) + 1
                ws.Column(idx).Style.WrapText = True
            Next

            If pItens.Count > 0 Then
                ws.Cells(linhaCabecalho, 1, linhaCabecalho + pItens.Count, colunasExcel.Length).AutoFilter = True
            End If
            ws.View.FreezePanes(linhaCabecalho + 1, 1)

            pkg.SaveAs(New FileInfo(caminhoFinal))
        End Using

        Return caminhoFinal

    End Function

    Private Shared Function Alinhamento(pCol As Coluna) As ExcelHorizontalAlignment
        Return If(pCol.Centralizada, ExcelHorizontalAlignment.Center, ExcelHorizontalAlignment.Left)
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
