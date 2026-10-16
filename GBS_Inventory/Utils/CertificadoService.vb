Imports System.Data
Imports System.Globalization
Imports System.IO
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO

''' <summary>
''' Certificados de teste da remessa: UM PDF com capa (resumo em ingles) seguida dos
''' PDFs individuais de diagnostico, na ordem das linhas recebidas (ordem de UID).
''' PdfSharp 1.50 (MIT).
''' </summary>
Public Class CertificadoService

    Public Class Resultado
        Public Caminho As String
        Public CertificadosIncluidos As Integer
        Public CertificadosComFalha As New List(Of String)   ' UIDs cujo PDF nao pode ser importado
    End Class

    Private Const Margem As Double = 40
    Private Const AlturaLinha As Double = 17

    ' Colunas da tabela-resumo da capa: header, largura (pt), centralizada
    Private Shared ReadOnly ColunasCapa As (Header As String, Largura As Double, Centro As Boolean)() = {
        ("Internal UID", 70, False),
        ("Service Tag", 78, False),
        ("Model", 150, False),
        ("Verdict", 92, True),
        ("Battery %", 52, True),
        ("SSD %", 42, True),
        ("Test Date", 48, True)
    }

    ''' <param name="pLinhas">Linhas do PROC_LIST_BY_SHIPMENT a incluir (capa + certificados).</param>
    ''' <param name="pPdfs">ID_DIAGNOSTICO -> caminho do PDF individual ja extraido do banco.</param>
    Public Shared Function GerarPdfCombinado(pCabecalho As DataRow,
                                             pLinhas As List(Of DataRow),
                                             pPdfs As Dictionary(Of Long, String),
                                             pCaminhoSaida As String,
                                             Optional pLogoPath As String = "") As Resultado

        Dim res As New Resultado With {.Caminho = pCaminhoSaida}

        Using doc As New PdfDocument()
            doc.Info.Title = "GBS Test Certificates - " & Texto(pCabecalho, "CODIGO_REMESSA")
            doc.Info.Author = "Global Business Solutions Inc."

            DesenharCapa(doc, pCabecalho, pLinhas, pLogoPath)

            For Each row As DataRow In pLinhas
                If IsDBNull(row("ID_DIAGNOSTICO")) Then Continue For
                Dim idDiag As Long = CLng(row("ID_DIAGNOSTICO"))
                Dim caminho As String = Nothing
                If Not pPdfs.TryGetValue(idDiag, caminho) OrElse Not File.Exists(caminho) Then Continue For

                Try
                    Using origem As PdfDocument = PdfReader.Open(caminho, PdfDocumentOpenMode.Import)
                        For Each pag As PdfPage In origem.Pages
                            doc.AddPage(pag)
                        Next
                    End Using
                    res.CertificadosIncluidos += 1
                Catch ex As Exception
                    ' Um PDF ilegivel nao derruba o documento: entra uma pagina de aviso no lugar
                    Dim uid As String = Texto(row, "INTERNAL_UID")
                    res.CertificadosComFalha.Add(uid)
                    PaginaAviso(doc, $"Certificate for Internal UID {uid} ({Texto(row, "SERVICE_TAG")}) could not be embedded.",
                                "Original file: " & Path.GetFileName(caminho) & " — " & ex.Message)
                End Try
            Next

            Dim pasta As String = Path.GetDirectoryName(pCaminhoSaida)
            If Not Directory.Exists(pasta) Then Directory.CreateDirectory(pasta)
            doc.Save(pCaminhoSaida)
        End Using

        Return res

    End Function

#Region "Capa"

    Private Shared Sub DesenharCapa(doc As PdfDocument, pCab As DataRow, pLinhas As List(Of DataRow), pLogoPath As String)

        Dim opt As New XPdfFontOptions(PdfFontEncoding.Unicode)
        Dim fTitulo As New XFont("Arial", 20, XFontStyle.Bold, opt)
        Dim fEmpresa As New XFont("Arial", 11, XFontStyle.Bold, opt)
        Dim fRotulo As New XFont("Arial", 9, XFontStyle.Bold, opt)
        Dim fTexto As New XFont("Arial", 9, XFontStyle.Regular, opt)
        Dim fTabela As New XFont("Arial", 7.5, XFontStyle.Regular, opt)
        Dim fTabelaB As New XFont("Arial", 7.5, XFontStyle.Bold, opt)
        Dim fRodape As New XFont("Arial", 7.5, XFontStyle.Regular, opt)

        Dim cinzaTexto As New XSolidBrush(XColor.FromArgb(107, 114, 128))
        Dim escuro As New XSolidBrush(XColor.FromArgb(17, 24, 39))
        Dim cabecalhoTabela As New XSolidBrush(XColor.FromArgb(55, 65, 81))
        Dim zebra As New XSolidBrush(XColor.FromArgb(249, 250, 251))
        Dim linhaPen As New XPen(XColor.FromArgb(229, 231, 235), 0.5)

        Dim larguraTabela As Double = ColunasCapa.Sum(Function(c) c.Largura)
        Dim paginas As New List(Of PdfPage)

        Dim page As PdfPage = NovaPagina(doc, paginas)
        Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
        Dim y As Double = Margem

        ' ── Logo / empresa ────────────────────────────────────────────
        Dim logo As String = ObterLogoPath(pLogoPath)
        Dim xTexto As Double = Margem
        If logo <> "" Then
            Try
                Using img As XImage = XImage.FromFile(logo)
                    Dim h As Double = 46
                    Dim w As Double = img.PixelWidth * h / Math.Max(1, img.PixelHeight)
                    gfx.DrawImage(img, Margem, y, w, h)
                    xTexto = Margem + w + 14
                End Using
            Catch
                ' sem logo: segue so com o nome
            End Try
        End If
        gfx.DrawString("Global Business Solutions Inc.", fEmpresa, escuro, xTexto, y + 16)
        gfx.DrawString("19 Weston St., Lexington, MA 02421  |  www.gbsolinc.com  |  contact@gbsolinc.com",
                       fRodape, cinzaTexto, xTexto, y + 32)
        y += 70

        gfx.DrawString("Hardware Test Certificates", fTitulo, escuro, Margem, y)
        y += 26

        ' ── Dados da remessa ─────────────────────────────────────────
        Dim dataEnvio As String = If(IsDBNull(pCab("DATA_ENVIO")), "-",
                                     CDate(pCab("DATA_ENVIO")).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture))
        Dim campos As (String, String)() = {
            ("Customer:", Texto(pCab, "CLIENTE_NOME")),
            ("Shipment:", Texto(pCab, "CODIGO_REMESSA")),
            ("Ship date:", dataEnvio),
            ("Issued:", DateTime.Now.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture)),
            ("Summary:", Resumo(pLinhas))
        }
        For Each c In campos
            gfx.DrawString(c.Item1, fRotulo, escuro, Margem, y)
            gfx.DrawString(If(c.Item2 = "", "-", c.Item2), fTexto, escuro, Margem + 62, y)
            y += 14
        Next
        y += 6
        gfx.DrawString("Each device listed below was tested with the GBS laptop diagnostic tool. The individual " &
                       "test reports follow this page, in the same order.", fRodape, cinzaTexto, Margem, y)
        y += 14

        ' ── Tabela-resumo (continua nas paginas seguintes se precisar) ──
        Dim desenharCabecalho As Action =
            Sub()
                gfx.DrawRectangle(cabecalhoTabela, Margem, y, larguraTabela, AlturaLinha)
                Dim x As Double = Margem
                For Each col In ColunasCapa
                    Celula(gfx, col.Header, fTabelaB, XBrushes.White, x, y, col.Largura, col.Centro)
                    x += col.Largura
                Next
                y += AlturaLinha
            End Sub

        desenharCabecalho()
        Dim limite As Double = page.Height.Point - Margem - 20

        For i As Integer = 0 To pLinhas.Count - 1
            If y + AlturaLinha > limite Then
                gfx.Dispose()
                page = NovaPagina(doc, paginas)
                gfx = XGraphics.FromPdfPage(page)
                y = Margem
                gfx.DrawString("Hardware Test Certificates — " & Texto(pCab, "CODIGO_REMESSA") & " (continued)",
                               fRotulo, escuro, Margem, y)
                y += 12
                desenharCabecalho()
            End If

            Dim r As DataRow = pLinhas(i)
            If i Mod 2 = 1 Then gfx.DrawRectangle(zebra, Margem, y, larguraTabela, AlturaLinha)

            Dim codigo As String = Texto(r, "VERDICT")
            Dim valores As String() = {
                Texto(r, "INTERNAL_UID"),
                Texto(r, "SERVICE_TAG"),
                Texto(r, "MODEL"),
                If(codigo = "", "NOT TESTED", DiagnosticoReportService.VerdictTexto(codigo)),
                Percentual(r, "BATTERY_HEALTH"),
                Percentual(r, "SSD_LIFE"),
                If(IsDBNull(r("DATA_TESTE")), "-", CDate(r("DATA_TESTE")).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture))
            }

            Dim x As Double = Margem
            For c As Integer = 0 To ColunasCapa.Length - 1
                If c = 3 Then
                    Dim cor As System.Drawing.Color = CorVerdictCapa(codigo)
                    gfx.DrawRectangle(New XSolidBrush(XColor.FromArgb(cor.R, cor.G, cor.B)), x + 1, y + 1.5, ColunasCapa(c).Largura - 2, AlturaLinha - 3)
                    Celula(gfx, valores(c), fTabelaB, escuro, x, y, ColunasCapa(c).Largura, True)
                Else
                    Celula(gfx, valores(c), fTabela, escuro, x, y, ColunasCapa(c).Largura, ColunasCapa(c).Centro)
                End If
                x += ColunasCapa(c).Largura
            Next
            gfx.DrawLine(linhaPen, Margem, y + AlturaLinha, Margem + larguraTabela, y + AlturaLinha)
            y += AlturaLinha
        Next
        gfx.Dispose()

        ' Rodape "Page x of y" nas paginas da capa
        For p As Integer = 0 To paginas.Count - 1
            Using g As XGraphics = XGraphics.FromPdfPage(paginas(p), XGraphicsPdfPageOptions.Append)
                Dim txt As String = "Global Business Solutions Inc. — Test Certificates " & Texto(pCab, "CODIGO_REMESSA") &
                                    "   |   Summary page " & (p + 1) & " of " & paginas.Count
                g.DrawString(txt, fRodape, cinzaTexto, New XRect(0, paginas(p).Height.Point - Margem + 10, paginas(p).Width.Point, 12),
                             XStringFormats.Center)
            End Using
        Next

    End Sub

    Private Shared Function NovaPagina(doc As PdfDocument, pLista As List(Of PdfPage)) As PdfPage
        Dim p As PdfPage = doc.AddPage()
        p.Size = PageSize.Letter
        pLista.Add(p)
        Return p
    End Function

    ''' <summary>Texto numa celula, cortado com "..." se nao couber.</summary>
    Private Shared Sub Celula(gfx As XGraphics, pTexto As String, pFonte As XFont, pCor As XBrush,
                              x As Double, y As Double, w As Double, pCentro As Boolean)
        Dim t As String = If(pTexto, "")
        Dim maxW As Double = w - 6
        If gfx.MeasureString(t, pFonte).Width > maxW Then
            Do While t.Length > 1 AndAlso gfx.MeasureString(t & "...", pFonte).Width > maxW
                t = t.Substring(0, t.Length - 1)
            Loop
            t &= "..."
        End If
        Dim rect As New XRect(x + 3, y, maxW, AlturaLinha)
        gfx.DrawString(t, pFonte, pCor, rect, If(pCentro, XStringFormats.Center, XStringFormats.CenterLeft))
    End Sub

    Private Shared Sub PaginaAviso(doc As PdfDocument, pTitulo As String, pDetalhe As String)
        Dim p As PdfPage = doc.AddPage()
        p.Size = PageSize.Letter
        Using g As XGraphics = XGraphics.FromPdfPage(p)
            Dim opt As New XPdfFontOptions(PdfFontEncoding.Unicode)
            g.DrawString(pTitulo, New XFont("Arial", 12, XFontStyle.Bold, opt), XBrushes.Black, Margem, 80)
            g.DrawString(pDetalhe, New XFont("Arial", 8, XFontStyle.Regular, opt), XBrushes.Gray, Margem, 100)
        End Using
    End Sub

#End Region

#Region "Utilitarios"

    ''' <summary>"12 machines: 10 passed, 1 failed, 1 not tested" (categorias zeradas omitidas).</summary>
    Public Shared Function Resumo(pLinhas As IEnumerable(Of DataRow)) As String
        Dim total As Integer = 0, ok As Integer = 0, falha As Integer = 0, incompleto As Integer = 0, semTeste As Integer = 0
        For Each r As DataRow In pLinhas
            total += 1
            Select Case Texto(r, "VERDICT").ToUpperInvariant()
                Case "PASS", "PASS_WITH_WARNINGS" : ok += 1
                Case "FAIL" : falha += 1
                Case "INCOMPLETE" : incompleto += 1
                Case Else : semTeste += 1
            End Select
        Next
        Dim partes As New List(Of String) From {ok & " passed"}
        If falha > 0 Then partes.Add(falha & " failed")
        If incompleto > 0 Then partes.Add(incompleto & " incomplete")
        If semTeste > 0 Then partes.Add(semTeste & " not tested")
        Return total & If(total = 1, " machine: ", " machines: ") & String.Join(", ", partes)
    End Function

    Public Shared Function Aprovado(pRow As DataRow) As Boolean
        Dim v As String = Texto(pRow, "VERDICT").ToUpperInvariant()
        Return v = "PASS" OrElse v = "PASS_WITH_WARNINGS"
    End Function

    ''' <summary>Mesmas cores da aba Diagnostics; sem teste = cinza.</summary>
    Public Shared Function CorVerdictCapa(pCodigo As String) As System.Drawing.Color
        If String.IsNullOrEmpty(pCodigo) Then Return System.Drawing.Color.FromArgb(209, 213, 219)
        Return DiagnosticoReportService.CorVerdict(pCodigo)
    End Function

    Private Shared Function Percentual(r As DataRow, pCol As String) As String
        If IsDBNull(r(pCol)) Then Return DiagnosticoReportService.TextoNulo
        Return Convert.ToDecimal(r(pCol)).ToString(DiagnosticoReportService.FormatoPercentual, CultureInfo.InvariantCulture)
    End Function

    Private Shared Function Texto(r As DataRow, pCol As String) As String
        If r Is Nothing OrElse Not r.Table.Columns.Contains(pCol) OrElse IsDBNull(r(pCol)) Then Return ""
        Return r(pCol).ToString().Trim()
    End Function

    Private Shared Function ObterLogoPath(pLogoPath As String) As String
        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then Return pLogoPath
        For Each candidato As String In {"C:\GBS\logo.png", "C:\GBS\logo.jpg", "C:\GBS\logo.jpeg"}
            If File.Exists(candidato) Then Return candidato
        Next
        Return ""
    End Function

#End Region

End Class
