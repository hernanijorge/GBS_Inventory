Imports System.Globalization
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Text
Imports GBS_Inventory.Models

''' <summary>
''' Gera arquivos de invoice em Word (.doc html) e PDF simples.
''' </summary>
Public Class InvoiceDocumentoService

    Public Shared Sub GerarArquivos(pInvoice As Invoice,
                                    pCliente As Cliente,
                                    pItens As List(Of InvoiceItem),
                                    pLogoPath As String,
                                    pWordPath As String,
                                    pPdfPath As String)

        Dim pastaWord As String = Path.GetDirectoryName(pWordPath)
        Dim pastaPdf As String = Path.GetDirectoryName(pPdfPath)

        If Not String.IsNullOrWhiteSpace(pastaWord) AndAlso Not Directory.Exists(pastaWord) Then
            Directory.CreateDirectory(pastaWord)
        End If

        If Not String.IsNullOrWhiteSpace(pastaPdf) AndAlso Not Directory.Exists(pastaPdf) Then
            Directory.CreateDirectory(pastaPdf)
        End If

        Dim logoEfetivo As String = ObterLogoPathEfetivo(pLogoPath)
        Dim html As String = MontarHtml(pInvoice, pCliente, pItens, logoEfetivo)
        File.WriteAllText(pWordPath, html, Encoding.UTF8)

        Dim linhasPdf As List(Of String) = MontarLinhasPdf(pInvoice, pCliente, pItens)
        SimplePdfWriter.WriteTextPdf(pPdfPath, linhasPdf, logoEfetivo)

    End Sub

    Private Shared Function MontarHtml(pInvoice As Invoice,
                                       pCliente As Cliente,
                                       pItens As List(Of InvoiceItem),
                                       pLogoPath As String) As String

        Dim sb As New StringBuilder()
        Dim totalItens As Decimal = 0D

        sb.AppendLine("<html>")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset='utf-8' />")
        sb.AppendLine("<style>")
        sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;font-size:11pt;color:#1f2937;}")
        sb.AppendLine("table{width:100%;border-collapse:collapse;margin-top:14px;}")
        sb.AppendLine("th,td{border:1px solid #d1d5db;padding:6px;text-align:left;}")
        sb.AppendLine("th{background:#f3f4f6;}")
        sb.AppendLine(".right{text-align:right;}")
        sb.AppendLine(".top{display:flex;justify-content:space-between;align-items:flex-start;}")
        sb.AppendLine(".small{font-size:10pt;color:#4b5563;}")
        sb.AppendLine("</style>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")

        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Dim bytes As Byte() = File.ReadAllBytes(pLogoPath)
            Dim b64 As String = Convert.ToBase64String(bytes)
            Dim mime As String = ObterMimeLogo(pLogoPath)
            sb.AppendLine("<div style='margin-bottom:8px;'><img src='data:" & mime & ";base64," & b64 & "' style='max-height:80px;max-width:240px;' /></div>")
        End If

        sb.AppendLine("<div>")
        sb.AppendLine("<h2>Global Business Solution - Invoice</h2>")
        sb.AppendLine("<div class='small'>Invoice Number: " & HtmlEncode(pInvoice.InvoiceNumber) & "</div>")
        sb.AppendLine("<div class='small'>Issue Date: " & pInvoice.IssueDate.ToString("yyyy-MM-dd") & "</div>")
        sb.AppendLine("<div class='small'>Due Date: " & If(pInvoice.DueDate.HasValue, pInvoice.DueDate.Value.ToString("yyyy-MM-dd"), "-") & "</div>")
        sb.AppendLine("</div>")

        sb.AppendLine("<h3>Bill To</h3>")
        sb.AppendLine("<div><strong>" & HtmlEncode(pCliente.NomeRazao) & "</strong></div>")
        sb.AppendLine("<div>" & HtmlEncode(pCliente.Documento) & "</div>")
        sb.AppendLine("<div>" & HtmlEncode(pCliente.Email) & " | " & HtmlEncode(pCliente.Telefone) & "</div>")
        sb.AppendLine("<div>" & HtmlEncode(pCliente.Endereco1) & " " & HtmlEncode(pCliente.Endereco2) & "</div>")
        sb.AppendLine("<div>" & HtmlEncode(pCliente.Cidade) & " - " & HtmlEncode(pCliente.Estado) & " " & HtmlEncode(pCliente.ZipCode) & " / " & HtmlEncode(pCliente.Pais) & "</div>")

        sb.AppendLine("<table>")
        sb.AppendLine("<tr><th>Description</th><th class='right'>Qty</th><th class='right'>Unit (USD)</th><th class='right'>Line Total (USD)</th></tr>")

        For Each item As InvoiceItem In pItens
            item.RecalcularTotal()
            totalItens += item.LineTotalUsd

            sb.AppendLine("<tr>")
            sb.AppendLine("<td>" & HtmlEncode(item.Description) & "</td>")
            sb.AppendLine("<td class='right'>" & item.Qty.ToString("0.##", CultureInfo.InvariantCulture) & "</td>")
            sb.AppendLine("<td class='right'>" & item.UnitPriceUsd.ToString("0.00", CultureInfo.InvariantCulture) & "</td>")
            sb.AppendLine("<td class='right'>" & item.LineTotalUsd.ToString("0.00", CultureInfo.InvariantCulture) & "</td>")
            sb.AppendLine("</tr>")
        Next

        sb.AppendLine("</table>")

        Dim totalFinal As Decimal = totalItens - pInvoice.DiscountUsd + pInvoice.ShippingUsd + pInvoice.TaxUsd

        sb.AppendLine("<table style='width:360px;margin-left:auto;margin-top:16px;'>")
        sb.AppendLine("<tr><td>Subtotal</td><td class='right'>" & totalItens.ToString("0.00", CultureInfo.InvariantCulture) & "</td></tr>")
        sb.AppendLine("<tr><td>Discount</td><td class='right'>-" & pInvoice.DiscountUsd.ToString("0.00", CultureInfo.InvariantCulture) & "</td></tr>")
        sb.AppendLine("<tr><td>Shipping</td><td class='right'>" & pInvoice.ShippingUsd.ToString("0.00", CultureInfo.InvariantCulture) & "</td></tr>")
        sb.AppendLine("<tr><td>Tax</td><td class='right'>" & pInvoice.TaxUsd.ToString("0.00", CultureInfo.InvariantCulture) & "</td></tr>")
        sb.AppendLine("<tr><td><strong>Total</strong></td><td class='right'><strong>" & totalFinal.ToString("0.00", CultureInfo.InvariantCulture) & "</strong></td></tr>")
        sb.AppendLine("</table>")

        If Not String.IsNullOrWhiteSpace(pInvoice.Notes) Then
            sb.AppendLine("<h4>Notes</h4>")
            sb.AppendLine("<div>" & HtmlEncode(pInvoice.Notes).Replace(vbCrLf, "<br/>").Replace(vbLf, "<br/>") & "</div>")
        End If

        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()

    End Function

    Private Shared Function MontarLinhasPdf(pInvoice As Invoice,
                                            pCliente As Cliente,
                                            pItens As List(Of InvoiceItem)) As List(Of String)

        Dim linhas As New List(Of String)()
        Dim subtotal As Decimal = 0D

        linhas.Add("Global Business Solution - Invoice")
        linhas.Add("Invoice Number: " & pInvoice.InvoiceNumber)
        linhas.Add("Issue Date: " & pInvoice.IssueDate.ToString("yyyy-MM-dd"))
        linhas.Add("Due Date: " & If(pInvoice.DueDate.HasValue, pInvoice.DueDate.Value.ToString("yyyy-MM-dd"), "-"))
        linhas.Add("")
        linhas.Add("Bill To: " & pCliente.NomeRazao)
        linhas.Add("Document: " & pCliente.Documento)
        linhas.Add("Email: " & pCliente.Email & " | Phone: " & pCliente.Telefone)
        linhas.Add("Address: " & pCliente.Endereco1 & " " & pCliente.Endereco2)
        linhas.Add("City/State: " & pCliente.Cidade & " - " & pCliente.Estado & " " & pCliente.ZipCode & " / " & pCliente.Pais)

        linhas.Add("")
        linhas.Add("Items:")

        For Each item As InvoiceItem In pItens
            item.RecalcularTotal()
            subtotal += item.LineTotalUsd
            linhas.Add("- " & item.Description &
                       " | Qty: " & item.Qty.ToString("0.##", CultureInfo.InvariantCulture) &
                       " | Unit: " & item.UnitPriceUsd.ToString("0.00", CultureInfo.InvariantCulture) &
                       " | Total: " & item.LineTotalUsd.ToString("0.00", CultureInfo.InvariantCulture))
        Next

        Dim total As Decimal = subtotal - pInvoice.DiscountUsd + pInvoice.ShippingUsd + pInvoice.TaxUsd

        linhas.Add("")
        linhas.Add("Subtotal: " & subtotal.ToString("0.00", CultureInfo.InvariantCulture))
        linhas.Add("Discount: -" & pInvoice.DiscountUsd.ToString("0.00", CultureInfo.InvariantCulture))
        linhas.Add("Shipping: " & pInvoice.ShippingUsd.ToString("0.00", CultureInfo.InvariantCulture))
        linhas.Add("Tax: " & pInvoice.TaxUsd.ToString("0.00", CultureInfo.InvariantCulture))
        linhas.Add("TOTAL: " & total.ToString("0.00", CultureInfo.InvariantCulture))

        If Not String.IsNullOrWhiteSpace(pInvoice.Notes) Then
            linhas.Add("")
            linhas.Add("Notes:")
            For Each linha As String In pInvoice.Notes.Replace(vbCrLf, vbLf).Split(ControlChars.Lf)
                linhas.Add(linha)
            Next
        End If

        Return linhas

    End Function

    Private Shared Function HtmlEncode(pTexto As String) As String
        If pTexto Is Nothing Then Return ""

        Return pTexto.Replace("&", "&amp;").
                      Replace("<", "&lt;").
                      Replace(">", "&gt;").
                      Replace("""", "&quot;")
    End Function

    Private Shared Function ObterMimeLogo(pLogoPath As String) As String

        Dim ext As String = Path.GetExtension(pLogoPath).ToLowerInvariant()

        Select Case ext
            Case ".png" : Return "image/png"
            Case ".jpg", ".jpeg" : Return "image/jpeg"
            Case ".gif" : Return "image/gif"
            Case ".bmp" : Return "image/bmp"
            Case Else : Return "image/jpeg"
        End Select

    End Function

    Private Shared Function ObterLogoPathEfetivo(pLogoPath As String) As String

        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Return pLogoPath
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

Friend Class SimplePdfWriter

    Public Shared Sub WriteTextPdf(pCaminho As String, pLinhas As List(Of String), pLogoPath As String)

        Dim linhasExpandida As New List(Of String)()

        For Each linha As String In pLinhas
            If linha Is Nothing Then
                linhasExpandida.Add("")
            ElseIf linha.Length <= 105 Then
                linhasExpandida.Add(linha)
            Else
                Dim i As Integer = 0
                While i < linha.Length
                    Dim tam As Integer = Math.Min(105, linha.Length - i)
                    linhasExpandida.Add(linha.Substring(i, tam))
                    i += tam
                End While
            End If
        Next

        Dim logoJpeg As Byte() = Nothing
        Dim logoWidthPx As Integer = 0
        Dim logoHeightPx As Integer = 0
        Dim temLogo As Boolean = False

        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            logoJpeg = ConverterImagemParaJpeg(pLogoPath, logoWidthPx, logoHeightPx)
            temLogo = (logoJpeg IsNot Nothing AndAlso logoJpeg.Length > 0 AndAlso logoWidthPx > 0 AndAlso logoHeightPx > 0)
        End If

        Dim sbConteudo As New StringBuilder()

        If temLogo Then
            Dim drawW As Decimal = 130D
            Dim drawH As Decimal = drawW * CDec(logoHeightPx) / CDec(logoWidthPx)
            If drawH > 70D Then
                drawH = 70D
                drawW = drawH * CDec(logoWidthPx) / CDec(logoHeightPx)
            End If

            Dim posX As Decimal = 40D
            Dim posY As Decimal = 842D - 40D - drawH

            sbConteudo.AppendLine("q")
            sbConteudo.AppendLine(drawW.ToString("0.##", CultureInfo.InvariantCulture) & " 0 0 " &
                                  drawH.ToString("0.##", CultureInfo.InvariantCulture) & " " &
                                  posX.ToString("0.##", CultureInfo.InvariantCulture) & " " &
                                  posY.ToString("0.##", CultureInfo.InvariantCulture) & " cm")
            sbConteudo.AppendLine("/Im1 Do")
            sbConteudo.AppendLine("Q")
        End If

        sbConteudo.AppendLine("BT")
        sbConteudo.AppendLine("/F1 10 Tf")
        Dim yInicialTexto As String = If(temLogo, "700", "800")
        sbConteudo.AppendLine("40 " & yInicialTexto & " Td")

        Dim primeira As Boolean = True
        Dim limiteLinhas As Integer = Math.Min(If(temLogo, 50, 55), linhasExpandida.Count)

        For i As Integer = 0 To limiteLinhas - 1
            Dim textoEscapado As String = EscapePdfText(linhasExpandida(i))

            If primeira Then
                sbConteudo.AppendLine("(" & textoEscapado & ") Tj")
                primeira = False
            Else
                sbConteudo.AppendLine("0 -14 Td")
                sbConteudo.AppendLine("(" & textoEscapado & ") Tj")
            End If
        Next

        sbConteudo.AppendLine("ET")

        Dim bytesConteudo As Byte() = Encoding.ASCII.GetBytes(sbConteudo.ToString())

        Using ms As New MemoryStream()
            Dim enc As Encoding = Encoding.ASCII
            Dim offsets As New Dictionary(Of Integer, Integer)()

            Dim cabecalho As Byte() = enc.GetBytes("%PDF-1.4" & vbLf)
            ms.Write(cabecalho, 0, cabecalho.Length)

            offsets(1) = CInt(ms.Position)
            EscreverAscii(ms, "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj" & vbLf)

            offsets(2) = CInt(ms.Position)
            EscreverAscii(ms, "2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj" & vbLf)

            offsets(3) = CInt(ms.Position)
            Dim recursos As String
            If temLogo Then
                recursos = "/Resources << /Font << /F1 5 0 R >> /XObject << /Im1 6 0 R >> >>"
            Else
                recursos = "/Resources << /Font << /F1 5 0 R >> >>"
            End If
            EscreverAscii(ms, "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] " & recursos & " /Contents 4 0 R >> endobj" & vbLf)

            offsets(4) = CInt(ms.Position)
            EscreverAscii(ms, "4 0 obj << /Length " & bytesConteudo.Length.ToString() & " >> stream" & vbLf)
            ms.Write(bytesConteudo, 0, bytesConteudo.Length)
            EscreverAscii(ms, vbLf & "endstream endobj" & vbLf)

            offsets(5) = CInt(ms.Position)
            EscreverAscii(ms, "5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj" & vbLf)

            Dim maxObj As Integer = 5

            If temLogo Then
                offsets(6) = CInt(ms.Position)
                EscreverAscii(ms,
                              "6 0 obj << /Type /XObject /Subtype /Image /Width " & logoWidthPx.ToString() &
                              " /Height " & logoHeightPx.ToString() &
                              " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " & logoJpeg.Length.ToString() &
                              " >> stream" & vbLf)
                ms.Write(logoJpeg, 0, logoJpeg.Length)
                EscreverAscii(ms, vbLf & "endstream endobj" & vbLf)
                maxObj = 6
            End If

            Dim xrefPos As Integer = CInt(ms.Position)
            Dim sbXref As New StringBuilder()
            sbXref.AppendLine("xref")
            sbXref.AppendLine("0 " & (maxObj + 1).ToString())
            sbXref.AppendLine("0000000000 65535 f ")

            For i As Integer = 1 To maxObj
                sbXref.AppendLine(offsets(i).ToString("0000000000") & " 00000 n ")
            Next

            sbXref.AppendLine("trailer << /Size " & (maxObj + 1).ToString() & " /Root 1 0 R >>")
            sbXref.AppendLine("startxref")
            sbXref.AppendLine(xrefPos.ToString())
            sbXref.AppendLine("%%EOF")

            Dim bXref As Byte() = enc.GetBytes(sbXref.ToString())
            ms.Write(bXref, 0, bXref.Length)

            File.WriteAllBytes(pCaminho, ms.ToArray())
        End Using

    End Sub

    Private Shared Sub EscreverAscii(pStream As Stream, pTexto As String)
        Dim b As Byte() = Encoding.ASCII.GetBytes(pTexto)
        pStream.Write(b, 0, b.Length)
    End Sub

    Private Shared Function ConverterImagemParaJpeg(pLogoPath As String,
                                                     ByRef pWidth As Integer,
                                                     ByRef pHeight As Integer) As Byte()

        Try
            Using img As Image = Image.FromFile(pLogoPath)
                pWidth = img.Width
                pHeight = img.Height

                Using ms As New MemoryStream()
                    img.Save(ms, ImageFormat.Jpeg)
                    Return ms.ToArray()
                End Using
            End Using
        Catch
            pWidth = 0
            pHeight = 0
            Return Nothing
        End Try

    End Function

    Private Shared Function EscapePdfText(pTexto As String) As String
        If pTexto Is Nothing Then Return ""

        Return pTexto.Replace("\", "\\").
                      Replace("(", "\(").
                      Replace(")", "\)")
    End Function

End Class
