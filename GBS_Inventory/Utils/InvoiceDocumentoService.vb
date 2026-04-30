Imports System.Globalization
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Text
Imports GBS_Inventory.Models
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.draw
Imports PdfFont = iTextSharp.text.Font
Imports PdfRect = iTextSharp.text.Rectangle

''' <summary>
''' Gera arquivos de invoice em Word (.doc html) e PDF (iTextSharp).
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

        PdfInvoiceWriter.GerarPdf(pPdfPath, pInvoice, pCliente, pItens, logoEfetivo)

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

Friend Class PdfInvoiceWriter

    Private Shared ReadOnly ColorGrayHeader As BaseColor = New BaseColor(55, 65, 81)
    Private Shared ReadOnly ColorRowAlt As BaseColor = New BaseColor(249, 250, 251)
    Private Shared ReadOnly ColorSeparator As BaseColor = New BaseColor(209, 213, 219)
    Private Shared ReadOnly ColorMuted As BaseColor = New BaseColor(107, 114, 128)
    Private Shared ReadOnly ColorDark As BaseColor = New BaseColor(31, 41, 55)
    Private Shared ReadOnly ColorLabel As BaseColor = New BaseColor(75, 85, 99)

    Public Shared Sub GerarPdf(pCaminho As String,
                               pInvoice As Invoice,
                               pCliente As Cliente,
                               pItens As List(Of InvoiceItem),
                               pLogoPath As String)

        Dim bf As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)
        Dim bfBold As BaseFont = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)

        Dim fNormal As New PdfFont(bf, 9, PdfFont.NORMAL, BaseColor.BLACK)
        Dim fBold As New PdfFont(bfBold, 9, PdfFont.BOLD, BaseColor.BLACK)
        Dim fSmall As New PdfFont(bf, 8, PdfFont.NORMAL, ColorLabel)
        Dim fTitle As New PdfFont(bfBold, 14, PdfFont.BOLD, ColorDark)
        Dim fSubtitle As New PdfFont(bf, 11, PdfFont.NORMAL, ColorDark)
        Dim fTableHeader As New PdfFont(bfBold, 9, PdfFont.BOLD, BaseColor.WHITE)
        Dim fTotalLabel As New PdfFont(bfBold, 10, PdfFont.BOLD, BaseColor.BLACK)

        Dim doc As New Document(PageSize.A4, 40, 40, 70, 60)
        Dim footerEvent As New InvoiceFooterEvent(pItens.Count, bf)

        Using fs As New FileStream(pCaminho, FileMode.Create)
            Dim writer As PdfWriter = PdfWriter.GetInstance(doc, fs)
            writer.PageEvent = footerEvent
            doc.Open()

            ' ── Header: logo + title ──────────────────────────────────────
            Dim tblHeader As New PdfPTable(2)
            tblHeader.WidthPercentage = 100
            tblHeader.SetWidths(New Single() {0.45F, 0.55F})
            tblHeader.DefaultCell.Border = PdfRect.NO_BORDER
            tblHeader.SpacingAfter = 6

            Dim cellLogo As New PdfPCell()
            cellLogo.Border = PdfRect.NO_BORDER
            cellLogo.VerticalAlignment = Element.ALIGN_MIDDLE

            If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
                Try
                    Dim logo As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(pLogoPath)
                    Dim maxW As Single = 130F
                    Dim maxH As Single = 50F
                    If logo.Width > maxW OrElse logo.Height > maxH Then
                        logo.ScaleToFit(maxW, maxH)
                    End If
                    logo.Alignment = Element.ALIGN_LEFT
                    cellLogo.AddElement(logo)
                Catch
                    cellLogo.AddElement(New Phrase("Global Business Solution", fTitle))
                End Try
            Else
                cellLogo.AddElement(New Phrase("Global Business Solution", fTitle))
            End If
            tblHeader.AddCell(cellLogo)

            Dim cellTitle As New PdfPCell()
            cellTitle.Border = PdfRect.NO_BORDER
            cellTitle.HorizontalAlignment = Element.ALIGN_RIGHT
            cellTitle.VerticalAlignment = Element.ALIGN_MIDDLE
            cellTitle.AddElement(New Paragraph("Global Business Solution", fTitle) With {.Alignment = Element.ALIGN_RIGHT})
            cellTitle.AddElement(New Paragraph("INVOICE", fSubtitle) With {.Alignment = Element.ALIGN_RIGHT})
            tblHeader.AddCell(cellTitle)

            doc.Add(tblHeader)

            ' ── Separator ─────────────────────────────────────────────────
            Dim sep As New LineSeparator(0.5F, 100F, ColorSeparator, Element.ALIGN_CENTER, -2)
            doc.Add(New Chunk(sep))
            doc.Add(Chunk.NEWLINE)

            ' ── Invoice info + Bill To ────────────────────────────────────
            Dim tblInfo As New PdfPTable(2)
            tblInfo.WidthPercentage = 100
            tblInfo.SetWidths(New Single() {0.5F, 0.5F})
            tblInfo.DefaultCell.Border = PdfRect.NO_BORDER
            tblInfo.SpacingAfter = 10

            Dim cellInvoiceInfo As New PdfPCell()
            cellInvoiceInfo.Border = PdfRect.NO_BORDER
            cellInvoiceInfo.AddElement(CriarLinhaInfo("Invoice #:", pInvoice.InvoiceNumber, fBold, fNormal))
            cellInvoiceInfo.AddElement(CriarLinhaInfo("Issue Date:", pInvoice.IssueDate.ToString("yyyy-MM-dd"), fBold, fNormal))
            cellInvoiceInfo.AddElement(CriarLinhaInfo("Due Date:", If(pInvoice.DueDate.HasValue, pInvoice.DueDate.Value.ToString("yyyy-MM-dd"), "-"), fBold, fNormal))
            tblInfo.AddCell(cellInvoiceInfo)

            Dim cellBillTo As New PdfPCell()
            cellBillTo.Border = PdfRect.NO_BORDER
            Dim pBillTo As New Paragraph()
            pBillTo.Add(New Chunk("Bill To" & vbLf, fBold))
            pBillTo.Add(New Chunk(If(pCliente.NomeRazao, "") & vbLf, fNormal))
            If Not String.IsNullOrWhiteSpace(pCliente.Documento) Then
                pBillTo.Add(New Chunk(pCliente.Documento & vbLf, fSmall))
            End If
            If Not String.IsNullOrWhiteSpace(pCliente.Email) Then
                pBillTo.Add(New Chunk(pCliente.Email & vbLf, fSmall))
            End If
            If Not String.IsNullOrWhiteSpace(pCliente.Endereco1) Then
                pBillTo.Add(New Chunk((pCliente.Endereco1 & " " & pCliente.Endereco2).Trim() & vbLf, fSmall))
            End If
            If Not String.IsNullOrWhiteSpace(pCliente.Cidade) Then
                Dim localidade As String = (pCliente.Cidade & " - " & pCliente.Estado & " " & pCliente.ZipCode & " / " & pCliente.Pais).Trim()
                pBillTo.Add(New Chunk(localidade, fSmall))
            End If
            cellBillTo.AddElement(pBillTo)
            tblInfo.AddCell(cellBillTo)

            doc.Add(tblInfo)

            ' ── Items Table ───────────────────────────────────────────────
            Dim tblItens As New PdfPTable(4)
            tblItens.WidthPercentage = 100
            tblItens.SetWidths(New Single() {0.55F, 0.1F, 0.175F, 0.175F})
            tblItens.SpacingAfter = 12
            tblItens.HeaderRows = 1

            Dim headerCols As String() = {"Description", "Qty", "Unit (USD)", "Total (USD)"}
            For Each col As String In headerCols
                Dim hCell As New PdfPCell(New Phrase(col, fTableHeader))
                hCell.BackgroundColor = ColorGrayHeader
                hCell.Padding = 5
                hCell.HorizontalAlignment = If(col = "Description", Element.ALIGN_LEFT, Element.ALIGN_RIGHT)
                tblItens.AddCell(hCell)
            Next

            Dim subtotal As Decimal = 0D
            Dim rowIdx As Integer = 0

            For Each item As InvoiceItem In pItens
                item.RecalcularTotal()
                subtotal += item.LineTotalUsd

                Dim bg As BaseColor = If(rowIdx Mod 2 = 0, BaseColor.WHITE, ColorRowAlt)

                Dim cDesc As New PdfPCell(New Phrase(If(item.Description, ""), fNormal))
                cDesc.BackgroundColor = bg : cDesc.Padding = 4 : cDesc.HorizontalAlignment = Element.ALIGN_LEFT
                tblItens.AddCell(cDesc)

                Dim cQty As New PdfPCell(New Phrase(item.Qty.ToString("0.##", CultureInfo.InvariantCulture), fNormal))
                cQty.BackgroundColor = bg : cQty.Padding = 4 : cQty.HorizontalAlignment = Element.ALIGN_RIGHT
                tblItens.AddCell(cQty)

                Dim cUnit As New PdfPCell(New Phrase(item.UnitPriceUsd.ToString("0.00", CultureInfo.InvariantCulture), fNormal))
                cUnit.BackgroundColor = bg : cUnit.Padding = 4 : cUnit.HorizontalAlignment = Element.ALIGN_RIGHT
                tblItens.AddCell(cUnit)

                Dim cLinha As New PdfPCell(New Phrase(item.LineTotalUsd.ToString("0.00", CultureInfo.InvariantCulture), fNormal))
                cLinha.BackgroundColor = bg : cLinha.Padding = 4 : cLinha.HorizontalAlignment = Element.ALIGN_RIGHT
                tblItens.AddCell(cLinha)

                rowIdx += 1
            Next

            doc.Add(tblItens)

            ' ── Totals ────────────────────────────────────────────────────
            Dim totalFinal As Decimal = subtotal - pInvoice.DiscountUsd + pInvoice.ShippingUsd + pInvoice.TaxUsd

            Dim tblTotals As New PdfPTable(2)
            tblTotals.HorizontalAlignment = Element.ALIGN_RIGHT
            tblTotals.TotalWidth = 260
            tblTotals.LockedWidth = True
            tblTotals.SpacingAfter = 14

            AdicionarLinhaTotal(tblTotals, "Subtotal", subtotal.ToString("0.00", CultureInfo.InvariantCulture), False, fNormal)
            AdicionarLinhaTotal(tblTotals, "Discount", "-" & pInvoice.DiscountUsd.ToString("0.00", CultureInfo.InvariantCulture), False, fNormal)
            AdicionarLinhaTotal(tblTotals, "Shipping", pInvoice.ShippingUsd.ToString("0.00", CultureInfo.InvariantCulture), False, fNormal)
            AdicionarLinhaTotal(tblTotals, "Tax", pInvoice.TaxUsd.ToString("0.00", CultureInfo.InvariantCulture), False, fNormal)
            AdicionarLinhaTotal(tblTotals, "TOTAL USD", totalFinal.ToString("0.00", CultureInfo.InvariantCulture), True, fTotalLabel)

            doc.Add(tblTotals)

            ' ── Notes ─────────────────────────────────────────────────────
            If Not String.IsNullOrWhiteSpace(pInvoice.Notes) Then
                doc.Add(New Paragraph("Notes", fBold) With {.SpacingAfter = 4})
                doc.Add(New Paragraph(pInvoice.Notes, fNormal) With {.SpacingAfter = 6})
            End If

            doc.Close()
        End Using

    End Sub

    Private Shared Function CriarLinhaInfo(pLabel As String, pValor As String,
                                           fLabel As PdfFont, fValor As PdfFont) As Paragraph
        Dim p As New Paragraph()
        p.Add(New Chunk(pLabel & " ", fLabel))
        p.Add(New Chunk(If(pValor, ""), fValor))
        p.SpacingAfter = 2
        Return p
    End Function

    Private Shared Sub AdicionarLinhaTotal(tbl As PdfPTable, pLabel As String, pValor As String,
                                           pDestaque As Boolean, fUsada As PdfFont)
        Dim bordaTipo As Integer = If(pDestaque, PdfRect.TOP_BORDER, PdfRect.NO_BORDER)

        Dim cLabel As New PdfPCell(New Phrase(pLabel, fUsada))
        cLabel.HorizontalAlignment = Element.ALIGN_LEFT
        cLabel.Padding = 4
        cLabel.Border = bordaTipo
        tbl.AddCell(cLabel)

        Dim cValor As New PdfPCell(New Phrase(pValor, fUsada))
        cValor.HorizontalAlignment = Element.ALIGN_RIGHT
        cValor.Padding = 4
        cValor.Border = bordaTipo
        tbl.AddCell(cValor)
    End Sub

End Class

Friend Class InvoiceFooterEvent
    Inherits PdfPageEventHelper

    Private ReadOnly _itemCount As Integer
    Private ReadOnly _bf As BaseFont

    Public Sub New(pItemCount As Integer, pBf As BaseFont)
        _itemCount = pItemCount
        _bf = pBf
    End Sub

    Public Overrides Sub OnEndPage(writer As PdfWriter, document As Document)
        Dim cb As PdfContentByte = writer.DirectContent
        Dim texto As String = "Items: " & _itemCount.ToString() & " | Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm")
        Dim xCenter As Single = (document.Left + document.Right) / 2
        Dim yPos As Single = document.Bottom - 20

        cb.BeginText()
        cb.SetFontAndSize(_bf, 8)
        cb.SetColorFill(New BaseColor(107, 114, 128))
        cb.ShowTextAligned(Element.ALIGN_CENTER, texto, xCenter, yPos, 0)
        cb.EndText()
    End Sub

End Class
