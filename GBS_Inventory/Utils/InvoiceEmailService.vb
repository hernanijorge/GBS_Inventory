Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Text

''' <summary>
''' Envio de email de invoice com anexos (pdf/word/logo).
''' Configuracao via App.config.
''' </summary>
Public Class InvoiceEmailService

    Public Shared Sub EnviarInvoice(pEmailDestino As String,
                                    pAssunto As String,
                                    pCorpo As String,
                                    pAnexos As List(Of String),
                                    Optional pCC As String = "",
                                    Optional pIsHtml As Boolean = False,
                                    Optional pLogoInlinePath As String = "",
                                    Optional pAdicionarAssinatura As Boolean = False)

        Dim smtpHost As String = ConfigurationManager.AppSettings("SmtpHost")
        Dim smtpPort As Integer = ObterInt("SmtpPort", 587)
        Dim smtpSsl As Boolean = ObterBool("SmtpEnableSsl", True)
        Dim smtpUser As String = ConfigurationManager.AppSettings("SmtpUser")
        Dim smtpPass As String = ConfigurationManager.AppSettings("SmtpPassword")
        Dim emailFrom As String = ConfigurationManager.AppSettings("SmtpFrom")

        If String.IsNullOrWhiteSpace(smtpHost) OrElse String.IsNullOrWhiteSpace(emailFrom) Then
            Throw New Exception("SMTP nao configurado. Preencha SmtpHost e SmtpFrom no App.config.")
        End If

        If String.IsNullOrWhiteSpace(pEmailDestino) Then
            Throw New Exception("Email do cliente nao informado.")
        End If

        Using msg As New MailMessage()
            msg.From = New MailAddress(emailFrom)
            msg.To.Add(pEmailDestino)
            If Not String.IsNullOrWhiteSpace(pCC) AndAlso pCC.Contains("@") Then msg.CC.Add(pCC)
            msg.Subject = pAssunto

            If pIsHtml OrElse pAdicionarAssinatura Then
                ConfigurarCorpoHtml(msg, pCorpo, pLogoInlinePath, pAdicionarAssinatura)
            Else
                msg.Body       = pCorpo
                msg.IsBodyHtml = False
            End If

            If pAnexos IsNot Nothing Then
                For Each arq As String In pAnexos
                    If Not String.IsNullOrWhiteSpace(arq) AndAlso IO.File.Exists(arq) Then
                        msg.Attachments.Add(New Attachment(arq))
                    End If
                Next
            End If

            Using cli As New SmtpClient(smtpHost, smtpPort)
                cli.EnableSsl = smtpSsl

                If Not String.IsNullOrWhiteSpace(smtpUser) Then
                    cli.Credentials = New NetworkCredential(smtpUser, smtpPass)
                End If

                cli.Send(msg)
            End Using
        End Using

    End Sub

    Private Shared Sub ConfigurarCorpoHtml(pMensagem As MailMessage,
                                           pCorpo As String,
                                           pLogoInlinePath As String,
                                           pAdicionarAssinatura As Boolean)

        Dim htmlBody As String = If(pCorpo, "")
        Dim logoPath As String = ObterLogoPath(pLogoInlinePath)
        Dim logoCid As String = ""

        If Not String.IsNullOrWhiteSpace(logoPath) AndAlso File.Exists(logoPath) Then
            logoCid = "gbs-logo-" & Guid.NewGuid().ToString("N")
        End If

        If pAdicionarAssinatura Then
            Dim assinatura As String = MontarAssinaturaHtml(If(String.IsNullOrWhiteSpace(logoCid), "", "cid:" & logoCid))
            Dim posBody As Integer = htmlBody.IndexOf("</body>", StringComparison.OrdinalIgnoreCase)
            If posBody >= 0 Then
                htmlBody = htmlBody.Insert(posBody, assinatura)
            Else
                htmlBody &= assinatura
            End If
        End If

        pMensagem.Body = htmlBody
        pMensagem.IsBodyHtml = True

        Dim view As AlternateView = AlternateView.CreateAlternateViewFromString(
            htmlBody, Encoding.UTF8, MediaTypeNames.Text.Html)

        If Not String.IsNullOrWhiteSpace(logoCid) Then
            Dim logoResource As New LinkedResource(logoPath, ObterMime(logoPath))
            logoResource.ContentId = logoCid
            logoResource.TransferEncoding = TransferEncoding.Base64
            logoResource.ContentType.Name = Path.GetFileName(logoPath)
            view.LinkedResources.Add(logoResource)
        End If

        pMensagem.AlternateViews.Add(view)

    End Sub

    Private Shared Function MontarAssinaturaHtml(pLogoSrc As String) As String

        Dim logoHtml As String = ""
        If Not String.IsNullOrWhiteSpace(pLogoSrc) Then
            logoHtml =
                "<td style='padding-right:15px;vertical-align:middle;'>" &
                "<img src='" & pLogoSrc & "' style='height:50px;width:auto;border:0;display:block;' alt='GBS' />" &
                "</td>"
        End If

        Return "<br/><hr style='border:none;border-top:1px solid #e5e7eb;margin:20px 0;'/>" &
            "<table cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;'>" &
            "<tr>" &
            logoHtml &
            "<td style='vertical-align:middle;'>" &
            "<div style='color:#111827;font-size:10pt;margin-bottom:8px;'>Best regards,</div>" &
            "<div style='font-weight:600;color:#111827;font-size:11pt;margin-bottom:8px;'>Partha Gajula</div>" &
            "<div style='font-weight:600;color:#111827;font-size:10pt;'>Global Business Solutions Inc.</div>" &
            "<div style='color:#6b7280;font-size:9pt;'>19 Weston St., Lexington, MA 02421</div>" &
            "<div style='color:#6b7280;font-size:9pt;'>" &
            "<a href='https://www.gbsolinc.com/' style='color:#6b7280;text-decoration:none;'>www.gbsolinc.com</a>" &
            " | <a href='mailto:contact@gbsolinc.com' style='color:#6b7280;text-decoration:none;'>contact@gbsolinc.com</a>" &
            " | (339) 234-6101</div>" &
            "</td></tr></table>"

    End Function

    Private Shared Function ObterLogoPath(pLogoPath As String) As String

        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Return pLogoPath
        End If

        Dim configLogo As String = ConfigurationManager.AppSettings("InvoiceLogoPath")
        If Not String.IsNullOrWhiteSpace(configLogo) AndAlso File.Exists(configLogo) Then
            Return configLogo
        End If

        For Each candidato As String In {"C:\GBS\logo.png", "C:\GBS\logo.jpg", "C:\GBS\logo.jpeg"}
            If File.Exists(candidato) Then Return candidato
        Next

        Return ""

    End Function

    Private Shared Function ObterMime(pPath As String) As String

        Select Case Path.GetExtension(pPath).ToLowerInvariant()
            Case ".png"        : Return "image/png"
            Case ".jpg", ".jpeg" : Return MediaTypeNames.Image.Jpeg
            Case ".gif"        : Return MediaTypeNames.Image.Gif
            Case Else          : Return MediaTypeNames.Image.Jpeg
        End Select

    End Function

    Private Shared Function ObterInt(pKey As String, pPadrao As Integer) As Integer

        Dim valor As String = ConfigurationManager.AppSettings(pKey)
        Dim n As Integer

        If Integer.TryParse(valor, n) Then
            Return n
        End If

        Return pPadrao

    End Function

    Private Shared Function ObterBool(pKey As String, pPadrao As Boolean) As Boolean

        Dim valor As String = ConfigurationManager.AppSettings(pKey)
        Dim b As Boolean

        If Boolean.TryParse(valor, b) Then
            Return b
        End If

        Return pPadrao

    End Function

End Class
