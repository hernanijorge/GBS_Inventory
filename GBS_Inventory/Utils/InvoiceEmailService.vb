Imports System.Configuration
Imports System.Net
Imports System.Net.Mail

''' <summary>
''' Envio de email de invoice com anexos (pdf/word/logo).
''' Configuracao via App.config.
''' </summary>
Public Class InvoiceEmailService

    Public Shared Sub EnviarInvoice(pEmailDestino As String,
                                    pAssunto As String,
                                    pCorpo As String,
                                    pAnexos As List(Of String))

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
            msg.Subject = pAssunto
            msg.Body = pCorpo
            msg.IsBodyHtml = False

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
