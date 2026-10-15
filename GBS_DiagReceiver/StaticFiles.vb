Imports System.Net

''' <summary>
''' GET/HEAD /diag/&lt;arquivo&gt;: serve arquivos estaticos da pasta configurada
''' (DiagFolder, padrao .\diag) para o boot pela rede. Subpastas sao permitidas;
''' qualquer caminho que saia da pasta raiz e recusado.
''' </summary>
Public Class StaticFiles

    Private ReadOnly _raiz As String

    Public Sub New(pPasta As String)
        If Not Path.IsPathRooted(pPasta) Then
            pPasta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, pPasta)
        End If
        _raiz = Path.GetFullPath(pPasta).TrimEnd(Path.DirectorySeparatorChar) & Path.DirectorySeparatorChar
        If Not Directory.Exists(_raiz) Then Directory.CreateDirectory(_raiz)
    End Sub

    Public ReadOnly Property Raiz As String
        Get
            Return _raiz
        End Get
    End Property

    Public Sub Processar(ctx As HttpListenerContext, pRelativo As String)

        Dim origem As String = ctx.Request.RemoteEndPoint.ToString()
        Dim rel As String = Uri.UnescapeDataString(pRelativo).Replace("/"c, Path.DirectorySeparatorChar)

        Dim caminho As String
        Try
            caminho = Path.GetFullPath(Path.Combine(_raiz, rel))
        Catch ex As Exception
            Program.ResponderTexto(ctx, 400, "caminho invalido")
            Return
        End Try

        If Not caminho.StartsWith(_raiz, StringComparison.OrdinalIgnoreCase) OrElse Not File.Exists(caminho) Then
            ReceiverLog.Aviso($"404 GET /diag/{pRelativo} de {origem}")
            Program.ResponderTexto(ctx, 404, "nao encontrado")
            Return
        End If

        Dim info As New FileInfo(caminho)
        ctx.Response.StatusCode = 200
        ctx.Response.ContentType = TipoConteudo(info.Extension)
        ctx.Response.ContentLength64 = info.Length

        If ctx.Request.HttpMethod = "HEAD" Then
            ctx.Response.OutputStream.Close()
            Return
        End If

        Using fs As New FileStream(caminho, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024)
            fs.CopyTo(ctx.Response.OutputStream, 1024 * 1024)
        End Using
        ctx.Response.OutputStream.Close()

        ReceiverLog.Info($"GET /diag/{pRelativo} -> {info.Length:N0} bytes para {origem}")

    End Sub

    Private Shared Function TipoConteudo(pExt As String) As String
        Select Case pExt.ToLowerInvariant()
            Case ".txt", ".cfg", ".ipxe", ".sh", ".conf" : Return "text/plain; charset=utf-8"
            Case ".json" : Return "application/json"
            Case ".html", ".htm" : Return "text/html; charset=utf-8"
            Case ".iso" : Return "application/x-iso9660-image"
            Case Else : Return "application/octet-stream"
        End Select
    End Function

End Class
