Imports System.Configuration
Imports System.Net
Imports System.Threading

''' <summary>
''' GBS_DiagReceiver — recebe os resultados do diagnostico de laptops e grava no
''' Oracle XE (PACK_DIAGNOSTICO). Tambem serve a pasta .\diag para o boot pela rede.
'''
'''   GBS_DiagReceiver.exe              inicia o servidor HTTP
'''   GBS_DiagReceiver.exe --selftest   grava e le de volta BLOB + CLOB no Oracle e sai
''' </summary>
Public Module Program

    Private _listener As HttpListener

    Public Function Main(args As String()) As Integer

        Console.OutputEncoding = Encoding.UTF8
        Console.Title = "GBS DiagReceiver"

        If args.Any(Function(a) a.Equals("--selftest", StringComparison.OrdinalIgnoreCase)) Then
            Return SelfTest()
        End If

        Dim porta As Integer = LerInt("Port", 8090)
        Dim token As String = ConfigurationManager.AppSettings("Token")
        Dim maxBody As Long = CLng(LerInt("MaxBodyMB", 50)) * 1024 * 1024

        If String.IsNullOrWhiteSpace(token) Then
            ReceiverLog.Erro("Token nao configurado. Crie secrets.config ao lado do .exe (ver README). Abortando.")
            Return 1
        End If

        Dim api As New DiagnosticoApi(token, maxBody)
        Dim estaticos As New StaticFiles(If(ConfigurationManager.AppSettings("DiagFolder"), "diag"))

        _listener = IniciarListener(porta)
        If _listener Is Nothing Then Return 1

        ReceiverLog.Info("Pasta /diag: " & estaticos.Raiz)
        ReceiverLog.Info("Ctrl+C para encerrar.")

        AddHandler Console.CancelKeyPress, Sub(s, e)
                                               e.Cancel = True
                                               ReceiverLog.Info("Encerrando...")
                                               Try : _listener.Stop() : Catch : End Try
                                           End Sub

        While _listener.IsListening
            Dim ctx As HttpListenerContext
            Try
                ctx = _listener.GetContext()
            Catch ex As Exception When Not _listener.IsListening
                Exit While
            Catch ex As HttpListenerException
                ReceiverLog.Erro("GetContext: " & ex.Message)
                Continue While
            End Try
            ThreadPool.QueueUserWorkItem(Sub() Atender(ctx, api, estaticos))
        End While

        Return 0

    End Function

    ''' <summary>
    ''' Tenta http://+:porta/ (aceita a rede; exige urlacl ou admin). Se negar acesso,
    ''' cai para http://localhost:porta/ e avisa — so aceita conexoes desta maquina.
    ''' </summary>
    Private Function IniciarListener(pPorta As Integer) As HttpListener

        For Each prefixo As String In {$"http://+:{pPorta}/", $"http://localhost:{pPorta}/"}
            Dim l As New HttpListener()
            l.Prefixes.Add(prefixo)
            Try
                l.Start()
                ReceiverLog.Ok("Escutando em " & prefixo)
                If prefixo.Contains("localhost") Then
                    ReceiverLog.Aviso("Sem permissao para http://+:" & pPorta & "/ — aceitando SOMENTE conexoes locais.")
                    ReceiverLog.Aviso("Para liberar a rede, como Administrador: netsh http add urlacl url=http://+:" & pPorta & "/ user=Everyone")
                End If
                Return l
            Catch ex As HttpListenerException When ex.ErrorCode = 5
                l.Close()
            Catch ex As HttpListenerException
                l.Close()
                ReceiverLog.Erro($"Nao foi possivel escutar em {prefixo}: {ex.Message} (porta em uso?)")
                Return Nothing
            End Try
        Next
        ReceiverLog.Erro("Acesso negado ao HttpListener.")
        Return Nothing

    End Function

    Private Sub Atender(ctx As HttpListenerContext, api As DiagnosticoApi, estaticos As StaticFiles)

        Dim metodo As String = ctx.Request.HttpMethod
        Dim caminho As String = ctx.Request.Url.AbsolutePath

        Try
            If caminho.Equals("/api/diagnostico", StringComparison.OrdinalIgnoreCase) OrElse
               caminho.Equals("/api/diagnostico/", StringComparison.OrdinalIgnoreCase) Then
                If metodo <> "POST" Then
                    ResponderJson(ctx, 405, "{""error"":""use POST""}")
                Else
                    api.Processar(ctx)
                End If

            ElseIf caminho.StartsWith("/diag/", StringComparison.OrdinalIgnoreCase) Then
                If metodo <> "GET" AndAlso metodo <> "HEAD" Then
                    ResponderTexto(ctx, 405, "use GET")
                Else
                    estaticos.Processar(ctx, caminho.Substring("/diag/".Length))
                End If

            ElseIf caminho = "/" OrElse caminho.Equals("/health", StringComparison.OrdinalIgnoreCase) Then
                ResponderJson(ctx, 200, "{""status"":""ok"",""service"":""GBS_DiagReceiver""}")

            Else
                ResponderTexto(ctx, 404, "nao encontrado")
            End If

        Catch ex As Exception
            ReceiverLog.Erro($"{metodo} {caminho} de {ctx.Request.RemoteEndPoint}: {ex.Message}")
            Try
                ResponderJson(ctx, 500, "{""error"":" &
                    New Web.Script.Serialization.JavaScriptSerializer().Serialize(ex.Message) & "}")
            Catch
                ' resposta ja iniciada / cliente desconectou
            End Try
        End Try

    End Sub

#Region "Respostas"

    Public Sub ResponderJson(ctx As HttpListenerContext, pStatus As Integer, pJson As String)
        Responder(ctx, pStatus, "application/json; charset=utf-8", pJson)
    End Sub

    Public Sub ResponderTexto(ctx As HttpListenerContext, pStatus As Integer, pTexto As String)
        Responder(ctx, pStatus, "text/plain; charset=utf-8", pTexto)
    End Sub

    Private Sub Responder(ctx As HttpListenerContext, pStatus As Integer, pTipo As String, pCorpo As String)
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(pCorpo)
        ctx.Response.StatusCode = pStatus
        ctx.Response.ContentType = pTipo
        ctx.Response.ContentLength64 = bytes.Length
        ctx.Response.OutputStream.Write(bytes, 0, bytes.Length)
        ctx.Response.OutputStream.Close()
    End Sub

#End Region

#Region "Self-test"

    ''' <summary>
    ''' Grava um diagnostico de teste com PDF (BLOB ~300 KB) e JSON (CLOB > 32 KB, com acentos)
    ''' via PROC_INSERT, le de volta por PROC_GET_PDF / PROC_GET_JSON, compara e apaga.
    ''' </summary>
    Private Function SelfTest() As Integer

        ReceiverLog.Info("SELFTEST: BLOB + CLOB no Oracle (" &
                         GetType(Oracle.ManagedDataAccess.Client.OracleConnection).Assembly.GetName().Version.ToString() & ")")

        Dim repo As New DiagnosticoRepository()
        Dim pdf(300000) As Byte
        Dim rnd As New Random(42)
        rnd.NextBytes(pdf)

        Dim detalhes As New StringBuilder()
        For i As Integer = 1 To 3000
            detalhes.Append("Teste de memória — ação ").Append(i).Append("; ")
        Next

        Dim json As String =
            "{""tool_version"":""selftest"",""origin"":""WINDOWS"",""stamp"":""" & DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'") & """," &
            """verdict"":""PASS"",""device"":{""internal_uid"":""SELFTEST"",""service_tag"":""SELFTEST""}," &
            """checks"":[{""test"":""memtest"",""result"":""PASS"",""details"":""" & detalhes.ToString() & """}]," &
            """pdf_filename"":""selftest.pdf"",""pdf_base64"":""" & Convert.ToBase64String(pdf) & """}"

        Dim id As Long = 0
        Try
            Dim reg As DiagnosticoRepository.Registro = DiagnosticoApi.MontarRegistro(json)
            Dim res = repo.Inserir(reg)
            id = res.Id
            ReceiverLog.Info($"Gravado ID={id} (CLOB {reg.ResultJson.Length:N0} chars, BLOB {reg.PdfBytes.Length:N0} bytes)")

            Dim lido = repo.ObterPdf(id)
            Dim jsonLido As String = repo.ObterJson(id)

            Dim okPdf As Boolean = lido.Bytes IsNot Nothing AndAlso lido.Bytes.SequenceEqual(reg.PdfBytes)
            Dim okJson As Boolean = String.Equals(jsonLido, reg.ResultJson, StringComparison.Ordinal)
            Dim okNome As Boolean = lido.Nome = "selftest.pdf"

            ReceiverLog.Info($"BLOB lido: {If(lido.Bytes Is Nothing, 0, lido.Bytes.Length):N0} bytes, igual={okPdf}, nome={lido.Nome}")
            ReceiverLog.Info($"CLOB lido: {If(jsonLido Is Nothing, 0, jsonLido.Length):N0} chars, igual={okJson}")

            If okPdf AndAlso okJson AndAlso okNome Then
                ReceiverLog.Ok("SELFTEST PASSOU")
                Return 0
            End If
            ReceiverLog.Erro("SELFTEST FALHOU")
            Return 2

        Catch ex As Exception
            ReceiverLog.Erro("SELFTEST FALHOU: " & ex.ToString())
            Return 2
        Finally
            If id > 0 Then
                Try
                    repo.Excluir(id)
                    ReceiverLog.Info($"Registro de teste ID={id} removido.")
                Catch ex As Exception
                    ReceiverLog.Aviso("Nao foi possivel remover o registro de teste: " & ex.Message)
                End Try
            End If
        End Try

    End Function

#End Region

    Private Function LerInt(pChave As String, pPadrao As Integer) As Integer
        Dim n As Integer
        If Integer.TryParse(ConfigurationManager.AppSettings(pChave), n) AndAlso n > 0 Then Return n
        Return pPadrao
    End Function

End Module
