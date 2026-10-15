Imports System.Globalization
Imports System.Net
Imports System.Web.Script.Serialization

''' <summary>
''' POST /api/diagnostico
''' Header obrigatorio: X-GBS-Token. Corpo JSON:
''' {tool_version, origin, stamp, verdict,
'''  device:{internal_uid, service_tag, manufacturer, model, bios, cpu, ram_gib, storage,
'''          battery_health, battery_cycles, ssd_life, technician},
'''  checks:[{test, result, details}], pdf_filename, pdf_base64}
''' Resposta 200: {"id":N,"equipamento_encontrado":true|false}
''' </summary>
Public Class DiagnosticoApi

    Private Shared ReadOnly VerdictsValidos As String() = {"PASS", "PASS_WITH_WARNINGS", "FAIL", "INCOMPLETE"}
    Private Shared ReadOnly OrigensValidas As String() = {"UBUNTU-USB", "UBUNTU-PXE", "WINDOWS"}

    Private Const SeparadorItens As String = " | "
    Private Const TamanhoMaxItens As Integer = 1000

    Private ReadOnly _token As String
    Private ReadOnly _maxBodyBytes As Long
    Private ReadOnly _repo As New DiagnosticoRepository()

    ''' <summary>Erro de validacao do cliente (vira HTTP 4xx).</summary>
    Public Class ErroRequisicao
        Inherits Exception
        Public ReadOnly Status As Integer
        Public Sub New(pStatus As Integer, pMsg As String)
            MyBase.New(pMsg)
            Status = pStatus
        End Sub
    End Class

    Public Sub New(pToken As String, pMaxBodyBytes As Long)
        _token = If(pToken, "")
        _maxBodyBytes = pMaxBodyBytes
    End Sub

    Public Sub Processar(ctx As HttpListenerContext)

        Dim req As HttpListenerRequest = ctx.Request
        Dim origem As String = req.RemoteEndPoint.ToString()

        Try
            If Not TokenValido(req.Headers("X-GBS-Token")) Then
                ReceiverLog.Aviso($"401 token invalido de {origem}")
                Program.ResponderJson(ctx, 401, "{""error"":""token invalido ou ausente (X-GBS-Token)""}")
                Return
            End If

            Dim corpo As String = LerCorpo(req)
            Dim reg As DiagnosticoRepository.Registro = MontarRegistro(corpo)

            Dim res = _repo.Inserir(reg)
            Dim encontrado As Boolean = res.IdEquipamento.HasValue

            ReceiverLog.Ok($"POST de {origem}: ID={res.Id} UID={reg.InternalUid} TAG={reg.ServiceTag} " &
                           $"VERDICT={reg.Verdict} PDF={If(reg.PdfBytes Is Nothing, 0, reg.PdfBytes.Length)} bytes " &
                           If(encontrado, $"-> equipamento {res.IdEquipamento.Value}", "-> equipamento NAO encontrado"))

            Program.ResponderJson(ctx, 200,
                $"{{""id"":{res.Id},""equipamento_encontrado"":{If(encontrado, "true", "false")}}}")

        Catch ex As ErroRequisicao
            ReceiverLog.Aviso($"{ex.Status} de {origem}: {ex.Message}")
            Program.ResponderJson(ctx, ex.Status, "{""error"":" & New JavaScriptSerializer().Serialize(ex.Message) & "}")
        End Try

    End Sub

#Region "Parsing"

    Private Function LerCorpo(req As HttpListenerRequest) As String
        If Not req.HasEntityBody Then Throw New ErroRequisicao(400, "corpo vazio")
        If req.ContentLength64 > _maxBodyBytes Then
            Throw New ErroRequisicao(413, $"corpo maior que o limite ({_maxBodyBytes \ (1024 * 1024)} MB)")
        End If

        Using ms As New MemoryStream()
            Dim buf(81919) As Byte
            Dim n As Integer
            Do
                n = req.InputStream.Read(buf, 0, buf.Length)
                If n <= 0 Then Exit Do
                ms.Write(buf, 0, n)
                If ms.Length > _maxBodyBytes Then
                    Throw New ErroRequisicao(413, $"corpo maior que o limite ({_maxBodyBytes \ (1024 * 1024)} MB)")
                End If
            Loop
            Return Encoding.UTF8.GetString(ms.ToArray())
        End Using
    End Function

    ''' <summary>Converte o JSON recebido no registro a gravar (publico para o --selftest).</summary>
    Public Shared Function MontarRegistro(pCorpo As String) As DiagnosticoRepository.Registro

        Dim ser As New JavaScriptSerializer() With {.MaxJsonLength = Integer.MaxValue, .RecursionLimit = 64}
        Dim raiz As Dictionary(Of String, Object)
        Try
            raiz = TryCast(ser.DeserializeObject(pCorpo), Dictionary(Of String, Object))
        Catch ex As Exception
            Throw New ErroRequisicao(400, "JSON invalido: " & ex.Message)
        End Try
        If raiz Is Nothing Then Throw New ErroRequisicao(400, "JSON deve ser um objeto")

        Dim device As Dictionary(Of String, Object) = TryCast(Valor(raiz, "device"), Dictionary(Of String, Object))
        If device Is Nothing Then device = New Dictionary(Of String, Object)()

        Dim reg As New DiagnosticoRepository.Registro()

        reg.Verdict = Texto(raiz, "verdict").ToUpperInvariant()
        If Not VerdictsValidos.Contains(reg.Verdict) Then
            Throw New ErroRequisicao(400, "verdict invalido: '" & reg.Verdict & "' (use " & String.Join("/", VerdictsValidos) & ")")
        End If

        reg.Origin = Texto(raiz, "origin").ToUpperInvariant()
        If reg.Origin.Length > 0 AndAlso Not OrigensValidas.Contains(reg.Origin) Then
            Throw New ErroRequisicao(400, "origin invalido: '" & reg.Origin & "' (use " & String.Join("/", OrigensValidas) & ")")
        End If

        reg.ToolVersion = Texto(raiz, "tool_version")
        reg.StampUtc = Texto(raiz, "stamp")

        reg.InternalUid = Texto(device, "internal_uid")
        reg.ServiceTag = Texto(device, "service_tag")
        reg.Manufacturer = Texto(device, "manufacturer")
        reg.Model = Texto(device, "model")
        reg.Cpu = Texto(device, "cpu")
        reg.RamGib = Numero(device, "ram_gib")
        reg.Storage = Texto(device, "storage")
        reg.BatteryHealth = Numero(device, "battery_health")
        reg.BatteryCycles = Numero(device, "battery_cycles")
        reg.SsdLife = Numero(device, "ssd_life")
        reg.Technician = Texto(device, "technician")

        If reg.InternalUid.Length = 0 AndAlso reg.ServiceTag.Length = 0 Then
            Throw New ErroRequisicao(400, "device.internal_uid ou device.service_tag e obrigatorio")
        End If

        ' Itens por resultado (nomes dos checks separados por " | ")
        Dim falhas As New List(Of String), avisos As New List(Of String), naoTestados As New List(Of String)
        Dim checks As Object = Valor(raiz, "checks")
        If TypeOf checks Is Object() Then
            For Each item As Object In CType(checks, Object())
                Dim chk As Dictionary(Of String, Object) = TryCast(item, Dictionary(Of String, Object))
                If chk Is Nothing Then Continue For
                Dim nome As String = Texto(chk, "test")
                If nome.Length = 0 Then Continue For
                Select Case NormalizarResultado(Texto(chk, "result"))
                    Case "FAIL" : falhas.Add(nome)
                    Case "WARN" : avisos.Add(nome)
                    Case "NOT_TESTED" : naoTestados.Add(nome)
                End Select
            Next
        End If
        reg.FailItems = JuntarItens(falhas)
        reg.WarnItems = JuntarItens(avisos)
        reg.NotTestedItems = JuntarItens(naoTestados)

        ' PDF
        reg.PdfFilename = Texto(raiz, "pdf_filename")
        Dim b64 As String = Texto(raiz, "pdf_base64")
        If b64.Length > 0 Then
            Dim virgula As Integer = b64.IndexOf(","c)
            If b64.StartsWith("data:", StringComparison.OrdinalIgnoreCase) AndAlso virgula > 0 Then b64 = b64.Substring(virgula + 1)
            Try
                reg.PdfBytes = Convert.FromBase64String(b64)
            Catch ex As FormatException
                Throw New ErroRequisicao(400, "pdf_base64 invalido")
            End Try
            If reg.PdfFilename.Length = 0 Then
                reg.PdfFilename = "diag_" & If(reg.InternalUid.Length > 0, reg.InternalUid, reg.ServiceTag) & ".pdf"
            End If
        End If

        ' RESULT_JSON = JSON inteiro sem o pdf_base64
        raiz.Remove("pdf_base64")
        reg.ResultJson = ser.Serialize(raiz)

        Return reg

    End Function

    Private Shared Function NormalizarResultado(pResultado As String) As String
        Dim r As String = pResultado.ToUpperInvariant().Replace("-", "_").Replace(" ", "_")
        Select Case r
            Case "FAIL", "FAILED" : Return "FAIL"
            Case "WARN", "WARNING", "PASS_WITH_WARNINGS" : Return "WARN"
            Case "NOT_TESTED", "NOTTESTED", "SKIPPED", "SKIP" : Return "NOT_TESTED"
            Case Else : Return r
        End Select
    End Function

    Private Shared Function JuntarItens(pItens As List(Of String)) As String
        Dim s As String = String.Join(SeparadorItens, pItens)
        Return If(s.Length > TamanhoMaxItens, s.Substring(0, TamanhoMaxItens), s)
    End Function

    Private Shared Function Valor(pDic As Dictionary(Of String, Object), pChave As String) As Object
        Dim v As Object = Nothing
        If pDic.TryGetValue(pChave, v) Then Return v
        Return Nothing
    End Function

    Private Shared Function Texto(pDic As Dictionary(Of String, Object), pChave As String) As String
        Dim v As Object = Valor(pDic, pChave)
        If v Is Nothing Then Return ""
        If TypeOf v Is Decimal OrElse TypeOf v Is Double OrElse TypeOf v Is Integer OrElse TypeOf v Is Long Then
            Return Convert.ToString(v, CultureInfo.InvariantCulture)
        End If
        If TypeOf v Is Dictionary(Of String, Object) OrElse TypeOf v Is Object() Then
            Return New JavaScriptSerializer().Serialize(v)
        End If
        Return v.ToString().Trim()
    End Function

    ''' <summary>Aceita numero JSON ou string numerica ("87", "87.5", "87%"); vazio/invalido = Nothing.</summary>
    Private Shared Function Numero(pDic As Dictionary(Of String, Object), pChave As String) As Decimal?
        Dim v As Object = Valor(pDic, pChave)
        If v Is Nothing Then Return Nothing
        If TypeOf v Is Decimal OrElse TypeOf v Is Double OrElse TypeOf v Is Integer OrElse TypeOf v Is Long Then
            Return Convert.ToDecimal(v, CultureInfo.InvariantCulture)
        End If
        Dim s As String = v.ToString().Trim().TrimEnd("%"c).Trim()
        Dim d As Decimal
        If Decimal.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then Return d
        Return Nothing
    End Function

#End Region

    Private Function TokenValido(pRecebido As String) As Boolean
        If String.IsNullOrEmpty(_token) OrElse String.IsNullOrEmpty(pRecebido) Then Return False
        ' Comparacao em tempo constante
        Dim a As Byte() = Encoding.UTF8.GetBytes(pRecebido)
        Dim b As Byte() = Encoding.UTF8.GetBytes(_token)
        Dim dif As Integer = a.Length Xor b.Length
        For i As Integer = 0 To Math.Min(a.Length, b.Length) - 1
            dif = dif Or (CInt(a(i)) Xor CInt(b(i)))
        Next
        Return dif = 0
    End Function

End Class
