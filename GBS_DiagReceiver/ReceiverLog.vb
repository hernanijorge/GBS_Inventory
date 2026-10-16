Imports System.Configuration
Imports System.Globalization

''' <summary>
''' Log do receptor: console (quando interativo) + arquivo diario
''' &lt;LogFolder&gt;\DiagReceiver_yyyyMMdd.log (padrao C:\GBS\Logs), guardando LogRetentionDays
''' dias (padrao 90). Thread-safe. Falha ao gravar o arquivo nunca derruba o receptor.
''' </summary>
Public Module ReceiverLog

    Private ReadOnly _lock As New Object()
    Private _pasta As String
    Private _retencaoDias As Integer
    Private _diaAtual As Date = Date.MinValue
    Private _console As Boolean = True

    Private Const Prefixo As String = "DiagReceiver_"

    ''' <summary>Le a configuracao; chamar uma vez na inicializacao.</summary>
    Public Sub Configurar(pConsole As Boolean)
        _console = pConsole
        _pasta = ConfigurationManager.AppSettings("LogFolder")
        If String.IsNullOrWhiteSpace(_pasta) Then _pasta = "C:\GBS\Logs"
        If Not Integer.TryParse(ConfigurationManager.AppSettings("LogRetentionDays"), _retencaoDias) OrElse _retencaoDias < 1 Then
            _retencaoDias = 90
        End If
    End Sub

    Public ReadOnly Property ArquivoAtual As String
        Get
            Return Path.Combine(If(_pasta, "C:\GBS\Logs"), Prefixo & DateTime.Now.ToString("yyyyMMdd") & ".log")
        End Get
    End Property

    Public Sub Info(pMsg As String)
        Escrever("INFO ", pMsg, ConsoleColor.Gray)
    End Sub

    Public Sub Ok(pMsg As String)
        Escrever("OK   ", pMsg, ConsoleColor.Green)
    End Sub

    Public Sub Aviso(pMsg As String)
        Escrever("AVISO", pMsg, ConsoleColor.Yellow)
    End Sub

    Public Sub Erro(pMsg As String)
        Escrever("ERRO ", pMsg, ConsoleColor.Red)
    End Sub

    Private Sub Escrever(pNivel As String, pMsg As String, pCor As ConsoleColor)
        Dim linha As String = String.Format("{0:yyyy-MM-dd HH:mm:ss} {1} {2}", DateTime.Now, pNivel, pMsg)
        SyncLock _lock
            If _console Then
                Try
                    Dim corAnterior As ConsoleColor = Console.ForegroundColor
                    Console.ForegroundColor = pCor
                    Console.WriteLine(linha)
                    Console.ForegroundColor = corAnterior
                Catch
                    ' sem console (servico)
                End Try
            End If
            GravarArquivo(linha)
        End SyncLock
    End Sub

    Private Sub GravarArquivo(pLinha As String)
        If _pasta Is Nothing Then Configurar(_console)
        Try
            If DateTime.Today <> _diaAtual Then
                _diaAtual = DateTime.Today
                Directory.CreateDirectory(_pasta)
                LimparAntigos()
            End If
            Using fs As New FileStream(ArquivoAtual, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
                Using sw As New StreamWriter(fs, New UTF8Encoding(False))
                    sw.WriteLine(pLinha)
                End Using
            End Using
        Catch
            ' log em arquivo e "melhor esforco": nunca impede o atendimento
        End Try
    End Sub

    ''' <summary>Apaga DiagReceiver_yyyyMMdd.log mais antigos que a retencao (data pelo nome).</summary>
    Private Sub LimparAntigos()
        Dim limite As Date = DateTime.Today.AddDays(-_retencaoDias)
        For Each arq As String In Directory.GetFiles(_pasta, Prefixo & "*.log")
            Dim d As Date
            Dim nome As String = Path.GetFileNameWithoutExtension(arq).Substring(Prefixo.Length)
            If Date.TryParseExact(nome, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, d) AndAlso d < limite Then
                Try : File.Delete(arq) : Catch : End Try
            End If
        Next
    End Sub

End Module
