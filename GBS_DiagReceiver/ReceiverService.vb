Imports System.ServiceProcess

''' <summary>
''' Servico do Windows "GBS DiagReceiver" (nome interno GBSDiagReceiver).
''' O mesmo .exe roda como console quando aberto manualmente; quando iniciado pelo
''' Service Control Manager, Program.Main chama ServiceBase.Run com esta classe.
''' Instalacao: ver README (sc.exe create / failure).
''' </summary>
Public Class ReceiverService
    Inherits ServiceBase

    Public Const NomeServico As String = "GBSDiagReceiver"

    Public Sub New()
        ServiceName = NomeServico
        CanStop = True
        CanShutdown = True
        AutoLog = True
    End Sub

    Protected Overrides Sub OnStart(args As String())
        ' Servico roda em C:\Windows\System32: os caminhos relativos (diag, secrets.config)
        ' ja sao resolvidos pela pasta do .exe / do .config.
        Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory
        Try
            Program.Iniciar()
        Catch ex As Exception
            ReceiverLog.Erro("Falha ao iniciar o servico:" & vbCrLf & ex.ToString())
            ExitCode = 1064   ' ERROR_EXCEPTION_IN_SERVICE: SCM registra a falha
            Throw
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        Program.Parar()
    End Sub

    Protected Overrides Sub OnShutdown()
        Program.Parar()
    End Sub

End Class
