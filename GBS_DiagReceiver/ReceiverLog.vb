''' <summary>
''' Log simples no console (com timestamp e cor por nivel). Thread-safe.
''' </summary>
Public Module ReceiverLog

    Private ReadOnly _lock As New Object()

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
        SyncLock _lock
            Dim corAnterior As ConsoleColor = Console.ForegroundColor
            Console.ForegroundColor = pCor
            Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} {1} {2}", DateTime.Now, pNivel, pMsg)
            Console.ForegroundColor = corAnterior
        End SyncLock
    End Sub

End Module
