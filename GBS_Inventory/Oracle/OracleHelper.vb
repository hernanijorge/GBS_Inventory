Imports Oracle.ManagedDataAccess.Client
Imports System.Data

''' <summary>
''' Classe utilitária para execução de comandos Oracle (ODP.NET)
''' 
''' </summary>
Public Class OracleHelper

#Region "ExecuteNonQuery"

    ''' <summary>
    ''' Executa um comando sem retorno de dados (INSERT, UPDATE, DELETE, stored procedure)
    ''' usando a ConnectionString (cria conexão nova).
    ''' </summary>
    Public Shared Sub ExecuteNonQuery(ByVal connectionString As String,
                                      ByVal commandType As CommandType,
                                      ByVal commandText As String,
                                      ByVal ParamArray commandParameters() As OracleParameter)

        Dim oCmd As New OracleCommand()

        Using oCon As New OracleConnection(connectionString)
            PrepareCommand(oCmd, oCon, Nothing, commandType, commandText, commandParameters)
            oCmd.ExecuteNonQuery()
            oCmd.Parameters.Clear()
        End Using

    End Sub

    ''' <summary>
    ''' Executa um comando sem retorno de dados dentro de uma transação existente.
    ''' </summary>
    Public Shared Sub ExecuteNonQuery(ByVal transaction As OracleTransaction,
                                      ByVal commandType As CommandType,
                                      ByVal commandText As String,
                                      ByVal ParamArray commandParameters() As OracleParameter)

        If transaction Is Nothing Then
            Throw New ArgumentNullException("transaction", "Transação não pode ser nula.")
        End If

        Dim oCmd As New OracleCommand()
        PrepareCommand(oCmd, transaction.Connection, transaction, commandType, commandText, commandParameters)
        oCmd.ExecuteNonQuery()
        oCmd.Parameters.Clear()

    End Sub

#End Region

#Region "ExecuteDataset"

    ''' <summary>
    ''' Executa um SELECT ou cursor e retorna um DataSet preenchido.
    ''' </summary>
    Public Shared Function ExecuteDataset(ByVal connectionString As String,
                                          ByVal commandType As CommandType,
                                          ByVal commandText As String,
                                          ByVal ParamArray commandParameters() As OracleParameter) As DataSet

        Dim oCmd As New OracleCommand()
        Dim ds As New DataSet()

        Using oCon As New OracleConnection(connectionString)
            PrepareCommand(oCmd, oCon, Nothing, commandType, commandText, commandParameters)
            Dim oDa As New OracleDataAdapter(oCmd)
            oDa.Fill(ds)
            oCmd.Parameters.Clear()
        End Using

        Return ds

    End Function

    ''' <summary>
    ''' Executa um SELECT dentro de uma transação existente e retorna um DataSet preenchido.
    ''' Necessário para enxergar dados inseridos mas ainda não commitados na mesma transação.
    ''' </summary>
    Public Shared Function ExecuteDataset(ByVal transaction As OracleTransaction,
                                          ByVal commandType As CommandType,
                                          ByVal commandText As String,
                                          ByVal ParamArray commandParameters() As OracleParameter) As DataSet

        If transaction Is Nothing Then
            Throw New ArgumentNullException("transaction", "Transação não pode ser nula.")
        End If

        Dim oCmd As New OracleCommand()
        Dim ds As New DataSet()

        PrepareCommand(oCmd, transaction.Connection, transaction, commandType, commandText, commandParameters)
        Dim oDa As New OracleDataAdapter(oCmd)
        oDa.Fill(ds)
        oCmd.Parameters.Clear()

        Return ds

    End Function

#End Region

#Region "ExecuteReader"

    ''' <summary>
    ''' Executa um SELECT e retorna um OracleDataReader (com CloseConnection).
    ''' </summary>
    Public Shared Function ExecuteReader(ByVal connectionString As String,
                                         ByVal commandType As CommandType,
                                         ByVal commandText As String,
                                         ByVal ParamArray commandParameters() As OracleParameter) As OracleDataReader

        Dim oCmd As New OracleCommand()
        Dim oCon As New OracleConnection(connectionString)

        Try

            PrepareCommand(oCmd, oCon, Nothing, commandType, commandText, commandParameters)
            Return oCmd.ExecuteReader(CommandBehavior.CloseConnection)

        Catch ex As Exception

            oCon.Close()
            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

#Region "ExecuteScalar"

    ''' <summary>
    ''' Executa um comando e retorna o primeiro valor da primeira coluna.
    ''' </summary>
    Public Shared Function ExecuteScalar(ByVal connectionString As String,
                                         ByVal commandType As CommandType,
                                         ByVal commandText As String,
                                         ByVal ParamArray commandParameters() As OracleParameter) As Object

        Dim oCmd As New OracleCommand()
        Dim vResultado As Object

        Using oCon As New OracleConnection(connectionString)
            PrepareCommand(oCmd, oCon, Nothing, commandType, commandText, commandParameters)
            vResultado = oCmd.ExecuteScalar()
            oCmd.Parameters.Clear()
        End Using

        Return vResultado

    End Function

#End Region

#Region "Utilitário"

    Private Shared Sub PrepareCommand(ByVal command As OracleCommand,
                                      ByVal connection As OracleConnection,
                                      ByVal transaction As OracleTransaction,
                                      ByVal commandType As CommandType,
                                      ByVal commandText As String,
                                      ByVal commandParameters() As OracleParameter)

        If connection.State <> ConnectionState.Open Then
            connection.Open()
        End If

        command.Connection = connection
        command.CommandText = commandText
        command.CommandType = commandType
        command.BindByName = True

        If transaction IsNot Nothing Then
            command.Transaction = transaction
        End If

        If commandParameters IsNot Nothing Then
            For Each p As OracleParameter In commandParameters
                If p IsNot Nothing Then
                    command.Parameters.Add(p)
                End If
            Next
        End If

    End Sub

#End Region

End Class
