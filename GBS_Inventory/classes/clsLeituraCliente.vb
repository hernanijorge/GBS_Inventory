Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Classe de leitura de Clientes
''' </summary>
Public Class clsLeituraCliente

#Region "Atributos"

    Private ConnectionString As String

#End Region

#Region "Construtor"

    Public Sub New()

        Try
            Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
        Catch ex As Exception
            Throw New Exception(ex.ToString())
        End Try

    End Sub

#End Region

#Region "Consultas"

    Public Function selecionarCliente(pIdCliente As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try
            oPar(0).Value = pIdCliente
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_CLIENTE.PROC_SELECT_CLIENTE", oPar)
        Catch ex As Exception
            If IsErroObjetoCliente(ex) Then
                Return selecionarClienteSemPackage(pIdCliente)
            End If
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function selecionarClientes(pAtivo As String) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ATIVO", OracleDbType.Char, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try
            oPar(0).Value = If(String.IsNullOrWhiteSpace(pAtivo), DBNull.Value, CObj(pAtivo))
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_CLIENTE.PROC_SELECT_CLIENTES", oPar)
        Catch ex As Exception
            If IsErroObjetoCliente(ex) Then
                Return selecionarClientesSemPackage(pAtivo)
            End If
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function selecionarClientesFiltro(pPesquisa As String) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_PESQUISA", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try
            oPar(0).Value = If(String.IsNullOrWhiteSpace(pPesquisa), DBNull.Value, CObj(pPesquisa))
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_CLIENTE.PROC_SELECT_CLIENTES_FILTRO", oPar)
        Catch ex As Exception
            If IsErroObjetoCliente(ex) Then
                Return selecionarClientesFiltroSemPackage(pPesquisa)
            End If
            Throw New Exception(ex.Message)
        End Try

    End Function

#End Region

#Region "Fallback sem package"

    Private Function selecionarClienteSemPackage(pIdCliente As Integer) As DataSet

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(0).Value = pIdCliente

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                                               "SELECT * FROM TBL_CLIENTE WHERE ID_CLIENTE = :P_ID", oPar)
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Private Function selecionarClientesSemPackage(pAtivo As String) As DataSet

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("P_ATIVO", OracleDbType.Char, ParameterDirection.Input)
        oPar(0).Value = If(String.IsNullOrWhiteSpace(pAtivo), DBNull.Value, CObj(pAtivo))

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                                               "SELECT * " &
                                               "  FROM TBL_CLIENTE " &
                                               " WHERE (:P_ATIVO IS NULL OR ATIVO = UPPER(:P_ATIVO)) " &
                                               " ORDER BY NOME_RAZAO", oPar)
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Private Function selecionarClientesFiltroSemPackage(pPesquisa As String) As DataSet

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_PESQUISA", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_PESQ", OracleDbType.Varchar2, ParameterDirection.Input)

        Dim sPesquisa As String = If(String.IsNullOrWhiteSpace(pPesquisa), Nothing, pPesquisa.Trim())
        oPar(0).Value = If(String.IsNullOrWhiteSpace(sPesquisa), DBNull.Value, CObj(sPesquisa))
        oPar(1).Value = If(String.IsNullOrWhiteSpace(sPesquisa), DBNull.Value, CObj("%" & sPesquisa.ToUpper() & "%"))

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                                               "SELECT * " &
                                               "  FROM TBL_CLIENTE " &
                                               " WHERE (:P_PESQUISA IS NULL " &
                                               "        OR UPPER(NOME_RAZAO) LIKE :P_PESQ " &
                                               "        OR UPPER(NVL(NOME_FANTASIA, '')) LIKE :P_PESQ " &
                                               "        OR UPPER(NVL(DOCUMENTO, '')) LIKE :P_PESQ " &
                                               "        OR UPPER(NVL(EMAIL, '')) LIKE :P_PESQ) " &
                                               " ORDER BY NOME_RAZAO", oPar)
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

#End Region

#Region "Auxiliares"

    Private Function IsErroObjetoCliente(ex As Exception) As Boolean

        Dim msg As String = ex.ToString().ToUpperInvariant()

        Return msg.Contains("PLS-00201") OrElse
               msg.Contains("ORA-06550") OrElse
               msg.Contains("PACK_CLIENTE") OrElse
               msg.Contains("PROC_SELECT_CLIENT")

    End Function

    Private Function mensagemErroEstrutural(ex As Exception) As String

        Dim msg As String = ex.ToString().ToUpperInvariant()

        If msg.Contains("ORA-00942") Then
            Return "Estrutura de clientes/invoice ainda nao instalada no Oracle. Execute os scripts 07_DDL_CADASTROS_INVOICE.sql, 08_PACK_CLIENTE.sql e 09_PACK_INVOICE.sql."
        End If

        Return ex.Message

    End Function

#End Region

End Class
