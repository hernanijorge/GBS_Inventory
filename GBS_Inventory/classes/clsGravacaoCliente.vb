Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravacao de Clientes
''' </summary>
Public Class clsGravacaoCliente

#Region "Atributos"

    Private ConnectionString As String
    Private oTransacao As OracleTransaction
    Private oConexao As OracleConnection

#End Region

#Region "Construtor"

    Public Sub New()

        Try
            Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        oConexao = Nothing

    End Sub

#End Region

#Region "Transacao"

    Public Function beginTransacao() As Boolean

        Dim oCon As New OracleConnection(Me.ConnectionString)

        Try
            oCon.Open()
            Me.oConexao = oCon
            Me.oTransacao = Me.oConexao.BeginTransaction()
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function commitTransacao() As Boolean

        Try
            Me.oTransacao.Commit()
            Me.oConexao.Close()
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function rollbackTransacao() As Boolean

        Try
            Me.oTransacao.Rollback()
            Me.oConexao.Close()
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

#End Region

#Region "Metodos Transacionais"

    Public Function incluirCliente(pCliente As Cliente) As Integer

        Dim oPar(13) As OracleParameter

        oPar(0) = New OracleParameter("V_NOME_RAZAO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_NOME_FANTASIA", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_DOCUMENTO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("V_EMAIL", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4) = New OracleParameter("V_TELEFONE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5) = New OracleParameter("V_ENDERECO1", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6) = New OracleParameter("V_ENDERECO2", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(7) = New OracleParameter("V_CIDADE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8) = New OracleParameter("V_ESTADO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9) = New OracleParameter("V_ZIP_CODE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_PAIS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("V_ATIVO", OracleDbType.Char, ParameterDirection.Input)
        oPar(12) = New OracleParameter("V_OBSERVACOES", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(13) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Output)

        Dim vIdGerado As OracleDecimal

        Try

            oPar(0).Value = pCliente.NomeRazao
            oPar(1).Value = If(String.IsNullOrWhiteSpace(pCliente.NomeFantasia), DBNull.Value, CObj(pCliente.NomeFantasia))
            oPar(2).Value = If(String.IsNullOrWhiteSpace(pCliente.Documento), DBNull.Value, CObj(pCliente.Documento))
            oPar(3).Value = If(String.IsNullOrWhiteSpace(pCliente.Email), DBNull.Value, CObj(pCliente.Email))
            oPar(4).Value = If(String.IsNullOrWhiteSpace(pCliente.Telefone), DBNull.Value, CObj(pCliente.Telefone))
            oPar(5).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco1), DBNull.Value, CObj(pCliente.Endereco1))
            oPar(6).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco2), DBNull.Value, CObj(pCliente.Endereco2))
            oPar(7).Value = If(String.IsNullOrWhiteSpace(pCliente.Cidade), DBNull.Value, CObj(pCliente.Cidade))
            oPar(8).Value = If(String.IsNullOrWhiteSpace(pCliente.Estado), DBNull.Value, CObj(pCliente.Estado))
            oPar(9).Value = If(String.IsNullOrWhiteSpace(pCliente.ZipCode), DBNull.Value, CObj(pCliente.ZipCode))
            oPar(10).Value = If(String.IsNullOrWhiteSpace(pCliente.Pais), "USA", pCliente.Pais)
            oPar(11).Value = If(String.IsNullOrWhiteSpace(pCliente.Ativo), "Y", pCliente.Ativo)
            oPar(12).Value = If(String.IsNullOrWhiteSpace(pCliente.Observacoes), DBNull.Value, CObj(pCliente.Observacoes))

            Try
                OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_CLIENTE.PROC_INSERT_CLIENTE", oPar)
                vIdGerado = CType(oPar(13).Value, OracleDecimal)
                Return CInt(vIdGerado.Value)
            Catch exPkg As Exception
                If IsErroObjetoCliente(exPkg) Then
                    Return incluirClienteSemPackage(pCliente)
                End If
                Throw
            End Try

        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Public Function alterarCliente(pCliente As Cliente) As Boolean

        Dim oPar(13) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_NOME_RAZAO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_NOME_FANTASIA", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("V_DOCUMENTO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4) = New OracleParameter("V_EMAIL", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5) = New OracleParameter("V_TELEFONE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6) = New OracleParameter("V_ENDERECO1", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(7) = New OracleParameter("V_ENDERECO2", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8) = New OracleParameter("V_CIDADE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9) = New OracleParameter("V_ESTADO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_ZIP_CODE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("V_PAIS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(12) = New OracleParameter("V_ATIVO", OracleDbType.Char, ParameterDirection.Input)
        oPar(13) = New OracleParameter("V_OBSERVACOES", OracleDbType.Varchar2, ParameterDirection.Input)

        Try

            oPar(0).Value = pCliente.IdCliente
            oPar(1).Value = pCliente.NomeRazao
            oPar(2).Value = If(String.IsNullOrWhiteSpace(pCliente.NomeFantasia), DBNull.Value, CObj(pCliente.NomeFantasia))
            oPar(3).Value = If(String.IsNullOrWhiteSpace(pCliente.Documento), DBNull.Value, CObj(pCliente.Documento))
            oPar(4).Value = If(String.IsNullOrWhiteSpace(pCliente.Email), DBNull.Value, CObj(pCliente.Email))
            oPar(5).Value = If(String.IsNullOrWhiteSpace(pCliente.Telefone), DBNull.Value, CObj(pCliente.Telefone))
            oPar(6).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco1), DBNull.Value, CObj(pCliente.Endereco1))
            oPar(7).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco2), DBNull.Value, CObj(pCliente.Endereco2))
            oPar(8).Value = If(String.IsNullOrWhiteSpace(pCliente.Cidade), DBNull.Value, CObj(pCliente.Cidade))
            oPar(9).Value = If(String.IsNullOrWhiteSpace(pCliente.Estado), DBNull.Value, CObj(pCliente.Estado))
            oPar(10).Value = If(String.IsNullOrWhiteSpace(pCliente.ZipCode), DBNull.Value, CObj(pCliente.ZipCode))
            oPar(11).Value = If(String.IsNullOrWhiteSpace(pCliente.Pais), "USA", pCliente.Pais)
            oPar(12).Value = If(String.IsNullOrWhiteSpace(pCliente.Ativo), "Y", pCliente.Ativo)
            oPar(13).Value = If(String.IsNullOrWhiteSpace(pCliente.Observacoes), DBNull.Value, CObj(pCliente.Observacoes))

            Try
                OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_CLIENTE.PROC_UPDATE_CLIENTE", oPar)
            Catch exPkg As Exception
                If IsErroObjetoCliente(exPkg) Then
                    alterarClienteSemPackage(pCliente)
                Else
                    Throw
                End If
            End Try
            Return True

        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Public Function excluirCliente(pIdCliente As Integer) As Boolean

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try
            oPar(0).Value = pIdCliente
            Try
                OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_CLIENTE.PROC_DELETE_CLIENTE", oPar)
            Catch exPkg As Exception
                If IsErroObjetoCliente(exPkg) Then
                    excluirClienteSemPackage(pIdCliente)
                Else
                    Throw
                End If
            End Try
            Return True
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

#End Region

#Region "Fallback sem package"

    Private Function incluirClienteSemPackage(pCliente As Cliente) As Integer

        Dim id As Integer = obterProximoIdCliente()
        Dim oPar(13) As OracleParameter

        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input) : oPar(0).Value = id
        oPar(1) = New OracleParameter("P_NOME_RAZAO", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(1).Value = pCliente.NomeRazao
        oPar(2) = New OracleParameter("P_NOME_FANTASIA", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(2).Value = If(String.IsNullOrWhiteSpace(pCliente.NomeFantasia), DBNull.Value, CObj(pCliente.NomeFantasia))
        oPar(3) = New OracleParameter("P_DOCUMENTO", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(3).Value = If(String.IsNullOrWhiteSpace(pCliente.Documento), DBNull.Value, CObj(pCliente.Documento))
        oPar(4) = New OracleParameter("P_EMAIL", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(4).Value = If(String.IsNullOrWhiteSpace(pCliente.Email), DBNull.Value, CObj(pCliente.Email))
        oPar(5) = New OracleParameter("P_TELEFONE", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(5).Value = If(String.IsNullOrWhiteSpace(pCliente.Telefone), DBNull.Value, CObj(pCliente.Telefone))
        oPar(6) = New OracleParameter("P_ENDERECO1", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(6).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco1), DBNull.Value, CObj(pCliente.Endereco1))
        oPar(7) = New OracleParameter("P_ENDERECO2", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(7).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco2), DBNull.Value, CObj(pCliente.Endereco2))
        oPar(8) = New OracleParameter("P_CIDADE", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(8).Value = If(String.IsNullOrWhiteSpace(pCliente.Cidade), DBNull.Value, CObj(pCliente.Cidade))
        oPar(9) = New OracleParameter("P_ESTADO", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(9).Value = If(String.IsNullOrWhiteSpace(pCliente.Estado), DBNull.Value, CObj(pCliente.Estado))
        oPar(10) = New OracleParameter("P_ZIP", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(10).Value = If(String.IsNullOrWhiteSpace(pCliente.ZipCode), DBNull.Value, CObj(pCliente.ZipCode))
        oPar(11) = New OracleParameter("P_PAIS", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(11).Value = If(String.IsNullOrWhiteSpace(pCliente.Pais), "USA", pCliente.Pais)
        oPar(12) = New OracleParameter("P_ATIVO", OracleDbType.Char, ParameterDirection.Input) : oPar(12).Value = If(String.IsNullOrWhiteSpace(pCliente.Ativo), "Y", pCliente.Ativo)
        oPar(13) = New OracleParameter("P_OBS", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(13).Value = If(String.IsNullOrWhiteSpace(pCliente.Observacoes), DBNull.Value, CObj(pCliente.Observacoes))

        OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                                     "INSERT INTO TBL_CLIENTE " &
                                     "(ID_CLIENTE, NOME_RAZAO, NOME_FANTASIA, DOCUMENTO, EMAIL, TELEFONE, " &
                                     " ENDERECO1, ENDERECO2, CIDADE, ESTADO, ZIP_CODE, PAIS, ATIVO, OBSERVACOES, DATA_CADASTRO) " &
                                     "VALUES (:P_ID, UPPER(TRIM(:P_NOME_RAZAO)), :P_NOME_FANTASIA, :P_DOCUMENTO, LOWER(TRIM(:P_EMAIL)), :P_TELEFONE, " &
                                     " :P_ENDERECO1, :P_ENDERECO2, :P_CIDADE, :P_ESTADO, :P_ZIP, :P_PAIS, UPPER(:P_ATIVO), :P_OBS, SYSDATE)", oPar)

        Return id

    End Function

    Private Sub alterarClienteSemPackage(pCliente As Cliente)

        Dim oPar(13) As OracleParameter

        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input) : oPar(0).Value = pCliente.IdCliente
        oPar(1) = New OracleParameter("P_NOME_RAZAO", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(1).Value = pCliente.NomeRazao
        oPar(2) = New OracleParameter("P_NOME_FANTASIA", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(2).Value = If(String.IsNullOrWhiteSpace(pCliente.NomeFantasia), DBNull.Value, CObj(pCliente.NomeFantasia))
        oPar(3) = New OracleParameter("P_DOCUMENTO", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(3).Value = If(String.IsNullOrWhiteSpace(pCliente.Documento), DBNull.Value, CObj(pCliente.Documento))
        oPar(4) = New OracleParameter("P_EMAIL", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(4).Value = If(String.IsNullOrWhiteSpace(pCliente.Email), DBNull.Value, CObj(pCliente.Email))
        oPar(5) = New OracleParameter("P_TELEFONE", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(5).Value = If(String.IsNullOrWhiteSpace(pCliente.Telefone), DBNull.Value, CObj(pCliente.Telefone))
        oPar(6) = New OracleParameter("P_ENDERECO1", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(6).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco1), DBNull.Value, CObj(pCliente.Endereco1))
        oPar(7) = New OracleParameter("P_ENDERECO2", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(7).Value = If(String.IsNullOrWhiteSpace(pCliente.Endereco2), DBNull.Value, CObj(pCliente.Endereco2))
        oPar(8) = New OracleParameter("P_CIDADE", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(8).Value = If(String.IsNullOrWhiteSpace(pCliente.Cidade), DBNull.Value, CObj(pCliente.Cidade))
        oPar(9) = New OracleParameter("P_ESTADO", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(9).Value = If(String.IsNullOrWhiteSpace(pCliente.Estado), DBNull.Value, CObj(pCliente.Estado))
        oPar(10) = New OracleParameter("P_ZIP", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(10).Value = If(String.IsNullOrWhiteSpace(pCliente.ZipCode), DBNull.Value, CObj(pCliente.ZipCode))
        oPar(11) = New OracleParameter("P_PAIS", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(11).Value = If(String.IsNullOrWhiteSpace(pCliente.Pais), "USA", pCliente.Pais)
        oPar(12) = New OracleParameter("P_ATIVO", OracleDbType.Char, ParameterDirection.Input) : oPar(12).Value = If(String.IsNullOrWhiteSpace(pCliente.Ativo), "Y", pCliente.Ativo)
        oPar(13) = New OracleParameter("P_OBS", OracleDbType.Varchar2, ParameterDirection.Input) : oPar(13).Value = If(String.IsNullOrWhiteSpace(pCliente.Observacoes), DBNull.Value, CObj(pCliente.Observacoes))

        OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                                     "UPDATE TBL_CLIENTE " &
                                     "   SET NOME_RAZAO = UPPER(TRIM(:P_NOME_RAZAO)), " &
                                     "       NOME_FANTASIA = :P_NOME_FANTASIA, " &
                                     "       DOCUMENTO = :P_DOCUMENTO, " &
                                     "       EMAIL = LOWER(TRIM(:P_EMAIL)), " &
                                     "       TELEFONE = :P_TELEFONE, " &
                                     "       ENDERECO1 = :P_ENDERECO1, " &
                                     "       ENDERECO2 = :P_ENDERECO2, " &
                                     "       CIDADE = :P_CIDADE, " &
                                     "       ESTADO = :P_ESTADO, " &
                                     "       ZIP_CODE = :P_ZIP, " &
                                     "       PAIS = :P_PAIS, " &
                                     "       ATIVO = UPPER(:P_ATIVO), " &
                                     "       OBSERVACOES = :P_OBS, " &
                                     "       DATA_ALTERACAO = SYSDATE " &
                                     " WHERE ID_CLIENTE = :P_ID", oPar)

    End Sub

    Private Sub excluirClienteSemPackage(pIdCliente As Integer)

        Dim totalInvoice As Integer = 0

        Using cmd As New OracleCommand("SELECT COUNT(*) FROM TBL_INVOICE WHERE ID_CLIENTE = :P_ID", Me.oConexao)
            cmd.Transaction = Me.oTransacao
            cmd.Parameters.Add(New OracleParameter("P_ID", OracleDbType.Int32, pIdCliente, ParameterDirection.Input))
            totalInvoice = Convert.ToInt32(cmd.ExecuteScalar())
        End Using

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(0).Value = pIdCliente

        If totalInvoice > 0 Then
            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                                         "UPDATE TBL_CLIENTE SET ATIVO = 'N', DATA_ALTERACAO = SYSDATE WHERE ID_CLIENTE = :P_ID", oPar)
        Else
            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                                         "DELETE FROM TBL_CLIENTE WHERE ID_CLIENTE = :P_ID", oPar)
        End If

    End Sub

    Private Function obterProximoIdCliente() As Integer

        Using cmd As New OracleCommand("SELECT SEQ_CLIENTE.NEXTVAL FROM DUAL", Me.oConexao)
            cmd.Transaction = Me.oTransacao
            Return Convert.ToInt32(cmd.ExecuteScalar())
        End Using

    End Function

#End Region

#Region "Auxiliares"

    Private Function IsErroObjetoCliente(ex As Exception) As Boolean

        Dim msg As String = ex.ToString().ToUpperInvariant()

        Return msg.Contains("PLS-00201") OrElse
               msg.Contains("ORA-06550") OrElse
               msg.Contains("PACK_CLIENTE")

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
