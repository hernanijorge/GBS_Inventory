Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravação de Remessas
''' </summary>
Public Class clsGravacaoRemessa

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

            Throw New Exception(ex.ToString)

        End Try

        oConexao = Nothing

    End Sub

#End Region

#Region "Transação"

    Public Function beginTransacao() As Boolean

        Dim oCon As New OracleConnection(Me.ConnectionString)

        Try

            oCon.Open()

            Me.oConexao = oCon

            Me.oTransacao = Me.oConexao.BeginTransaction

            Return True

        Catch ex As Exception

            Throw New Exception(ex.ToString)

            Return False

        End Try

    End Function

    Public Function commitTransacao() As Boolean

        Try

            Me.oTransacao.Commit()

            Me.oConexao.Close()

            Return True

        Catch ex As Exception

            Return False

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function rollbackTransacao() As Boolean

        Try

            Me.oTransacao.Rollback()

            Me.oConexao.Close()

            Return True

        Catch ex As Exception

            Return False

            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

#Region "Métodos Transacionais"

    Public Function incluirRemessa(pRemessa As Remessa, ByRef pRemessaRef As String) As Integer

        ' Gera código único: GBS-SH-YYYYMMDDHH24MISS
        Dim sCodigo As String = "GBS-SH-" & DateTime.Now.ToString("yyyyMMddHHmmss")
        If Not String.IsNullOrEmpty(pRemessa.RemessaRef) Then sCodigo = pRemessa.RemessaRef
        pRemessaRef = sCodigo

        Dim oPar(4) As OracleParameter

        oPar(0) = New OracleParameter("P_CODIGO_REMESSA", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_DESTINATARIO",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_CARRIER",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("P_TRACKING",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4) = New OracleParameter("P_OBSERVACAO",     OracleDbType.Varchar2, ParameterDirection.Input)

        Dim vIdGerado As Integer = 0

        Try

            oPar(0).Value = sCodigo
            oPar(1).Value = If(String.IsNullOrEmpty(pRemessa.RecipientName), "N/A", pRemessa.RecipientName)
            oPar(2).Value = If(String.IsNullOrEmpty(pRemessa.Carrier), "OTHER", pRemessa.Carrier)
            oPar(3).Value = If(String.IsNullOrEmpty(pRemessa.TrackingNumber), DBNull.Value, CObj(pRemessa.TrackingNumber))
            oPar(4).Value = If(String.IsNullOrEmpty(pRemessa.Notes), DBNull.Value, CObj(pRemessa.Notes))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_REMESSA.PROC_INSERT_REMESSA", oPar)

            ' Obtém o ID gerado
            Dim oPar2(0) As OracleParameter
            oPar2(0) = New OracleParameter("P_COD", OracleDbType.Varchar2, ParameterDirection.Input)
            oPar2(0).Value = sCodigo

            Dim ds As DataSet = OracleHelper.ExecuteDataset(Me.oTransacao, CommandType.Text,
                "SELECT ID_REMESSA FROM TBL_REMESSA WHERE CODIGO_REMESSA = :P_COD AND ROWNUM = 1", oPar2)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                vIdGerado = CInt(ds.Tables(0).Rows(0)("ID_REMESSA"))
            End If

            ' Cliente automatico pelo Recipient (nome / nome fantasia / aliases), na mesma transacao
            If vIdGerado > 0 Then autoVincularClientes(vIdGerado)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return vIdGerado

    End Function

    Public Function atualizarStatusRemessa(pIdRemessa As Integer, pStatus As String, pDataEntrega As Date?) As Boolean

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("P_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_ID",     OracleDbType.Int32,    ParameterDirection.Input)

        Try

            oPar(0).Value = pStatus.ToUpper()
            oPar(1).Value = pIdRemessa

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "UPDATE TBL_REMESSA SET STATUS_REMESSA = :P_STATUS WHERE ID_REMESSA = :P_ID", oPar)

            If pStatus.ToUpper() = "DELIVERED" Then
                OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                    "UPDATE TBL_REMESSA_ITEM SET STATUS_ITEM = 'DELIVERED' WHERE ID_REMESSA = :P_ID", oPar)

                OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                    "UPDATE TBL_EQUIPAMENTO SET STATUS = 'SOLD', DATA_ATUALIZACAO = SYSDATE " &
                    " WHERE ID_EQUIPAMENTO IN (SELECT ID_EQUIPAMENTO FROM TBL_REMESSA_ITEM WHERE ID_REMESSA = :P_ID)", oPar)
            End If

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function incluirItemRemessa(pItem As ItemRemessa) As Boolean

        Try

            Dim oParInsert(3) As OracleParameter
            oParInsert(0) = New OracleParameter("P_ID_REMESSA",    OracleDbType.Int32,    ParameterDirection.Input)
            oParInsert(1) = New OracleParameter("P_ID_EQUIP",      OracleDbType.Int32,    ParameterDirection.Input)
            oParInsert(2) = New OracleParameter("P_SALE_PRICE",    OracleDbType.Decimal,  ParameterDirection.Input)
            oParInsert(3) = New OracleParameter("P_NOTES",         OracleDbType.Varchar2, ParameterDirection.Input)
            oParInsert(0).Value = pItem.IdRemessa
            oParInsert(1).Value = pItem.IdEquipamento
            oParInsert(2).Value = If(pItem.SalePriceUsd.HasValue, CObj(pItem.SalePriceUsd.Value), DBNull.Value)
            oParInsert(3).Value = If(String.IsNullOrEmpty(pItem.Notes), DBNull.Value, CObj(pItem.Notes))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "INSERT INTO TBL_REMESSA_ITEM (ID_REMESSA_ITEM, ID_REMESSA, ID_EQUIPAMENTO, STATUS_ITEM, SALE_PRICE_USD, NOTES, DATA_CADASTRO) " &
                "VALUES (SEQ_REMESSA_ITEM.NEXTVAL, :P_ID_REMESSA, :P_ID_EQUIP, 'PENDING', :P_SALE_PRICE, :P_NOTES, SYSDATE)", oParInsert)

            Dim oParUpdate(0) As OracleParameter
            oParUpdate(0) = New OracleParameter("P_ID_EQUIP", OracleDbType.Int32, ParameterDirection.Input)
            oParUpdate(0).Value = pItem.IdEquipamento

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "UPDATE TBL_EQUIPAMENTO SET STATUS = 'SHIPPED', DATA_ATUALIZACAO = SYSDATE " &
                " WHERE ID_EQUIPAMENTO = :P_ID_EQUIP", oParUpdate)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function excluirRemessa(pIdRemessa As Integer) As Boolean

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdRemessa

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "UPDATE TBL_EQUIPAMENTO SET STATUS = 'IN_STOCK', DATA_ATUALIZACAO = SYSDATE " &
                " WHERE ID_EQUIPAMENTO IN (SELECT ID_EQUIPAMENTO FROM TBL_REMESSA_ITEM WHERE ID_REMESSA = :P_ID)", oPar)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "DELETE FROM TBL_REMESSA_ITEM WHERE ID_REMESSA = :P_ID", oPar)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "DELETE FROM TBL_REMESSA WHERE ID_REMESSA = :P_ID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    ' Replaces the whole client alias mapping (TBL_CLIENT_ALIAS). Runs inside the open
    ' transaction; names arrive already normalized (UPPER/TRIM) from the mapping dialog.
    Public Function substituirAliases(pMapa As Dictionary(Of String, List(Of String)), pUsuario As String) As Boolean

        Try

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text, "DELETE FROM TBL_CLIENT_ALIAS")

            For Each par As KeyValuePair(Of String, List(Of String)) In pMapa
                For Each sAlias As String In par.Value
                    Dim oPar(2) As OracleParameter
                    oPar(0) = New OracleParameter("P_CANONICAL", OracleDbType.Varchar2, ParameterDirection.Input) With {.Value = par.Key}
                    oPar(1) = New OracleParameter("P_ALIAS",     OracleDbType.Varchar2, ParameterDirection.Input) With {.Value = sAlias}
                    oPar(2) = New OracleParameter("P_USER",      OracleDbType.Varchar2, ParameterDirection.Input) With {
                                  .Value = If(String.IsNullOrWhiteSpace(pUsuario), DBNull.Value, CObj(pUsuario))}
                    OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                        "INSERT INTO TBL_CLIENT_ALIAS (ID_ALIAS, CANONICAL_NAME, ALIAS_NAME, CREATED_AT, CREATED_BY)" &
                        " VALUES (SEQ_CLIENT_ALIAS.NEXTVAL, :P_CANONICAL, :P_ALIAS, SYSDATE, :P_USER)", oPar)
                Next
            Next

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function cancelarRemessa(pIdRemessa As Integer) As Boolean

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdRemessa

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "UPDATE TBL_EQUIPAMENTO SET STATUS = 'IN_STOCK', DATA_ATUALIZACAO = SYSDATE " &
                " WHERE ID_EQUIPAMENTO IN (SELECT ID_EQUIPAMENTO FROM TBL_REMESSA_ITEM WHERE ID_REMESSA = :P_ID)", oPar)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "DELETE FROM TBL_REMESSA_ITEM WHERE ID_REMESSA = :P_ID", oPar)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "DELETE FROM TBL_REMESSA WHERE ID_REMESSA = :P_ID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

#End Region

#Region "Auto-link de cliente"

    ' Cliente ativo cujo nome, nome fantasia ou um dos ALIASES (";"-separados, gravados em
    ' maiusculas) e igual ao DESTINATARIO da remessa (UPPER/TRIM).
    Private Const SqlMatchCliente As String =
        "SELECT {0} FROM TBL_CLIENTE C" &
        " WHERE C.ATIVO = 'Y'" &
        "   AND (UPPER(TRIM(C.NOME_RAZAO)) = UPPER(TRIM(R.DESTINATARIO))" &
        "        OR UPPER(TRIM(C.NOME_FANTASIA)) = UPPER(TRIM(R.DESTINATARIO))" &
        "        OR INSTR(';' || UPPER(C.ALIASES) || ';', ';' || UPPER(TRIM(R.DESTINATARIO)) || ';') > 0)"

    ''' <summary>
    ''' Preenche TBL_REMESSA.ID_CLIENTE onde ainda e NULL e o DESTINATARIO bate com EXATAMENTE um
    ''' cliente ativo. Nunca sobrescreve um ID_CLIENTE ja definido. pIdRemessa = Nothing: todas.
    ''' Usa a transacao aberta (criacao da remessa) ou abre uma propria (botao Auto-link).
    ''' </summary>
    Public Function autoVincularClientes(pIdRemessa As Integer?) As (Vinculadas As Integer, SemMatch As Integer, Ambiguas As Integer)

        Dim filtro As String = If(pIdRemessa.HasValue, " AND R.ID_REMESSA = :P_ID_REMESSA", "")
        Dim sqlUpdate As String =
            "UPDATE TBL_REMESSA R SET R.ID_CLIENTE = (" & String.Format(SqlMatchCliente, "MIN(C.ID_CLIENTE)") & ")" &
            " WHERE R.ID_CLIENTE IS NULL" & filtro &
            "   AND (" & String.Format(SqlMatchCliente, "COUNT(*)") & ") = 1"
        Dim sqlRestantes As String =
            "SELECT NVL(SUM(CASE WHEN N = 0 THEN 1 ELSE 0 END), 0) AS SEM_MATCH," &
            "       NVL(SUM(CASE WHEN N > 1 THEN 1 ELSE 0 END), 0) AS AMBIGUAS" &
            "  FROM (SELECT (" & String.Format(SqlMatchCliente, "COUNT(*)") & ") AS N" &
            "          FROM TBL_REMESSA R WHERE R.ID_CLIENTE IS NULL" & filtro & ")"

        Dim propria As OracleConnection = Nothing
        Dim tx As OracleTransaction = Me.oTransacao
        Dim emTransacaoExterna As Boolean = tx IsNot Nothing AndAlso tx.Connection IsNot Nothing
        Try
            If Not emTransacaoExterna Then
                propria = New OracleConnection(Me.ConnectionString)
                propria.Open()
                tx = propria.BeginTransaction()
            End If

            Dim vinculadas As Integer
            Using cmd As New OracleCommand(sqlUpdate, tx.Connection)
                cmd.Transaction = tx
                cmd.BindByName = True
                If pIdRemessa.HasValue Then cmd.Parameters.Add("P_ID_REMESSA", OracleDbType.Int32).Value = pIdRemessa.Value
                vinculadas = cmd.ExecuteNonQuery()
            End Using

            Dim semMatch As Integer = 0, ambiguas As Integer = 0
            Using cmd As New OracleCommand(sqlRestantes, tx.Connection)
                cmd.Transaction = tx
                cmd.BindByName = True
                If pIdRemessa.HasValue Then cmd.Parameters.Add("P_ID_REMESSA", OracleDbType.Int32).Value = pIdRemessa.Value
                Using rd As OracleDataReader = cmd.ExecuteReader()
                    If rd.Read() Then
                        semMatch = Convert.ToInt32(rd("SEM_MATCH"))
                        ambiguas = Convert.ToInt32(rd("AMBIGUAS"))
                    End If
                End Using
            End Using

            If Not emTransacaoExterna Then tx.Commit()
            Return (vinculadas, semMatch, ambiguas)

        Catch ex As Exception
            If Not emTransacaoExterna AndAlso tx IsNot Nothing Then
                Try : tx.Rollback() : Catch : End Try
            End If
            Throw New Exception(ex.ToString)
        Finally
            If propria IsNot Nothing Then propria.Dispose()
        End Try

    End Function

#End Region

#Region "Cliente da remessa"

    ' Define (ou limpa, com Nothing) o cliente da remessa — TBL_REMESSA.ID_CLIENTE.
    ' Fora de transacao: so esta coluna e alterada; status e itens nao sao tocados.
    Public Sub atualizarCliente(pIdRemessa As Integer, pIdCliente As Integer?)

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_ID_CLIENTE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_ID_REMESSA", OracleDbType.Int32, ParameterDirection.Input)
        oPar(0).Value = If(pIdCliente.HasValue, CObj(pIdCliente.Value), DBNull.Value)
        oPar(1).Value = pIdRemessa

        Try
            OracleHelper.ExecuteNonQuery(Me.ConnectionString, CommandType.Text,
                "UPDATE TBL_REMESSA SET ID_CLIENTE = :P_ID_CLIENTE WHERE ID_REMESSA = :P_ID_REMESSA", oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Sub

#End Region

End Class
