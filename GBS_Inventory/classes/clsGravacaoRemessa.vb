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

End Class
