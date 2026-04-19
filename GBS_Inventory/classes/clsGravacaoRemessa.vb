Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravação de Remessas (FedEx, UPS, USPS)
''' 
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

        Dim oPar(14) As OracleParameter

        oPar(0)  = New OracleParameter("V_DIRECTION",          OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1)  = New OracleParameter("V_CARRIER",            OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2)  = New OracleParameter("V_TRACKING_NUMBER",    OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3)  = New OracleParameter("V_SENDER_NAME",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4)  = New OracleParameter("V_SENDER_ADDRESS",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5)  = New OracleParameter("V_RECIPIENT_NAME",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6)  = New OracleParameter("V_RECIPIENT_ADDRESS",  OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(7)  = New OracleParameter("V_WEIGHT_LBS",         OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(8)  = New OracleParameter("V_SHIPPING_COST_USD",  OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(9)  = New OracleParameter("V_INSURANCE_USD",      OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_SERVICE_LEVEL",      OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("V_ESTIMATED_DELIVERY", OracleDbType.Date,     ParameterDirection.Input)
        oPar(12) = New OracleParameter("V_NOTES",              OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(13) = New OracleParameter("V_ID",                 OracleDbType.Int32,    ParameterDirection.Output)
        oPar(14) = New OracleParameter("V_REF",                OracleDbType.Varchar2, 40, Nothing, ParameterDirection.Output)

        Dim vIdGerado As OracleDecimal

        Try

            oPar(0).Value  = pRemessa.Direction
            oPar(1).Value  = pRemessa.Carrier
            oPar(2).Value  = If(String.IsNullOrEmpty(pRemessa.TrackingNumber), DBNull.Value, CObj(pRemessa.TrackingNumber))
            oPar(3).Value  = If(String.IsNullOrEmpty(pRemessa.SenderName), DBNull.Value, CObj(pRemessa.SenderName))
            oPar(4).Value  = If(String.IsNullOrEmpty(pRemessa.SenderAddress), DBNull.Value, CObj(pRemessa.SenderAddress))
            oPar(5).Value  = If(String.IsNullOrEmpty(pRemessa.RecipientName), DBNull.Value, CObj(pRemessa.RecipientName))
            oPar(6).Value  = If(String.IsNullOrEmpty(pRemessa.RecipientAddress), DBNull.Value, CObj(pRemessa.RecipientAddress))
            oPar(7).Value  = If(pRemessa.WeightLbs.HasValue, CObj(pRemessa.WeightLbs.Value), DBNull.Value)
            oPar(8).Value  = If(pRemessa.ShippingCostUsd.HasValue, CObj(pRemessa.ShippingCostUsd.Value), DBNull.Value)
            oPar(9).Value  = If(pRemessa.InsuranceUsd.HasValue, CObj(pRemessa.InsuranceUsd.Value), DBNull.Value)
            oPar(10).Value = If(String.IsNullOrEmpty(pRemessa.ServiceLevel), DBNull.Value, CObj(pRemessa.ServiceLevel))
            oPar(11).Value = If(pRemessa.EstimatedDelivery.HasValue, CObj(pRemessa.EstimatedDelivery.Value), DBNull.Value)
            oPar(12).Value = If(String.IsNullOrEmpty(pRemessa.Notes), DBNull.Value, CObj(pRemessa.Notes))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_REMESSA.PROC_INSERT_REMESSA", oPar)

            vIdGerado    = CType(oPar(13).Value, OracleDecimal)
            pRemessaRef  = oPar(14).Value.ToString()

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return CInt(vIdGerado.Value)

    End Function

    Public Function atualizarStatusRemessa(pIdRemessa As Integer, pStatus As String, pDataEntrega As Date?) As Boolean

        Dim oPar(2) As OracleParameter

        oPar(0) = New OracleParameter("V_ID",           OracleDbType.Int32,    ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_STATUS",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_DATA_ENTREGA", OracleDbType.Date,     ParameterDirection.Input)

        Try

            oPar(0).Value = pIdRemessa
            oPar(1).Value = pStatus
            oPar(2).Value = If(pDataEntrega.HasValue, CObj(pDataEntrega.Value), DBNull.Value)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_REMESSA.PROC_UPDATE_STATUS_REMESSA", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function incluirItemRemessa(pItem As ItemRemessa) As Boolean

        Dim oPar(5) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_REMESSA",     OracleDbType.Int32,    ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_ID_EQUIPAMENTO", OracleDbType.Int32,    ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_INTERNAL_UID",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("V_CONDITION",      OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4) = New OracleParameter("V_SALE_PRICE_USD", OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(5) = New OracleParameter("V_NOTES",          OracleDbType.Varchar2, ParameterDirection.Input)

        Try

            oPar(0).Value = pItem.IdRemessa
            oPar(1).Value = pItem.IdEquipamento
            oPar(2).Value = pItem.InternalUID
            oPar(3).Value = pItem.ConditionAtShip
            oPar(4).Value = If(pItem.SalePriceUsd.HasValue, CObj(pItem.SalePriceUsd.Value), DBNull.Value)
            oPar(5).Value = If(String.IsNullOrEmpty(pItem.Notes), DBNull.Value, CObj(pItem.Notes))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_REMESSA.PROC_INSERT_REMESSA_ITEM", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function excluirRemessa(pIdRemessa As Integer) As Boolean

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdRemessa

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_REMESSA.PROC_DELETE_REMESSA", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

#End Region

End Class
