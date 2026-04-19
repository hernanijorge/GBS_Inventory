Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravação de Upgrades (RAM, SSD, etc)
''' 
''' </summary>
Public Class clsGravacaoUpgrade

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

    Public Function incluirUpgrade(pUpgrade As Upgrade) As Integer

        Dim oPar(10) As OracleParameter

        oPar(0)  = New OracleParameter("V_ID_EQUIPAMENTO", OracleDbType.Int32,    ParameterDirection.Input)
        oPar(1)  = New OracleParameter("V_INTERNAL_UID",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2)  = New OracleParameter("V_COMPONENT_TYPE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3)  = New OracleParameter("V_VALUE_BEFORE",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4)  = New OracleParameter("V_VALUE_AFTER",    OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5)  = New OracleParameter("V_PART_SERIAL",    OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6)  = New OracleParameter("V_SOURCE_ORIGEM",  OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(7)  = New OracleParameter("V_COST_USD",       OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(8)  = New OracleParameter("V_TECHNICIAN",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9)  = New OracleParameter("V_NOTES",          OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_ID",             OracleDbType.Int32,    ParameterDirection.Output)

        Dim vIdGerado As OracleDecimal

        Try

            oPar(0).Value = pUpgrade.IdEquipamento
            oPar(1).Value = pUpgrade.InternalUID
            oPar(2).Value = pUpgrade.ComponentType
            oPar(3).Value = If(String.IsNullOrEmpty(pUpgrade.ValueBefore), DBNull.Value, CObj(pUpgrade.ValueBefore))
            oPar(4).Value = pUpgrade.ValueAfter
            oPar(5).Value = If(String.IsNullOrEmpty(pUpgrade.PartSerial), DBNull.Value, CObj(pUpgrade.PartSerial))
            oPar(6).Value = pUpgrade.SourceOrigem
            oPar(7).Value = If(pUpgrade.CostUsd.HasValue, CObj(pUpgrade.CostUsd.Value), DBNull.Value)
            oPar(8).Value = If(String.IsNullOrEmpty(pUpgrade.Technician), DBNull.Value, CObj(pUpgrade.Technician))
            oPar(9).Value = If(String.IsNullOrEmpty(pUpgrade.Notes), DBNull.Value, CObj(pUpgrade.Notes))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_INSERT_UPGRADE", oPar)

            vIdGerado = CType(oPar(10).Value, OracleDecimal)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return CInt(vIdGerado.Value)

    End Function

    Public Function excluirUpgrade(pIdUpgrade As Integer) As Boolean

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdUpgrade

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_DELETE_UPGRADE", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

#End Region

End Class
