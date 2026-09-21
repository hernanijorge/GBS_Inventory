Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravação de Equipamentos
''' INSERT, UPDATE, DELETE via PACK_EQUIPAMENTO
''' </summary>
Public Class clsGravacaoEquipamento

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

    ' Extrai número de uma string como "8GB" ou "256" → Integer
    Private Function ParseGb(s As String) As Integer
        If String.IsNullOrEmpty(s) Then Return 0
        Dim v As Integer = 0
        Dim m As System.Text.RegularExpressions.Match =
            System.Text.RegularExpressions.Regex.Match(s, "\d+")
        If m.Success Then Integer.TryParse(m.Value, v)
        Return v
    End Function

    Private Function DbConditionStatus(pValor As String) As Object
        If String.IsNullOrWhiteSpace(pValor) Then Return "GOOD"
        Return pValor.Trim().ToUpperInvariant()
    End Function

    Public Function incluirEquipamento(pEquipamento As Equipamento) As Integer

        Dim oPar(11) As OracleParameter

        oPar(0)  = New OracleParameter("P_INTERNAL_UID",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1)  = New OracleParameter("P_SERIAL_NUMBER",      OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2)  = New OracleParameter("P_MODELO",             OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3)  = New OracleParameter("P_MARCA",              OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4)  = New OracleParameter("P_PROCESSADOR",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5)  = New OracleParameter("P_RAM_GB",             OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(6)  = New OracleParameter("P_STORAGE_GB",         OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(7)  = New OracleParameter("P_CONDITION_STATUS",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8)  = New OracleParameter("P_STATUS_EQUIPAMENTO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9)  = New OracleParameter("P_OBSERVACAO",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("P_BATTERY_CHECK",      OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("P_SCREEN_SIZE",        OracleDbType.Varchar2, ParameterDirection.Input)

        Try

            oPar(0).Value  = pEquipamento.InternalUID
            oPar(1).Value  = If(String.IsNullOrEmpty(pEquipamento.SerialNumber), DBNull.Value, CObj(pEquipamento.SerialNumber))
            oPar(2).Value  = pEquipamento.Model
            oPar(3).Value  = pEquipamento.Manufacturer
            oPar(4).Value  = If(String.IsNullOrEmpty(pEquipamento.CpuModel), DBNull.Value, CObj(pEquipamento.CpuModel))
            oPar(5).Value  = ParseGb(pEquipamento.RamGb)
            oPar(6).Value  = ParseGb(pEquipamento.StorageGb)
            oPar(7).Value  = DbConditionStatus(pEquipamento.ConditionStatus)
            oPar(8).Value  = If(String.IsNullOrEmpty(pEquipamento.Status), "IN_STOCK", pEquipamento.Status)
            oPar(9).Value  = If(String.IsNullOrEmpty(pEquipamento.Notes), DBNull.Value, CObj(pEquipamento.Notes))
            oPar(10).Value = If(String.IsNullOrEmpty(pEquipamento.BatteryCheck), DBNull.Value, CObj(pEquipamento.BatteryCheck))
            oPar(11).Value = If(String.IsNullOrWhiteSpace(pEquipamento.ScreenSize), DBNull.Value, CObj(pEquipamento.ScreenSize.Trim()))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_INSERT", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return 0

    End Function

    Public Function alterarEquipamento(pEquipamento As Equipamento) As Boolean

        Dim oPar(10) As OracleParameter

        oPar(0) = New OracleParameter("P_INTERNAL_UID",  OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_SERIAL_NUMBER", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_MODEL",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("P_MARCA",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4) = New OracleParameter("P_PROCESSADOR",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5) = New OracleParameter("P_RAM_GB",        OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(6) = New OracleParameter("P_STORAGE_GB",    OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(7) = New OracleParameter("P_CONDITION_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8) = New OracleParameter("P_STATUS",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9) = New OracleParameter("P_OBSERVACAO",    OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("P_SCREEN_SIZE",  OracleDbType.Varchar2, ParameterDirection.Input)

        Try

            oPar(0).Value = pEquipamento.InternalUID
            oPar(1).Value = If(String.IsNullOrEmpty(pEquipamento.SerialNumber), DBNull.Value, CObj(pEquipamento.SerialNumber))
            oPar(2).Value = pEquipamento.Model
            oPar(3).Value = pEquipamento.Manufacturer
            oPar(4).Value = If(String.IsNullOrEmpty(pEquipamento.CpuModel), DBNull.Value, CObj(pEquipamento.CpuModel))
            oPar(5).Value = ParseGb(pEquipamento.RamGb)
            oPar(6).Value = ParseGb(pEquipamento.StorageGb)
            oPar(7).Value = DbConditionStatus(pEquipamento.ConditionStatus)
            oPar(8).Value = If(String.IsNullOrEmpty(pEquipamento.Status), "IN_STOCK", pEquipamento.Status)
            oPar(9).Value = If(String.IsNullOrEmpty(pEquipamento.Notes), DBNull.Value, CObj(pEquipamento.Notes))
            oPar(10).Value = If(String.IsNullOrWhiteSpace(pEquipamento.ScreenSize), DBNull.Value, CObj(pEquipamento.ScreenSize.Trim()))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_UPSERT_EQUIPAMENTO", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function atualizarStatus(pIdEquipamento As Integer, pStatus As String) As Boolean

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("P_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_ID",     OracleDbType.Int32,    ParameterDirection.Input)

        Try

            oPar(0).Value = pStatus
            oPar(1).Value = pIdEquipamento

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "UPDATE TBL_EQUIPAMENTO SET STATUS = :P_STATUS, DATA_ATUALIZACAO = SYSDATE WHERE ID_EQUIPAMENTO = :P_ID",
                oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function excluirEquipamento(pIdEquipamento As Integer) As Boolean

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdEquipamento

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "DELETE FROM TBL_REMESSA_ITEM WHERE ID_EQUIPAMENTO = :P_ID", oPar)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "DELETE FROM TBL_EQUIPAMENTO_UPGRADE WHERE ID_EQUIPAMENTO = :P_ID", oPar)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.Text,
                "DELETE FROM TBL_EQUIPAMENTO WHERE ID_EQUIPAMENTO = :P_ID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    ''' <summary>
    ''' UPSERT em massa (usado pela importação de planilha Excel)
    ''' Retorna "INSERTED", "UPDATED" ou "ERROR: xxx"
    ''' </summary>
    Public Function upsertEquipamento(pEquipamento As Equipamento) As String

        Dim oPar(10) As OracleParameter

        oPar(0) = New OracleParameter("P_INTERNAL_UID",  OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_SERIAL_NUMBER", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_MODEL",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("P_MARCA",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4) = New OracleParameter("P_PROCESSADOR",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5) = New OracleParameter("P_RAM_GB",        OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(6) = New OracleParameter("P_STORAGE_GB",    OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(7) = New OracleParameter("P_CONDITION_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8) = New OracleParameter("P_STATUS",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9) = New OracleParameter("P_OBSERVACAO",    OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("P_SCREEN_SIZE",  OracleDbType.Varchar2, ParameterDirection.Input)

        Try

            oPar(0).Value = pEquipamento.InternalUID
            oPar(1).Value = If(String.IsNullOrEmpty(pEquipamento.SerialNumber), DBNull.Value, CObj(pEquipamento.SerialNumber))
            oPar(2).Value = pEquipamento.Model
            oPar(3).Value = pEquipamento.Manufacturer
            oPar(4).Value = If(String.IsNullOrEmpty(pEquipamento.CpuModel), DBNull.Value, CObj(pEquipamento.CpuModel))
            oPar(5).Value = ParseGb(pEquipamento.RamGb)
            oPar(6).Value = ParseGb(pEquipamento.StorageGb)
            oPar(7).Value = DbConditionStatus(pEquipamento.ConditionStatus)
            oPar(8).Value = If(String.IsNullOrEmpty(pEquipamento.Status), "IN_STOCK", pEquipamento.Status)
            oPar(9).Value = If(String.IsNullOrEmpty(pEquipamento.Notes), DBNull.Value, CObj(pEquipamento.Notes))
            oPar(10).Value = If(String.IsNullOrWhiteSpace(pEquipamento.ScreenSize), DBNull.Value, CObj(pEquipamento.ScreenSize.Trim()))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_UPSERT_EQUIPAMENTO", oPar)

            Return "INSERTED"

        Catch ex As Exception

            Return "ERROR: " & ex.Message

        End Try

    End Function

#End Region

End Class
