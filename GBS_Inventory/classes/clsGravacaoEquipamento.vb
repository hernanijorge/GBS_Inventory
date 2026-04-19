Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravação de Equipamentos
''' INSERT, UPDATE, DELETE via PACK_EQUIPAMENTO
''' 
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

    Public Function incluirEquipamento(pEquipamento As Equipamento) As Integer

        Dim oPar(19) As OracleParameter

        oPar(0)  = New OracleParameter("V_INTERNAL_UID",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1)  = New OracleParameter("V_MANUFACTURER",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2)  = New OracleParameter("V_MODEL",                OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3)  = New OracleParameter("V_SERIAL_NUMBER",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4)  = New OracleParameter("V_CPU_FAMILY",           OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5)  = New OracleParameter("V_CPU_MODEL",            OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6)  = New OracleParameter("V_CPU_SPEED_GHZ",        OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(7)  = New OracleParameter("V_RAM_GB",               OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8)  = New OracleParameter("V_STORAGE_GB",           OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9)  = New OracleParameter("V_HARD_DRIVE_TYPE",      OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_RESOLUTION",           OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("V_GRAPHICS",             OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(12) = New OracleParameter("V_DEVICE_TYPE",          OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(13) = New OracleParameter("V_CONDITION_STATUS",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(14) = New OracleParameter("V_NOTES",                OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(15) = New OracleParameter("V_SOURCE_BATCH",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(16) = New OracleParameter("V_STATUS",               OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(17) = New OracleParameter("V_ID_EMPRESA",           OracleDbType.Int32,    ParameterDirection.Input)
        oPar(18) = New OracleParameter("V_ID_USUARIO_CADASTRO",  OracleDbType.Int32,    ParameterDirection.Input)
        oPar(19) = New OracleParameter("V_ID",                   OracleDbType.Int32,    ParameterDirection.Output)

        Dim vIdGerado As OracleDecimal

        Try

            oPar(0).Value  = pEquipamento.InternalUID
            oPar(1).Value  = pEquipamento.Manufacturer
            oPar(2).Value  = pEquipamento.Model
            oPar(3).Value  = pEquipamento.SerialNumber
            oPar(4).Value  = If(String.IsNullOrEmpty(pEquipamento.CpuFamily), DBNull.Value, CObj(pEquipamento.CpuFamily))
            oPar(5).Value  = If(String.IsNullOrEmpty(pEquipamento.CpuModel), DBNull.Value, CObj(pEquipamento.CpuModel))
            oPar(6).Value  = If(pEquipamento.CpuSpeedGhz.HasValue, CObj(pEquipamento.CpuSpeedGhz.Value), DBNull.Value)
            oPar(7).Value  = If(String.IsNullOrEmpty(pEquipamento.RamGb), DBNull.Value, CObj(pEquipamento.RamGb))
            oPar(8).Value  = If(String.IsNullOrEmpty(pEquipamento.StorageGb), DBNull.Value, CObj(pEquipamento.StorageGb))
            oPar(9).Value  = If(String.IsNullOrEmpty(pEquipamento.HardDriveType), DBNull.Value, CObj(pEquipamento.HardDriveType))
            oPar(10).Value = If(String.IsNullOrEmpty(pEquipamento.Resolution), DBNull.Value, CObj(pEquipamento.Resolution))
            oPar(11).Value = If(String.IsNullOrEmpty(pEquipamento.Graphics), DBNull.Value, CObj(pEquipamento.Graphics))
            oPar(12).Value = pEquipamento.DeviceType
            oPar(13).Value = pEquipamento.ConditionStatus
            oPar(14).Value = If(String.IsNullOrEmpty(pEquipamento.Notes), DBNull.Value, CObj(pEquipamento.Notes))
            oPar(15).Value = If(String.IsNullOrEmpty(pEquipamento.SourceBatch), DBNull.Value, CObj(pEquipamento.SourceBatch))
            oPar(16).Value = pEquipamento.Status
            oPar(17).Value = pEquipamento.IdEmpresa
            oPar(18).Value = If(pEquipamento.IdUsuarioCadastro.HasValue, CObj(pEquipamento.IdUsuarioCadastro.Value), DBNull.Value)

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_INSERT_EQUIPAMENTO", oPar)

            vIdGerado = CType(oPar(19).Value, OracleDecimal)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return CInt(vIdGerado.Value)

    End Function

    Public Function alterarEquipamento(pEquipamento As Equipamento) As Boolean

        Dim oPar(14) As OracleParameter

        oPar(0)  = New OracleParameter("V_MANUFACTURER",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1)  = New OracleParameter("V_MODEL",            OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2)  = New OracleParameter("V_CPU_FAMILY",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3)  = New OracleParameter("V_CPU_MODEL",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4)  = New OracleParameter("V_CPU_SPEED_GHZ",    OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(5)  = New OracleParameter("V_RAM_GB",           OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6)  = New OracleParameter("V_STORAGE_GB",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(7)  = New OracleParameter("V_HARD_DRIVE_TYPE",  OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8)  = New OracleParameter("V_RESOLUTION",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9)  = New OracleParameter("V_GRAPHICS",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_DEVICE_TYPE",      OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("V_CONDITION_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(12) = New OracleParameter("V_NOTES",            OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(13) = New OracleParameter("V_STATUS",           OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(14) = New OracleParameter("V_ID",               OracleDbType.Int32,    ParameterDirection.Input)

        Try

            oPar(0).Value  = pEquipamento.Manufacturer
            oPar(1).Value  = pEquipamento.Model
            oPar(2).Value  = If(String.IsNullOrEmpty(pEquipamento.CpuFamily), DBNull.Value, CObj(pEquipamento.CpuFamily))
            oPar(3).Value  = If(String.IsNullOrEmpty(pEquipamento.CpuModel), DBNull.Value, CObj(pEquipamento.CpuModel))
            oPar(4).Value  = If(pEquipamento.CpuSpeedGhz.HasValue, CObj(pEquipamento.CpuSpeedGhz.Value), DBNull.Value)
            oPar(5).Value  = If(String.IsNullOrEmpty(pEquipamento.RamGb), DBNull.Value, CObj(pEquipamento.RamGb))
            oPar(6).Value  = If(String.IsNullOrEmpty(pEquipamento.StorageGb), DBNull.Value, CObj(pEquipamento.StorageGb))
            oPar(7).Value  = If(String.IsNullOrEmpty(pEquipamento.HardDriveType), DBNull.Value, CObj(pEquipamento.HardDriveType))
            oPar(8).Value  = If(String.IsNullOrEmpty(pEquipamento.Resolution), DBNull.Value, CObj(pEquipamento.Resolution))
            oPar(9).Value  = If(String.IsNullOrEmpty(pEquipamento.Graphics), DBNull.Value, CObj(pEquipamento.Graphics))
            oPar(10).Value = pEquipamento.DeviceType
            oPar(11).Value = pEquipamento.ConditionStatus
            oPar(12).Value = If(String.IsNullOrEmpty(pEquipamento.Notes), DBNull.Value, CObj(pEquipamento.Notes))
            oPar(13).Value = pEquipamento.Status
            oPar(14).Value = pEquipamento.IdEquipamento

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_UPDATE_EQUIPAMENTO", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function atualizarStatus(pIdEquipamento As Integer, pStatus As String) As Boolean

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_ID",     OracleDbType.Int32,    ParameterDirection.Input)

        Try

            oPar(0).Value = pStatus
            oPar(1).Value = pIdEquipamento

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_UPDATE_STATUS", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

    Public Function excluirEquipamento(pIdEquipamento As Integer) As Boolean

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdEquipamento

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_DELETE_EQUIPAMENTO", oPar)

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

        Dim oPar(13) As OracleParameter

        oPar(0)  = New OracleParameter("V_INTERNAL_UID",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1)  = New OracleParameter("V_MANUFACTURER",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2)  = New OracleParameter("V_MODEL",            OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3)  = New OracleParameter("V_SERIAL_NUMBER",    OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(4)  = New OracleParameter("V_CPU_FAMILY",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(5)  = New OracleParameter("V_CPU_MODEL",        OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6)  = New OracleParameter("V_CPU_SPEED_GHZ",    OracleDbType.Decimal,  ParameterDirection.Input)
        oPar(7)  = New OracleParameter("V_RAM_GB",           OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8)  = New OracleParameter("V_STORAGE_GB",       OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(9)  = New OracleParameter("V_HARD_DRIVE_TYPE",  OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_CONDITION_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("V_NOTES",            OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(12) = New OracleParameter("V_SOURCE_BATCH",     OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(13) = New OracleParameter("V_RESULTADO",        OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Output)

        Dim vResultado As String = ""

        Try

            oPar(0).Value  = pEquipamento.InternalUID
            oPar(1).Value  = pEquipamento.Manufacturer
            oPar(2).Value  = pEquipamento.Model
            oPar(3).Value  = pEquipamento.SerialNumber
            oPar(4).Value  = If(String.IsNullOrEmpty(pEquipamento.CpuFamily), DBNull.Value, CObj(pEquipamento.CpuFamily))
            oPar(5).Value  = If(String.IsNullOrEmpty(pEquipamento.CpuModel), DBNull.Value, CObj(pEquipamento.CpuModel))
            oPar(6).Value  = If(pEquipamento.CpuSpeedGhz.HasValue, CObj(pEquipamento.CpuSpeedGhz.Value), DBNull.Value)
            oPar(7).Value  = If(String.IsNullOrEmpty(pEquipamento.RamGb), DBNull.Value, CObj(pEquipamento.RamGb))
            oPar(8).Value  = If(String.IsNullOrEmpty(pEquipamento.StorageGb), DBNull.Value, CObj(pEquipamento.StorageGb))
            oPar(9).Value  = If(String.IsNullOrEmpty(pEquipamento.HardDriveType), DBNull.Value, CObj(pEquipamento.HardDriveType))
            oPar(10).Value = pEquipamento.ConditionStatus
            oPar(11).Value = If(String.IsNullOrEmpty(pEquipamento.Notes), DBNull.Value, CObj(pEquipamento.Notes))
            oPar(12).Value = If(String.IsNullOrEmpty(pEquipamento.SourceBatch), DBNull.Value, CObj(pEquipamento.SourceBatch))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_UPSERT_EQUIPAMENTO", oPar)

            vResultado = oPar(13).Value.ToString()

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return vResultado

    End Function

#End Region

End Class
