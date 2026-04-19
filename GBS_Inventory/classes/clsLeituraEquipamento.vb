Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Classe de leitura de Equipamentos
''' SELECTs via PACK_EQUIPAMENTO
''' 
''' </summary>
Public Class clsLeituraEquipamento

#Region "Atributos"

    Private ConnectionString As String

#End Region

#Region "Construtor"

    Public Sub New()

        Try

            Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Sub

#End Region

#Region "Consultas"

    Public Function selecionarPorId(pIdEquipamento As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID",     OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdEquipamento

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_EQUIPAMENTO", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    ''' <summary>
    ''' Busca CRÍTICA do scanner — localiza equipamento pelo Internal UID
    ''' </summary>
    Public Function selecionarPorUID(pInternalUID As String) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_INTERNAL_UID", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR",       OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pInternalUID

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_EQUIPAMENTO_UID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarPorSerial(pSerial As String) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_SERIAL", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pSerial

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_EQUIPAMENTO_SERIAL", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarTodos(pIdEmpresa As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_EMPRESA", OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR",  OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdEmpresa

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_EQUIPAMENTOS", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarPorFiltro(pPesquisa As String, pManufacturer As String, pCondition As String, pStatus As String) As DataSet

        Dim oPar(4) As OracleParameter

        oPar(0) = New OracleParameter("V_PESQUISA",     OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_MANUFACTURER", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_CONDITION",    OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(3) = New OracleParameter("V_STATUS",       OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(4) = New OracleParameter("V_CURSOR",       OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = If(String.IsNullOrEmpty(pPesquisa),     DBNull.Value, CObj(pPesquisa))
            oPar(1).Value = If(String.IsNullOrEmpty(pManufacturer), DBNull.Value, CObj(pManufacturer))
            oPar(2).Value = If(String.IsNullOrEmpty(pCondition),    DBNull.Value, CObj(pCondition))
            oPar(3).Value = If(String.IsNullOrEmpty(pStatus),       DBNull.Value, CObj(pStatus))

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_POR_FILTRO", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function obterDashboardTotais(pIdEmpresa As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_EMPRESA", OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR",  OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdEmpresa

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_DASHBOARD_TOTAIS", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function obterResumoWhatsApp(pIdEmpresa As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_EMPRESA", OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR",  OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdEmpresa

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_RESUMO_WHATSAPP", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

End Class
