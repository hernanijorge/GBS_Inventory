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

        oPar(0) = New OracleParameter("P_ID",     OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdEquipamento

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_ID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    ''' <summary>
    ''' Busca CRÍTICA do scanner — localiza equipamento pelo Internal UID
    ''' </summary>
    Public Function selecionarPorUID(pInternalUID As String) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("P_INTERNAL_UID", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_CURSOR",       OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pInternalUID

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_UID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarPorSerial(pSerial As String) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("P_SERIAL", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pSerial

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_SERIAL", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarTodos(pIdEmpresa As Integer) As DataSet

        Dim oPar(2) As OracleParameter

        oPar(0) = New OracleParameter("P_FILTRO",             OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_STATUS_EQUIPAMENTO", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_CURSOR",             OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = DBNull.Value
            oPar(1).Value = DBNull.Value

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_FILTRO", oPar)

        Catch ex As Exception
            If IsErroObjetoEquipamento(ex) Then
                Return selecionarFiltroSemPackage(Nothing, Nothing)
            End If
            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarPorFiltro(pPesquisa As String, pManufacturer As String, pCondition As String, pStatus As String) As DataSet

        Dim oPar(2) As OracleParameter

        oPar(0) = New OracleParameter("P_FILTRO",             OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_STATUS_EQUIPAMENTO", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_CURSOR",             OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = If(String.IsNullOrEmpty(pPesquisa), DBNull.Value, CObj(pPesquisa))
            oPar(1).Value = If(String.IsNullOrEmpty(pStatus),   DBNull.Value, CObj(pStatus))

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_SELECT_FILTRO", oPar)

        Catch ex As Exception
            If IsErroObjetoEquipamento(ex) Then
                Return selecionarFiltroSemPackage(pPesquisa, pStatus)
            End If
            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function obterDashboardTotais(pIdEmpresa As Integer) As DataSet

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_DASHBOARD", oPar)

        Catch ex As Exception
            If IsErroObjetoEquipamento(ex) Then
                Return obterDashboardTotaisSemPackage()
            End If
            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function obterResumoWhatsApp(pIdEmpresa As Integer) As DataSet

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_EQUIPAMENTO.PROC_RESUMO_WHATSAPP", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

#Region "Fallback sem package"

    Private Function selecionarFiltroSemPackage(pPesquisa As String, pStatus As String) As DataSet

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_FILTRO", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(0).Value = If(String.IsNullOrWhiteSpace(pPesquisa), DBNull.Value, CObj(pPesquisa))
        oPar(1).Value = If(String.IsNullOrWhiteSpace(pStatus), DBNull.Value, CObj(pStatus))

        Dim sqlPt As String =
            "SELECT E.ID_EQUIPAMENTO, E.INTERNAL_UID, E.SERIAL_NUMBER, " &
            "       E.MARCA, E.MODEL, E.PROCESSADOR, E.RAM_GB, E.STORAGE_GB, " &
            "       E.STATUS, CAST(NULL AS VARCHAR2(4000)) AS OBSERVACAO, E.DATA_CADASTRO, CAST(NULL AS DATE) AS DATA_ATUALIZACAO, " &
            "       E.STATUS AS STATUS_DESCRICAO " &
            "  FROM TBL_EQUIPAMENTO E " &
            " WHERE (:P_FILTRO IS NULL " &
            "        OR UPPER(E.INTERNAL_UID) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.SERIAL_NUMBER) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.MODEL) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.MARCA) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.PROCESSADOR) LIKE '%' || UPPER(:P_FILTRO) || '%') " &
            "   AND (:P_STATUS IS NULL OR UPPER(E.STATUS) = UPPER(:P_STATUS)) " &
            " ORDER BY E.DATA_CADASTRO DESC"

        Dim sqlEn As String =
            "SELECT E.ID_EQUIPAMENTO, E.INTERNAL_UID, E.SERIAL_NUMBER, " &
            "       E.MANUFACTURER AS MARCA, E.MODEL, E.CPU_MODEL AS PROCESSADOR, E.RAM_GB, E.STORAGE_GB, " &
            "       E.STATUS, CAST(NULL AS VARCHAR2(4000)) AS OBSERVACAO, E.DATA_CADASTRO, CAST(NULL AS DATE) AS DATA_ATUALIZACAO, " &
            "       E.STATUS AS STATUS_DESCRICAO " &
            "  FROM TBL_EQUIPAMENTO E " &
            " WHERE (:P_FILTRO IS NULL " &
            "        OR UPPER(E.INTERNAL_UID) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.SERIAL_NUMBER) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.MODEL) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.MANUFACTURER) LIKE '%' || UPPER(:P_FILTRO) || '%' " &
            "        OR UPPER(E.CPU_MODEL) LIKE '%' || UPPER(:P_FILTRO) || '%') " &
            "   AND (:P_STATUS IS NULL OR UPPER(E.STATUS) = UPPER(:P_STATUS)) " &
            " ORDER BY E.DATA_CADASTRO DESC"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlPt, oPar)
        Catch ex As Exception
            If IsErroEstruturaColuna(ex) Then
                Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlEn, oPar)
            End If
            Throw
        End Try

    End Function

    Private Function obterDashboardTotaisSemPackage() As DataSet

        Dim sqlStatus As String =
            "SELECT (SELECT COUNT(*) FROM TBL_EQUIPAMENTO) AS TOTAL_UNIDADES, " &
            "       (SELECT COUNT(*) FROM TBL_EQUIPAMENTO WHERE UPPER(STATUS) IN ('IN_STOCK','EM ESTOQUE')) AS EM_ESTOQUE, " &
            "       0 AS CONDICAO_BOA, " &
            "       (SELECT COUNT(*) FROM TBL_EQUIPAMENTO_UPGRADE WHERE DATA_UPGRADE >= TRUNC(SYSDATE) - 30) AS UPGRADES_30D, " &
            "       (SELECT COUNT(*) FROM TBL_REMESSA WHERE UPPER(NVL(STATUS_REMESSA, 'X')) NOT IN ('DELIVERED','RETURNED','ENTREGUE','DEVOLVIDA')) AS REMESSAS_ATIVAS " &
            "  FROM DUAL"

        Dim sqlStatusEquipamento As String =
            "SELECT (SELECT COUNT(*) FROM TBL_EQUIPAMENTO) AS TOTAL_UNIDADES, " &
            "       (SELECT COUNT(*) FROM TBL_EQUIPAMENTO WHERE UPPER(STATUS_EQUIPAMENTO) IN ('IN_STOCK','EM ESTOQUE')) AS EM_ESTOQUE, " &
            "       0 AS CONDICAO_BOA, " &
            "       (SELECT COUNT(*) FROM TBL_EQUIPAMENTO_UPGRADE WHERE DATA_UPGRADE >= TRUNC(SYSDATE) - 30) AS UPGRADES_30D, " &
            "       (SELECT COUNT(*) FROM TBL_REMESSA WHERE UPPER(NVL(STATUS_REMESSA, 'X')) NOT IN ('DELIVERED','RETURNED','ENTREGUE','DEVOLVIDA')) AS REMESSAS_ATIVAS " &
            "  FROM DUAL"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlStatus)
        Catch ex As Exception
            If IsErroEstruturaColuna(ex) Then
                Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlStatusEquipamento)
            End If
            Throw
        End Try

    End Function

#End Region

#Region "Auxiliares"

    Private Function IsErroObjetoEquipamento(ex As Exception) As Boolean

        Dim msg As String = ex.ToString().ToUpperInvariant()
        Return msg.Contains("PACK_EQUIPAMENTO") OrElse
               msg.Contains("PLS-00201") OrElse
               msg.Contains("PLS-00302") OrElse
               msg.Contains("ORA-06550")

    End Function

    Private Function IsErroEstruturaColuna(ex As Exception) As Boolean

        Dim msg As String = ex.ToString().ToUpperInvariant()
        Return msg.Contains("ORA-00904") OrElse msg.Contains("ORA-00942")

    End Function

#End Region

End Class
