Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Classe de leitura de Upgrades
''' </summary>
Public Class clsLeituraUpgrade

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

    Public Function selecionarUpgradesEquipamento(pIdEquipamento As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_EQUIPAMENTO", OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR",         OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdEquipamento

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_SELECT_UPGRADES_EQUIP", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarUpgradesRecentes(pDias As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_DIAS",   OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pDias

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_SELECT_UPGRADES_RECENTES", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    ''' <summary>
    ''' Upgrades dos últimos pDias dias com nome do cliente (DESTINATARIO da remessa).
    ''' Equipamentos sem remessa aparecem com CUSTOMER = 'Unassigned'.
    ''' </summary>
    Public Function selecionarUpgradesComCliente(pDias As Integer) As DataSet

        Dim sql As String =
            "SELECT NVL(R.DESTINATARIO, 'Unassigned')        AS CUSTOMER," &
            "       NVL(R.CODIGO_REMESSA, '—')               AS SHIPMENT_REF," &
            "       E.INTERNAL_UID," &
            "       E.MARCA," &
            "       E.MODEL," &
            "       E.SERIAL_NUMBER," &
            "       U.DATA_UPGRADE," &
            "       U.TIPO_UPGRADE                           AS COMPONENT_TYPE," &
            "       CASE U.TIPO_UPGRADE" &
            "           WHEN 'RAM' THEN TO_CHAR(U.RAM_ANTERIOR_GB)     || ' GB'" &
            "           WHEN 'SSD' THEN TO_CHAR(U.STORAGE_ANTERIOR_GB) || ' GB'" &
            "           WHEN 'HDD' THEN TO_CHAR(U.STORAGE_ANTERIOR_GB) || ' GB'" &
            "           ELSE NULL" &
            "       END                                      AS VALUE_BEFORE," &
            "       CASE U.TIPO_UPGRADE" &
            "           WHEN 'RAM' THEN CASE WHEN U.RAM_NOVA_GB     > 0 THEN TO_CHAR(U.RAM_NOVA_GB)     || ' GB' END" &
            "           WHEN 'SSD' THEN CASE WHEN U.STORAGE_NOVO_GB > 0 THEN TO_CHAR(U.STORAGE_NOVO_GB) || ' GB' END" &
            "           WHEN 'HDD' THEN CASE WHEN U.STORAGE_NOVO_GB > 0 THEN TO_CHAR(U.STORAGE_NOVO_GB) || ' GB' END" &
            "           ELSE NULL" &
            "       END                                      AS VALUE_AFTER," &
            "       U.TECNICO                                AS TECHNICIAN," &
            "       U.OBSERVACAO                             AS NOTES" &
            "  FROM TBL_EQUIPAMENTO_UPGRADE U" &
            "  JOIN TBL_EQUIPAMENTO E ON E.ID_EQUIPAMENTO = U.ID_EQUIPAMENTO" &
            "  LEFT JOIN TBL_REMESSA_ITEM RI ON RI.ID_EQUIPAMENTO = E.ID_EQUIPAMENTO" &
            "  LEFT JOIN TBL_REMESSA      R  ON R.ID_REMESSA      = RI.ID_REMESSA" &
            " WHERE U.DATA_UPGRADE >= SYSDATE - :V_DIAS" &
            " ORDER BY NVL(R.DESTINATARIO, 'Unassigned'), E.INTERNAL_UID, U.DATA_UPGRADE"

        Dim par As New OracleParameter("V_DIAS", OracleDbType.Int32, ParameterDirection.Input)
        par.Value = pDias

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sql,
                                               New OracleParameter() {par})
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

#End Region

End Class
